# -*- coding: utf-8 -*-
"""
Checkpoint/delta history for YAML profile edits stored via Revit Extensible Storage.
Only the first entry for a YAML path stores the full text; subsequent edits keep diffs.
"""

import base64
import difflib
import io
import json
import os
import zlib
from datetime import datetime

import System
import clr
try:
    clr.AddReference("RevitAPIUI")
except Exception:
    pass
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder
from Autodesk.Revit.DB.Events import DocumentChangedEventArgs, UndoOperation
from System import Guid, String, EventHandler  # noqa: E402

try:
    from Autodesk.Revit.UI.Events import ApplicationUndoRedoEventArgs
except Exception:  # pragma: no cover
    ApplicationUndoRedoEventArgs = None


try:
    from pyrevit import script as _script_logger  # noqa: E402
    _logger = _script_logger.get_logger()
except Exception:
    _logger = None


class ExtensibleStorage(object):
    """
    Tracks YAML edits per-project by storing history entries (deltas) inside the RVT.
    """

    SCHEMA_GUID = Guid("9f6633b1-d77f-49ef-9390-5111fbb16d82")
    SCHEMA_NAME = "CED_YamlHistory"
    HISTORY_FIELD_NAME = "HistoryJson"
    META_FIELD_NAME = "MetadataJson"

    DIFF_FORMAT = "ndiff"
    TRANSACTION_PREFIX = "YAML_HISTORY::"

    _schema = None
    _history_field = None
    _meta_field = None
    _undo_handler_registered = False
    _undo_handler_delegate = None
    _doc_handler = None
    _ui_handler = None
    _entry_cache = {}
    _recent_entries = {}

    @classmethod
    def _log(cls, message):
        try:
            if _logger:
                _logger.info(message)
            else:
                print("[ExtensibleStorage] " + message)
        except Exception:
            pass

    @classmethod
    def capture_change(cls, doc, yaml_path, previous_text, new_text, action, description=None, force_checkpoint=False):
        """
        Record a YAML mutation (add/edit/delete/etc.).
        """
        if doc is None:
            raise ValueError("Document reference is required for ExtensibleStorage writes.")
        payload = cls._read_storage(doc)
        entries = payload.setdefault("entries", [])
        meta = payload.setdefault("meta", {"next_seq": 1})

        normalized_path = cls._normalize_path(yaml_path)
        now_utc = datetime.utcnow()
        timestamp = now_utc.strftime("%Y-%m-%dT%H:%M:%SZ")

        base_map = meta.setdefault("base_text", {})
        if normalized_path not in base_map:
            base_map[normalized_path] = cls._compress_text(previous_text or "")

        active = meta.setdefault("active_yaml", {})
        if normalized_path:
            # keep active pointer aligned with the path receiving edits so future reads
            # (e.g. Add/Delete buttons) see the latest YAML text without rehydrating history
            if not active.get("normalized"):
                active["normalized"] = normalized_path
            if not active.get("path"):
                active["path"] = yaml_path or active.get("path")
            if active.get("normalized") == normalized_path:
                active["path"] = yaml_path or active.get("path")
                active["text"] = new_text or ""

        entry = {
            "seq": meta.get("next_seq", len(entries) + 1),
            "timestamp": timestamp,
            "user": cls._current_user(doc),
            "yaml_path": yaml_path or "",
            "yaml_path_norm": normalized_path,
            "action": action or "",
            "description": description or "",
            "entry_type": "delta",
            "prev_content": cls._compress_text(previous_text or ""),
            "new_content": cls._compress_text(new_text or ""),
        }
        entries.append(entry)
        meta["next_seq"] = entry["seq"] + 1
        txn_name = cls._make_transaction_name(entry["seq"], yaml_path, action)
        cls._log("capture_change: seq={} action='{}' yaml='{}'".format(entry["seq"], action or "", yaml_path or ""))
        cls._write_storage(doc, payload, txn_name)
        cls._remember_entry(doc, entry)
        cls._remember_recent_entry(doc, entry, txn_name)
        cls._log("Recent entry recorded seq={} stack-size={}".format(
            entry["seq"],
            len(cls._recent_entries.get(cls._doc_cache_key(doc), []) or []),
        ))
        return entry["seq"]

    @classmethod
    def list_history(cls, doc, yaml_path=None):
        """
        Return history entries (newest first). Each entry dict contains metadata but not decoded content.
        """
        payload = cls._read_storage(doc)
        entries = payload.get("entries", [])
        normalized = cls._normalize_path(yaml_path) if yaml_path else None
        filtered = []
        for entry in entries:
            if normalized and entry.get("yaml_path_norm") != normalized:
                continue
            summary = {
                "seq": entry.get("seq"),
                "timestamp": entry.get("timestamp"),
                "user": entry.get("user"),
                "yaml_path": entry.get("yaml_path"),
                "action": entry.get("action"),
                "description": entry.get("description"),
                "entry_type": entry.get("entry_type"),
            }
            filtered.append(summary)
        filtered.sort(key=lambda e: e["seq"], reverse=True)
        return filtered

    @classmethod
    def get_entry_detail(cls, doc, seq, yaml_path):
        payload = cls._read_storage(doc)
        normalized = cls._normalize_path(yaml_path)
        for entry in payload.get("entries", []):
            if entry.get("seq") == seq and entry.get("yaml_path_norm") == normalized:
                detail = dict(entry)
                detail["diff_text"] = cls._decompress_text(entry.get("content"))
                return detail
        return None

    @classmethod
    def reconstruct_entry(cls, doc, seq, yaml_path, base_text):
        """
        Rebuild the YAML text for a specific entry sequence number.
        Returns (yaml_path, reconstructed_text).
        base_text must represent the YAML content *before* the first recorded entry for that path.
        """
        payload = cls._read_storage(doc)
        entries = cls._merge_with_cache(doc, payload.get("entries", []), None)
        entry = cls._find_entry(entries, seq)
        if not entry:
            raise ValueError("Could not find history entry seq={} for path '{}'.".format(seq, yaml_path))
        text = cls._decompress_text(entry.get("new_content"))
        return entry.get("yaml_path"), text

    @classmethod
    def revert_to_entry(cls, doc, seq, yaml_path, writer_callback, base_text):
        """
        Restore the YAML file to the state stored at entry seq.
        writer_callback(path, text) handles disk writes and any external logging.
        """
        payload = cls._read_storage(doc)
        normalized = cls._normalize_path(yaml_path)
        entries = cls._merge_with_cache(doc, payload.get("entries", []), normalized)
        entry = cls._find_entry(entries, seq)
        if not entry:
            raise ValueError("History entry seq={} not found.".format(seq))
        text = cls._decompress_text(entry.get("new_content"))
        if callable(writer_callback):
            writer_callback(yaml_path, text)
        return yaml_path, text

    @classmethod
    def seed_active_yaml(cls, doc, yaml_path, raw_text):
        if doc is None or not yaml_path:
            raise ValueError("Document and YAML path are required.")
        payload = cls._read_storage(doc)
        normalized = cls._normalize_path(yaml_path)
        entries = [
            entry
            for entry in payload.get("entries", [])
            if entry.get("yaml_path_norm") != normalized
        ]
        payload["entries"] = entries
        meta = payload.setdefault("meta", {"next_seq": 1})
        base_map = meta.setdefault("base_text", {})
        base_map[normalized] = cls._compress_text(raw_text or "")
        meta["active_yaml"] = {
            "path": yaml_path,
            "normalized": normalized,
            "text": raw_text or "",
        }
        key = cls._doc_cache_key(doc)
        cls._entry_cache.pop(key, None)
        cls._recent_entries.pop(key, None)
        cls._write_storage(doc, payload, "Initialize YAML History")

    @classmethod
    def get_active_yaml(cls, doc):
        payload = cls._read_storage(doc)
        meta = payload.get("meta", {})
        active = meta.get("active_yaml") or {}
        path = active.get("path")
        normalized = active.get("normalized") or (cls._normalize_path(path) if path else None)
        text = None
        if path and normalized:
            entries = cls._merge_with_cache(doc, payload.get("entries", []), normalized)
            latest_text = None
            if entries:
                latest_text = cls._decompress_text(entries[-1].get("new_content"))
            stored_text = active.get("text")
            if latest_text:
                text = latest_text
                cls._log("get_active_yaml returning latest entry for {} (len={})".format(path, len(latest_text or "")))
            elif stored_text:
                text = stored_text
                cls._log("get_active_yaml returning stored text for {} (len={})".format(path, len(stored_text or "")))
            else:
                base_map = meta.get("base_text") or {}
                text = cls._decompress_text(base_map.get(normalized))
                cls._log("get_active_yaml returning base text for {} (len={})".format(path, len(text or "")))
        return path, normalized, text

    @classmethod
    def update_active_yaml(cls, doc, yaml_path, previous_text, new_text, action, description):
        if not yaml_path:
            raise ValueError("Active YAML path is not set.")
        payload = cls._read_storage(doc)
        normalized = cls._normalize_path(yaml_path)
        meta = payload.setdefault("meta", {})
        base_map = meta.setdefault("base_text", {})
        if normalized not in base_map:
            base_map[normalized] = cls._compress_text(previous_text or "")
        active = meta.setdefault("active_yaml", {"path": yaml_path, "normalized": normalized})
        active["path"] = yaml_path
        active["normalized"] = normalized
        active["text"] = new_text or ""
        entry = cls.capture_change(
            doc,
            yaml_path,
            previous_text or "",
            new_text or "",
            action,
            description=description or "",
        )

    @classmethod
    def update_active_text_only(cls, doc, yaml_path, new_text):
        if doc is None or not yaml_path:
            return
        payload = cls._read_storage(doc)
        normalized = cls._normalize_path(yaml_path)
        meta = payload.setdefault("meta", {})
        active = meta.setdefault("active_yaml", {"path": yaml_path, "normalized": normalized})
        active["path"] = yaml_path
        active["normalized"] = normalized
        active["text"] = new_text or ""
        cls._write_storage(doc, payload, "ACTIVE_YAML_REFRESH")

    # ---------------------------------------------------------------------- #
    # Undo / Redo integration
    # ---------------------------------------------------------------------- #
    @classmethod
    def ensure_undo_handler(cls):
        if cls._undo_handler_registered:
            return
        handlers_registered = False
        try:
            uiapp = __revit__
            app = getattr(uiapp, "Application", None)
            if app is None:
                cls._log("DocumentChanged handler unavailable; missing Application reference.")
            else:
                handler = EventHandler[DocumentChangedEventArgs](cls._on_document_changed)
                app.DocumentChanged += handler
                cls._doc_handler = handler
                handlers_registered = True
                cls._log("DocumentChanged handler registered.")
        except Exception:
            pass
        if ApplicationUndoRedoEventArgs is not None:
            try:
                uiapp = __revit__
                handler = System.EventHandler[ApplicationUndoRedoEventArgs](cls._on_undo_redo)
                uiapp.UndoRedo += handler
                cls._ui_handler = handler
                handlers_registered = True
                cls._log("UndoRedo handler registered.")
            except Exception:
                pass
        else:
            cls._log("UndoRedo handler unavailable; ApplicationUndoRedoEventArgs missing.")
        cls._undo_handler_registered = handlers_registered

    @classmethod
    def _on_document_changed(cls, sender, args):
        operation = getattr(args, "Operation", None)
        cls._log("DocumentChanged event operation={!r}".format(operation))
        if operation not in (UndoOperation.Undo, UndoOperation.Redo):
            return
        names = []
        get_names = getattr(args, "GetTransactionNames", None)
        if callable(get_names):
            try:
                names = list(get_names() or [])
            except Exception:
                names = []
        if not names:
            name = None
            get_single = getattr(args, "GetTransactionName", None)
            if callable(get_single):
                try:
                    name = get_single()
                except Exception:
                    name = None
            if not name:
                name = getattr(args, "TransactionName", None)
            if name:
                names = [name]
        if not names:
            cls._log("DocumentChanged event has no transaction names.")
            return
        try:
            cls._log("DocumentChanged transaction names: {}".format(
                ", ".join([str(n) for n in names]) or "<empty>"
            ))
        except Exception:
            pass
        doc = None
        try:
            doc = args.GetDocument()
        except Exception:
            doc = getattr(args, "Document", None)
        if doc is None:
            try:
                doc = __revit__.ActiveUIDocument.Document
            except Exception:
                doc = None
        if doc is None:
            cls._log("DocumentChanged: could not resolve document.")
            return
        handled = False
        for name in names:
            if cls._process_transaction_name(doc, name, operation, "DocumentChanged"):
                handled = True
        if not handled:
            cls._process_recent_entry(doc, operation, "DocumentChanged")

    @classmethod
    def _on_undo_redo(cls, sender, args):
        operation = getattr(args, "Operation", None)
        cls._log("UndoRedo event operation={!r}".format(operation))
        names = []
        get_names = getattr(args, "GetTransactionNames", None)
        if callable(get_names):
            try:
                names = list(get_names() or [])
            except Exception:
                names = []
        if not names:
            name = None
            get_single = getattr(args, "GetTransactionName", None)
            if callable(get_single):
                try:
                    name = get_single()
                except Exception:
                    name = None
            if not name:
                name = getattr(args, "TransactionName", None)
            if name:
                names = [name]
        if not names:
            cls._log("UndoRedo event has no transaction names.")
            return
        try:
            cls._log("UndoRedo transaction names: {}".format(
                ", ".join([str(n) for n in names]) or "<empty>"
            ))
        except Exception:
            pass
        doc = None
        try:
            doc = getattr(args, "Document", None)
        except Exception:
            doc = None
        if doc is None:
            try:
                doc = __revit__.ActiveUIDocument.Document
            except Exception:
                doc = None
        if doc is None:
            cls._log("UndoRedo event: could not resolve document.")
            return
        handled = False
        for name in names:
            if cls._process_transaction_name(doc, name, operation, "UndoRedo"):
                handled = True
        if not handled:
            cls._process_recent_entry(doc, operation, "UndoRedo")

    @classmethod
    def _process_transaction_name(cls, doc, name, operation, source):
        if not name:
            cls._log("Transaction hook called with empty name via {}".format(source))
            return False
        if name == "ACTIVE_YAML_REFRESH":
            cls._log("Transaction hook ignoring active YAML refresh via {}".format(source))
            return True
        cls._log("Transaction hook evaluating '{}' via {}".format(name, source))
        seq, normalized = cls._resolve_transaction_tokens(doc, name)
        if seq is None or normalized is None:
            return False
        payload_data = cls._read_storage(doc)
        entries = cls._merge_with_cache(doc, payload_data.get("entries", []), normalized)
        if not entries:
            cls._log("DocumentChanged: no entries for {}".format(normalized))
            return False
        entry = cls._find_entry(entries, seq)
        if not entry:
            cls._log("DocumentChanged: seq {} not found.".format(seq))
            return False
        text = None
        if operation == UndoOperation.Undo:
            text = cls._decompress_text(entry.get("prev_content"))
        elif operation == UndoOperation.Redo:
            text = cls._decompress_text(entry.get("new_content"))
        if text is None:
            meta = payload_data.get("meta", {})
            base_text = cls._get_base_text(meta, normalized)
            try:
                target_seq = seq if operation == UndoOperation.Redo else cls._previous_sequence(entries, seq)
                text = cls._reconstruct_text(entries, base_text, target_seq)
            except Exception:
                text = None
        if text is None:
            cls._log("DocumentChanged: failed to reconstruct text for seq {}".format(seq))
            return False
        yaml_path = cls._resolve_yaml_path(entries, normalized)
        cls._write_yaml_file(yaml_path, text)
        cls._log(
            "DocumentChanged: rewrote '{}' via {}".format(
                yaml_path,
                "Undo" if operation == UndoOperation.Undo else "Redo",
            )
        )
        cls._forget_recent_entry(doc, seq, normalized)
        return True

    @classmethod
    def _process_recent_entry(cls, doc, operation, source):
        key = cls._doc_cache_key(doc)
        if not key:
            cls._log("Recent entry fallback skipped; missing doc key.")
            return False
        stack = cls._recent_entries.get(key)
        if not stack:
            cls._log("Recent entry fallback skipped; stack empty.")
            return False
        entry = stack.pop()
        cls._log("Recent entry fallback via {} using seq {}; stack-size={}.".format(source, entry.get("seq"), len(stack)))
        return cls._process_transaction_name(doc, entry.get("txn_name"), operation, source + "::fallback")


    # ---------------------------------------------------------------------- #
    # Internal helpers
    # ---------------------------------------------------------------------- #
    @classmethod
    def _find_entry_index(cls, entries, seq, normalized_path):
        for idx, entry in enumerate(entries):
            if entry.get("seq") == seq and entry.get("yaml_path_norm") == normalized_path:
                return idx
        return None

    @classmethod
    def _entries_for_path(cls, entries, normalized_path):
        relevant = [
            entry for entry in entries
            if entry.get("yaml_path_norm") == normalized_path
        ]
        relevant.sort(key=lambda e: e.get("seq") or 0)
        return relevant

    @classmethod
    def _find_entry(cls, entries, seq):
        for entry in entries:
            if entry.get("seq") == seq:
                return entry
        return None

    @classmethod
    def _previous_sequence(cls, entries, seq):
        prev = None
        for entry in entries:
            current_seq = entry.get("seq")
            if current_seq == seq:
                return prev
            prev = current_seq
        return prev

    @classmethod
    def _get_base_text(cls, meta, normalized_path):
        base_map = (meta or {}).get("base_text") or {}
        compressed = base_map.get(normalized_path)
        return cls._decompress_text(compressed) if compressed else ""

    @classmethod
    def _reconstruct_text(cls, entries, base_text, target_seq):
        text = base_text or ""
        if target_seq is None:
            return text
        for entry in entries:
            diff_text = cls._decompress_text(entry.get("content")) if entry.get("content") else None
            if diff_text:
                text = cls._apply_diff(text, diff_text)
            else:
                text = cls._decompress_text(entry.get("new_content"))
            if entry.get("seq") == target_seq:
                break
        return text

    @classmethod
    def _resolve_yaml_path(cls, entries, normalized_path):
        for entry in reversed(entries):
            if entry.get("yaml_path_norm") == normalized_path:
                path = entry.get("yaml_path")
                if path:
                    return path
        return normalized_path

    @classmethod
    def _write_yaml_file(cls, yaml_path, text):
        cls._log("YAML rewrites are managed in Extensible Storage; skipping disk write for '{}'.".format(yaml_path or "<unknown>"))

    @classmethod
    def _schema_and_fields(cls):
        if cls._schema:
            return cls._schema, cls._history_field, cls._meta_field
        schema = Schema.Lookup(cls.SCHEMA_GUID)
        if schema is None:
            builder = SchemaBuilder(cls.SCHEMA_GUID)
            builder.SetSchemaName(cls.SCHEMA_NAME)
            builder.SetDocumentation("Stores YAML history deltas.")
            history_field = builder.AddSimpleField(cls.HISTORY_FIELD_NAME, String)
            meta_field = builder.AddSimpleField(cls.META_FIELD_NAME, String)
            doc_field = builder.AddSimpleField("DocGuid", String)
            schema = builder.Finish()
        history_field = schema.GetField(cls.HISTORY_FIELD_NAME)
        meta_field = schema.GetField(cls.META_FIELD_NAME)
        cls._schema = schema
        cls._history_field = history_field
        cls._meta_field = meta_field
        return schema, history_field, meta_field

    @classmethod
    def _read_storage(cls, doc):
        schema, history_field, meta_field = cls._schema_and_fields()
        payload = {"entries": [], "meta": {"next_seq": 1}}
        project_info = getattr(doc, "ProjectInformation", None)
        if project_info is None:
            return payload
        entity = project_info.GetEntity(schema)
        needed_guid = cls._normalize_guid(doc)
        if not entity or not entity.IsValid():
            return payload
        doc_guid_field = schema.GetField("DocGuid")
        doc_guid = entity.Get[str](doc_guid_field) if doc_guid_field else None
        if doc_guid and doc_guid != needed_guid:
            return payload
        history_json = entity.Get[str](history_field)
        meta_json = entity.Get[str](meta_field)
        if history_json:
            try:
                payload["entries"] = json.loads(history_json)
            except Exception:
                payload["entries"] = []
        if meta_json:
            try:
                payload["meta"] = json.loads(meta_json)
            except Exception:
                pass
        if "meta" not in payload:
            payload["meta"] = {"next_seq": 1}
        return payload

    @classmethod
    def _write_storage(cls, doc, payload, transaction_name=None):
        schema, history_field, meta_field = cls._schema_and_fields()
        history_json = json.dumps(payload.get("entries", []))
        meta_json = json.dumps(payload.get("meta", {}))
        project_info = getattr(doc, "ProjectInformation", None)
        if project_info is None:
            raise RuntimeError("ProjectInformation element is required for ExtensibleStorage writes.")

        def _apply():
            entity = project_info.GetEntity(schema)
            if not entity or not entity.IsValid():
                entity = Entity(schema)
            entity.Set[str](history_field, history_json)
            entity.Set[str](meta_field, meta_json)
            doc_guid_field = schema.GetField("DocGuid")
            if doc_guid_field:
                entity.Set[str](doc_guid_field, cls._normalize_guid(doc))
            project_info.SetEntity(entity)

        # Always wrap in our own transaction so Undo stack records it
        t = Transaction(doc, transaction_name or "YAML Change")
        t.Start()
        try:
            cls._log("ExtensibleStorage write txn={} entries={} active_path={}".format(
                transaction_name or "YAML Change",
                len(payload.get("entries", [])),
                (payload.get("meta", {}).get("active_yaml") or {}).get("path"),
            ))
            _apply()
            t.Commit()
        except Exception as ex:
            cls._log("ExtensibleStorage write failed: {}".format(ex))
            t.RollBack()
            raise

    @classmethod
    def _forget_recent_entry(cls, doc, seq, normalized_path):
        key = cls._doc_cache_key(doc)
        if not key:
            return
        stack = cls._recent_entries.get(key)
        if not stack:
            return
        filtered = [
            entry for entry in stack
            if entry.get("seq") != seq or entry.get("yaml_path_norm") != normalized_path
        ]
        if filtered:
            cls._recent_entries[key] = filtered
        else:
            cls._recent_entries.pop(key, None)

    @classmethod
    def _remember_txn_suffix(cls, seq, suffix):
        cls._transaction_suffix = getattr(cls, "_transaction_suffix", {})
        cls._transaction_suffix[seq] = suffix

    @classmethod
    def _resolve_transaction_tokens(cls, doc, name):
        prefix_index = name.find(cls.TRANSACTION_PREFIX)
        if prefix_index != -1:
            payload = name[prefix_index + len(cls.TRANSACTION_PREFIX):]
            payload = payload.split("]", 1)[0]
            parts = payload.split("::", 1)
            if len(parts) == 2:
                try:
                    seq = int(parts[0])
                    encoded_path = parts[1]
                    normalized = cls._decode_path(encoded_path)
                    return seq, normalized
                except Exception:
                    pass
        # fallback: try to resolve by looking up last suffix for any seq
        recent_stack = cls._recent_entries.get(cls._doc_cache_key(doc) or "", [])
        if recent_stack:
            entry = recent_stack[-1]
            return entry.get("seq"), entry.get("yaml_path_norm")
        return None, None

    @classmethod
    def _remember_entry(cls, doc, entry):
        key = cls._doc_cache_key(doc)
        if not key:
            return
        cache = cls._entry_cache.setdefault(key, [])
        cache.append(dict(entry))
        if len(cache) > 200:
            del cache[:-200]

    @classmethod
    def _remember_recent_entry(cls, doc, entry, txn_name):
        key = cls._doc_cache_key(doc)
        if not key:
            return
        stack = cls._recent_entries.setdefault(key, [])
        stored = dict(entry)
        stored["txn_name"] = txn_name
        stack.append(stored)
        if len(stack) > 50:
            del stack[:-50]

    @classmethod
    def _merge_with_cache(cls, doc, entries, normalized_path):
        if normalized_path:
            merged = cls._entries_for_path(entries, normalized_path)
        else:
            merged = list(entries or [])
            merged.sort(key=lambda e: e.get("seq") or 0)
        cache_entries = cls._cached_entries(doc, normalized_path)
        if not cache_entries:
            return merged
        existing = {entry.get("seq") for entry in merged}
        for cached in cache_entries:
            seq = cached.get("seq")
            if seq in existing:
                continue
            merged.append(dict(cached))
            existing.add(seq)
        merged.sort(key=lambda e: e.get("seq") or 0)
        return merged

    @classmethod
    def _cached_entries(cls, doc, normalized_path):
        key = cls._doc_cache_key(doc)
        if not key:
            return []
        cached = cls._entry_cache.get(key, [])
        if not cached:
            return []
        if not normalized_path:
            return list(cached)
        return [
            dict(entry)
            for entry in cached
            if entry.get("yaml_path_norm") == normalized_path
        ]

    @classmethod
    def _doc_cache_key(cls, doc):
        if doc is None:
            return None
        return cls._normalize_guid(doc)

    @classmethod
    def _compress_text(cls, text):
        data = text.encode("utf-8")
        return base64.b64encode(zlib.compress(data)).decode("ascii")

    @classmethod
    def _decompress_text(cls, payload):
        if not payload:
            return ""
        raw = base64.b64decode(payload.encode("ascii"))
        return zlib.decompress(raw).decode("utf-8")

    @classmethod
    def _compute_diff(cls, old_text, new_text):
        old_lines = old_text.splitlines(True)
        new_lines = new_text.splitlines(True)
        diff_lines = difflib.ndiff(old_lines, new_lines)
        return "".join(diff_lines)

    @classmethod
    def _apply_diff(cls, base_text, diff_text):
        diff_lines = diff_text.splitlines(True)
        restored = difflib.restore(diff_lines, 2)  # 2 -> new version
        return "".join(restored)

    @classmethod
    def _encode_path(cls, path):
        try:
            return base64.urlsafe_b64encode(path.encode("utf-8")).decode("ascii")
        except Exception:
            return path

    @classmethod
    def _decode_path(cls, token):
        try:
            return base64.urlsafe_b64decode(token.encode("ascii")).decode("utf-8")
        except Exception:
            return token

    @classmethod
    def _make_transaction_name(cls, seq, yaml_path, action):
        normalized = cls._normalize_path(yaml_path)
        encoded = cls._encode_path(normalized)
        base_name = (action or "YAML Change").strip() or "YAML Change"
        safe_action = base_name.replace("\n", " ")
        suffix = "{}{}::{}".format(cls.TRANSACTION_PREFIX, seq, encoded)
        cls._remember_txn_suffix(seq, suffix)
        return safe_action

    @classmethod
    def _normalize_path(cls, path):
        if not path:
            return ""
        try:
            normalized = os.path.abspath(path)
        except Exception:
            normalized = path
        return normalized.replace("\\", "/").lower()

    @classmethod
    def _normalize_guid(cls, doc):
        try:
            path = getattr(doc, "PathName", "") or ""
            unique = getattr(doc, "Title", "") or ""
            return "{}::{}".format(path, unique)
        except Exception:
            return "unknown"

    @classmethod
    def _current_user(cls, doc):
        try:
            revit_user = doc.Application.Username
            if revit_user:
                return revit_user
        except Exception:
            pass
        return os.getenv("USERNAME") or os.getenv("USER") or "unknown"


__all__ = ["ExtensibleStorage"]

try:
    __revit__  # type: ignore # noqa
except Exception:
    __revit__ = None

if __revit__:
    ExtensibleStorage.ensure_undo_handler()
