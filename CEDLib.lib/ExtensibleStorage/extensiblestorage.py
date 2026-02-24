# -*- coding: utf-8 -*-
"""
Active YAML storage and user settings stored via Revit Extensible Storage.
"""

import json
import os
from datetime import datetime

import System
import clr
try:
    clr.AddReference("RevitAPI")
except Exception:
    pass
try:
    clr.AddReference("RevitAPIUI")
except Exception:
    pass
try:
    import Autodesk.Revit.DB as DB
except Exception:  # pragma: no cover
    DB = None

if DB:
    Transaction = DB.Transaction
    FilteredElementCollector = DB.FilteredElementCollector
else:  # pragma: no cover
    Transaction = None
    FilteredElementCollector = None

DataStorage = None
if DB:
    try:
        DataStorage = DB.DataStorage
    except Exception:
        DataStorage = None
if DataStorage is None:
    try:
        DataStorage = System.Type.GetType("Autodesk.Revit.DB.DataStorage, RevitAPI")
    except Exception:
        DataStorage = None
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder
from System import Guid, String, Array  # noqa: E402

try:
    from Autodesk.Revit.UI.Events import DocumentSynchronizedWithCentralEventArgs
except Exception:  # pragma: no cover
    DocumentSynchronizedWithCentralEventArgs = None


try:
    from pyrevit import script as _script_logger  # noqa: E402
    _logger = _script_logger.get_logger()
except Exception:
    _logger = None


class ExtensibleStorage(object):
    """
    Stores the active YAML payload, user settings, and editor locks inside the RVT.
    """

    SCHEMA_NAME = "CED_YamlHistory"
    HISTORY_FIELD_NAME = "HistoryJson"
    META_FIELD_NAME = "MetadataJson"
    HISTORY_CHUNKS_FIELD_NAME = "HistoryJsonChunks"
    META_CHUNKS_FIELD_NAME = "MetadataJsonChunks"
    SCHEMA_VERSION = 3
    MAX_ES_STRING = 16 * 1024 * 1024
    CHUNK_SIZE = 8 * 1024 * 1024

    USER_SETTINGS_KEY = "user_settings"

    _schema_cache = {}
    _undo_handler_registered = False
    _sync_handler = None
    _datastorage_not_found_logged = False

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
    def seed_active_yaml(cls, doc, yaml_path, raw_text):
        if doc is None or not yaml_path:
            raise ValueError("Document and YAML path are required.")
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        meta.pop("base_text", None)
        meta.pop("next_seq", None)
        meta["active_yaml"] = {
            "path": yaml_path,
            "normalized": cls._normalize_path(yaml_path),
            "text": raw_text or "",
        }
        cls._write_storage(doc, payload, "Initialize Active YAML")

    @classmethod
    def acquire_editor_lock(cls, doc, user):
        if doc is None:
            return None
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        lock = meta.get("editor_lock")
        normalized_user = user or cls._current_user(doc)
        if lock:
            holder = lock.get("user")
            if holder and holder != normalized_user:
                return dict(lock)
        meta["editor_lock"] = {
            "user": normalized_user,
            "timestamp": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        cls._write_storage(doc, payload, "YAML_EDITOR_LOCK")
        return None

    @classmethod
    def release_editor_lock(cls, doc, user):
        if doc is None:
            return
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        lock = meta.get("editor_lock")
        normalized_user = user or cls._current_user(doc)
        if lock and (not lock.get("user") or lock.get("user") == normalized_user):
            meta.pop("editor_lock", None)
            cls._write_storage(doc, payload, "YAML_EDITOR_UNLOCK")

    @classmethod
    def current_editor_lock(cls, doc):
        if doc is None:
            return None
        payload = cls._read_storage(doc)
        meta = payload.get("meta", {})
        lock = meta.get("editor_lock")
        return dict(lock) if isinstance(lock, dict) else None

    @classmethod
    def get_user_setting(cls, doc, setting_key, default=None, user=None):
        if doc is None or not setting_key:
            return default
        payload = cls._read_storage(doc)
        meta = payload.get("meta", {})
        settings = meta.get(cls.USER_SETTINGS_KEY) or {}
        if not isinstance(settings, dict):
            return default
        setting_map = settings.get(setting_key)
        if not isinstance(setting_map, dict):
            return default
        normalized_user = user or cls._current_user(doc)
        if not normalized_user:
            return default
        if normalized_user not in setting_map:
            return default
        return setting_map.get(normalized_user)

    @classmethod
    def set_user_setting(cls, doc, setting_key, value, user=None, transaction_name=None):
        if doc is None or not setting_key:
            return False
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        settings = meta.get(cls.USER_SETTINGS_KEY)
        if not isinstance(settings, dict):
            settings = {}
            meta[cls.USER_SETTINGS_KEY] = settings
        setting_map = settings.get(setting_key)
        if not isinstance(setting_map, dict):
            setting_map = {}
            settings[setting_key] = setting_map
        normalized_user = user or cls._current_user(doc)
        if not normalized_user:
            return False
        setting_map[normalized_user] = value
        txn_name = transaction_name or "USER_SETTING::{}".format(setting_key)
        cls._write_storage(doc, payload, txn_name)
        return True

    @classmethod
    def get_active_yaml(cls, doc):
        payload = cls._read_storage(doc)
        meta = payload.get("meta", {})
        active = meta.get("active_yaml") or {}
        path = active.get("path")
        normalized = active.get("normalized") or (cls._normalize_path(path) if path else None)
        text = active.get("text")
        if path and text is not None:
            cls._log("get_active_yaml returning stored text for {} (len={})".format(path, len(text or "")))
        return path, normalized, text

    @classmethod
    def update_active_yaml(cls, doc, yaml_path, previous_text, new_text, action, description):
        if not yaml_path:
            raise ValueError("Active YAML path is not set.")
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        meta.pop("base_text", None)
        meta.pop("next_seq", None)
        active = meta.setdefault("active_yaml", {})
        active["path"] = yaml_path
        active["normalized"] = cls._normalize_path(yaml_path)
        active["text"] = new_text or ""
        txn_name = action or "YAML Update"
        cls._write_storage(doc, payload, txn_name)

    @classmethod
    def update_active_text_only(cls, doc, yaml_path, new_text):
        if doc is None or not yaml_path:
            return
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        meta.pop("base_text", None)
        meta.pop("next_seq", None)
        active = meta.setdefault("active_yaml", {})
        active["path"] = yaml_path
        active["normalized"] = cls._normalize_path(yaml_path)
        active["text"] = new_text or ""
        cls._write_storage(doc, payload, "ACTIVE_YAML_REFRESH")

    # ---------------------------------------------------------------------- #
    # Sync integration
    # ---------------------------------------------------------------------- #
    @classmethod
    def ensure_undo_handler(cls):
        if cls._undo_handler_registered:
            return
        handlers_registered = False
        if DocumentSynchronizedWithCentralEventArgs is not None:
            try:
                uiapp = __revit__
                handler = System.EventHandler[DocumentSynchronizedWithCentralEventArgs](cls._on_doc_sync)
                uiapp.DocumentSynchronizedWithCentral += handler
                cls._sync_handler = handler
                handlers_registered = True
                cls._log("DocumentSynchronized handler registered.")
            except Exception:
                pass
        cls._undo_handler_registered = handlers_registered

    @classmethod
    def _on_doc_sync(cls, sender, args):
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
            return
        user = cls._current_user(doc)
        cls.release_editor_lock(doc, user)
    @classmethod
    def _schema_and_fields(cls, doc):
        return cls._schema_and_fields_versioned(doc, prefer_chunked=False)

    @classmethod
    def _schema_and_fields_versioned(cls, doc, prefer_chunked=False, force_version=None):
        doc_key = getattr(doc, "Title", "unknown")
        cache = cls._schema_cache.get(doc_key, {})
        if force_version == 3 and cache.get("v3"):
            return cache["v3"]
        if not force_version:
            if not prefer_chunked and cache.get("default"):
                return cache["default"]
            if prefer_chunked:
                if cache.get("v3"):
                    return cache["v3"]
                if cache.get("v2"):
                    return cache["v2"]

        if force_version == 3:
            schema_v3 = Schema.Lookup(_make_doc_guid(doc, version=3))
            if schema_v3 is None:
                schema_v3 = cls._build_schema_v3(doc)
            packed = cls._pack_schema(schema_v3, version=3)
            cache["v3"] = packed
            cache["default"] = packed
            cls._schema_cache[doc_key] = cache
            return packed

        schema_v3 = Schema.Lookup(_make_doc_guid(doc, version=3))
        if schema_v3 is not None:
            packed = cls._pack_schema(schema_v3, version=3)
            cache["v3"] = packed
            cache["default"] = packed
            cls._schema_cache[doc_key] = cache
            return packed

        schema_v2 = Schema.Lookup(_make_doc_guid(doc, version=2))
        if schema_v2 is not None and not prefer_chunked:
            packed = cls._pack_schema(schema_v2, version=2)
            cache["v2"] = packed
            cache["default"] = packed
            cls._schema_cache[doc_key] = cache
            return packed

        schema_v1 = Schema.Lookup(_make_doc_guid(doc, version=1))
        if schema_v1 is not None and not prefer_chunked:
            packed = cls._pack_schema(schema_v1, version=1)
            cache["v1"] = packed
            cache["default"] = packed
            cls._schema_cache[doc_key] = cache
            return packed

        schema_v3 = cls._build_schema_v3(doc)
        packed = cls._pack_schema(schema_v3, version=3)
        cache["v3"] = packed
        cache["default"] = packed
        cls._schema_cache[doc_key] = cache
        return packed

    @classmethod
    def _build_schema_v2(cls, doc):
        builder = SchemaBuilder(_make_doc_guid(doc, version=2))
        builder.SetSchemaName("{}_v2".format(cls.SCHEMA_NAME))
        builder.SetDocumentation("Stores active YAML payload and metadata (chunked).")
        builder.AddSimpleField(cls.HISTORY_FIELD_NAME, String)
        builder.AddSimpleField(cls.META_FIELD_NAME, String)
        try:
            builder.AddArrayField(cls.HISTORY_CHUNKS_FIELD_NAME, String)
            builder.AddArrayField(cls.META_CHUNKS_FIELD_NAME, String)
        except Exception:
            # If array fields are unavailable, continue with only simple fields.
            pass
        builder.AddSimpleField("DocGuid", String)
        return builder.Finish()

    @classmethod
    def _build_schema_v3(cls, doc):
        builder = SchemaBuilder(_make_doc_guid(doc, version=3))
        builder.SetSchemaName("{}_v3".format(cls.SCHEMA_NAME))
        builder.SetDocumentation("Stores active YAML payload and metadata (chunked, fixed schema GUID).")
        builder.AddSimpleField(cls.HISTORY_FIELD_NAME, String)
        builder.AddSimpleField(cls.META_FIELD_NAME, String)
        try:
            builder.AddArrayField(cls.HISTORY_CHUNKS_FIELD_NAME, String)
            builder.AddArrayField(cls.META_CHUNKS_FIELD_NAME, String)
        except Exception:
            # If array fields are unavailable, continue with only simple fields.
            pass
        builder.AddSimpleField("DocGuid", String)
        return builder.Finish()

    @classmethod
    def _pack_schema(cls, schema, version):
        history_field = schema.GetField(cls.HISTORY_FIELD_NAME)
        meta_field = schema.GetField(cls.META_FIELD_NAME)
        history_chunks_field = None
        meta_chunks_field = None
        if version >= 2:
            history_chunks_field = schema.GetField(cls.HISTORY_CHUNKS_FIELD_NAME)
            meta_chunks_field = schema.GetField(cls.META_CHUNKS_FIELD_NAME)
        doc_field = schema.GetField("DocGuid")
        return (schema, history_field, meta_field, history_chunks_field, meta_chunks_field, doc_field, version)

    @classmethod
    def _read_chunked_field(cls, entity, base_field, chunks_field):
        chunks = cls._get_chunk_list(entity, chunks_field)
        if chunks:
            return "".join(chunks)
        if base_field:
            try:
                return entity.Get[str](base_field)
            except Exception:
                return None
        return None

    @classmethod
    def _write_chunked_field(cls, entity, base_field, chunks_field, value):
        value = value or ""
        if chunks_field and len(value) > cls.MAX_ES_STRING:
            chunks = cls._split_chunks(value, cls.CHUNK_SIZE)
            cls._set_chunk_list(entity, chunks_field, chunks)
            if base_field:
                entity.Set[str](base_field, "")
            return
        if chunks_field:
            cls._set_chunk_list(entity, chunks_field, [])
        if base_field:
            entity.Set[str](base_field, value)

    @classmethod
    def _split_chunks(cls, text, chunk_size):
        if not text:
            return []
        size = max(1, int(chunk_size or 1))
        return [text[i:i + size] for i in range(0, len(text), size)]

    @classmethod
    def _get_chunk_list(cls, entity, chunks_field):
        if not chunks_field:
            return None
        try:
            from System.Collections.Generic import IList, List
        except Exception:
            IList = None
            List = None
        if IList:
            try:
                chunks = entity.Get[IList[String]](chunks_field)
                if chunks is not None:
                    return list(chunks)
            except Exception:
                pass
        if List:
            try:
                chunks = entity.Get[List[String]](chunks_field)
                if chunks is not None:
                    return list(chunks)
            except Exception:
                pass
        try:
            chunks = entity.Get[Array[String]](chunks_field)
            if chunks is not None:
                return list(chunks)
        except Exception:
            pass
        return None

    @classmethod
    def _set_chunk_list(cls, entity, chunks_field, chunks):
        if not chunks_field:
            return
        chunks = chunks or []
        try:
            from System.Collections.Generic import List
        except Exception:
            List = None
        if List:
            try:
                list_obj = List[String]()
                for chunk in chunks:
                    list_obj.Add(chunk)
                entity.Set[List[String]](chunks_field, list_obj)
                return
            except Exception:
                pass
        try:
            arr = Array[String](chunks)
            entity.Set[Array[String]](chunks_field, arr)
        except Exception:
            # Last resort: store as empty to avoid oversize strings.
            try:
                entity.Set[str](chunks_field, "")
            except Exception:
                pass

    @classmethod
    def _read_storage(cls, doc):
        schema, history_field, meta_field, history_chunks_field, meta_chunks_field, doc_field, version = cls._schema_and_fields(doc)
        payload = {"entries": [], "meta": {}}
        needed_guid = cls._normalize_guid(doc)
        storage_elem = cls._find_data_storage(doc, schema, doc_field, needed_guid)
        entity = None
        if storage_elem is not None:
            entity = storage_elem.GetEntity(schema)
        if not entity or not entity.IsValid():
            return payload
        doc_guid = entity.Get[str](doc_field) if doc_field else None
        if doc_guid and doc_guid != needed_guid:
            return payload
        history_json = cls._read_chunked_field(entity, history_field, history_chunks_field)
        meta_json = cls._read_chunked_field(entity, meta_field, meta_chunks_field)
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
            payload["meta"] = {}
        return payload

    @classmethod
    def _write_storage(cls, doc, payload, transaction_name=None):
        history_json = json.dumps(payload.get("entries", []))
        meta_json = json.dumps(payload.get("meta", {}))
        needs_chunking = (
            len(history_json or "") > cls.MAX_ES_STRING
            or len(meta_json or "") > cls.MAX_ES_STRING
        )
        schema, history_field, meta_field, history_chunks_field, meta_chunks_field, doc_field, version = cls._schema_and_fields_versioned(
            doc,
            prefer_chunked=needs_chunking,
            force_version=3,
        )
        if cls._resolve_datastorage_type() is None:
            version = None
            try:
                version = doc.Application.VersionNumber
            except Exception:
                version = None
            db_loaded = DB is not None
            raise RuntimeError(
                "DataStorage is required for ExtensibleStorage writes. Revit API version: {}. DB loaded: {}.".format(
                    version or "unknown",
                    db_loaded,
                )
            )
        storage_elem = cls._find_data_storage(doc, schema, doc_field, cls._normalize_guid(doc))

        def _apply():
            target_elem = storage_elem
            if target_elem is None:
                target_elem = cls._get_or_create_data_storage(doc, schema, doc_field, cls._normalize_guid(doc))
            if target_elem is None:
                raise RuntimeError("DataStorage element is required for ExtensibleStorage writes.")
            entity = target_elem.GetEntity(schema)
            if not entity or not entity.IsValid():
                entity = Entity(schema)
            cls._write_chunked_field(entity, history_field, history_chunks_field, history_json)
            cls._write_chunked_field(entity, meta_field, meta_chunks_field, meta_json)
            if doc_field:
                entity.Set[str](doc_field, cls._normalize_guid(doc))
            target_elem.SetEntity(entity)

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
    def _resolve_datastorage_type(cls):
        global DataStorage, DB
        if DataStorage is not None:
            return DataStorage
        if DB:
            try:
                DataStorage = DB.DataStorage
            except Exception:
                DataStorage = None
        if DataStorage is None:
            try:
                DataStorage = System.Type.GetType("Autodesk.Revit.DB.DataStorage, RevitAPI")
            except Exception:
                DataStorage = None
        if DataStorage is None:
            candidates = []
            matches = []
            try:
                for asm in System.AppDomain.CurrentDomain.GetAssemblies():
                    asm_name = None
                    try:
                        asm_name = asm.GetName().Name
                    except Exception:
                        asm_name = None
                    ds_type = None
                    try:
                        ds_type = asm.GetType("Autodesk.Revit.DB.DataStorage")
                    except Exception:
                        ds_type = None
                    if ds_type is None:
                        try:
                            ds_type = asm.GetType("Autodesk.Revit.DB.DataStorage", False)
                        except Exception:
                            ds_type = None
                    if ds_type is None:
                        try:
                            types = asm.GetExportedTypes()
                        except Exception:
                            try:
                                types = asm.GetTypes()
                            except Exception:
                                types = None
                        if types:
                            base_element = None
                            if DB:
                                try:
                                    base_element = DB.Element
                                except Exception:
                                    base_element = None
                            for t in types:
                                try:
                                    if t.Name != "DataStorage":
                                        continue
                                    ns = t.Namespace or ""
                                    if asm_name:
                                        matches.append("{} ({})".format(ns or "<no-namespace>", asm_name))
                                    is_revit_db = ns.startswith("Autodesk.Revit.DB")
                                    is_element = False
                                    if base_element is not None:
                                        try:
                                            is_element = t.IsSubclassOf(base_element)
                                        except Exception:
                                            is_element = False
                                    if is_revit_db or is_element:
                                        ds_type = t
                                        break
                                    if ns:
                                        candidates.append(ns)
                                except Exception:
                                    pass
                    if ds_type is not None:
                        DataStorage = ds_type
                        break
            except Exception:
                pass
            if DataStorage is None and not cls._datastorage_not_found_logged:
                cls._datastorage_not_found_logged = True
                try:
                    assemblies = [asm.GetName().Name for asm in System.AppDomain.CurrentDomain.GetAssemblies()]
                    revit_assemblies = [name for name in assemblies if "Revit" in name]
                    candidates = sorted(set(candidates))[:8]
                    matches = sorted(set(matches))[:8]
                    cls._log(
                        "DataStorage type not found. Loaded assemblies: {}. DataStorage candidates: {}. Matches: {}.".format(
                            ", ".join(revit_assemblies) or "<none>",
                            ", ".join(candidates) or "<none>",
                            ", ".join(matches) or "<none>",
                        )
                    )
                except Exception:
                    cls._log("DataStorage type not found; unable to enumerate loaded assemblies.")
        return DataStorage

    @classmethod
    def _find_data_storage(cls, doc, schema, doc_field=None, needed_guid=None):
        ds_type = cls._resolve_datastorage_type()
        if doc is None or ds_type is None:
            return None
        try:
            collector = FilteredElementCollector(doc).OfClass(ds_type)
        except Exception:
            return None
        for storage_elem in collector:
            try:
                entity = storage_elem.GetEntity(schema)
            except Exception:
                entity = None
            if not entity or not entity.IsValid():
                continue
            if doc_field and needed_guid:
                try:
                    doc_guid = entity.Get[str](doc_field)
                except Exception:
                    doc_guid = None
                if doc_guid and doc_guid != needed_guid:
                    continue
            return storage_elem
        return None

    @classmethod
    def _get_or_create_data_storage(cls, doc, schema, doc_field=None, needed_guid=None):
        ds_type = cls._resolve_datastorage_type()
        if ds_type is None:
            return None
        storage_elem = cls._find_data_storage(doc, schema, doc_field, needed_guid)
        if storage_elem is not None:
            return storage_elem
        try:
            return ds_type.Create(doc)
        except Exception:
            pass
        try:
            create = ds_type.GetMethod("Create")
        except Exception:
            create = None
        if create is None:
            return None
        try:
            args = Array[System.Object]([doc])
        except Exception:
            args = None
        try:
            if args is not None:
                return create.Invoke(None, args)
            return create.Invoke(None, [doc])
        except Exception:
            return None

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


def _make_doc_guid(doc, version=1):
    """Generate document-specific GUID to avoid schema conflicts between models."""
    return _make_doc_guid_versioned(doc, version=version)


def _make_doc_guid_versioned(doc, version):
    import hashlib
    import uuid
    if version == 3:
        return Guid("4a2f6b98-2b5e-4d1b-9e7e-0c92fd18b6d4")
    title = getattr(doc, "Title", "unknown")
    salt = "9f6633b1d77f49ef93905111fbb16d82"
    if version == 2:
        salt = "e7f1c8d43e7f4c4f9b4ebcc2f4b54c51"
    hash_bytes = hashlib.md5((title + salt).encode("utf-8")).digest()
    guid_str = str(uuid.UUID(bytes=hash_bytes))
    return Guid(guid_str)


__all__ = ["ExtensibleStorage"]

try:
    __revit__  # type: ignore # noqa
except Exception:
    __revit__ = None

if __revit__:
    ExtensibleStorage.ensure_undo_handler()
