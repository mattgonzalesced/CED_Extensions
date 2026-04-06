# -*- coding: utf-8 -*-
"""
Active YAML storage and user settings stored via Revit Extensible Storage.
"""

import base64
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
from System import Guid, String, Array, Boolean, Int32, Int64, Double  # noqa: E402

try:
    from Autodesk.Revit.UI.Events import DocumentSynchronizedWithCentralEventArgs
except Exception:  # pragma: no cover
    DocumentSynchronizedWithCentralEventArgs = None


try:
    from pyrevit import script as _script_logger  # noqa: E402
    _logger = _script_logger.get_logger()
except Exception:
    _logger = None

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str


class ExtensibleStorage(object):
    """
    Stores the active YAML payload, user settings, and editor locks inside the RVT.
    """

    SCHEMA_NAME = "CED_YamlHistory"
    HISTORY_FIELD_NAME = "HistoryJson"
    META_FIELD_NAME = "MetadataJson"
    HISTORY_CHUNKS_FIELD_NAME = "HistoryJsonChunks"
    META_CHUNKS_FIELD_NAME = "MetadataJsonChunks"
    SCHEMA_VERSION = 4
    MAX_ES_STRING = 16 * 1024 * 1024
    CHUNK_SIZE = 8 * 1024 * 1024

    USER_SETTINGS_KEY = "user_settings"
    PROJECT_DATA_KEY = "project_data"
    STRING_MAP_FIELD_NAME = "StringMap"
    BOOL_MAP_FIELD_NAME = "BoolMap"
    INT_MAP_FIELD_NAME = "IntMap"
    DOUBLE_MAP_FIELD_NAME = "DoubleMap"

    PAYLOAD_HISTORY_MAP_KEY = "__ced.payload.history_json"
    PAYLOAD_META_MAP_KEY = "__ced.payload.meta_json"
    MAP_CHUNK_COUNT_SUFFIX = ".__chunk_count"
    MAP_CHUNK_KEY_PREFIX = ".__chunk__"
    USER_SETTING_MAP_PREFIX = "__ced.user_setting__::"
    DOUBLE_STRING_FALLBACK_PREFIX = "__ced.double__::"

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
    def get_project_data(cls, doc, storage_id, default=None):
        """
        Read project-scoped payload stored under a logical storage ID.
        """
        if doc is None or not storage_id:
            return default
        payload = cls._read_storage(doc)
        meta = payload.get("meta", {})
        project_data = meta.get(cls.PROJECT_DATA_KEY) or {}
        if not isinstance(project_data, dict):
            return default
        if storage_id not in project_data:
            return default
        return project_data.get(storage_id)

    @classmethod
    def set_project_data(cls, doc, storage_id, value, transaction_name=None):
        """
        Write project-scoped payload under a logical storage ID.
        """
        if doc is None or not storage_id:
            return False
        payload = cls._read_storage(doc)
        meta = payload.setdefault("meta", {})
        project_data = meta.get(cls.PROJECT_DATA_KEY)
        if not isinstance(project_data, dict):
            project_data = {}
            meta[cls.PROJECT_DATA_KEY] = project_data
        project_data[storage_id] = value
        txn_name = transaction_name or "PROJECT_DATA::{}".format(storage_id)
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
        needed_guid = cls._normalize_guid(doc)

        if force_version == 4 and cache.get("v4"):
            return cache["v4"]
        if force_version == 3 and cache.get("v3"):
            return cache["v3"]
        if not force_version and cache.get("default"):
            cached_default = cache.get("default")
            try:
                cached_schema = cached_default[0]
                cached_doc_field = cached_default[5]
                cached_version = cached_default[10] if len(cached_default) > 10 else None
                # If we previously cached v3, but v4 data now exists, promote reads to v4.
                if cached_version != 4:
                    packed_v4 = cache.get("v4")
                    if packed_v4 is None:
                        schema_v4 = Schema.Lookup(_make_doc_guid(doc, version=4))
                        if schema_v4 is not None:
                            packed_v4 = cls._pack_schema(schema_v4, version=4)
                            cache["v4"] = packed_v4
                    if packed_v4 is not None:
                        if cls._find_data_storage(doc, packed_v4[0], packed_v4[5], needed_guid) is not None:
                            cache["default"] = packed_v4
                            cls._schema_cache[doc_key] = cache
                            return packed_v4
                if cls._find_data_storage(doc, cached_schema, cached_doc_field, needed_guid) is not None:
                    return cached_default
            except Exception:
                pass

        if force_version == 4:
            schema_v4 = Schema.Lookup(_make_doc_guid(doc, version=4))
            if schema_v4 is None:
                schema_v4 = cls._build_schema_v4(doc)
            packed = cls._pack_schema(schema_v4, version=4)
            cache["v4"] = packed
            cache["default"] = packed
            cls._schema_cache[doc_key] = cache
            return packed

        if force_version == 3:
            schema_v3 = Schema.Lookup(_make_doc_guid(doc, version=3))
            if schema_v3 is None:
                schema_v3 = cls._build_schema_v3(doc)
            packed = cls._pack_schema(schema_v3, version=3)
            cache["v3"] = packed
            cache["default"] = packed
            cls._schema_cache[doc_key] = cache
            return packed

        packed_v4 = None
        schema_v4 = Schema.Lookup(_make_doc_guid(doc, version=4))
        if schema_v4 is not None:
            packed_v4 = cls._pack_schema(schema_v4, version=4)
            cache["v4"] = packed_v4
            if cls._find_data_storage(doc, schema_v4, packed_v4[5], needed_guid) is not None:
                cache["default"] = packed_v4
                cls._schema_cache[doc_key] = cache
                return packed_v4

        packed_v3 = None
        schema_v3 = Schema.Lookup(_make_doc_guid(doc, version=3))
        if schema_v3 is not None:
            packed_v3 = cls._pack_schema(schema_v3, version=3)
            cache["v3"] = packed_v3
            if cls._find_data_storage(doc, schema_v3, packed_v3[5], needed_guid) is not None:
                cache["default"] = packed_v3
                cls._schema_cache[doc_key] = cache
                return packed_v3

        # No existing data for this document; prefer v4 for new writes.
        if packed_v4 is not None:
            cache["default"] = packed_v4
            cls._schema_cache[doc_key] = cache
            return packed_v4

        schema_v4 = cls._build_schema_v4(doc)
        packed = cls._pack_schema(schema_v4, version=4)
        cache["v4"] = packed
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
    def _build_schema_v4(cls, doc):
        builder = SchemaBuilder(_make_doc_guid(doc, version=4))
        builder.SetSchemaName("{}_v4".format(cls.SCHEMA_NAME))
        builder.SetDocumentation("Stores active YAML payload and metadata in typed map fields.")
        builder.AddMapField(cls.STRING_MAP_FIELD_NAME, String, String)
        builder.AddMapField(cls.BOOL_MAP_FIELD_NAME, String, Boolean)
        cls._add_int_map_field(builder)
        double_map_builder = builder.AddMapField(cls.DOUBLE_MAP_FIELD_NAME, String, Double)
        cls._set_double_field_units(double_map_builder)
        builder.AddSimpleField("DocGuid", String)
        return builder.Finish()

    @classmethod
    def _add_int_map_field(cls, schema_builder):
        try:
            return schema_builder.AddMapField(cls.INT_MAP_FIELD_NAME, String, Int32)
        except Exception:
            return schema_builder.AddMapField(cls.INT_MAP_FIELD_NAME, String, Int64)

    @classmethod
    def _set_double_field_units(cls, field_builder):
        if field_builder is None:
            return
        # Revit requires a measurable spec for Double fields in Extensible Storage.
        try:
            spec_id = getattr(getattr(DB, "SpecTypeId", None), "Number", None)
            if spec_id is not None and hasattr(field_builder, "SetSpec"):
                field_builder.SetSpec(spec_id)
                return
        except Exception:
            pass
        try:
            unit_type = getattr(getattr(DB, "UnitType", None), "UT_Number", None)
            if unit_type is not None and hasattr(field_builder, "SetUnitType"):
                field_builder.SetUnitType(unit_type)
                return
        except Exception:
            pass
        try:
            unit_type = getattr(getattr(DB, "UnitType", None), "UT_Custom", None)
            if unit_type is not None and hasattr(field_builder, "SetUnitType"):
                field_builder.SetUnitType(unit_type)
        except Exception:
            pass

    @classmethod
    def _safe_get_field(cls, schema, field_name):
        if schema is None or not field_name:
            return None
        try:
            return schema.GetField(field_name)
        except Exception:
            return None

    @classmethod
    def _pack_schema(cls, schema, version):
        history_field = cls._safe_get_field(schema, cls.HISTORY_FIELD_NAME)
        meta_field = cls._safe_get_field(schema, cls.META_FIELD_NAME)
        history_chunks_field = cls._safe_get_field(schema, cls.HISTORY_CHUNKS_FIELD_NAME)
        meta_chunks_field = cls._safe_get_field(schema, cls.META_CHUNKS_FIELD_NAME)
        doc_field = cls._safe_get_field(schema, "DocGuid")
        string_map_field = cls._safe_get_field(schema, cls.STRING_MAP_FIELD_NAME)
        bool_map_field = cls._safe_get_field(schema, cls.BOOL_MAP_FIELD_NAME)
        int_map_field = cls._safe_get_field(schema, cls.INT_MAP_FIELD_NAME)
        double_map_field = cls._safe_get_field(schema, cls.DOUBLE_MAP_FIELD_NAME)
        return (
            schema,
            history_field,
            meta_field,
            history_chunks_field,
            meta_chunks_field,
            doc_field,
            string_map_field,
            bool_map_field,
            int_map_field,
            double_map_field,
            version,
        )

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
    def _read_map_field(cls, entity, map_field, value_type):
        if not map_field:
            return {}
        try:
            from System.Collections.Generic import IDictionary, Dictionary
        except Exception:
            IDictionary = None
            Dictionary = None

        iface_type = None
        dict_type = None
        try:
            if value_type == String:
                iface_type = IDictionary[String, String] if IDictionary else None
                dict_type = Dictionary[String, String] if Dictionary else None
            elif value_type == Boolean:
                iface_type = IDictionary[String, Boolean] if IDictionary else None
                dict_type = Dictionary[String, Boolean] if Dictionary else None
            elif value_type == Int32:
                iface_type = IDictionary[String, Int32] if IDictionary else None
                dict_type = Dictionary[String, Int32] if Dictionary else None
            elif value_type == Int64:
                iface_type = IDictionary[String, Int64] if IDictionary else None
                dict_type = Dictionary[String, Int64] if Dictionary else None
            elif value_type == Double:
                iface_type = IDictionary[String, Double] if IDictionary else None
                dict_type = Dictionary[String, Double] if Dictionary else None
        except Exception:
            iface_type = None
            dict_type = None

        if iface_type is None and dict_type is None:
            return {}

        map_data = None
        if iface_type is not None:
            try:
                map_data = entity.Get[iface_type](map_field)
            except Exception:
                map_data = None
        if map_data is None and dict_type is not None:
            try:
                map_data = entity.Get[dict_type](map_field)
            except Exception:
                map_data = None
        if map_data is None:
            return {}

        result = {}
        try:
            for key in map_data.Keys:
                result[str(key)] = map_data[key]
            return result
        except Exception:
            pass
        try:
            for pair in map_data:
                result[str(pair.Key)] = pair.Value
        except Exception:
            return {}
        return result

    @classmethod
    def _write_map_field(cls, entity, map_field, value_type, values):
        if not map_field:
            return
        values = values or {}
        try:
            from System.Collections.Generic import IDictionary, Dictionary
        except Exception:
            IDictionary = None
            Dictionary = None
        iface_type = None
        dict_type = None
        if value_type == String:
            iface_type = IDictionary[String, String] if IDictionary else None
            dict_type = Dictionary[String, String] if Dictionary else None
        elif value_type == Boolean:
            iface_type = IDictionary[String, Boolean] if IDictionary else None
            dict_type = Dictionary[String, Boolean] if Dictionary else None
        elif value_type == Int32:
            iface_type = IDictionary[String, Int32] if IDictionary else None
            dict_type = Dictionary[String, Int32] if Dictionary else None
        elif value_type == Int64:
            iface_type = IDictionary[String, Int64] if IDictionary else None
            dict_type = Dictionary[String, Int64] if Dictionary else None
        elif value_type == Double:
            iface_type = IDictionary[String, Double] if IDictionary else None
            dict_type = Dictionary[String, Double] if Dictionary else None
        if dict_type is None:
            raise RuntimeError("Map write type is unavailable for value type: {}".format(value_type))

        map_obj = dict_type()
        for key, value in values.items():
            if key is None:
                continue
            typed_key = str(key)
            if value_type == String:
                map_obj[typed_key] = str(value or "")
            elif value_type == Boolean:
                map_obj[typed_key] = bool(value)
            elif value_type == Int32:
                map_obj[typed_key] = Int32(int(value))
            elif value_type == Int64:
                map_obj[typed_key] = Int64(int(value))
            elif value_type == Double:
                map_obj[typed_key] = float(value)
            else:
                map_obj[typed_key] = value

        set_error_1 = None
        set_error_2 = None
        if iface_type is not None:
            try:
                entity.Set[iface_type](map_field, map_obj)
                return
            except Exception as ex1:
                set_error_1 = ex1
        try:
            entity.Set[dict_type](map_field, map_obj)
            return
        except Exception as ex2:
            set_error_2 = ex2
        raise RuntimeError(
            "Failed to write map field '{}'. Errors: {} | {}".format(
                map_field.FieldName,
                set_error_1,
                set_error_2,
            )
        )

    @classmethod
    def _read_int_map_field(cls, entity, map_field):
        int_map = cls._read_map_field(entity, map_field, Int32)
        if int_map:
            return int_map
        return cls._read_map_field(entity, map_field, Int64)

    @classmethod
    def _write_int_map_field(cls, entity, map_field, values):
        errors = []
        try:
            cls._write_map_field(entity, map_field, Int32, values)
            return
        except Exception as ex32:
            errors.append(ex32)
        try:
            cls._write_map_field(entity, map_field, Int64, values)
            return
        except Exception as ex64:
            errors.append(ex64)
        raise RuntimeError(
            "Failed to write int map field '{}': {}".format(
                map_field.FieldName if map_field else "<unknown>",
                " | ".join([str(err) for err in errors]),
            )
        )

    @classmethod
    def _map_chunk_count_key(cls, base_key):
        return "{}{}".format(base_key, cls.MAP_CHUNK_COUNT_SUFFIX)

    @classmethod
    def _map_chunk_key(cls, base_key, index):
        return "{}{}{}".format(base_key, cls.MAP_CHUNK_KEY_PREFIX, int(index))

    @classmethod
    def _clear_chunked_map_value(cls, string_map, int_map, base_key):
        string_map.pop(base_key, None)
        count_key = cls._map_chunk_count_key(base_key)
        count = int_map.pop(count_key, None)
        try:
            count = int(count or 0)
        except Exception:
            count = 0
        for idx in range(max(0, count)):
            string_map.pop(cls._map_chunk_key(base_key, idx), None)

    @classmethod
    def _write_payload_json_to_maps(cls, string_map, int_map, base_key, value):
        value = value or ""
        cls._clear_chunked_map_value(string_map, int_map, base_key)
        if len(value) <= cls.MAX_ES_STRING:
            string_map[base_key] = value
            return
        chunks = cls._split_chunks(value, cls.CHUNK_SIZE)
        int_map[cls._map_chunk_count_key(base_key)] = int(len(chunks))
        for idx, chunk in enumerate(chunks):
            string_map[cls._map_chunk_key(base_key, idx)] = chunk

    @classmethod
    def _read_payload_json_from_maps(cls, string_map, int_map, base_key):
        count = int_map.get(cls._map_chunk_count_key(base_key))
        try:
            count = int(count or 0)
        except Exception:
            count = 0
        if count > 0:
            parts = []
            for idx in range(count):
                parts.append(string_map.get(cls._map_chunk_key(base_key, idx), ""))
            return "".join(parts)
        return string_map.get(base_key, "")

    @classmethod
    def _encode_setting_storage_key(cls, setting_key, user):
        payload = json.dumps([setting_key or "", user or ""], separators=(",", ":"))
        encoded = base64.b64encode(payload.encode("utf-8")).decode("ascii")
        return "{}{}".format(cls.USER_SETTING_MAP_PREFIX, encoded)

    @classmethod
    def _encode_double_fallback(cls, value):
        try:
            return "{}{}".format(cls.DOUBLE_STRING_FALLBACK_PREFIX, repr(float(value)))
        except Exception:
            return "{}0.0".format(cls.DOUBLE_STRING_FALLBACK_PREFIX)

    @classmethod
    def _decode_double_fallback(cls, value):
        if not isinstance(value, basestring):
            return None
        if not value.startswith(cls.DOUBLE_STRING_FALLBACK_PREFIX):
            return None
        raw = value[len(cls.DOUBLE_STRING_FALLBACK_PREFIX):]
        try:
            return float(raw)
        except Exception:
            return None

    @classmethod
    def _decode_setting_storage_key(cls, encoded_key):
        if not encoded_key or not encoded_key.startswith(cls.USER_SETTING_MAP_PREFIX):
            return None, None
        suffix = encoded_key[len(cls.USER_SETTING_MAP_PREFIX):]
        try:
            decoded = base64.b64decode(suffix.encode("ascii"))
            payload = json.loads(decoded.decode("utf-8"))
            if isinstance(payload, list) and len(payload) == 2:
                return str(payload[0]), str(payload[1])
        except Exception:
            pass
        return None, None

    @classmethod
    def _typed_user_settings_maps_from_meta(cls, meta):
        settings = {}
        if isinstance(meta, dict):
            settings = meta.get(cls.USER_SETTINGS_KEY) or {}
        if not isinstance(settings, dict):
            settings = {}

        str_map = {}
        bool_map = {}
        int_map = {}
        dbl_map = {}
        for setting_key, setting_map in settings.items():
            if not isinstance(setting_map, dict):
                continue
            for user, value in setting_map.items():
                map_key = cls._encode_setting_storage_key(setting_key, user)
                if isinstance(value, bool):
                    bool_map[map_key] = bool(value)
                elif isinstance(value, int):
                    int_map[map_key] = int(value)
                elif isinstance(value, float):
                    dbl_map[map_key] = float(value)
                elif value is None:
                    continue
                else:
                    str_map[map_key] = str(value)
        return str_map, bool_map, int_map, dbl_map

    @classmethod
    def _overlay_user_settings_from_maps(cls, meta, str_map, bool_map, int_map, dbl_map):
        if not isinstance(meta, dict):
            return
        settings = meta.get(cls.USER_SETTINGS_KEY) or {}
        if not isinstance(settings, dict):
            settings = {}

        def _apply(raw_map, decode_doubles=False):
            for key, value in (raw_map or {}).items():
                setting_key, user = cls._decode_setting_storage_key(key)
                if not setting_key or not user:
                    continue
                if decode_doubles:
                    decoded = cls._decode_double_fallback(value)
                    if decoded is not None:
                        value = decoded
                setting_map = settings.get(setting_key)
                if not isinstance(setting_map, dict):
                    setting_map = {}
                    settings[setting_key] = setting_map
                setting_map[user] = value

        _apply(str_map, decode_doubles=True)
        _apply(bool_map)
        _apply(int_map)
        _apply(dbl_map)
        if settings:
            meta[cls.USER_SETTINGS_KEY] = settings

    @classmethod
    def _read_storage(cls, doc):
        (
            schema,
            history_field,
            meta_field,
            history_chunks_field,
            meta_chunks_field,
            doc_field,
            string_map_field,
            bool_map_field,
            int_map_field,
            double_map_field,
            version,
        ) = cls._schema_and_fields(doc)
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

        str_map = {}
        bool_map = {}
        int_map = {}
        dbl_map = {}
        if version >= 4:
            str_map = cls._read_map_field(entity, string_map_field, String)
            bool_map = cls._read_map_field(entity, bool_map_field, Boolean)
            int_map = cls._read_int_map_field(entity, int_map_field)
            dbl_map = cls._read_map_field(entity, double_map_field, Double)
            history_json = cls._read_payload_json_from_maps(str_map, int_map, cls.PAYLOAD_HISTORY_MAP_KEY)
            meta_json = cls._read_payload_json_from_maps(str_map, int_map, cls.PAYLOAD_META_MAP_KEY)
        else:
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
        if version >= 4:
            cls._overlay_user_settings_from_maps(payload["meta"], str_map, bool_map, int_map, dbl_map)
        return payload

    @classmethod
    def _write_storage(cls, doc, payload, transaction_name=None):
        history_json = json.dumps(payload.get("entries", []))
        meta_json = json.dumps(payload.get("meta", {}))
        needs_chunking = (
            len(history_json or "") > cls.MAX_ES_STRING
            or len(meta_json or "") > cls.MAX_ES_STRING
        )
        (
            schema,
            history_field,
            meta_field,
            history_chunks_field,
            meta_chunks_field,
            doc_field,
            string_map_field,
            bool_map_field,
            int_map_field,
            double_map_field,
            version,
        ) = cls._schema_and_fields_versioned(
            doc,
            prefer_chunked=needs_chunking,
            force_version=4,
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

            if version >= 4:
                string_map = {}
                bool_map = {}
                int_map = {}
                dbl_map = {}

                cls._write_payload_json_to_maps(string_map, int_map, cls.PAYLOAD_HISTORY_MAP_KEY, history_json)
                cls._write_payload_json_to_maps(string_map, int_map, cls.PAYLOAD_META_MAP_KEY, meta_json)

                setting_str_map, setting_bool_map, setting_int_map, setting_dbl_map = cls._typed_user_settings_maps_from_meta(
                    payload.get("meta", {})
                )
                string_map.update(setting_str_map)
                bool_map.update(setting_bool_map)
                int_map.update(setting_int_map)
                dbl_map.update(setting_dbl_map)

                if dbl_map:
                    try:
                        cls._write_map_field(entity, double_map_field, Double, dbl_map)
                    except Exception as ex:
                        for key, value in dbl_map.items():
                            string_map[key] = cls._encode_double_fallback(value)
                        cls._log(
                            "DoubleMap write failed; storing {} value(s) in StringMap fallback. Error: {}".format(
                                len(dbl_map),
                                ex,
                            )
                        )

                cls._write_map_field(entity, string_map_field, String, string_map)
                cls._write_map_field(entity, bool_map_field, Boolean, bool_map)
                cls._write_int_map_field(entity, int_map_field, int_map)
            else:
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
    if version == 4:
        return Guid("2f3a4aa6-5f43-4f89-a16f-0b9f7c68da82")
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
