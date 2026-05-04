# -*- coding: utf-8 -*-
"""
Shared 4-map ExtensibleStorage primitives for MEPRFP Automation 2.0.

The two project-level stores (equipment YAML and space classifications)
share the same schema *shape* — four typed map fields plus a DocGuid
housekeeping string. This module owns the schema builder, the
type-aware map read/write helpers, and the Int64-first / Int32-fallback
plumbing that keeps Revit 2026 happy.

Schema shape::

    StringMap   Map<String, String>
    BoolMap     Map<String, Boolean>
    IntMap      Map<String, Int64>     (try Int64 first, Int32 fallback)
    DoubleMap   Map<String, Double>    (uses SpecTypeId.Number where supported)
    DocGuid     simple String

Each store passes its own ``SCHEMA_GUID`` and ``SCHEMA_NAME`` into the
helpers below; the field names are fixed so both stores share a single
read/write code path.
"""

import clr  # noqa: F401  -- required before importing Autodesk.Revit.DB

from Autodesk.Revit.DB import (  # noqa: E402
    Document as _RevitDocument,  # noqa: F401  (re-exported for type hints)
)
from Autodesk.Revit.DB.ExtensibleStorage import (  # noqa: E402
    AccessLevel,
    Entity,
    Schema,
    SchemaBuilder,
)
from Autodesk.Revit.DB import SpecTypeId  # noqa: E402
from System import Boolean, Double, Int32, Int64, String  # noqa: E402
from System.Collections.Generic import (  # noqa: E402
    Dictionary,
    IDictionary,
)


# ---------------------------------------------------------------------
# Field names (fixed across both stores)
# ---------------------------------------------------------------------

FIELD_STRING_MAP = "StringMap"
FIELD_BOOL_MAP = "BoolMap"
FIELD_INT_MAP = "IntMap"
FIELD_DOUBLE_MAP = "DoubleMap"
FIELD_DOC_GUID = "DocGuid"


class StorageError(Exception):
    """Raised on schema build failures or unrecoverable read/write errors."""


# ---------------------------------------------------------------------
# Schema builder
# ---------------------------------------------------------------------

def get_or_create_schema(schema_guid, schema_name, documentation=""):
    """Return the v4 schema, building it on first use.

    The IntMap value type is tried as ``Int64`` first (preferred for
    Revit 2026 forward) and falls back to ``Int32`` if the running
    Revit rejects it. The schema-build attempt happens inside a single
    ``Schema.Lookup`` guard so a previous-session schema with either
    width is reused exactly as-is.

    The DoubleMap field tries ``SpecTypeId.Number`` for its display
    spec; older Revits without that token fall back to no spec (still
    works, just doesn't carry a unit hint).
    """
    schema = Schema.Lookup(schema_guid)
    if schema is not None:
        return schema

    last_err = None
    for int_type in (Int64, Int32):
        try:
            return _try_build_schema(
                schema_guid, schema_name, documentation, int_type,
            )
        except Exception as exc:
            last_err = exc
            # If schema was partially built then Lookup might now find
            # it (rare race). Try once more.
            existing = Schema.Lookup(schema_guid)
            if existing is not None:
                return existing
            continue
    raise StorageError(
        "Failed to build ExtensibleStorage schema {}: {}".format(
            schema_name, last_err,
        )
    )


def _try_build_schema(schema_guid, schema_name, documentation, int_type):
    builder = SchemaBuilder(schema_guid)
    builder.SetSchemaName(schema_name)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    if documentation:
        builder.SetDocumentation(documentation)

    builder.AddMapField(FIELD_STRING_MAP, String, String)
    builder.AddMapField(FIELD_BOOL_MAP, String, Boolean)
    builder.AddMapField(FIELD_INT_MAP, String, int_type)
    double_field = builder.AddMapField(FIELD_DOUBLE_MAP, String, Double)
    _set_double_spec(double_field)
    builder.AddSimpleField(FIELD_DOC_GUID, String)
    return builder.Finish()


def _set_double_spec(field_builder):
    if field_builder is None:
        return
    spec = getattr(SpecTypeId, "Number", None)
    if spec is None:
        return
    setter = getattr(field_builder, "SetSpec", None)
    if setter is None:
        return
    try:
        setter(spec)
    except Exception:
        pass


# ---------------------------------------------------------------------
# Int storage-type detection
# ---------------------------------------------------------------------

def _int_value_type(schema):
    """Return the actual storage type the IntMap was built with.

    Newer Revits get Int64; older builds may have had to fall back to
    Int32. We inspect the schema field directly so a project written
    in one Revit version reads correctly in another.
    """
    field = schema.GetField(FIELD_INT_MAP) if schema is not None else None
    if field is None:
        return Int64  # default for fresh-build path
    try:
        sub_type = field.SubType
    except Exception:
        sub_type = None
    # In modern Revit, ForgeTypeId-based sub_type doesn't directly tell
    # us Int32 vs Int64. ValueType is the canonical accessor.
    try:
        vt = field.ValueType
    except Exception:
        vt = None
    if vt is None:
        return Int64
    name = getattr(vt, "Name", "") or ""
    return Int32 if "32" in name else Int64


# ---------------------------------------------------------------------
# Map read / write
# ---------------------------------------------------------------------

def _read_map(entity, field_name, value_type):
    """Read one map field. Returns ``{}`` on missing / error."""
    field = entity.Schema.GetField(field_name)
    if field is None:
        return {}
    iface = IDictionary[String, value_type]
    try:
        net_map = entity.Get[iface](field)
    except Exception:
        try:
            net_map = entity.Get[Dictionary[String, value_type]](field)
        except Exception:
            return {}
    if net_map is None:
        return {}
    out = {}
    try:
        for key in net_map.Keys:
            out[str(key)] = net_map[key]
        return out
    except Exception:
        try:
            for pair in net_map:
                out[str(pair.Key)] = pair.Value
            return out
        except Exception:
            return {}


def _write_map(entity, field_name, value_type, py_dict):
    """Write one map field. ``py_dict`` is ``{str: value}``."""
    field = entity.Schema.GetField(field_name)
    if field is None:
        return False

    net_map = Dictionary[String, value_type]()
    for k, v in (py_dict or {}).items():
        if k is None:
            continue
        key = String(str(k))
        if value_type == String:
            net_map[key] = String("" if v is None else str(v))
        elif value_type == Boolean:
            net_map[key] = Boolean(bool(v))
        elif value_type == Int64:
            net_map[key] = Int64(int(v) if v is not None else 0)
        elif value_type == Int32:
            net_map[key] = Int32(int(v) if v is not None else 0)
        elif value_type == Double:
            net_map[key] = Double(float(v) if v is not None else 0.0)
        else:
            net_map[key] = v

    iface = IDictionary[String, value_type]
    try:
        entity.Set[iface](field, net_map)
        return True
    except Exception:
        try:
            entity.Set[Dictionary[String, value_type]](field, net_map)
            return True
        except Exception:
            return False


# ---------------------------------------------------------------------
# Public read/write API
# ---------------------------------------------------------------------

def read_maps(entity):
    """Return ``{string_map, bool_map, int_map, double_map, doc_guid}``.

    Empty dicts on missing fields. Caller decides which keys it cares
    about.
    """
    if entity is None or not entity.IsValid():
        return None
    schema = entity.Schema
    int_type = _int_value_type(schema)

    string_map = _read_map(entity, FIELD_STRING_MAP, String)
    bool_map = _read_map(entity, FIELD_BOOL_MAP, Boolean)
    int_map_raw = _read_map(entity, FIELD_INT_MAP, int_type)
    # Coerce to plain Python int so callers don't have to think about
    # the underlying CLR type.
    int_map = {k: int(v) for k, v in int_map_raw.items()}
    double_map = {k: float(v) for k, v in _read_map(
        entity, FIELD_DOUBLE_MAP, Double,
    ).items()}

    doc_guid = ""
    try:
        guid_field = schema.GetField(FIELD_DOC_GUID)
        if guid_field is not None:
            doc_guid = entity.Get[String](guid_field) or ""
    except Exception:
        doc_guid = ""

    return {
        "string_map": string_map,
        "bool_map": bool_map,
        "int_map": int_map,
        "double_map": double_map,
        "doc_guid": doc_guid,
    }


def build_entity(schema, string_map=None, bool_map=None, int_map=None,
                 double_map=None, doc_guid=""):
    """Build a fresh ``Entity`` from python dicts for each map."""
    int_type = _int_value_type(schema)
    entity = Entity(schema)
    _write_map(entity, FIELD_STRING_MAP, String, string_map or {})
    _write_map(entity, FIELD_BOOL_MAP, Boolean, bool_map or {})
    _write_map(entity, FIELD_INT_MAP, int_type, int_map or {})
    _write_map(entity, FIELD_DOUBLE_MAP, Double, double_map or {})
    try:
        guid_field = schema.GetField(FIELD_DOC_GUID)
        if guid_field is not None:
            entity.Set[String](guid_field, String(doc_guid or ""))
    except Exception:
        pass
    return entity


# ---------------------------------------------------------------------
# Project-info helpers
# ---------------------------------------------------------------------

def project_info_or_raise(doc):
    pi = doc.ProjectInformation
    if pi is None:
        raise StorageError("Document has no ProjectInformation element")
    return pi


def get_entity(doc, schema):
    """Return the stored entity on ProjectInformation, or None."""
    pi = project_info_or_raise(doc)
    entity = pi.GetEntity(schema)
    if entity is None or not entity.IsValid():
        return None
    return entity


def set_entity(doc, entity):
    """Persist the entity on ProjectInformation. Caller manages txn."""
    pi = project_info_or_raise(doc)
    pi.SetEntity(entity)


def delete_entity(doc, schema):
    pi = project_info_or_raise(doc)
    pi.DeleteEntity(schema)
