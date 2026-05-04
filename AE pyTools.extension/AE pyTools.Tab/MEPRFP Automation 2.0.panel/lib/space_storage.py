# -*- coding: utf-8 -*-
"""
Extensible Storage primitives for per-project Space classifications.

Space *templates* (``space_buckets`` and ``space_profiles``) live in the
shared YAML payload alongside ``equipment_definitions`` — they are the
exportable, reusable configuration. Per-project *classifications*
(which Space element belongs to which bucket) are project state, not
template, and therefore live in this separate Extensible Storage entity
on ``ProjectInformation`` so an export of the YAML doesn't drag one
project's space assignments into another.

Layout (v4 — current). Same 4-map shape as ``storage.py`` so both
stores share machinery. The classification list is JSON-encoded into
``StringMap['json_text']``; ``IntMap['store_version']`` carries the
layout version. JSON is dependency-free here (no PyYAML in this hot
path) and stays human-readable when poked with the Revit Lookup
add-in.

Reads fall back to the legacy v1 schema (single ``JsonText`` simple
field under GUID ``b5e8c1a2-…``) if the v4 entity is missing.
"""

import clr  # noqa: F401  -- needed before importing Autodesk.Revit.DB
import json

from Autodesk.Revit.DB.ExtensibleStorage import (  # noqa: E402
    Schema,
)
from System import Guid, Int32, String  # noqa: E402

import _es_v4  # noqa: E402


# ---------------------------------------------------------------------
# Schema GUIDs
# ---------------------------------------------------------------------

# v4 (current) — 4-map layout.
SCHEMA_GUID_STR = "c1f5d4a8-6e3b-4d92-8a47-f1e9c2b5a8d6"
SCHEMA_GUID = Guid(SCHEMA_GUID_STR)
SCHEMA_NAME = "MEPRFP_Automation_2_SpaceClassifications_v4"
SCHEMA_DOC = (
    "MEPRFP Automation 2.0 per-project Space classifications. JSON list "
    "lives in StringMap['json_text']."
)

# v1 (legacy in-2.0) — simple-fields layout. Read-only fallback.
LEGACY_V1_SCHEMA_GUID_STR = "b5e8c1a2-2d6f-4a17-9c3d-7e4b1f0a8d6e"
LEGACY_V1_SCHEMA_GUID = Guid(LEGACY_V1_SCHEMA_GUID_STR)
LEGACY_V1_SCHEMA_NAME = "MEPRFP_Automation_2_SpaceClassifications"
LEGACY_V1_FIELD_STORE_VERSION = "StoreVersion"
LEGACY_V1_FIELD_JSON_TEXT = "JsonText"
LEGACY_V1_FIELD_LAST_MODIFIED_UTC = "LastModifiedUtc"

STORE_LAYOUT_VERSION = 1


# ---------------------------------------------------------------------
# Map keys
# ---------------------------------------------------------------------

KEY_JSON_TEXT = "json_text"
KEY_LAST_MODIFIED_UTC = "last_modified_utc"
KEY_STORE_VERSION = "store_version"


# ---------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------

class SpaceStorageError(_es_v4.StorageError):
    pass


# ---------------------------------------------------------------------
# Schema accessors
# ---------------------------------------------------------------------

def get_or_create_schema():
    return _es_v4.get_or_create_schema(SCHEMA_GUID, SCHEMA_NAME, SCHEMA_DOC)


def _legacy_v1_schema():
    return Schema.Lookup(LEGACY_V1_SCHEMA_GUID)


# ---------------------------------------------------------------------
# Read
# ---------------------------------------------------------------------

def read_payload(doc):
    """Return the stored payload as a dict, or ``None`` if no entity exists.

    Tries the v4 entity first; falls back to legacy v1 if v4 is
    missing. Returned shape::

        {
          "store_version": int,
          "json_text": str,
          "last_modified_utc": str,
          "_legacy_v1": bool,
        }
    """
    payload = _read_v4(doc)
    if payload is not None:
        payload["_legacy_v1"] = False
        return payload

    payload = _read_legacy_v1(doc)
    if payload is not None:
        payload["_legacy_v1"] = True
        return payload

    return None


def _read_v4(doc):
    schema = get_or_create_schema()
    entity = _es_v4.get_entity(doc, schema)
    if entity is None:
        return None
    maps = _es_v4.read_maps(entity)
    if maps is None:
        return None
    sm = maps["string_map"]
    im = maps["int_map"]
    return {
        "store_version": int(im.get(KEY_STORE_VERSION) or STORE_LAYOUT_VERSION),
        "json_text": sm.get(KEY_JSON_TEXT) or "",
        "last_modified_utc": sm.get(KEY_LAST_MODIFIED_UTC) or "",
    }


def _read_legacy_v1(doc):
    schema = _legacy_v1_schema()
    if schema is None:
        return None
    pi = _es_v4.project_info_or_raise(doc)
    entity = pi.GetEntity(schema)
    if entity is None or not entity.IsValid():
        return None
    return {
        "store_version": int(entity.Get[Int32](LEGACY_V1_FIELD_STORE_VERSION) or 0),
        "json_text": entity.Get[String](LEGACY_V1_FIELD_JSON_TEXT) or "",
        "last_modified_utc": entity.Get[String](LEGACY_V1_FIELD_LAST_MODIFIED_UTC) or "",
    }


# ---------------------------------------------------------------------
# Write
# ---------------------------------------------------------------------

def write_payload(doc, json_text, last_modified_utc):
    """Persist a JSON-encoded classification list. Caller manages txn."""
    schema = get_or_create_schema()
    entity = _es_v4.build_entity(
        schema,
        string_map={
            KEY_JSON_TEXT: json_text or "",
            KEY_LAST_MODIFIED_UTC: last_modified_utc or "",
        },
        int_map={
            KEY_STORE_VERSION: STORE_LAYOUT_VERSION,
        },
    )
    _es_v4.set_entity(doc, entity)


def clear_payload(doc):
    """Delete the v4 classifications entity. Caller manages the txn.

    The legacy v1 entity (if present) is left untouched. Use
    ``clear_legacy_v1_payload`` to remove it explicitly.
    """
    schema = get_or_create_schema()
    _es_v4.delete_entity(doc, schema)


def clear_legacy_v1_payload(doc):
    schema = _legacy_v1_schema()
    if schema is None:
        return
    pi = _es_v4.project_info_or_raise(doc)
    try:
        pi.DeleteEntity(schema)
    except Exception:
        pass


# ---------------------------------------------------------------------
# JSON codec  (unchanged across v1 -> v4)
# ---------------------------------------------------------------------

def encode(classifications):
    """Encode a list of classification dicts to JSON text."""
    safe = []
    for entry in classifications or ():
        if not isinstance(entry, dict):
            continue
        safe.append({
            "space_element_id": _coerce_int(entry.get("space_element_id")),
            "bucket_id": str(entry.get("bucket_id") or ""),
            "space_name": str(entry.get("space_name") or ""),
        })
    return json.dumps(safe, indent=2, sort_keys=True)


def decode(text):
    """Parse the JSON text back into a list of classification dicts."""
    if not text or not text.strip():
        return []
    try:
        data = json.loads(text)
    except (ValueError, TypeError) as exc:
        raise SpaceStorageError(
            "Failed to decode space classifications JSON: {}".format(exc)
        )
    if not isinstance(data, list):
        return []
    out = []
    for entry in data:
        if not isinstance(entry, dict):
            continue
        out.append({
            "space_element_id": _coerce_int(entry.get("space_element_id")),
            "bucket_id": str(entry.get("bucket_id") or ""),
            "space_name": str(entry.get("space_name") or ""),
        })
    return out


def _coerce_int(value):
    if value is None:
        return None
    try:
        return int(value)
    except (ValueError, TypeError):
        return None


# ---------------------------------------------------------------------
# Convenience
# ---------------------------------------------------------------------

def has_v4_entity(doc):
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema is None:
        return False
    return _es_v4.get_entity(doc, schema) is not None


def has_legacy_v1_entity(doc):
    schema = _legacy_v1_schema()
    if schema is None:
        return False
    pi = _es_v4.project_info_or_raise(doc)
    entity = pi.GetEntity(schema)
    return entity is not None and entity.IsValid()
