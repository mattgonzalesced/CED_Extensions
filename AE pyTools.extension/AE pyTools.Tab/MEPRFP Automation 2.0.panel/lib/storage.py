# -*- coding: utf-8 -*-
"""
Extensible Storage primitives for the MEPRFP 2.0 active YAML store.

The schema is independent from the original MEP Automation panel. Tools
in the 2.0 panel only ever read and write this schema, and the legacy
panel never sees it. The schema layout itself is versioned via
STORE_LAYOUT_VERSION so we can migrate without changing the GUID later.
"""

import clr  # noqa: F401  -- needed before importing Autodesk.Revit.DB

from Autodesk.Revit.DB.ExtensibleStorage import (
    Schema,
    SchemaBuilder,
    AccessLevel,
    Entity,
)
from System import Guid, Int32, String


SCHEMA_GUID_STR = "a7d4e2f1-9c3b-4e8a-b6d5-f3c1a8e9b2d4"
SCHEMA_GUID = Guid(SCHEMA_GUID_STR)
SCHEMA_NAME = "MEPRFP_Automation_2_YamlStore"
SCHEMA_DOC = "MEPRFP Automation 2.0 active YAML storage"

FIELD_STORE_VERSION = "StoreVersion"
FIELD_YAML_TEXT = "YamlText"
FIELD_SOURCE_PATH = "SourcePath"
FIELD_SCHEMA_VERSION = "SchemaVersion"
FIELD_LAST_MODIFIED_UTC = "LastModifiedUtc"

STORE_LAYOUT_VERSION = 1


class StorageError(Exception):
    pass


def get_or_create_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema is not None:
        return schema
    builder = SchemaBuilder(SCHEMA_GUID)
    builder.SetSchemaName(SCHEMA_NAME)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.SetDocumentation(SCHEMA_DOC)
    builder.AddSimpleField(FIELD_STORE_VERSION, Int32)
    builder.AddSimpleField(FIELD_YAML_TEXT, String)
    builder.AddSimpleField(FIELD_SOURCE_PATH, String)
    builder.AddSimpleField(FIELD_SCHEMA_VERSION, Int32)
    builder.AddSimpleField(FIELD_LAST_MODIFIED_UTC, String)
    return builder.Finish()


def _project_info(doc):
    pi = doc.ProjectInformation
    if pi is None:
        raise StorageError("Document has no ProjectInformation element")
    return pi


def read_payload(doc):
    """Return the stored payload as a dict, or None if no entity exists."""
    schema = get_or_create_schema()
    pi = _project_info(doc)
    entity = pi.GetEntity(schema)
    if entity is None or not entity.IsValid():
        return None
    return {
        "store_version": entity.Get[Int32](FIELD_STORE_VERSION),
        "yaml_text": entity.Get[String](FIELD_YAML_TEXT) or "",
        "source_path": entity.Get[String](FIELD_SOURCE_PATH) or "",
        "schema_version": entity.Get[Int32](FIELD_SCHEMA_VERSION),
        "last_modified_utc": entity.Get[String](FIELD_LAST_MODIFIED_UTC) or "",
    }


def write_payload(doc, yaml_text, source_path, schema_version, last_modified_utc):
    """Persist the payload onto ProjectInformation.

    The caller is responsible for opening a Revit transaction; this function
    only mutates the model and will fail if no transaction is active.
    """
    schema = get_or_create_schema()
    pi = _project_info(doc)
    entity = Entity(schema)
    entity.Set[Int32](FIELD_STORE_VERSION, STORE_LAYOUT_VERSION)
    entity.Set[String](FIELD_YAML_TEXT, yaml_text or "")
    entity.Set[String](FIELD_SOURCE_PATH, source_path or "")
    entity.Set[Int32](FIELD_SCHEMA_VERSION, int(schema_version))
    entity.Set[String](FIELD_LAST_MODIFIED_UTC, last_modified_utc or "")
    pi.SetEntity(entity)


def clear_payload(doc):
    """Delete the stored entity. Caller manages the transaction."""
    schema = get_or_create_schema()
    pi = _project_info(doc)
    pi.DeleteEntity(schema)
