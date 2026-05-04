# -*- coding: utf-8 -*-
"""
High-level orchestration for Import / Export workflows.

These functions are pure logic — file I/O, YAML parsing, schema
validation, migration, and Extensible Storage writes. They do no UI.
Pushbutton scripts import them and handle the user prompts.

Import flow:  read file -> parse YAML -> validate input version
              -> migrate to internal v100 -> dump canonical text -> store

Export flow:  read store -> write file (verbatim text already in v100)
"""

import io
import datetime

from pyrevit import revit

import schema as _schema
import schema_migrations as _migrations
import yaml_io
import storage
import space_storage


def _utc_now_iso():
    return datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")


def _read_text_file(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return f.read()


def _write_text_file(path, text):
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(text)


def has_active_payload(doc):
    return storage.read_payload(doc) is not None


def get_active_payload(doc):
    return storage.read_payload(doc)


def load_active_data(doc):
    """Parse the stored YAML text and return the resulting dict.

    Returns ``{}`` if no payload exists or the stored text is empty.
    """
    payload = storage.read_payload(doc)
    if payload is None:
        return {}
    text = payload.get("yaml_text") or ""
    if not text.strip():
        return {}
    return yaml_io.parse(text)


def save_active_data(doc, data, action="MEPRFP 2.0 edit"):
    """Re-serialise ``data`` to YAML and persist to Extensible Storage.

    Caller manages the Revit transaction.
    """
    canonical = yaml_io.dump(data)
    payload = storage.read_payload(doc) or {}
    storage.write_payload(
        doc=doc,
        yaml_text=canonical,
        source_path=payload.get("source_path") or "",
        schema_version=_schema.INTERNAL_VERSION,
        last_modified_utc=_utc_now_iso(),
    )


def import_yaml_file(doc, source_path):
    """Read ``source_path``, validate, migrate to v100, and persist.

    Returns a dict with summary fields::

        {
            "source_path": str,
            "input_schema_version": int,    # what the file declared
            "stored_schema_version": int,   # always INTERNAL_VERSION
            "byte_count": int,
            "blank": bool,
        }

    Raises:
        IOError / OSError                          on read failure
        yaml_io.YamlError                          on parse failure
        schema.SchemaVersionError                  on schema-version check
        schema_migrations.MigrationError           on migration failure
    """
    raw_text = _read_text_file(source_path)
    is_blank = not raw_text.strip()

    if is_blank:
        # Synthesise a minimal v100 document so a blank import still
        # produces a stamped, parseable export.
        data = {
            "schema_version": _schema.INTERNAL_VERSION,
            "equipment_definitions": [],
        }
        canonical = yaml_io.dump(data)
        input_version = _schema.INTERNAL_VERSION
        stored_version = _schema.INTERNAL_VERSION
    else:
        data = yaml_io.parse(raw_text)
        input_version = _schema.validate_schema_versions(data, allow_empty=False)
        data = _migrations.migrate_to_internal(data, source_version=input_version)
        canonical = yaml_io.dump(data)
        stored_version = _schema.INTERNAL_VERSION

    with revit.Transaction("Import YAML File (MEPRFP 2.0)", doc=doc):
        storage.write_payload(
            doc=doc,
            yaml_text=canonical,
            source_path=source_path,
            schema_version=stored_version,
            last_modified_utc=_utc_now_iso(),
        )

    return {
        "source_path": source_path,
        "input_schema_version": input_version,
        "stored_schema_version": stored_version,
        "byte_count": len(canonical),
        "blank": is_blank,
    }


# ---------------------------------------------------------------------
# Spaces (Stage 6)
#
# Templates (``space_buckets`` and ``space_profiles``) live in the same
# YAML payload as ``equipment_definitions``, so they round-trip through
# the existing import/export. Per-project ``classifications`` live in a
# separate Extensible Storage entity managed by ``space_storage`` so an
# export of the YAML doesn't drag one project's space assignments into
# another.
# ---------------------------------------------------------------------

def load_space_buckets(doc):
    """Return ``data['space_buckets']`` as a list of dicts (never None)."""
    data = load_active_data(doc)
    raw = data.get("space_buckets") if isinstance(data, dict) else None
    return list(raw or [])


def save_space_buckets(doc, buckets, action="MEPRFP 2.0 edit space buckets"):
    """Persist ``space_buckets`` into the active YAML payload.

    Caller manages the Revit transaction.
    """
    data = load_active_data(doc) or {}
    data["space_buckets"] = list(buckets or [])
    data.setdefault("schema_version", _schema.INTERNAL_VERSION)
    save_active_data(doc, data, action=action)


def load_space_profiles(doc):
    """Return ``data['space_profiles']`` as a list of dicts (never None)."""
    data = load_active_data(doc)
    raw = data.get("space_profiles") if isinstance(data, dict) else None
    return list(raw or [])


def save_space_profiles(doc, profiles, action="MEPRFP 2.0 edit space profiles"):
    """Persist ``space_profiles`` into the active YAML payload.

    Caller manages the Revit transaction.
    """
    data = load_active_data(doc) or {}
    data["space_profiles"] = list(profiles or [])
    data.setdefault("schema_version", _schema.INTERNAL_VERSION)
    save_active_data(doc, data, action=action)


def load_classifications(doc):
    """Return per-project space classifications as a list of dicts."""
    payload = space_storage.read_payload(doc)
    if payload is None:
        return []
    return space_storage.decode(payload.get("json_text") or "")


def save_classifications(doc, classifications):
    """Persist per-project classifications. Caller manages txn."""
    text = space_storage.encode(classifications)
    space_storage.write_payload(
        doc=doc,
        json_text=text,
        last_modified_utc=_utc_now_iso(),
    )


def export_yaml_file(doc, save_path):
    """Read the active payload from storage and write it to ``save_path``.

    Returns a dict with summary fields::

        {
            "save_path": str,
            "byte_count": int,
            "schema_version": int,
            "source_path": str,
            "last_modified_utc": str,
        }

    Raises:
        StorageError      if no active payload exists
        IOError / OSError on write failure
    """
    payload = storage.read_payload(doc)
    if payload is None:
        raise storage.StorageError(
            "No active YAML in this project. Import a file first."
        )

    text = payload["yaml_text"] or ""
    _write_text_file(save_path, text)

    return {
        "save_path": save_path,
        "byte_count": len(text),
        "schema_version": payload.get("schema_version"),
        "source_path": payload.get("source_path"),
        "last_modified_utc": payload.get("last_modified_utc"),
    }
