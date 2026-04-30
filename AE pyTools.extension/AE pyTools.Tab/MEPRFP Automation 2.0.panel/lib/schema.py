# -*- coding: utf-8 -*-
"""
Schema-version validation for equipment-definition YAML payloads.

The ``schema_version`` field appears at the document root and on each
``equipment_definitions[*]`` entry. Both must agree, and the value must
be in ``SUPPORTED_INPUT_VERSIONS``.

MEPRFP 2.0's *internal* version is ``INTERNAL_VERSION`` (100). Legacy
v3/v4 imports are migrated forward by ``schema_migrations.migrate_to_internal``
before storage. Re-imports of an already-exported v100 file are a no-op.
"""

import schema_migrations as _migrations

# Re-export so callers don't have to know two module names.
INTERNAL_VERSION = _migrations.INTERNAL_VERSION
SUPPORTED_INPUT_VERSIONS = _migrations.SUPPORTED_INPUT_VERSIONS
DEFAULT_SCHEMA_VERSION = INTERNAL_VERSION


class SchemaVersionError(ValueError):
    pass


def collect_schema_versions(data):
    """Return all schema_version values found in a parsed YAML dict."""
    versions = []
    if not isinstance(data, dict):
        return versions
    if "schema_version" in data:
        versions.append(data["schema_version"])
    defs = data.get("equipment_definitions") or []
    if isinstance(defs, list):
        for entry in defs:
            if isinstance(entry, dict) and "schema_version" in entry:
                versions.append(entry["schema_version"])
    return versions


def normalize_version(value):
    if value is None:
        return None
    try:
        return int(str(value).strip())
    except (ValueError, AttributeError):
        return None


def validate_schema_versions(data, allow_empty=True):
    """
    Validate every schema_version marker in the payload.

    Returns the single distinct version (int) if the payload is consistent.
    Raises SchemaVersionError otherwise.

    If ``allow_empty`` is True and the payload has no markers (e.g. blank
    file or an empty equipment_definitions list), returns
    ``DEFAULT_SCHEMA_VERSION``.
    """
    raw = collect_schema_versions(data)
    if not raw:
        if allow_empty:
            return DEFAULT_SCHEMA_VERSION
        raise SchemaVersionError("YAML has no schema_version markers")

    invalid = [v for v in raw if normalize_version(v) is None]
    if invalid:
        raise SchemaVersionError(
            "YAML has invalid schema_version values: {}".format(invalid)
        )

    distinct = sorted({normalize_version(v) for v in raw})
    unsupported = [v for v in distinct if v not in SUPPORTED_INPUT_VERSIONS]
    if unsupported:
        raise SchemaVersionError(
            "Unsupported schema_version(s): {} (supported: {})".format(
                unsupported, list(SUPPORTED_INPUT_VERSIONS)
            )
        )
    if len(distinct) > 1:
        raise SchemaVersionError(
            "Inconsistent schema_versions across payload: {}".format(distinct)
        )
    return distinct[0]
