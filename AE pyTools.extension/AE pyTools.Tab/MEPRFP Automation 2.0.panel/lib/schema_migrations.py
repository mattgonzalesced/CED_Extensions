# -*- coding: utf-8 -*-
"""
Schema-version migrations for equipment-definition YAML payloads.

MEPRFP 2.0 uses internal version 100. Legacy versions 3 and 4 are
accepted on import and converted forward. The shape didn't actually
change between v3 and v4 — they're version markers — so the migration
to v100 is currently a stamp-only operation. Future shape changes will
add real migration steps here.

Each migration is a function ``data -> data`` that mutates the input
dict in place and returns it. The orchestrator picks the right chain
based on the detected source version.
"""

import copy


SUPPORTED_INPUT_VERSIONS = (3, 4, 100)
INTERNAL_VERSION = 100


class MigrationError(Exception):
    pass


def detect_version(data):
    """Return the dominant ``schema_version`` in the payload, or None."""
    if not isinstance(data, dict):
        return None
    versions = []
    if "schema_version" in data:
        versions.append(data["schema_version"])
    defs = data.get("equipment_definitions") or []
    if isinstance(defs, list):
        for entry in defs:
            if isinstance(entry, dict) and "schema_version" in entry:
                versions.append(entry["schema_version"])
    if not versions:
        return None
    distinct = []
    for v in versions:
        try:
            distinct.append(int(str(v).strip()))
        except (ValueError, TypeError):
            continue
    if not distinct:
        return None
    # The validator (in schema.py) already rejects mixed versions, so
    # any value will do as the "detected" one.
    return distinct[0]


def _stamp_version(data, version):
    if not isinstance(data, dict):
        return data
    data["schema_version"] = version
    defs = data.get("equipment_definitions")
    if isinstance(defs, list):
        for entry in defs:
            if isinstance(entry, dict):
                entry["schema_version"] = version
    return data


def _migrate_v3_to_v4(data):
    # No shape changes between v3 and v4 in the legacy data; placeholder
    # for future field renames or restructures.
    return _stamp_version(data, 4)


def _migrate_v4_to_v100(data):
    # No shape changes yet. Future 2.0-specific structural changes go here.
    return _stamp_version(data, 100)


def migrate_to_internal(data, source_version=None):
    """Return ``data`` migrated to ``INTERNAL_VERSION``.

    ``data`` is mutated in place AND returned. If you need a clean copy,
    pass ``copy.deepcopy(data)``.

    ``source_version`` may be supplied to skip detection. If omitted and
    the payload has no version markers, the call raises MigrationError.
    """
    if data is None:
        return data
    if not isinstance(data, dict):
        raise MigrationError(
            "Cannot migrate non-dict payload (got {})".format(type(data).__name__)
        )

    if source_version is None:
        source_version = detect_version(data)

    if source_version is None:
        raise MigrationError(
            "Cannot migrate payload without schema_version markers"
        )
    if source_version not in SUPPORTED_INPUT_VERSIONS:
        raise MigrationError(
            "Unsupported source schema_version {} (supported: {})".format(
                source_version, list(SUPPORTED_INPUT_VERSIONS)
            )
        )

    if source_version == 3:
        data = _migrate_v3_to_v4(data)
        source_version = 4
    if source_version == 4:
        data = _migrate_v4_to_v100(data)
        source_version = 100
    if source_version == 100:
        # Re-stamp defensively in case some entries lacked the marker.
        data = _stamp_version(data, 100)
    return data


def migrate_to_internal_copy(data, source_version=None):
    """Non-mutating variant — returns a deep-copy that's been migrated."""
    if data is None:
        return data
    return migrate_to_internal(copy.deepcopy(data), source_version=source_version)
