# -*- coding: utf-8 -*-
"""Tests for schema_migrations.py and the schema validator."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import schema as _schema
import schema_migrations as _migrations


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def _doc(version, n=2):
    return {
        "schema_version": version,
        "equipment_definitions": [
            {"id": "EQ-{:03d}".format(i), "schema_version": version}
            for i in range(1, n + 1)
        ],
    }


def test_validate_valid():
    print("\n[migrations] validate_schema_versions valid")
    _check("v3 ok", _schema.validate_schema_versions(_doc(3)) == 3)
    _check("v4 ok", _schema.validate_schema_versions(_doc(4)) == 4)
    _check("v100 ok", _schema.validate_schema_versions(_doc(100)) == 100)


def test_validate_blank_default():
    print("\n[migrations] validate blank")
    _check("blank dict -> default v100",
           _schema.validate_schema_versions({}) == _schema.INTERNAL_VERSION)
    _check("blank dict no allow -> raises",
           _raises(lambda: _schema.validate_schema_versions({}, allow_empty=False),
                   _schema.SchemaVersionError))


def test_validate_inconsistent():
    print("\n[migrations] validate inconsistent")
    bad = _doc(3)
    bad["equipment_definitions"][1]["schema_version"] = 4
    _check("mixed versions raise",
           _raises(lambda: _schema.validate_schema_versions(bad),
                   _schema.SchemaVersionError))


def test_validate_unsupported():
    print("\n[migrations] validate unsupported")
    _check("v2 raises",
           _raises(lambda: _schema.validate_schema_versions(_doc(2)),
                   _schema.SchemaVersionError))
    _check("v999 raises",
           _raises(lambda: _schema.validate_schema_versions(_doc(999)),
                   _schema.SchemaVersionError))


def test_validate_invalid_value():
    print("\n[migrations] validate invalid value")
    bad = {"schema_version": "not-a-number",
           "equipment_definitions": [{"id": "EQ-1", "schema_version": "abc"}]}
    _check("string version raises",
           _raises(lambda: _schema.validate_schema_versions(bad),
                   _schema.SchemaVersionError))


def test_migrate_v3_to_internal():
    print("\n[migrations] migrate v3 -> v100")
    data = _doc(3, n=3)
    out = _migrations.migrate_to_internal(data, source_version=3)
    _check("root stamped 100", out["schema_version"] == 100)
    _check("all entries stamped 100",
           all(e["schema_version"] == 100 for e in out["equipment_definitions"]))


def test_migrate_v4_to_internal():
    print("\n[migrations] migrate v4 -> v100")
    data = _doc(4, n=3)
    out = _migrations.migrate_to_internal(data, source_version=4)
    _check("root stamped 100", out["schema_version"] == 100)
    _check("all entries stamped 100",
           all(e["schema_version"] == 100 for e in out["equipment_definitions"]))


def test_migrate_v100_idempotent():
    print("\n[migrations] migrate v100 (idempotent)")
    data = _doc(100, n=3)
    out = _migrations.migrate_to_internal(data, source_version=100)
    _check("still v100", out["schema_version"] == 100)
    _check("entries still v100",
           all(e["schema_version"] == 100 for e in out["equipment_definitions"]))


def test_migrate_detect():
    print("\n[migrations] detect_version + auto-migrate")
    data = _doc(3, n=2)
    out = _migrations.migrate_to_internal(data)  # no source_version
    _check("auto-detected and migrated", out["schema_version"] == 100)


def test_migrate_no_version():
    print("\n[migrations] migrate without version markers")
    _check("missing version raises",
           _raises(lambda: _migrations.migrate_to_internal({}),
                   _migrations.MigrationError))


def test_migrate_unsupported():
    print("\n[migrations] migrate unsupported version")
    _check("v2 raises",
           _raises(lambda: _migrations.migrate_to_internal(_doc(2), source_version=2),
                   _migrations.MigrationError))


def test_migrate_copy_does_not_mutate():
    print("\n[migrations] migrate_to_internal_copy does not mutate input")
    src = _doc(3, n=2)
    out = _migrations.migrate_to_internal_copy(src, source_version=3)
    _check("source unchanged", src["schema_version"] == 3)
    _check("copy migrated", out["schema_version"] == 100)


def test_migrate_partial_markers():
    """If some entries have schema_version and others don't, migration
    should stamp every entry uniformly."""
    print("\n[migrations] partial markers")
    data = {
        "schema_version": 4,
        "equipment_definitions": [
            {"id": "EQ-1", "schema_version": 4},
            {"id": "EQ-2"},   # missing
            {"id": "EQ-3", "schema_version": 4},
        ],
    }
    out = _migrations.migrate_to_internal(data, source_version=4)
    _check("all entries stamped 100",
           all(e["schema_version"] == 100 for e in out["equipment_definitions"]))


def _raises(callable_, exc_type):
    try:
        callable_()
    except exc_type:
        return True
    except Exception:
        return False
    return False


def run():
    test_validate_valid()
    test_validate_blank_default()
    test_validate_inconsistent()
    test_validate_unsupported()
    test_validate_invalid_value()
    test_migrate_v3_to_internal()
    test_migrate_v4_to_internal()
    test_migrate_v100_idempotent()
    test_migrate_detect()
    test_migrate_no_version()
    test_migrate_unsupported()
    test_migrate_copy_does_not_mutate()
    test_migrate_partial_markers()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[migrations] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
