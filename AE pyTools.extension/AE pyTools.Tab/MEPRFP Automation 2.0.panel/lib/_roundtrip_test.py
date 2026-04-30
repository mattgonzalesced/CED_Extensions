# -*- coding: utf-8 -*-
"""Standalone test for yaml_io + schema. Not invoked at runtime."""

import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import yaml_io
import schema as _schema
import schema_migrations as _migrations


def _read(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return f.read()


def test_roundtrip(path):
    print("\n=== {} ===".format(os.path.basename(path)))
    text = _read(path)
    print("input bytes: {}".format(len(text)))

    data = yaml_io.parse(text)
    n_defs = len(data.get("equipment_definitions") or [])
    print("equipment_definitions: {}".format(n_defs))

    version = _schema.validate_schema_versions(data, allow_empty=False)
    print("schema_version: {}".format(version))

    dumped = yaml_io.dump(data)
    print("dumped bytes: {}".format(len(dumped)))

    data2 = yaml_io.parse(dumped)
    if data == data2:
        print("round-trip data equal: OK")
    else:
        print("round-trip data MISMATCH")
        # surface first divergent top-level key
        keys = set(data) | set(data2)
        for k in keys:
            if data.get(k) != data2.get(k):
                print("  diff at key: {}".format(k))
                break
        return False

    version2 = _schema.validate_schema_versions(data2, allow_empty=False)
    if version != version2:
        print("schema_version mismatch: {} -> {}".format(version, version2))
        return False
    print("schema_version stable: OK")
    return True


def test_blank():
    print("\n=== blank input ===")
    data = yaml_io.parse("")
    if data != {}:
        print("blank parse failed: {!r}".format(data))
        return False
    v = _schema.validate_schema_versions(data, allow_empty=True)
    print("blank default version: {}".format(v))
    return v == _schema.DEFAULT_SCHEMA_VERSION


def test_migration_pipeline(path):
    """End-to-end: parse -> validate input -> migrate -> dump -> reparse -> verify v100."""
    print("\n=== migration pipeline: {} ===".format(os.path.basename(path)))
    text = _read(path)
    data = yaml_io.parse(text)
    input_version = _schema.validate_schema_versions(data, allow_empty=False)
    print("input version: {}".format(input_version))
    if input_version not in (3, 4):
        print("  not v3/v4, skipping migration test")
        return True

    migrated = _migrations.migrate_to_internal(data, source_version=input_version)
    if migrated.get("schema_version") != _schema.INTERNAL_VERSION:
        print("  root not stamped: {}".format(migrated.get("schema_version")))
        return False
    bad = [
        e for e in migrated.get("equipment_definitions") or []
        if isinstance(e, dict) and e.get("schema_version") != _schema.INTERNAL_VERSION
    ]
    if bad:
        print("  {} entries not stamped".format(len(bad)))
        return False
    print("  all {} entries stamped to v{}".format(
        len(migrated["equipment_definitions"]), _schema.INTERNAL_VERSION
    ))

    # Round-trip the migrated data.
    dumped = yaml_io.dump(migrated)
    reparsed = yaml_io.parse(dumped)
    if reparsed != migrated:
        print("  migration round-trip MISMATCH")
        return False
    out_version = _schema.validate_schema_versions(reparsed, allow_empty=False)
    if out_version != _schema.INTERNAL_VERSION:
        print("  reparsed version {} != INTERNAL_VERSION".format(out_version))
        return False
    print("  reparsed at v{}: OK".format(out_version))
    return True


def test_quoted_hash_value():
    """Real data has '#' inside single-quoted values like 'RO# RECEPTACLE'."""
    print("\n=== quoted hash in value ===")
    src = (
        "equipment_definitions:\n"
        "  - id: EQ-001\n"
        "    name: Foo\n"
        "    schema_version: 4\n"
        "    note: 'RO# RECEPTACLE'\n"
    )
    data = yaml_io.parse(src)
    note = data["equipment_definitions"][0]["note"]
    if note != "RO# RECEPTACLE":
        print("value mangled: {!r}".format(note))
        return False
    dumped = yaml_io.dump(data)
    data2 = yaml_io.parse(dumped)
    if data == data2:
        print("quoted-hash round-trip: OK")
        return True
    print("quoted-hash mismatch")
    return False


if __name__ == "__main__":
    samples = []
    # lib -> panel -> AE pyTools.Tab -> AE pyTools.extension -> CED_Extensions
    repo_root = os.path.normpath(os.path.join(HERE, "..", "..", "..", ".."))
    for name in os.listdir(repo_root):
        if name.lower().endswith(".yaml") and "profiles" in name.lower():
            samples.append(os.path.join(repo_root, name))
    samples.sort()

    all_ok = True
    if not test_blank():
        all_ok = False
    if not test_quoted_hash_value():
        all_ok = False
    for s in samples[:2]:
        if not test_roundtrip(s):
            all_ok = False
        if not test_migration_pipeline(s):
            all_ok = False

    print("\n=== {} ===".format("PASS" if all_ok else "FAIL"))
    sys.exit(0 if all_ok else 1)
