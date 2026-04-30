# -*- coding: utf-8 -*-
"""Offline tests for the alias-based merge_workflow.py."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import merge_workflow
import truth_groups


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def _make_doc():
    return {
        "schema_version": 100,
        "equipment_definitions": [
            {"id": "EQ-001", "name": "Foo : Master", "schema_version": 100,
             "linked_sets": [{"id": "SET-001"}]},
            {"id": "EQ-002", "name": "Foo : V1", "schema_version": 100,
             "linked_sets": [{"id": "SET-002"}]},
            {"id": "EQ-003", "name": "Foo : V2", "schema_version": 100,
             "linked_sets": [{"id": "SET-003"}]},
        ],
    }


def _profile(doc, pid):
    for p in doc["equipment_definitions"]:
        if p.get("id") == pid:
            return p
    return None


# ---------------------------------------------------------------------
# Alias add / remove
# ---------------------------------------------------------------------

def test_add_alias_basic():
    print("\n[merge] add_alias basic")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    _check("first add returns True", merge_workflow.add_alias(src, "Foo : V1") is True)
    _check("alias in list", "Foo : V1" in merge_workflow.aliases(src))


def test_add_alias_dedup():
    print("\n[merge] add_alias dedup (case-insensitive)")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    merge_workflow.add_alias(src, "Foo : V1")
    _check("same case skipped", merge_workflow.add_alias(src, "Foo : V1") is False)
    _check("different case skipped", merge_workflow.add_alias(src, "FOO : v1") is False)
    _check("trimmed-whitespace skipped",
           merge_workflow.add_alias(src, "  Foo : V1  ") is False)
    _check("only one entry stored", len(merge_workflow.aliases(src)) == 1)


def test_add_alias_empty_skipped():
    print("\n[merge] empty alias skipped")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    _check("empty string -> False", merge_workflow.add_alias(src, "") is False)
    _check("whitespace -> False", merge_workflow.add_alias(src, "   ") is False)
    _check("None -> False", merge_workflow.add_alias(src, None) is False)
    _check("no entries", merge_workflow.aliases(src) == [])


def test_remove_alias():
    print("\n[merge] remove_alias")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    merge_workflow.add_alias(src, "Foo : V1")
    merge_workflow.add_alias(src, "Foo : V2")
    _check("remove existing", merge_workflow.remove_alias(src, "Foo : V1") is True)
    _check("remove case-insensitive",
           merge_workflow.remove_alias(src, "foo : v2") is True)
    _check("remove missing -> False",
           merge_workflow.remove_alias(src, "nope") is False)
    _check("list empty", merge_workflow.aliases(src) == [])


def test_add_aliases_bulk():
    print("\n[merge] add_aliases (bulk)")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    added, skipped = merge_workflow.add_aliases(
        src, ["A", "B", "a", "C", "", "B"]
    )
    _check("3 added (A, B, C)", added == 3)
    _check("3 skipped (a dup, empty, B dup)", skipped == 3)


def test_find_alias_owner():
    print("\n[merge] find_alias_owner")
    doc = _make_doc()
    merge_workflow.add_alias(_profile(doc, "EQ-001"), "OnlyOnEQ001")
    merge_workflow.add_alias(_profile(doc, "EQ-002"), "OnlyOnEQ002")

    owner = merge_workflow.find_alias_owner(doc, "OnlyOnEQ002")
    _check("found owner of OnlyOnEQ002", owner is not None and owner.get("id") == "EQ-002")
    _check("missing alias -> None",
           merge_workflow.find_alias_owner(doc, "Nope") is None)


def test_all_alias_entries():
    print("\n[merge] all_alias_entries")
    doc = _make_doc()
    merge_workflow.add_alias(_profile(doc, "EQ-001"), "X")
    merge_workflow.add_alias(_profile(doc, "EQ-001"), "Y")
    merge_workflow.add_alias(_profile(doc, "EQ-002"), "Z")
    entries = merge_workflow.all_alias_entries(doc)
    _check("three entries total", len(entries) == 3)
    aliases_only = [a for _src, a in entries]
    _check("contents", set(aliases_only) == {"X", "Y", "Z"})


# ---------------------------------------------------------------------
# Legacy migration
# ---------------------------------------------------------------------

def _make_legacy_doc():
    """Doc shaped like the old data-duplication model."""
    return {
        "schema_version": 100,
        "equipment_definitions": [
            {
                "id": "EQ-001",
                "name": "Foo : Master",
                "schema_version": 100,
                "linked_sets": [{"id": "SET-001"}],
            },
            {
                "id": "EQ-002",
                "name": "Foo : V1",
                "schema_version": 100,
                "linked_sets": [{"id": "SET-002"}],
                "ced_truth_source_id": "EQ-001",
                "ced_truth_source_name": "Foo : Master",
            },
            {
                "id": "EQ-003",
                "name": "Foo : V2",
                "schema_version": 100,
                "linked_sets": [{"id": "SET-003"}],
                "ced_truth_source_id": "EQ-001",
                "ced_truth_source_name": "Foo : Master",
            },
        ],
    }


def test_has_legacy_members():
    print("\n[merge] has_legacy_members")
    doc = _make_doc()
    _check("clean doc -> False", merge_workflow.has_legacy_members(doc) is False)
    legacy = _make_legacy_doc()
    _check("legacy doc -> True", merge_workflow.has_legacy_members(legacy) is True)


def test_migrate_legacy_members():
    print("\n[merge] migrate_legacy_members")
    doc = _make_legacy_doc()
    report = merge_workflow.migrate_legacy_members(doc)
    src = _profile(doc, "EQ-001")
    aliases_on_src = merge_workflow.aliases(src)
    _check("two aliases added to source",
           set(aliases_on_src) == {"Foo : V1", "Foo : V2"},
           "got {}".format(aliases_on_src))
    _check("members had ced_truth_source cleared",
           not truth_groups.is_group_member(_profile(doc, "EQ-002")))
    _check("report.aliases_added == 2", report.aliases_added == 2)
    _check("report.members_cleared == 2", report.members_cleared == 2)
    _check("no unresolved", report.unresolved_members == [])


def test_migrate_unresolved_source():
    """A member whose ced_truth_source_id points at a missing source
    should still get its tag cleared but report unresolved."""
    print("\n[merge] migrate with unresolved source")
    doc = {
        "schema_version": 100,
        "equipment_definitions": [
            {"id": "EQ-002", "name": "Member only", "schema_version": 100,
             "ced_truth_source_id": "EQ-DOES-NOT-EXIST",
             "ced_truth_source_name": "Ghost"},
        ],
    }
    report = merge_workflow.migrate_legacy_members(doc)
    _check("0 aliases added", report.aliases_added == 0)
    _check("1 member cleared", report.members_cleared == 1)
    _check("1 unresolved", len(report.unresolved_members) == 1)


def test_delete_profiles_by_id():
    print("\n[merge] delete_profiles_by_id")
    doc = _make_doc()
    removed = merge_workflow.delete_profiles_by_id(doc, ["EQ-002", "EQ-XX"])
    _check("1 removed", removed == 1)
    ids = [p.get("id") for p in doc["equipment_definitions"]]
    _check("EQ-002 gone", "EQ-002" not in ids)
    _check("EQ-001 + EQ-003 still there", "EQ-001" in ids and "EQ-003" in ids)


# ---------------------------------------------------------------------
# Bulk CSV
# ---------------------------------------------------------------------

def test_bulk_add_aliases_from_csv(tmp_path_dir):
    print("\n[merge] bulk_add_aliases_from_csv")
    csv = os.path.join(tmp_path_dir, "bulk.csv")
    with open(csv, "w", encoding="utf-8") as f:
        f.write("source,target\n")
        f.write("EQ-001,Foo : V1\n")
        f.write("EQ-001,Foo : V2\n")
        f.write("Foo : Master,DupTest\n")    # source-by-name
        f.write("Foo : Master,DupTest\n")    # already-added dup
        f.write("Missing : Source,X\n")     # not found
        f.write(",empty source\n")          # skipped silently
    doc = _make_doc()
    results = merge_workflow.bulk_add_aliases_from_csv(doc, csv)
    src = _profile(doc, "EQ-001")
    aliases_on_src = merge_workflow.aliases(src)
    _check("3 distinct aliases on EQ-001",
           set(aliases_on_src) == {"Foo : V1", "Foo : V2", "DupTest"},
           "got {}".format(aliases_on_src))
    ok_count = sum(1 for r in results if r.ok)
    fail_count = sum(1 for r in results if not r.ok)
    _check("3 ok rows", ok_count == 3)
    _check("2 fail/skip rows", fail_count == 2)


def test_bulk_csv_missing_columns(tmp_path_dir):
    print("\n[merge] bulk csv missing required columns")
    csv = os.path.join(tmp_path_dir, "bad.csv")
    with open(csv, "w", encoding="utf-8") as f:
        f.write("col1,col2\n")
        f.write("a,b\n")
    doc = _make_doc()
    raised = False
    try:
        merge_workflow.bulk_add_aliases_from_csv(doc, csv)
    except merge_workflow.MergeError:
        raised = True
    _check("MergeError on missing columns", raised)


# ---------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------

def run():
    test_add_alias_basic()
    test_add_alias_dedup()
    test_add_alias_empty_skipped()
    test_remove_alias()
    test_add_aliases_bulk()
    test_find_alias_owner()
    test_all_alias_entries()
    test_has_legacy_members()
    test_migrate_legacy_members()
    test_migrate_unresolved_source()
    test_delete_profiles_by_id()

    import tempfile
    tmpdir = tempfile.mkdtemp(prefix="merge_test_")
    try:
        test_bulk_add_aliases_from_csv(tmpdir)
        test_bulk_csv_missing_columns(tmpdir)
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[merge] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
