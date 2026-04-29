# -*- coding: utf-8 -*-
"""Offline tests for merge_workflow.py — pure-logic eligibility +
renumbering. Doesn't touch the Revit API."""

from __future__ import print_function

import copy
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


# ---------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------

def _make_doc():
    return {
        "schema_version": 100,
        "equipment_definitions": [
            {
                "id": "EQ-001",
                "name": "Foo : Master",
                "schema_version": 100,
                "parent_filter": {
                    "category": "Mechanical Equipment",
                    "family_name_pattern": "Foo",
                    "type_name_pattern": "Master",
                    "parameter_filters": {},
                },
                "linked_sets": [
                    {
                        "id": "SET-001",
                        "name": "Foo : Master Types",
                        "linked_element_definitions": [
                            {
                                "id": "SET-001-LED-001",
                                "label": "X : Y",
                                "is_group": False,
                                "parameters": {},
                                "offsets": [{"x_inches": 1.0, "y_inches": 0.0,
                                             "z_inches": 0.0, "rotation_deg": 0.0}],
                                "annotations": [
                                    {
                                        "id": "SET-001-LED-001-ANN-001",
                                        "kind": "tag",
                                        "label": "TagA",
                                        "parameters": {},
                                        "offsets": {"x_inches": 0.0, "y_inches": 12.0,
                                                    "z_inches": 0.0, "rotation_deg": 0.0},
                                    },
                                    {
                                        "id": "SET-001-LED-001-ANN-002",
                                        "kind": "keynote",
                                        "label": "Keynote",
                                        "parameters": {},
                                        "offsets": {"x_inches": 0.0, "y_inches": 0.0,
                                                    "z_inches": 0.0, "rotation_deg": 0.0},
                                    },
                                ],
                            },
                            {
                                "id": "SET-001-LED-002",
                                "label": "X : Z",
                                "is_group": False,
                                "parameters": {},
                                "offsets": [{"x_inches": 5.0, "y_inches": 0.0,
                                             "z_inches": 0.0, "rotation_deg": 0.0}],
                                "annotations": [],
                            },
                        ],
                    },
                ],
                "equipment_properties": {"flag": "1"},
                "allow_parentless": False,
                "allow_unmatched_parents": True,
                "prompt_on_parent_mismatch": False,
            },
            {
                "id": "EQ-002",
                "name": "Foo : Variant1",
                "schema_version": 100,
                "parent_filter": {"category": "", "family_name_pattern": "",
                                  "type_name_pattern": "", "parameter_filters": {}},
                "linked_sets": [
                    {
                        "id": "SET-002",
                        "name": "Foo : Variant1 Types",
                        "linked_element_definitions": [
                            {"id": "SET-002-LED-001", "label": "old",
                             "is_group": False, "parameters": {},
                             "offsets": [{}], "annotations": []},
                        ],
                    },
                ],
                "equipment_properties": {},
                "allow_parentless": False,
                "allow_unmatched_parents": True,
                "prompt_on_parent_mismatch": False,
            },
            {
                "id": "EQ-003",
                "name": "Foo : Variant2",
                "schema_version": 100,
                "parent_filter": {"category": "", "family_name_pattern": "",
                                  "type_name_pattern": "", "parameter_filters": {}},
                "linked_sets": [],
                "equipment_properties": {},
                "allow_parentless": False,
                "allow_unmatched_parents": True,
                "prompt_on_parent_mismatch": False,
            },
        ],
    }


def _profile(doc, pid):
    for p in doc["equipment_definitions"]:
        if p.get("id") == pid:
            return p
    return None


# ---------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------

def test_eligibility_no_existing_groups():
    print("\n[merge] eligibility (clean state)")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    targets = merge_workflow.eligible_targets(doc, src)
    target_ids = {t["id"] for t in targets}
    _check("EQ-002 + EQ-003 eligible", target_ids == {"EQ-002", "EQ-003"})

    sources = merge_workflow.eligible_sources(doc)
    _check("All three eligible as sources",
           {p["id"] for p in sources} == {"EQ-001", "EQ-002", "EQ-003"})


def test_self_merge_forbidden():
    print("\n[merge] self-merge forbidden")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    ok, reason = merge_workflow.can_be_target(doc, src, src)
    _check("Cannot merge into self", ok is False)


def test_cannot_target_existing_member():
    print("\n[merge] target already merged")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    target = _profile(doc, "EQ-002")
    merge_workflow.merge_into(doc, src, target)
    # Now try to target EQ-002 from a different source.
    other = _profile(doc, "EQ-003")
    ok, reason = merge_workflow.can_be_target(doc, other, target)
    _check("Already-member target rejected", ok is False)


def test_cannot_target_existing_source():
    print("\n[merge] target is itself a source for others")
    doc = _make_doc()
    # EQ-001 -> EQ-002 makes EQ-001 a source.
    merge_workflow.merge_into(doc, _profile(doc, "EQ-001"), _profile(doc, "EQ-002"))
    # Try to merge EQ-001 into EQ-003. EQ-001 should be eligible as source
    # but should NOT be eligible as a target if someone tried.
    src = _profile(doc, "EQ-003")
    cand = _profile(doc, "EQ-001")
    ok, reason = merge_workflow.can_be_target(doc, src, cand)
    _check("Source-for-others rejected as target", ok is False)


def test_cannot_source_member():
    print("\n[merge] member cannot be a source")
    doc = _make_doc()
    merge_workflow.merge_into(doc, _profile(doc, "EQ-001"), _profile(doc, "EQ-002"))
    member = _profile(doc, "EQ-002")
    ok, reason = merge_workflow.can_be_source(doc, member)
    _check("Member rejected as source", ok is False)


def test_renumber_isolation():
    print("\n[merge] renumbered ids don't collide with anything in the doc")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    target = _profile(doc, "EQ-002")
    merge_workflow.merge_into(doc, src, target)

    # Collect all SET ids across the doc post-merge.
    all_set_ids = []
    for p in doc["equipment_definitions"]:
        for s in p.get("linked_sets") or []:
            all_set_ids.append(s["id"])
    _check("All SET ids unique post-merge",
           len(set(all_set_ids)) == len(all_set_ids),
           "got {}".format(all_set_ids))


def test_renumber_internal_consistency():
    """LED ids must reference the parent SET id; ANN ids must reference
    the parent LED id."""
    print("\n[merge] LED + ANN ids stay nested under their parent")
    doc = _make_doc()
    merge_workflow.merge_into(doc, _profile(doc, "EQ-001"), _profile(doc, "EQ-002"))
    target = _profile(doc, "EQ-002")
    for s in target.get("linked_sets") or []:
        sid = s["id"]
        for led in s.get("linked_element_definitions") or []:
            lid = led["id"]
            _check("LED {} starts with set id".format(lid),
                   lid.startswith(sid + "-LED-"))
            for ann in led.get("annotations") or []:
                aid = ann["id"]
                _check("ANN {} starts with led id".format(aid),
                       aid.startswith(lid + "-ANN-"))


def test_target_keeps_id_and_name():
    print("\n[merge] target retains its own id + name")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    target = _profile(doc, "EQ-002")
    target_id = target["id"]
    target_name = target["name"]
    merge_workflow.merge_into(doc, src, target)
    _check("Target id unchanged", target["id"] == target_id)
    _check("Target name unchanged", target["name"] == target_name)


def test_target_marked_as_member():
    print("\n[merge] target tagged with truth source")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    target = _profile(doc, "EQ-002")
    merge_workflow.merge_into(doc, src, target)
    _check("ced_truth_source_id set",
           truth_groups.truth_source_id(target) == "EQ-001")
    _check("ced_truth_source_name set",
           truth_groups.truth_source_name(target) == "Foo : Master")
    _check("is_group_member True", truth_groups.is_group_member(target))


def test_unmerge_clears_lineage():
    print("\n[merge] unmerge clears truth source")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    target = _profile(doc, "EQ-002")
    merge_workflow.merge_into(doc, src, target)

    merge_workflow.unmerge(doc, target)
    _check("truth_source_id cleared",
           truth_groups.truth_source_id(target) is None)
    _check("truth_source_name cleared",
           truth_groups.truth_source_name(target) is None)
    # Structural content remains.
    _check("linked_sets still present",
           len(target.get("linked_sets") or []) > 0)


def test_merge_many():
    print("\n[merge] merge_many returns succeeded + failed lists")
    doc = _make_doc()
    src = _profile(doc, "EQ-001")
    targets = [_profile(doc, "EQ-002"), _profile(doc, "EQ-003")]
    succeeded, failed = merge_workflow.merge_many(doc, src, targets)
    _check("Both succeed", len(succeeded) == 2 and len(failed) == 0)


def test_renumber_uses_global_counter():
    """The renumber base should account for ALL existing SET ids in the
    doc, not just the source's. So merging EQ-001 (with SET-001) into
    EQ-003 — when EQ-002 already has SET-002 — should produce SET-003,
    not SET-002 (which is taken by EQ-002's set)."""
    print("\n[merge] renumber respects all existing SET ids")
    doc = _make_doc()
    merge_workflow.merge_into(doc, _profile(doc, "EQ-001"), _profile(doc, "EQ-003"))
    # EQ-003 should have a new SET id that's >= SET-003
    target = _profile(doc, "EQ-003")
    sets = target.get("linked_sets") or []
    sid = sets[0]["id"]
    n = int(sid.split("-")[-1])
    _check("renumber > 2 (SET-001 in source, SET-002 in EQ-002)",
           n >= 3, "got {}".format(sid))


def run():
    test_eligibility_no_existing_groups()
    test_self_merge_forbidden()
    test_cannot_target_existing_member()
    test_cannot_target_existing_source()
    test_cannot_source_member()
    test_renumber_isolation()
    test_renumber_internal_consistency()
    test_target_keeps_id_and_name()
    test_target_marked_as_member()
    test_unmerge_clears_lineage()
    test_merge_many()
    test_renumber_uses_global_counter()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[merge] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
