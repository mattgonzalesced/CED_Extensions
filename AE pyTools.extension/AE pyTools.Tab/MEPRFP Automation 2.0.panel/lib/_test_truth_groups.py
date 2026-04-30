# -*- coding: utf-8 -*-
"""Tests for truth_groups.py. Pure-Python, runs outside Revit."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import truth_groups as tg


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def _make_profile(profile_id, name=None, source_id=None, source_name=None):
    p = {"id": profile_id, "name": name or profile_id}
    if source_id:
        p["ced_truth_source_id"] = source_id
    if source_name:
        p["ced_truth_source_name"] = source_name
    return p


def test_basic_accessors():
    print("\n[truth_groups] basic accessors")
    src = _make_profile("EQ-001")
    member = _make_profile("EQ-002", source_id="EQ-001", source_name="EQ-001")
    _check("source has no truth source", tg.truth_source_id(src) is None)
    _check("source not a member", tg.is_group_member(src) is False)
    _check("member has source id", tg.truth_source_id(member) == "EQ-001")
    _check("member is_group_member", tg.is_group_member(member) is True)
    _check("non-dict guarded", tg.truth_source_id(None) is None)


def test_set_clear_truth_source():
    print("\n[truth_groups] set/clear")
    p = _make_profile("EQ-002")
    tg.set_truth_source(p, "EQ-001", "Source One")
    _check("set source id", p["ced_truth_source_id"] == "EQ-001")
    _check("set source name", p["ced_truth_source_name"] == "Source One")
    tg.clear_truth_source(p)
    _check("cleared id", "ced_truth_source_id" not in p)
    _check("cleared name", "ced_truth_source_name" not in p)


def test_set_validates():
    print("\n[truth_groups] set validation")
    try:
        tg.set_truth_source(None, "EQ-001", "x")
    except TypeError:
        _check("non-dict raises TypeError", True)
    else:
        _check("non-dict raises TypeError", False)
    try:
        tg.set_truth_source({}, "", "x")
    except ValueError:
        _check("empty source raises ValueError", True)
    else:
        _check("empty source raises ValueError", False)


def test_find_group():
    print("\n[truth_groups] find source / members")
    profiles = [
        _make_profile("EQ-001", "Source"),
        _make_profile("EQ-002", source_id="EQ-001"),
        _make_profile("EQ-003", source_id="EQ-001"),
        _make_profile("EQ-004", source_id="EQ-099"),  # different group
        _make_profile("EQ-005"),                       # no group
    ]
    src = tg.find_group_source(profiles, "EQ-001")
    _check("find_group_source returns source", src is profiles[0])
    members = tg.find_group_members(profiles, "EQ-001")
    _check("two members for EQ-001", len(members) == 2)
    _check("members are EQ-002 / EQ-003",
           {m["id"] for m in members} == {"EQ-002", "EQ-003"})

    members_99 = tg.find_group_members(profiles, "EQ-099")
    _check("one member for EQ-099", len(members_99) == 1)

    members_none = tg.find_group_members(profiles, "EQ-NONE")
    _check("zero members for unknown group", members_none == [])


def test_group_members_by_source():
    print("\n[truth_groups] group_members_by_source")
    profiles = [
        _make_profile("EQ-001", "Source"),
        _make_profile("EQ-002", source_id="EQ-001"),
        _make_profile("EQ-003", source_id="EQ-001"),
        _make_profile("EQ-004", source_id="EQ-099"),
        _make_profile("EQ-005"),
        # An entry where the truth_source_id == its own id (the source itself)
        # — should NOT be listed as a member of its own group.
        _make_profile("EQ-006", source_id="EQ-006"),
    ]
    by_source = tg.group_members_by_source(profiles)
    _check("two sources with members", set(by_source.keys()) == {"EQ-001", "EQ-099"})
    _check("EQ-001 has 2 members", len(by_source["EQ-001"]) == 2)
    _check("self-reference excluded", "EQ-006" not in by_source)


def run():
    test_basic_accessors()
    test_set_clear_truth_source()
    test_set_validates()
    test_find_group()
    test_group_members_by_source()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[truth_groups] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
