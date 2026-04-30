# -*- coding: utf-8 -*-
"""Offline tests for the pure-logic parts of placement.py.

We can't import placement.py directly here because it imports the
Revit API at the top. The matching helpers are duplicated here so the
test stays runnable in plain CPython 3 with no Revit. If placement.py's
matching rules change, mirror the change in this file.
"""

from __future__ import print_function

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


# Mirrors of placement.py functions
_TRAILING_SUFFIX_RE = re.compile(r"_\d+$")


def strip_trailing_suffix(name):
    if not name:
        return ""
    return _TRAILING_SUFFIX_RE.sub("", str(name))


def normalize_name(name):
    return strip_trailing_suffix((name or "").strip()).lower()


def collect_profile_aliases(profile):
    if not isinstance(profile, dict):
        return set()
    props = profile.get("equipment_properties") or {}
    if not isinstance(props, dict):
        return set()
    raw = props.get("cad_aliases")
    if raw is None:
        return set()
    items = []
    if isinstance(raw, list):
        items = [str(x) for x in raw if x is not None]
    elif isinstance(raw, str):
        items = [s for s in raw.split(",")]
    else:
        items = [str(raw)]
    out = set()
    for item in items:
        norm = normalize_name(item)
        if norm:
            out.add(norm)
    return out


def profile_family_names(profile):
    if not isinstance(profile, dict):
        return set()
    out = set()
    pf = profile.get("parent_filter") or {}
    if isinstance(pf, dict):
        fam = pf.get("family_name_pattern")
        if fam:
            out.add(normalize_name(fam))
    name = profile.get("name") or ""
    if " : " in name:
        fam, _ = name.split(" : ", 1)
        if fam:
            out.add(normalize_name(fam))
    return {n for n in out if n}


# ---- tests ---------------------------------------------------------

def test_strip_trailing_suffix():
    print("\n[placement] strip_trailing_suffix")
    _check("no suffix", strip_trailing_suffix("AC_BLOCK") == "AC_BLOCK")
    _check("_1 suffix", strip_trailing_suffix("AC_BLOCK_1") == "AC_BLOCK")
    _check("_42 suffix", strip_trailing_suffix("AC_BLOCK_42") == "AC_BLOCK")
    _check("multi-digit", strip_trailing_suffix("AC_BLOCK_2321321") == "AC_BLOCK")
    _check("_2A leaves it",
           strip_trailing_suffix("AC_BLOCK_2A") == "AC_BLOCK_2A")
    _check("trailing underscore only",
           strip_trailing_suffix("AC_BLOCK_") == "AC_BLOCK_")
    _check("leading underscore", strip_trailing_suffix("_LEAD_5") == "_LEAD")
    _check("digits-only stays (no underscore prefix)",
           strip_trailing_suffix("123") == "123")
    _check("empty stays empty", strip_trailing_suffix("") == "")
    _check("None -> empty", strip_trailing_suffix(None) == "")


def test_normalize_name():
    print("\n[placement] normalize_name")
    _check("lowercase", normalize_name("AC_BLOCK") == "ac_block")
    _check("strip + lowercase",
           normalize_name("AC_BLOCK_5") == "ac_block")
    _check("whitespace trimmed",
           normalize_name("  AC_BLOCK_3 ") == "ac_block")


def test_collect_profile_aliases():
    print("\n[placement] collect_profile_aliases")
    p1 = {"equipment_properties": {"cad_aliases": "BLOCK_A, BLOCK_B"}}
    aliases = collect_profile_aliases(p1)
    _check("comma-separated",
           aliases == {"block_a", "block_b"},
           "got {}".format(aliases))

    p2 = {"equipment_properties": {"cad_aliases": ["X1", "X2", " "]}}
    _check("list form", collect_profile_aliases(p2) == {"x1", "x2"})

    p3 = {"equipment_properties": {"cad_aliases": "AC_BLOCK_42"}}
    _check("alias normalized via strip",
           collect_profile_aliases(p3) == {"ac_block"})

    _check("missing aliases -> empty", collect_profile_aliases({}) == set())
    _check("non-dict profile -> empty",
           collect_profile_aliases(None) == set())


def test_profile_family_names():
    print("\n[placement] profile_family_names")
    p = {
        "name": "ME_Air Curtain_CED : Mars Air PH1284-2E",
        "parent_filter": {"family_name_pattern": "ME_Air Curtain_CED"},
    }
    names = profile_family_names(p)
    _check("collects family from filter + name split",
           "me_air curtain_ced" in names, "got {}".format(names))

    p2 = {"name": "Foo_Bar_5 : Default"}
    _check("name-only family stripped",
           profile_family_names(p2) == {"foo_bar"})


def test_match_linked_revit_logic():
    """A linked element with family 'Foo_Bar_3' should match a profile
    whose family pattern is 'Foo_Bar' (suffix stripped both sides)."""
    print("\n[placement] linked-revit match (suffix strip both sides)")

    profiles = [
        {
            "id": "EQ-001", "name": "Foo_Bar : Default",
            "parent_filter": {"family_name_pattern": "Foo_Bar"},
        },
        {
            "id": "EQ-002", "name": "Other : Default",
            "parent_filter": {"family_name_pattern": "Other"},
        },
    ]
    target_name = "Foo_Bar_3"

    target_key = normalize_name(target_name)
    matched = [p for p in profiles if target_key in profile_family_names(p)]
    _check("Foo_Bar_3 -> Foo_Bar", len(matched) == 1 and matched[0]["id"] == "EQ-001")

    target2 = "Foo_Bar"
    matched2 = [p for p in profiles
                if normalize_name(target2) in profile_family_names(p)]
    _check("Foo_Bar -> Foo_Bar", len(matched2) == 1 and matched2[0]["id"] == "EQ-001")

    target3 = "Different"
    matched3 = [p for p in profiles
                if normalize_name(target3) in profile_family_names(p)]
    _check("non-match yields empty", matched3 == [])


def test_match_cad_alias_logic():
    """A CSV / DWG block whose name matches any profile alias should match."""
    print("\n[placement] CAD alias match (suffix strip both sides)")

    profiles = [
        {
            "id": "EQ-001", "name": "ME_Air Curtain_CED : Mars Air",
            "equipment_properties": {"cad_aliases": "AC_BLOCK, AIR_CURTAIN_BLOCK"},
        },
        {
            "id": "EQ-002", "name": "Other : Default",
            "equipment_properties": {"cad_aliases": "OTHER_BLOCK"},
        },
        {"id": "EQ-003", "name": "No aliases", "equipment_properties": {}},
    ]

    cases = [
        ("AC_BLOCK", "EQ-001"),
        ("ac_block", "EQ-001"),
        ("AC_BLOCK_1", "EQ-001"),
        ("AC_BLOCK_42", "EQ-001"),
        ("AIR_CURTAIN_BLOCK_99", "EQ-001"),
        ("OTHER_BLOCK", "EQ-002"),
        ("UNRELATED", None),
    ]
    for block_name, expect_eq in cases:
        target_key = normalize_name(block_name)
        matched = [p for p in profiles if target_key in collect_profile_aliases(p)]
        if expect_eq is None:
            _check("'{}' does not match".format(block_name), matched == [])
        else:
            _check(
                "'{}' -> {}".format(block_name, expect_eq),
                len(matched) == 1 and matched[0]["id"] == expect_eq,
                "got {}".format([m["id"] for m in matched]),
            )


def run():
    test_strip_trailing_suffix()
    test_normalize_name()
    test_collect_profile_aliases()
    test_profile_family_names()
    test_match_linked_revit_logic()
    test_match_cad_alias_logic()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[placement] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
