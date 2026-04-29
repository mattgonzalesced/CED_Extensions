# -*- coding: utf-8 -*-
"""Pure-Python tests for the ID-generation helpers in capture.py.

We can't load capture.py directly here because it imports the Revit API
at the top. Instead we re-import the small ``_max_numeric_suffix`` and
``_next_*`` functions by extracting them into this test fixture form
manually — duplicating the logic so the test stays offline.

If capture.py's id rules change, also update this test."""

from __future__ import print_function

import os
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


def _max_numeric_suffix(strings, prefix):
    best = 0
    for s in strings:
        if not isinstance(s, str) or not s.startswith(prefix):
            continue
        rest = s[len(prefix):]
        try:
            n = int(rest)
        except ValueError:
            continue
        if n > best:
            best = n
    return best


def test_max_suffix():
    print("\n[capture-idgen] _max_numeric_suffix")
    _check("empty", _max_numeric_suffix([], "EQ-") == 0)
    _check("one", _max_numeric_suffix(["EQ-001"], "EQ-") == 1)
    _check("max",
           _max_numeric_suffix(["EQ-001", "EQ-007", "EQ-003"], "EQ-") == 7)
    _check("ignore non-matching prefix",
           _max_numeric_suffix(["SET-099", "EQ-001"], "EQ-") == 1)
    _check("ignore non-numeric tail",
           _max_numeric_suffix(["EQ-001", "EQ-foo"], "EQ-") == 1)


def test_next_eq_id():
    print("\n[capture-idgen] next EQ id")
    profiles = {"equipment_definitions": [
        {"id": "EQ-001"}, {"id": "EQ-005"},
    ]}
    profile_ids = [p.get("id") for p in profiles["equipment_definitions"]]
    n = _max_numeric_suffix(profile_ids, "EQ-") + 1
    next_id = "EQ-{:03d}".format(n)
    _check("EQ-006 next", next_id == "EQ-006")


def test_next_set_id():
    print("\n[capture-idgen] next SET id (across profiles)")
    profiles = {"equipment_definitions": [
        {"id": "EQ-001", "linked_sets": [{"id": "SET-001"}]},
        {"id": "EQ-002", "linked_sets": [{"id": "SET-009"}, {"id": "SET-010"}]},
    ]}
    seen = []
    for p in profiles["equipment_definitions"]:
        for s in p["linked_sets"]:
            seen.append(s["id"])
    n = _max_numeric_suffix(seen, "SET-") + 1
    next_id = "SET-{:03d}".format(n)
    _check("SET-011 next", next_id == "SET-011")


def test_next_led_id():
    print("\n[capture-idgen] next LED id (within set)")
    set_dict = {
        "id": "SET-005",
        "linked_element_definitions": [
            {"id": "SET-005-LED-001"},
            {"id": "SET-005-LED-003"},
        ],
    }
    prefix = "{}-LED-".format(set_dict["id"])
    led_ids = [l["id"] for l in set_dict["linked_element_definitions"]]
    n = _max_numeric_suffix(led_ids, prefix) + 1
    next_id = "{}{:03d}".format(prefix, n)
    _check("LED-004 next", next_id == "SET-005-LED-004")


def run():
    test_max_suffix()
    test_next_eq_id()
    test_next_set_id()
    test_next_led_id()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[capture-idgen] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
