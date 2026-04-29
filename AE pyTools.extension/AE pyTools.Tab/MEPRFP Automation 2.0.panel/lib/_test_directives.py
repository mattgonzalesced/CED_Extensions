# -*- coding: utf-8 -*-
"""Tests for directives.py."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import directives as _dir


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def test_classification():
    print("\n[directives] classification")
    _check("static int", _dir.directive_kind(120) == "static")
    _check("static str", _dir.directive_kind("foo") == "static")
    _check("parent dict",
           _dir.directive_kind({"parent_parameter": "PanelName"}) == "parent")
    _check("sibling dict",
           _dir.directive_kind({"sibling_parameter": "LED-1:CKT"}) == "sibling")


def test_constructors():
    print("\n[directives] constructors")
    p = _dir.parent_directive("PanelName")
    _check("parent shape", p == {"parent_parameter": "PanelName"})

    s = _dir.sibling_directive("SET-001-LED-002", "Voltage")
    _check("sibling shape",
           s == {"sibling_parameter": "SET-001-LED-002:Voltage"})

    try:
        _dir.parent_directive("")
    except _dir.DirectiveError:
        _check("parent empty raises", True)
    else:
        _check("parent empty raises", False)


def test_accessors():
    print("\n[directives] accessors")
    p = {"parent_parameter": "PanelName"}
    _check("parent name", _dir.parent_param_name(p) == "PanelName")

    s = {"sibling_parameter": "SET-001-LED-002:Voltage"}
    led, name = _dir.sibling_target(s)
    _check("sibling led", led == "SET-001-LED-002")
    _check("sibling name", name == "Voltage")

    bad = {"sibling_parameter": "no-colon-here"}
    _check("malformed sibling -> None", _dir.sibling_target(bad) is None)


def test_resolve_static():
    print("\n[directives] resolve static")
    found, value = _dir.resolve_expected_value(120, lambda *_: None, lambda *_: None)
    _check("static found", found is True)
    _check("static value", value == 120)


def test_resolve_parent():
    print("\n[directives] resolve parent")
    parent = {"PanelName": "PA-1", "Voltage": 120}
    parent_lookup = lambda name: parent.get(name)
    sibling_lookup = lambda *_: None

    p = _dir.parent_directive("PanelName")
    found, value = _dir.resolve_expected_value(p, parent_lookup, sibling_lookup)
    _check("parent found", found is True)
    _check("parent value", value == "PA-1")

    miss = _dir.parent_directive("DoesNotExist")
    found, _ = _dir.resolve_expected_value(miss, parent_lookup, sibling_lookup)
    _check("parent miss -> not found", found is False)


def test_resolve_sibling():
    print("\n[directives] resolve sibling")
    siblings = {("LED-2", "CKT"): "5"}
    sibling_lookup = lambda led, name: siblings.get((led, name))
    parent_lookup = lambda *_: None

    s = _dir.sibling_directive("LED-2", "CKT")
    found, value = _dir.resolve_expected_value(s, parent_lookup, sibling_lookup)
    _check("sibling found", found is True)
    _check("sibling value", value == "5")


def run():
    test_classification()
    test_constructors()
    test_accessors()
    test_resolve_static()
    test_resolve_parent()
    test_resolve_sibling()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[directives] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
