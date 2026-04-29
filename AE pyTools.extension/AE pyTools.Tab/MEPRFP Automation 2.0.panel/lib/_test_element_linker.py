# -*- coding: utf-8 -*-
"""Tests for element_linker.py. Pure-Python, runs outside Revit."""

from __future__ import print_function

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import element_linker as el


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def test_json_round_trip():
    print("\n[element_linker] JSON round-trip")
    a = el.ElementLinker(
        led_id="SET-001-LED-005",
        set_id="SET-001",
        location_ft=[1.5, 2.5, 3.5],
        rotation_deg=45.0,
        parent_rotation_deg=30.0,
        parent_element_id=12345,
        level_id=678,
        element_id=999,
        facing=[1.0, 0.0, 0.0],
    )
    text = a.to_json()
    b = el.ElementLinker.from_json(text)
    _check("equal after round-trip", a == b, "got {}".format(b.to_dict()))
    _check("codec version stamped", json.loads(text)["v"] == el.CODEC_VERSION)


def test_from_json_blank():
    print("\n[element_linker] from_json blank/None")
    _check("None -> None", el.ElementLinker.from_json(None) is None)
    _check("'' -> None", el.ElementLinker.from_json("") is None)
    _check("'   ' -> None", el.ElementLinker.from_json("   ") is None)


def test_from_json_bad_data():
    print("\n[element_linker] from_json error cases")
    try:
        el.ElementLinker.from_json("not valid json")
    except el.ElementLinkerError:
        _check("bad json raises", True)
    else:
        _check("bad json raises", False, "no exception")

    try:
        el.ElementLinker.from_json('{"v": 999}')
    except el.ElementLinkerError:
        _check("future codec version raises", True)
    else:
        _check("future codec version raises", False, "no exception")


def test_legacy_multiline_full():
    print("\n[element_linker] from_legacy_text (multiline, full)")
    text = (
        "Linked Element Definition ID: SET-001-LED-005\n"
        "Set Definition ID: SET-001\n"
        "Location XYZ (ft): 1.500000,2.500000,3.500000\n"
        "Rotation (deg): 45.000000\n"
        "Parent Rotation (deg): 30.000000\n"
        "Parent ElementId: 12345\n"
        "LevelId: 678\n"
        "ElementId: 999\n"
        "FacingOrientation: 1.000000,0.000000,0.000000"
    )
    linker = el.ElementLinker.from_legacy_text(text)
    _check("led_id", linker.led_id == "SET-001-LED-005")
    _check("set_id", linker.set_id == "SET-001")
    _check("location_ft", linker.location_ft == [1.5, 2.5, 3.5])
    _check("rotation_deg", linker.rotation_deg == 45.0)
    _check("parent_rotation_deg", linker.parent_rotation_deg == 30.0)
    _check("parent_element_id (int)", linker.parent_element_id == 12345)
    _check("level_id", linker.level_id == 678)
    _check("element_id", linker.element_id == 999)
    _check("facing", linker.facing == [1.0, 0.0, 0.0])


def test_legacy_multiline_empty_fields():
    print("\n[element_linker] from_legacy_text (empty fields -> None)")
    text = (
        "Linked Element Definition ID: SET-001-LED-002\n"
        "Set Definition ID: SET-001\n"
        "Location XYZ (ft): 60.000000,20.000000,3.000000\n"
        "Rotation (deg): 0.000000\n"
        "Parent Rotation (deg): \n"
        "Parent ElementId: \n"
        "LevelId: \n"
        "ElementId: 654320\n"
        "FacingOrientation: "
    )
    linker = el.ElementLinker.from_legacy_text(text)
    _check("parent_rotation_deg None", linker.parent_rotation_deg is None)
    _check("parent_element_id None", linker.parent_element_id is None)
    _check("level_id None", linker.level_id is None)
    _check("facing None", linker.facing is None)
    _check("element_id 654320", linker.element_id == 654320)


def test_legacy_not_found_literal():
    print("\n[element_linker] from_legacy_text ('Not found' -> None)")
    text = (
        "Linked Element Definition ID: SET-001-LED-001\n"
        "Parent_location: Not found\n"
        "ElementId: 1"
    )
    linker = el.ElementLinker.from_legacy_text(text)
    _check("parent_location_ft None", linker.parent_location_ft is None)
    _check("element_id 1", linker.element_id == 1)


def test_legacy_inline_format():
    print("\n[element_linker] from_legacy_text (inline comma format)")
    text = (
        "Linked Element Definition ID: SET-003-LED-001, "
        "Set Definition ID: SET-003, "
        "Host Name: Equipment-A, "
        "Location XYZ (ft): 155.123,80.654,8.500, "
        "Rotation (deg): 30.000000, "
        "Parent ElementId: 1300000, "
        "LevelId: 789456, "
        "ElementId: 654323, "
        "FacingOrientation: 0.866,0.500,0.000"
    )
    linker = el.ElementLinker.from_legacy_text(text)
    _check("led_id", linker.led_id == "SET-003-LED-001")
    _check("host_name", linker.host_name == "Equipment-A")
    _check("location_ft", linker.location_ft == [155.123, 80.654, 8.5])
    _check("parent_element_id", linker.parent_element_id == 1300000)
    _check("element_id", linker.element_id == 654323)


def test_legacy_alt_capitalization():
    print("\n[element_linker] from_legacy_text ('Parent Element ID' alt key)")
    text = (
        "Linked Element Definition ID: SET-001-LED-001\n"
        "Parent Element ID: 99\n"
        "ElementId: 1"
    )
    linker = el.ElementLinker.from_legacy_text(text)
    _check("alt parent key recognised", linker.parent_element_id == 99)


def test_legacy_blank():
    print("\n[element_linker] from_legacy_text blank")
    _check("'' -> None", el.ElementLinker.from_legacy_text("") is None)
    _check("None -> None", el.ElementLinker.from_legacy_text(None) is None)


def test_legacy_to_json_round_trip():
    """Reading legacy text and re-emitting as JSON should be lossless for
    every field the codec supports."""
    print("\n[element_linker] legacy -> JSON -> codec equivalence")
    text = (
        "Linked Element Definition ID: SET-002-LED-000\n"
        "Set Definition ID: SET-002\n"
        "Location XYZ (ft): 200.000000,100.000000,5.000000\n"
        "Rotation (deg): 90.000000\n"
        "Parent Rotation (deg): 90.000000\n"
        "Parent ElementId: 1500000\n"
        "LevelId: 789456\n"
        "ElementId: 654322\n"
        "FacingOrientation: 0.000000,1.000000,0.000000"
    )
    a = el.ElementLinker.from_legacy_text(text)
    js = a.to_json()
    b = el.ElementLinker.from_json(js)
    _check("legacy -> JSON -> back == legacy", a == b)


def run():
    test_json_round_trip()
    test_from_json_blank()
    test_from_json_bad_data()
    test_legacy_multiline_full()
    test_legacy_multiline_empty_fields()
    test_legacy_not_found_literal()
    test_legacy_inline_format()
    test_legacy_alt_capitalization()
    test_legacy_blank()
    test_legacy_to_json_round_trip()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[element_linker] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
