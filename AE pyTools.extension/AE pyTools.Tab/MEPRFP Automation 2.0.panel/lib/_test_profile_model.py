# -*- coding: utf-8 -*-
"""Tests for profile_model.py. Exercises wrappers against a synthetic
v4-shape profile that mirrors the structure observed in the real
HEB_profiles_V4_MODIFIED_*.yaml files."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import profile_model as pm


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


SAMPLE = {
    "schema_version": 100,
    "equipment_definitions": [
        {
            "id": "EQ-010",
            "name": "Sample Equipment : Default",
            "schema_version": 100,
            "prompt_on_parent_mismatch": False,
            "parent_filter": {
                "family_name_pattern": "Sample Equipment",
                "category": "Specialty Equipment",
                "parameter_filters": {},
                "type_name_pattern": "Default",
            },
            "linked_sets": [
                {
                    "id": "SET-010",
                    "name": "Sample Equipment : Default Types",
                    "linked_element_definitions": [
                        {
                            "id": "SET-010-LED-001",
                            "label": "EF-U_Receptacle_CED : Quad Wall",
                            "category": "Electrical Fixtures",
                            "is_group": True,
                            "parameters": {"Voltage_CED": 120},
                            "offsets": [
                                {
                                    "rotation_deg": -180.0,
                                    "x_inches": -29.4126,
                                    "y_inches": 11.5,
                                    "z_inches": 18.0,
                                }
                            ],
                            "tags": [
                                {
                                    "category_name": "Annotation Symbols",
                                    "family_name": "EF-Tag_Electrical Fixtures_CED",
                                    "type_name": "Elevation (Inches)",
                                    "parameters": {},
                                    "offsets": {
                                        "rotation_deg": 0.0,
                                        "x_inches": 0.0,
                                        "y_inches": 12.0,
                                        "z_inches": 0.0,
                                    },
                                }
                            ],
                            "keynotes": [],
                            "text_notes": [],
                        }
                    ],
                }
            ],
            "allow_parentless": True,
            "allow_unmatched_parents": True,
            "equipment_properties": {},
            "ced_truth_source_id": "EQ-010",
            "ced_truth_source_name": "Sample Equipment : Default",
        }
    ],
}


def test_document():
    print("\n[profile_model] ProfileDocument")
    doc = pm.ProfileDocument(SAMPLE)
    _check("schema_version 100", doc.schema_version == 100)
    _check("one profile", len(doc.profiles) == 1)
    p = doc.find_profile_by_id("EQ-010")
    _check("find_profile_by_id", p is not None and p.id == "EQ-010")
    _check("find_profile_by_name",
           doc.find_profile_by_name("Sample Equipment : Default") is not None)
    _check("find unknown -> None", doc.find_profile_by_id("EQ-999") is None)


def test_profile():
    print("\n[profile_model] Profile")
    p = pm.ProfileDocument(SAMPLE).profiles[0]
    _check("id", p.id == "EQ-010")
    _check("name", p.name == "Sample Equipment : Default")
    _check("schema_version", p.schema_version == 100)
    _check("allow_parentless", p.allow_parentless is True)
    _check("allow_unmatched_parents", p.allow_unmatched_parents is True)
    _check("prompt_on_parent_mismatch", p.prompt_on_parent_mismatch is False)
    _check("truth_source_id", p.truth_source_id == "EQ-010")
    _check("truth_source_name", p.truth_source_name == "Sample Equipment : Default")
    _check("equipment_properties dict", p.equipment_properties == {})


def test_parent_filter():
    print("\n[profile_model] ParentFilter")
    pf = pm.ProfileDocument(SAMPLE).profiles[0].parent_filter
    _check("category", pf.category == "Specialty Equipment")
    _check("family_name_pattern", pf.family_name_pattern == "Sample Equipment")
    _check("type_name_pattern", pf.type_name_pattern == "Default")
    _check("parameter_filters dict", pf.parameter_filters == {})


def test_linked_set_and_led():
    print("\n[profile_model] LinkedSet and LED")
    p = pm.ProfileDocument(SAMPLE).profiles[0]
    sets = p.linked_sets
    _check("one linked set", len(sets) == 1)
    s = sets[0]
    _check("set id", s.id == "SET-010")
    _check("set name", s.name == "Sample Equipment : Default Types")
    leds = s.leds
    _check("one LED", len(leds) == 1)
    led = leds[0]
    _check("led id", led.id == "SET-010-LED-001")
    _check("led label", led.label == "EF-U_Receptacle_CED : Quad Wall")
    _check("led category", led.category == "Electrical Fixtures")
    _check("led is_group", led.is_group is True)
    _check("led params", led.parameters == {"Voltage_CED": 120})


def test_offsets_list():
    print("\n[profile_model] LED offsets (list)")
    led = pm.ProfileDocument(SAMPLE).profiles[0].linked_sets[0].leds[0]
    offsets = led.offsets
    _check("one offset", len(offsets) == 1)
    o = offsets[0]
    _check("x_inches", abs(o.x_inches - (-29.4126)) < 1e-9)
    _check("y_inches", abs(o.y_inches - 11.5) < 1e-9)
    _check("z_inches", abs(o.z_inches - 18.0) < 1e-9)
    _check("rotation_deg", abs(o.rotation_deg - (-180.0)) < 1e-9)


def test_tags():
    print("\n[profile_model] Tag (single offset dict)")
    led = pm.ProfileDocument(SAMPLE).profiles[0].linked_sets[0].leds[0]
    tags = led.tags
    _check("one tag", len(tags) == 1)
    t = tags[0]
    _check("category_name", t.category_name == "Annotation Symbols")
    _check("family_name", t.family_name == "EF-Tag_Electrical Fixtures_CED")
    _check("type_name", t.type_name == "Elevation (Inches)")
    o = t.offset
    _check("tag offset is Offset", o is not None)
    _check("tag y_inches", abs(o.y_inches - 12.0) < 1e-9)


def test_to_dict_round_trip():
    print("\n[profile_model] to_dict round-trip")
    doc = pm.ProfileDocument(SAMPLE)
    _check("doc.to_dict() is the source dict", doc.to_dict() is SAMPLE)
    p = doc.profiles[0]
    _check("profile.to_dict() identity",
           p.to_dict() is SAMPLE["equipment_definitions"][0])


def test_mutation_propagates():
    """Wrappers own the dict; setter on Offset must propagate to source."""
    print("\n[profile_model] mutation propagates to source dict")
    sample = {
        "equipment_definitions": [{
            "id": "EQ-1",
            "linked_sets": [{
                "id": "SET-1",
                "linked_element_definitions": [{
                    "id": "SET-1-LED-1",
                    "offsets": [{"x_inches": 1.0, "y_inches": 2.0, "z_inches": 3.0, "rotation_deg": 0.0}],
                }],
            }],
        }],
    }
    led = pm.ProfileDocument(sample).profiles[0].linked_sets[0].leds[0]
    o = led.offsets[0]
    o.x_inches = 99.0
    _check("source dict updated",
           sample["equipment_definitions"][0]["linked_sets"][0]
                 ["linked_element_definitions"][0]["offsets"][0]["x_inches"] == 99.0)


def test_missing_optional_fields():
    print("\n[profile_model] missing optional fields")
    minimal = {
        "equipment_definitions": [{
            "id": "EQ-X",
            "name": "Minimal",
            "linked_sets": [],
        }],
    }
    p = pm.ProfileDocument(minimal).profiles[0]
    _check("schema_version absent -> None", p.schema_version is None)
    _check("truth_source_id absent -> None", p.truth_source_id is None)
    _check("allow_parentless default True", p.allow_parentless is True)
    _check("prompt_on_parent_mismatch default False",
           p.prompt_on_parent_mismatch is False)


def test_bool_string_coercion():
    """Legacy v1/v3 stored these as 'true'/'false' strings; wrappers
    must still return real booleans."""
    print("\n[profile_model] bool-string coercion")
    legacy = {
        "equipment_definitions": [{
            "id": "EQ-L",
            "allow_parentless": "true",
            "allow_unmatched_parents": "false",
            "prompt_on_parent_mismatch": "TRUE",
        }],
    }
    p = pm.ProfileDocument(legacy).profiles[0]
    _check("'true' -> True", p.allow_parentless is True)
    _check("'false' -> False", p.allow_unmatched_parents is False)
    _check("'TRUE' -> True", p.prompt_on_parent_mismatch is True)


def run():
    test_document()
    test_profile()
    test_parent_filter()
    test_linked_set_and_led()
    test_offsets_list()
    test_tags()
    test_to_dict_round_trip()
    test_mutation_propagates()
    test_missing_optional_fields()
    test_bool_string_coercion()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[profile_model] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
