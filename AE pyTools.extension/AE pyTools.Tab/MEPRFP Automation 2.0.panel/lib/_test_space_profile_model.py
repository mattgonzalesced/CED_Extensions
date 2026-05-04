# -*- coding: utf-8 -*-
"""Tests for space_profile_model.py."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from space_profile_model import (
    SpaceProfile,
    SpaceLinkedSet,
    SpaceLED,
    PlacementRule,
    PLACEMENT_KINDS,
    KIND_CENTER, KIND_NE, KIND_DOOR_RELATIVE, KIND_N, KIND_E,
    EDGE_KINDS, CORNER_KINDS,
    is_valid_placement_kind,
    wrap_profiles,
    find_profile_by_id,
    profiles_for_bucket,
    profiles_for_buckets,
)
from profile_model import Offset, Annotation


_FAILS = []


def _check(label, cond, detail=""):
    if cond:
        print("  PASS  {}".format(label))
    else:
        print("  FAIL  {}  {}".format(label, detail))
        _FAILS.append(label)


# ---------------------------------------------------------------------
# Placement rule
# ---------------------------------------------------------------------

def test_placement_kinds_inventory():
    print("\n[placement] kind inventory")
    _check("count", len(PLACEMENT_KINDS) == 10)
    _check("center in set", KIND_CENTER in PLACEMENT_KINDS)
    _check("door_relative in set", KIND_DOOR_RELATIVE in PLACEMENT_KINDS)
    _check("EDGE_KINDS == 4", len(EDGE_KINDS) == 4)
    _check("CORNER_KINDS == 4", len(CORNER_KINDS) == 4)
    _check("EDGE/CORNER disjoint", not (EDGE_KINDS & CORNER_KINDS))
    _check("validate ok", is_valid_placement_kind(KIND_NE))
    _check("validate bad",
           is_valid_placement_kind("up_and_to_the_left") is False)


def test_placement_rule_defaults():
    print("\n[placement] default rule")
    rule = PlacementRule({})
    _check("default kind=center", rule.kind == KIND_CENTER)
    _check("default inset=0", rule.inset_inches == 0.0)
    door = rule.door_offset_inches
    _check("door_offset has x", "x" in door and door["x"] == 0.0)
    _check("door_offset has y", "y" in door and door["y"] == 0.0)


def test_placement_rule_setters_round_trip():
    print("\n[placement] setters round-trip")
    d = {}
    rule = PlacementRule(d)
    rule.kind = KIND_NE
    rule.inset_inches = 18
    rule.door_offset_x_inches = 12
    rule.door_offset_y_inches = -6
    _check("dict kind", d.get("kind") == KIND_NE)
    _check("dict inset", d.get("inset_inches") == 18.0)
    door = d.get("door_offset_inches") or {}
    _check("dict door.x", door.get("x") == 12.0)
    _check("dict door.y", door.get("y") == -6.0)


def test_placement_rule_invalid_kind_falls_back_on_read():
    print("\n[placement] invalid kind read")
    rule = PlacementRule({"kind": ""})
    _check("blank kind reads as center", rule.kind == KIND_CENTER)


# ---------------------------------------------------------------------
# Space LED
# ---------------------------------------------------------------------

def test_space_led_basic():
    print("\n[led] basic fields + lazy placement_rule")
    raw = {
        "id": "L1",
        "label": "Receptacle : Wall",
        "category": "Electrical Fixtures",
    }
    led = SpaceLED(raw)
    _check("id", led.id == "L1")
    _check("label", led.label == "Receptacle : Wall")
    _check("category", led.category == "Electrical Fixtures")
    _check("not group", led.is_group is False)
    rule = led.placement_rule
    _check("placement_rule wrapped", isinstance(rule, PlacementRule))
    _check("rule defaults applied", rule.kind == KIND_CENTER)
    _check("rule mounted into led dict",
           "placement_rule" in raw)


def test_space_led_offsets_round_trip():
    print("\n[led] offsets list round-trip")
    raw = {"id": "L1"}
    led = SpaceLED(raw)
    offsets = led.offsets
    _check("starts empty", offsets == [])
    raw["offsets"] = [
        {"x_inches": 0, "y_inches": 0, "z_inches": 18},
        {"x_inches": 12, "y_inches": 0, "z_inches": 18},
    ]
    offs = led.offsets
    _check("two offsets read", len(offs) == 2)
    _check("first z=18", offs[0].z_inches == 18.0)
    _check("second x=12", offs[1].x_inches == 12.0)


def test_space_led_annotations():
    print("\n[led] annotations pass-through")
    raw = {
        "id": "L1",
        "annotations": [
            {"id": "A1", "kind": "tag", "label": "Recept Tag"},
        ],
    }
    led = SpaceLED(raw)
    annots = led.annotations
    _check("one annotation", len(annots) == 1)
    _check("is Annotation",
           isinstance(annots[0], Annotation))
    _check("kind passed through", annots[0].kind == "tag")


# ---------------------------------------------------------------------
# Profile + LinkedSet
# ---------------------------------------------------------------------

def test_space_profile_basic():
    print("\n[profile] basic + linked_sets")
    raw = {
        "id": "SP-001",
        "name": "HEB Bakery Default",
        "bucket_id": "BUCKET-001",
        "linked_sets": [
            {
                "id": "SET-1",
                "name": "Receptacles",
                "linked_element_definitions": [
                    {"id": "LED-A"}, {"id": "LED-B"},
                ],
            },
        ],
    }
    p = SpaceProfile(raw)
    _check("id", p.id == "SP-001")
    _check("name", p.name == "HEB Bakery Default")
    _check("bucket_id", p.bucket_id == "BUCKET-001")
    sets = p.linked_sets
    _check("one set", len(sets) == 1)
    _check("set is wrapped",
           isinstance(sets[0], SpaceLinkedSet))
    leds = sets[0].leds
    _check("two leds", len(leds) == 2)
    _check("led ids",
           [l.id for l in leds] == ["LED-A", "LED-B"])


def test_space_profile_setters():
    print("\n[profile] setters round-trip")
    d = {}
    p = SpaceProfile(d)
    p.id = "SP-9"
    p.name = "Custom"
    p.bucket_id = "BK-7"
    _check("dict id", d.get("id") == "SP-9")
    _check("dict name", d.get("name") == "Custom")
    _check("dict bucket_id", d.get("bucket_id") == "BK-7")


# ---------------------------------------------------------------------
# Collection helpers
# ---------------------------------------------------------------------

def test_find_profile_by_id():
    print("\n[helpers] find_profile_by_id")
    raw = [
        {"id": "SP-1", "name": "A"},
        {"id": "SP-2", "name": "B"},
    ]
    found = find_profile_by_id(raw, "SP-2")
    _check("returns wrapper", isinstance(found, SpaceProfile))
    _check("right one", found.name == "B")
    _check("missing -> None", find_profile_by_id(raw, "SP-99") is None)
    _check("blank -> None", find_profile_by_id(raw, "") is None)


def test_profiles_for_bucket():
    print("\n[helpers] profiles_for_bucket (single)")
    raw = [
        {"id": "SP-1", "bucket_id": "RR"},
        {"id": "SP-2", "bucket_id": "BK"},
        {"id": "SP-3", "bucket_id": "RR"},  # second restroom profile
        {"id": "SP-4"},  # no bucket
    ]
    rr = profiles_for_bucket(raw, "RR")
    _check("two RR profiles", len(rr) == 2)
    _check("preserves YAML order",
           [p.id for p in rr] == ["SP-1", "SP-3"])
    _check("bk one profile",
           [p.id for p in profiles_for_bucket(raw, "BK")] == ["SP-2"])
    _check("missing returns []",
           profiles_for_bucket(raw, "ZZ") == [])


def test_profiles_for_buckets_dedup():
    print("\n[helpers] profiles_for_buckets unions and dedups")
    raw = [
        {"id": "SP-1", "bucket_id": "RR"},
        {"id": "SP-2", "bucket_id": "WMN"},
        {"id": "SP-3", "bucket_id": "RR"},
        {"id": "SP-4", "bucket_id": "BK"},
    ]
    out = profiles_for_buckets(raw, ["RR", "WMN"])
    ids = [p.id for p in out]
    _check("RR + WMN -> 3 profiles", len(out) == 3)
    _check("YAML order preserved",
           ids == ["SP-1", "SP-2", "SP-3"])
    # Same bucket twice should not duplicate.
    out2 = profiles_for_buckets(raw, ["RR", "RR"])
    _check("dup bucket -> still two",
           [p.id for p in out2] == ["SP-1", "SP-3"])
    _check("empty bucket list -> []",
           profiles_for_buckets(raw, []) == [])


def test_wrap_profiles():
    print("\n[helpers] wrap_profiles tolerates non-dicts")
    raw = [
        {"id": "SP-1"},
        "not a dict",
        None,
        {"id": "SP-2"},
    ]
    out = wrap_profiles(raw)
    _check("two wrapped", len(out) == 2)
    _check("ids", [p.id for p in out] == ["SP-1", "SP-2"])


# ---------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------

def main():
    print("Running space_profile_model tests")
    test_placement_kinds_inventory()
    test_placement_rule_defaults()
    test_placement_rule_setters_round_trip()
    test_placement_rule_invalid_kind_falls_back_on_read()
    test_space_led_basic()
    test_space_led_offsets_round_trip()
    test_space_led_annotations()
    test_space_profile_basic()
    test_space_profile_setters()
    test_find_profile_by_id()
    test_profiles_for_bucket()
    test_profiles_for_buckets_dedup()
    test_wrap_profiles()

    print("")
    if _FAILS:
        print("FAILED: {} test(s) — {}".format(len(_FAILS), _FAILS))
        sys.exit(1)
    print("All space_profile_model tests passed.")


if __name__ == "__main__":
    main()
