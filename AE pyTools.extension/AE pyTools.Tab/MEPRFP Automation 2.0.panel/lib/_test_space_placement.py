# -*- coding: utf-8 -*-
"""Tests for space_placement (pure-logic anchor computation)."""

from __future__ import print_function

import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from space_placement import (
    SpaceGeometry,
    anchor_points,
    expand_led_placements,
    INCHES_PER_FOOT,
)
from space_profile_model import (
    PlacementRule,
    SpaceLED,
    KIND_CENTER, KIND_N, KIND_S, KIND_E, KIND_W,
    KIND_NE, KIND_NW, KIND_SE, KIND_SW,
    KIND_DOOR_RELATIVE,
)


_FAILS = []


def _check(label, cond, detail=""):
    if cond:
        print("  PASS  {}".format(label))
    else:
        print("  FAIL  {}  {}".format(label, detail))
        _FAILS.append(label)


def _close(a, b, eps=1e-6):
    return abs(a - b) < eps


def _make_geom(xmin=0, ymin=0, xmax=20, ymax=10, z=100, doors=None, name=""):
    return SpaceGeometry(
        bbox=((float(xmin), float(ymin)), (float(xmax), float(ymax))),
        floor_z=float(z),
        door_anchors=doors or [],
        name=name,
    )


# ---------------------------------------------------------------------
# center / edges / corners
# ---------------------------------------------------------------------

def test_center():
    print("\n[anchor] center")
    g = _make_geom()  # 20x10 box at z=100
    pts = anchor_points(PlacementRule({"kind": KIND_CENTER}), g)
    _check("one point", len(pts) == 1)
    x, y, z = pts[0]
    _check("center x", _close(x, 10.0))
    _check("center y", _close(y, 5.0))
    _check("center z=floor", _close(z, 100.0))


def test_edges_no_inset():
    print("\n[anchor] edges with zero inset land on bbox edge")
    g = _make_geom(0, 0, 20, 10, 100)
    cases = [
        (KIND_N, 10.0, 10.0),  # midpoint of north edge
        (KIND_S, 10.0, 0.0),
        (KIND_E, 20.0, 5.0),
        (KIND_W, 0.0, 5.0),
    ]
    for kind, ex, ey in cases:
        pts = anchor_points(PlacementRule({"kind": kind, "inset_inches": 0}), g)
        _check("{} one point".format(kind), len(pts) == 1)
        x, y, z = pts[0]
        _check("{} x".format(kind), _close(x, ex))
        _check("{} y".format(kind), _close(y, ey))
        _check("{} z".format(kind), _close(z, 100.0))


def test_edges_inset_inward():
    print("\n[anchor] inset_inches pushes IN-ward from the edge")
    g = _make_geom(0, 0, 20, 10, 100)
    inset_in = 12  # 1 ft
    cases = [
        (KIND_N, 10.0, 10.0 - 1.0),
        (KIND_S, 10.0, 0.0 + 1.0),
        (KIND_E, 20.0 - 1.0, 5.0),
        (KIND_W, 0.0 + 1.0, 5.0),
    ]
    for kind, ex, ey in cases:
        pts = anchor_points(
            PlacementRule({"kind": kind, "inset_inches": inset_in}), g
        )
        x, y, _ = pts[0]
        _check("{} inset x".format(kind), _close(x, ex))
        _check("{} inset y".format(kind), _close(y, ey))


def test_corners_inset_inward():
    print("\n[anchor] corners inset along BOTH axes")
    g = _make_geom(0, 0, 20, 10, 100)
    inset_in = 6  # 0.5 ft
    cases = [
        (KIND_NE, 20.0 - 0.5, 10.0 - 0.5),
        (KIND_NW, 0.0 + 0.5, 10.0 - 0.5),
        (KIND_SE, 20.0 - 0.5, 0.0 + 0.5),
        (KIND_SW, 0.0 + 0.5, 0.0 + 0.5),
    ]
    for kind, ex, ey in cases:
        pts = anchor_points(
            PlacementRule({"kind": kind, "inset_inches": inset_in}), g
        )
        x, y, _ = pts[0]
        _check("{} x".format(kind), _close(x, ex))
        _check("{} y".format(kind), _close(y, ey))


def test_unknown_kind_returns_empty():
    print("\n[anchor] unknown kind returns []")
    g = _make_geom()
    pts = anchor_points(PlacementRule({"kind": "no_such_kind"}), g)
    _check("no anchors", pts == [])


def test_no_geom_returns_empty():
    print("\n[anchor] missing geom or rule returns []")
    _check("None geom", anchor_points(PlacementRule({"kind": KIND_CENTER}), None) == [])
    _check("None rule", anchor_points(None, _make_geom()) == [])
    _check(
        "no bbox",
        anchor_points(
            PlacementRule({"kind": KIND_CENTER}),
            SpaceGeometry(bbox=None, floor_z=0),
        ) == [],
    )


# ---------------------------------------------------------------------
# door_relative
# ---------------------------------------------------------------------

def test_door_relative_zero_offset():
    print("\n[anchor] door_relative with zero offset returns door origins")
    # Two doors: one on south wall (inward = +Y), one on east wall (inward = -X).
    doors = [
        ((10.0, 0.0), (0.0, 1.0)),
        ((20.0, 5.0), (-1.0, 0.0)),
    ]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(PlacementRule({"kind": KIND_DOOR_RELATIVE}), g)
    _check("two anchors", len(pts) == 2)
    _check("first at door 1", _close(pts[0][0], 10.0) and _close(pts[0][1], 0.0))
    _check("second at door 2", _close(pts[1][0], 20.0) and _close(pts[1][1], 5.0))
    _check("z carries floor", _close(pts[0][2], 100.0))


def test_door_relative_x_pushes_inward():
    print("\n[anchor] door x-offset pushes along inward normal")
    # South-wall door, inward = +Y. x=12 in (1 ft) should bump y by +1.
    doors = [((10.0, 0.0), (0.0, 1.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({
            "kind": KIND_DOOR_RELATIVE,
            "door_offset_inches": {"x": 12, "y": 0},
        }), g,
    )
    x, y, _ = pts[0]
    _check("x unchanged", _close(x, 10.0))
    _check("y bumped +1ft", _close(y, 1.0))


def test_door_relative_y_pushes_sideways():
    print("\n[anchor] door y-offset pushes 90deg CCW from inward")
    # South-wall door, inward = +Y. Sideways (90 CCW) = -X.
    # y=12 in (1 ft) should shift x by -1.
    doors = [((10.0, 0.0), (0.0, 1.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({
            "kind": KIND_DOOR_RELATIVE,
            "door_offset_inches": {"x": 0, "y": 12},
        }), g,
    )
    x, y, _ = pts[0]
    _check("x shifted -1ft", _close(x, 9.0))
    _check("y unchanged", _close(y, 0.0))


def test_door_relative_combined_offset():
    print("\n[anchor] door x+y combined")
    # East-wall door, inward = -X. Sideways (90 CCW) = -Y inward => let me compute.
    # inward = (-1, 0); rotated 90 CCW => (0, -1). So +y goes -Y direction.
    doors = [((20.0, 5.0), (-1.0, 0.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({
            "kind": KIND_DOOR_RELATIVE,
            "door_offset_inches": {"x": 24, "y": 12},  # 2ft inward, 1ft side
        }), g,
    )
    x, y, _ = pts[0]
    # inward 2ft -> x = 20 - 2 = 18
    # side 1ft (CCW 90 from (-1,0) is (0,-1)) -> y = 5 - 1 = 4
    _check("x = 18", _close(x, 18.0))
    _check("y = 4", _close(y, 4.0))


def test_door_relative_unnormalized_input():
    print("\n[anchor] non-unit inward normals are normalised")
    # inward (3, 4) length 5; should be normalised to (0.6, 0.8).
    doors = [((0.0, 0.0), (3.0, 4.0))]
    g = _make_geom(-10, -10, 10, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({
            "kind": KIND_DOOR_RELATIVE,
            "door_offset_inches": {"x": 60, "y": 0},  # 5 ft inward
        }), g,
    )
    x, y, _ = pts[0]
    _check("x = 3.0 (5ft * 0.6)", _close(x, 3.0))
    _check("y = 4.0 (5ft * 0.8)", _close(y, 4.0))


def test_door_relative_no_doors_returns_empty():
    print("\n[anchor] door_relative with no doors -> []")
    g = _make_geom()
    pts = anchor_points(
        PlacementRule({"kind": KIND_DOOR_RELATIVE}), g,
    )
    _check("empty list", pts == [])


# ---------------------------------------------------------------------
# expand_led_placements
# ---------------------------------------------------------------------

def test_expand_led_no_offsets_yields_one_per_anchor():
    print("\n[expand] LED with no offsets -> one per anchor")
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": KIND_CENTER},
    })
    g = _make_geom(0, 0, 20, 10, 100)
    out = expand_led_placements(led, g)
    _check("one placement", len(out) == 1)
    x, y, z, rot = out[0]
    _check("at center", _close(x, 10.0) and _close(y, 5.0))
    _check("z=floor", _close(z, 100.0))
    _check("rotation 0", _close(rot, 0.0))


def test_expand_led_one_offset_lifts_z():
    print("\n[expand] z_inches offset lifts above floor")
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": KIND_CENTER},
        "offsets": [{"z_inches": 18}],   # 1.5 ft
    })
    g = _make_geom(0, 0, 20, 10, 100)
    out = expand_led_placements(led, g)
    _check("one placement", len(out) == 1)
    _, _, z, _ = out[0]
    _check("z = 101.5", _close(z, 101.5))


def test_expand_multiplies_anchors_by_offsets():
    print("\n[expand] anchors * offsets")
    # Two doors x two offsets = 4 placements.
    doors = [
        ((10.0, 0.0), (0.0, 1.0)),
        ((10.0, 10.0), (0.0, -1.0)),
    ]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": KIND_DOOR_RELATIVE},
        "offsets": [
            {"z_inches": 18},
            {"z_inches": 96, "rotation_deg": 90},
        ],
    })
    out = expand_led_placements(led, g)
    _check("4 placements", len(out) == 4)
    rotations = sorted(set(p[3] for p in out))
    _check("two distinct rotations", rotations == [0.0, 90.0])
    zs = sorted(set(round(p[2], 4) for p in out))
    _check("two distinct z (floor+1.5, floor+8)", zs == [101.5, 108.0])


def test_expand_unknown_kind_yields_empty():
    print("\n[expand] unknown rule kind -> []")
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": "no_such_kind"},
        "offsets": [{"z_inches": 18}],
    })
    g = _make_geom()
    out = expand_led_placements(led, g)
    _check("no placements", out == [])


# ---------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------

def main():
    print("Running space_placement tests")
    test_center()
    test_edges_no_inset()
    test_edges_inset_inward()
    test_corners_inset_inward()
    test_unknown_kind_returns_empty()
    test_no_geom_returns_empty()
    test_door_relative_zero_offset()
    test_door_relative_x_pushes_inward()
    test_door_relative_y_pushes_sideways()
    test_door_relative_combined_offset()
    test_door_relative_unnormalized_input()
    test_door_relative_no_doors_returns_empty()
    test_expand_led_no_offsets_yields_one_per_anchor()
    test_expand_led_one_offset_lifts_z()
    test_expand_multiplies_anchors_by_offsets()
    test_expand_unknown_kind_yields_empty()

    print("")
    if _FAILS:
        print("FAILED: {} test(s) -- {}".format(len(_FAILS), _FAILS))
        sys.exit(1)
    print("All space_placement tests passed.")


if __name__ == "__main__":
    main()
