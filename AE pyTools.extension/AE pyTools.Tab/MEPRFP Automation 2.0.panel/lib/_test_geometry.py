# -*- coding: utf-8 -*-
"""Tests for geometry.py. Pure-Python, runs outside Revit."""

from __future__ import print_function

import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import geometry


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def _approx_eq(a, b, eps=1e-6):
    if isinstance(a, (tuple, list)) and isinstance(b, (tuple, list)):
        if len(a) != len(b):
            return False
        return all(abs(float(x) - float(y)) <= eps for x, y in zip(a, b))
    return abs(float(a) - float(b)) <= eps


def test_unit_conversions():
    print("\n[geometry] unit conversions")
    _check("12 in == 1 ft", _approx_eq(geometry.feet_to_inches(1.0), 12.0))
    _check("36 in == 3 ft", _approx_eq(geometry.inches_to_feet(36.0), 3.0))


def test_normalize_angle():
    print("\n[geometry] normalize_angle")
    _check("0 -> 0", _approx_eq(geometry.normalize_angle(0), 0.0))
    _check("180 -> 180", _approx_eq(geometry.normalize_angle(180), 180.0))
    _check("181 -> -179", _approx_eq(geometry.normalize_angle(181), -179.0))
    _check("-180 -> 180", _approx_eq(geometry.normalize_angle(-180), 180.0))
    _check("360 -> 0", _approx_eq(geometry.normalize_angle(360), 0.0))
    _check("720 -> 0", _approx_eq(geometry.normalize_angle(720), 0.0))
    _check("-540 -> 180", _approx_eq(geometry.normalize_angle(-540), 180.0))


def test_rotate_xy():
    print("\n[geometry] rotate_xy")
    _check("rot 0 is identity", _approx_eq(geometry.rotate_xy((1, 0, 5), 0), (1, 0, 5)))
    _check("rot 90 of (1,0)", _approx_eq(geometry.rotate_xy((1, 0, 0), 90), (0, 1, 0)))
    _check("rot -90 of (1,0)", _approx_eq(geometry.rotate_xy((1, 0, 0), -90), (0, -1, 0)))
    _check("rot 180 of (1,0)", _approx_eq(geometry.rotate_xy((1, 0, 0), 180), (-1, 0, 0)))
    _check("z passes through", _approx_eq(geometry.rotate_xy((0, 0, 7), 90)[2], 7))


def test_offset_round_trip_zero_rotation():
    """With parent rotation 0, world delta == local offset."""
    print("\n[geometry] offset round-trip (parent rot = 0)")
    parent = (10.0, 5.0, 0.0)
    child = (12.0, 8.0, 3.0)
    offsets = geometry.compute_offsets_from_points(parent, 0.0, child, 45.0)
    _check("x_inches = 24", _approx_eq(offsets["x_inches"], 24.0))   # (12-10) ft = 24 in
    _check("y_inches = 36", _approx_eq(offsets["y_inches"], 36.0))   # (8-5) ft = 36 in
    _check("z_inches = 36", _approx_eq(offsets["z_inches"], 36.0))   # 3 ft = 36 in
    _check("rotation_deg = 45", _approx_eq(offsets["rotation_deg"], 45.0))

    target = geometry.target_point_from_offsets(parent, 0.0, offsets)
    _check("target == child", _approx_eq(target, child))
    _check("rot reverse",
           _approx_eq(geometry.child_rotation_from_offsets(0.0, offsets), 45.0))


def test_offset_round_trip_rotated_parent():
    """Rotating the parent after capture should still yield correct child world point."""
    print("\n[geometry] offset round-trip (parent rotates 90 after capture)")
    parent_at_capture = (10.0, 5.0, 0.0)
    child_at_capture = (12.0, 5.0, 0.0)  # 2 ft east of parent
    offsets = geometry.compute_offsets_from_points(
        parent_at_capture, 0.0, child_at_capture, 0.0
    )
    # Now rotate parent 90 deg in place. The child should land 2 ft NORTH.
    parent_now = parent_at_capture
    target = geometry.target_point_from_offsets(parent_now, 90.0, offsets)
    expected = (10.0, 7.0, 0.0)
    _check("rotated child north", _approx_eq(target, expected),
           "got {}".format(target))

    child_rot_now = geometry.child_rotation_from_offsets(90.0, offsets)
    _check("child rotation tracks parent", _approx_eq(child_rot_now, 90.0))


def test_angle_from_vector():
    print("\n[geometry] angle_from_vector")
    _check("(1,0) -> 0",   _approx_eq(geometry.angle_from_vector((1, 0)), 0.0))
    _check("(0,1) -> 90",  _approx_eq(geometry.angle_from_vector((0, 1)), 90.0))
    _check("(-1,0) -> 180",_approx_eq(geometry.angle_from_vector((-1, 0)), 180.0))
    _check("(0,-1) -> -90",_approx_eq(geometry.angle_from_vector((0, -1)), -90.0))
    _check("(0,0) -> 0",   _approx_eq(geometry.angle_from_vector((0, 0)), 0.0))


def test_distance_xy():
    print("\n[geometry] distance_xy")
    _check("3-4-5 triangle", _approx_eq(geometry.distance_xy((0, 0, 9), (3, 4, 0)), 5.0))


def run():
    test_unit_conversions()
    test_normalize_angle()
    test_rotate_xy()
    test_offset_round_trip_zero_rotation()
    test_offset_round_trip_rotated_parent()
    test_angle_from_vector()
    test_distance_xy()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[geometry] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
