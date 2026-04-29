# -*- coding: utf-8 -*-
"""
Pure-Python geometry math for parent-relative offsets.

Inputs and outputs use 3-tuples ``(x, y, z)`` instead of Revit's XYZ so
this module is fully testable outside Revit. ``links.py`` does the
``XYZ <-> tuple`` conversion at the Revit-API boundary.

Convention: world points are in feet (Revit internal units), stored
offsets are in inches. Rotations are in degrees normalised to (-180, 180].
"""

import math


INCHES_PER_FOOT = 12.0
_ZERO_TOLERANCE = 1e-9


# Centralised tolerance constants. Anything that compares positions or
# rotations should pull from here rather than hard-coding magic numbers.
class Tolerances(object):
    POSITION_FT = 1.0 / 256.0
    ROTATION_DEG = 0.01
    FAR_FROM_PARENT_FT = 10.0
    ZERO_VECTOR = _ZERO_TOLERANCE


def feet_to_inches(value_ft):
    return float(value_ft) * INCHES_PER_FOOT


def inches_to_feet(value_in):
    return float(value_in) / INCHES_PER_FOOT


def normalize_angle(angle_deg):
    """Wrap an angle into (-180, 180]."""
    if angle_deg is None:
        return 0.0
    a = float(angle_deg) % 360.0
    if a > 180.0:
        a -= 360.0
    elif a <= -180.0:
        a += 360.0
    return a


def rotate_xy(point, angle_deg):
    """Rotate the XY components of a 3-tuple by ``angle_deg``. Z passes through."""
    x, y, z = point
    rad = math.radians(angle_deg)
    cos_a = math.cos(rad)
    sin_a = math.sin(rad)
    return (
        x * cos_a - y * sin_a,
        x * sin_a + y * cos_a,
        z,
    )


def compute_offsets_from_points(parent_point, parent_rotation_deg,
                                child_point, child_rotation_deg):
    """Forward: world coordinates -> local offsets relative to parent.

    The parent rotation is *inverted* during this transform so that the
    stored offset is rotation-resilient: when a placement engine later
    applies the offset against a parent at a different rotation, the
    child lands in the correct position relative to the (rotated) parent.

    Returns dict with keys ``x_inches``, ``y_inches``, ``z_inches``,
    ``rotation_deg``. All values rounded to 6 decimals to match the
    legacy text format precision.
    """
    px, py, pz = parent_point
    cx, cy, cz = child_point
    delta = (cx - px, cy - py, cz - pz)
    local = rotate_xy(delta, -float(parent_rotation_deg or 0.0))
    rel_rot = normalize_angle(
        float(child_rotation_deg or 0.0) - float(parent_rotation_deg or 0.0)
    )
    return {
        "x_inches": round(feet_to_inches(local[0]), 6),
        "y_inches": round(feet_to_inches(local[1]), 6),
        "z_inches": round(feet_to_inches(local[2]), 6),
        "rotation_deg": round(rel_rot, 6),
    }


def target_point_from_offsets(parent_point, parent_rotation_deg, offsets):
    """Reverse: local offsets + parent state -> world child point."""
    local_ft = (
        inches_to_feet(offsets.get("x_inches", 0.0)),
        inches_to_feet(offsets.get("y_inches", 0.0)),
        inches_to_feet(offsets.get("z_inches", 0.0)),
    )
    world_delta = rotate_xy(local_ft, float(parent_rotation_deg or 0.0))
    px, py, pz = parent_point
    return (
        px + world_delta[0],
        py + world_delta[1],
        pz + world_delta[2],
    )


def child_rotation_from_offsets(parent_rotation_deg, offsets):
    """Reverse: parent rotation + offset rotation -> world child rotation."""
    return normalize_angle(
        float(parent_rotation_deg or 0.0)
        + float(offsets.get("rotation_deg", 0.0))
    )


def angle_from_vector(vec):
    """Return the XY angle of a vector in degrees, or 0.0 for a near-zero vector."""
    x, y = float(vec[0]), float(vec[1])
    if abs(x) < _ZERO_TOLERANCE and abs(y) < _ZERO_TOLERANCE:
        return 0.0
    return math.degrees(math.atan2(y, x))


def distance_xy(a, b):
    dx = float(a[0]) - float(b[0])
    dy = float(a[1]) - float(b[1])
    return math.sqrt(dx * dx + dy * dy)
