# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/orientation3D.py
from __future__ import absolute_import

import math

from Autodesk.Revit.DB import XYZ, LocationPoint


def _normalize(vec):
    if vec is None:
        return None
    try:
        mag = math.sqrt(vec.X * vec.X + vec.Y * vec.Y + vec.Z * vec.Z)
    except Exception:
        return None
    if mag < 1e-9:
        return None
    return XYZ(vec.X / mag, vec.Y / mag, vec.Z / mag)


def _instance_origin(inst):
    if inst is None:
        return None
    try:
        loc = getattr(inst, "Location", None)
        if isinstance(loc, LocationPoint):
            return loc.Point
    except Exception:
        pass
    return None


def _offset(point, direction, distance):
    if point is None or direction is None or distance is None:
        return None
    try:
        return XYZ(
            point.X + direction.X * distance,
            point.Y + direction.Y * distance,
            point.Z + direction.Z * distance,
        )
    except Exception:
        return None


def orient_instance_on_wall(inst, wall=None, space=None, logger=None, sample_ft=0.25):
    """
    Ensure a wall-hosted device faces the intended interior.

    - If a space is provided, we sample points in front/back of the device and flip
      so that the device's front points into the space.
    - Otherwise we fall back to wall orientation (exterior normal) and ensure the
      family instance faces the opposite/inward direction.
    """
    if inst is None:
        return

    facing = _normalize(getattr(inst, "FacingOrientation", None))
    if facing is None:
        return

    can_flip = bool(getattr(inst, "CanFlipFacing", False))
    origin = _instance_origin(inst)

    if space is not None and origin is not None and can_flip:
        try:
            front_pt = _offset(origin, facing, sample_ft)
            back_pt = _offset(origin, facing, -sample_ft)
            inside_front = bool(space.IsPointInSpace(front_pt)) if front_pt else False
            inside_back = bool(space.IsPointInSpace(back_pt)) if back_pt else False

            if inside_back and not inside_front:
                inst.FlipFacing()
                return
            if inside_front and not inside_back:
                return
        except Exception as ex:
            if logger:
                try:
                    logger.debug(u"[ORIENT] space probe failed: {}".format(ex))
                except Exception:
                    pass

    if wall is None or not can_flip:
        return

    wall_orient = _normalize(getattr(wall, "Orientation", None))
    if wall_orient is None:
        return

    # Wall.Orientation points toward the exterior; we want devices to face interior.
    interior_dir = XYZ(-wall_orient.X, -wall_orient.Y, -wall_orient.Z)

    try:
        dot = facing.X * interior_dir.X + facing.Y * interior_dir.Y + facing.Z * interior_dir.Z
    except Exception:
        dot = None

    if dot is not None and dot < 0.0:
        try:
            inst.FlipFacing()
        except Exception as ex:
            if logger:
                try:
                    logger.debug(u"[ORIENT] flip failed: {}".format(ex))
                except Exception:
                    pass
