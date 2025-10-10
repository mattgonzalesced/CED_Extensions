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


def orient_instance_on_wall(inst, wall=None, space=None, logger=None, sample_ft=0.5):
    """Best-effort flip so the device faces the interior."""
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

    if wall is None:
        return

    wall_orient = _normalize(getattr(wall, "Orientation", None))
    if wall_orient is None:
        return

    if can_flip:
        try:
            dot = facing.X * (-wall_orient.X) + facing.Y * (-wall_orient.Y) + facing.Z * (-wall_orient.Z)
        except Exception:
            dot = None
        if dot is not None and dot < 0.0:
            try:
                inst.FlipFacing()
            except Exception as ex:
                if logger:
                    try:
                        logger.debug(u"[ORIENT] flip (wall) failed: {}".format(ex))
                    except Exception:
                        pass
