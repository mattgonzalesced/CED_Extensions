# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/orientation3D.py
from __future__ import absolute_import

import math

from Autodesk.Revit.DB import XYZ, LocationPoint, ElementTransformUtils, Line


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


def _normalize_xy(vec):
    if vec is None:
        return None
    return _normalize(XYZ(vec.X, vec.Y, 0.0))


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
    """Rotate and/or flip a wall-hosted instance so it faces the space interior."""
    if inst is None:
        return

    doc = getattr(inst, "Document", None)
    if doc is None:
        return

    origin = _instance_origin(inst)
    if origin is None:
        return

    facing = _normalize(getattr(inst, "FacingOrientation", None))
    if facing is None:
        return

    can_flip = bool(getattr(inst, "CanFlipFacing", False))

    # Align with wall interior normal first
    if wall is not None:
        wall_orient = _normalize(getattr(wall, "Orientation", None))
        desired_xy = _normalize_xy(XYZ(-wall_orient.X, -wall_orient.Y, 0.0)) if wall_orient else None
        current_xy = _normalize_xy(facing)
        if desired_xy and current_xy:
            try:
                dot = current_xy.X * desired_xy.X + current_xy.Y * desired_xy.Y
                cross = current_xy.X * desired_xy.Y - current_xy.Y * desired_xy.X
                angle = math.atan2(cross, dot)
            except Exception:
                angle = None
            if angle is not None and abs(angle) > 1e-3:
                try:
                    axis = Line.CreateBound(origin, XYZ(origin.X, origin.Y, origin.Z + 1.0))
                    ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)
                    facing = _normalize(getattr(inst, "FacingOrientation", None)) or facing
                    current_xy = _normalize_xy(facing)
                except Exception as ex:
                    if logger:
                        try:
                            logger.debug(u"[ORIENT] rotate failed: {}".format(ex))
                        except Exception:
                            pass

    # After rotation, if space is available ensure facing points inward
    if space is not None and can_flip:
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

    return
