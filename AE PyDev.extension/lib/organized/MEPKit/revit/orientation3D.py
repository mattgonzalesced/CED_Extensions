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


def _wall_tangent_xy(wall, near_point=None):
    if wall is None:
        return None
    try:
        loc = getattr(wall, "Location", None)
        curve = getattr(loc, "Curve", None)
        if curve is None:
            return None
        t = 0.5
        if near_point is not None:
            try:
                res = curve.Project(near_point)
                if res:
                    t = res.Parameter
            except Exception:
                pass
        try:
            der = curve.ComputeDerivatives(t, True)
            return _normalize_xy(der.BasisX)
        except Exception:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return _normalize_xy(XYZ(p1.X - p0.X, p1.Y - p0.Y, p1.Z - p0.Z))
    except Exception:
        return None


def orient_instance_on_wall(inst, wall=None, space=None, logger=None, sample_ft=0.5):
    """Flip facing/hand so wall devices align with the interior."""
    if inst is None:
        return

    origin = _instance_origin(inst)

    facing = _normalize(getattr(inst, "FacingOrientation", None))
    if facing is None:
        return

    can_flip_facing = bool(getattr(inst, "CanFlipFacing", False))

    # Prefer the provided space for determining front/back
    if space is not None and origin is not None and can_flip_facing:
        try:
            front_pt = _offset(origin, facing, sample_ft)
            back_pt = _offset(origin, facing, -sample_ft)
            inside_front = bool(space.IsPointInSpace(front_pt)) if front_pt else False
            inside_back = bool(space.IsPointInSpace(back_pt)) if back_pt else False

            if inside_back and not inside_front:
                inst.FlipFacing()
                facing = _normalize(getattr(inst, "FacingOrientation", None)) or facing
            elif inside_front:
                pass
        except Exception as ex:
            if logger:
                try:
                    logger.debug(u"[ORIENT] space probe failed: {}".format(ex))
                except Exception:
                    pass

    # Fallback to wall orientation if space was unavailable
    if wall is not None:
        wall_orient = _normalize(getattr(wall, "Orientation", None))
        if wall_orient is not None and can_flip_facing:
            interior = XYZ(-wall_orient.X, -wall_orient.Y, -wall_orient.Z)
            try:
                dot = facing.X * interior.X + facing.Y * interior.Y + facing.Z * interior.Z
            except Exception:
                dot = None
            if dot is not None and dot < 0.0:
                try:
                    inst.FlipFacing()
                    facing = _normalize(getattr(inst, "FacingOrientation", None)) or facing
                except Exception as ex:
                    if logger:
                        try:
                            logger.debug(u"[ORIENT] wall flip failed: {}".format(ex))
                        except Exception:
                            pass

    # Align hand orientation with wall tangent so the device stays orthogonal to the face
    hand = _normalize_xy(getattr(inst, "HandOrientation", None))
    wall_tangent = _wall_tangent_xy(wall, origin) if wall is not None else None
    if hand is not None and wall_tangent is not None and getattr(inst, "CanFlipHand", False):
        try:
            dot = hand.X * wall_tangent.X + hand.Y * wall_tangent.Y
        except Exception:
            dot = None
        if dot is not None and dot < 0.0:
            try:
                inst.FlipHand()
            except Exception as ex:
                if logger:
                    try:
                        logger.debug(u"[ORIENT] hand flip failed: {}".format(ex))
                    except Exception:
                        pass

    return
