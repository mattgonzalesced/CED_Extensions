# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/orientation3D.py
from __future__ import absolute_import

import math

from Autodesk.Revit.DB import XYZ, LocationPoint, ElementTransformUtils, Wall


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


def _interior_direction(wall):
    if wall is None:
        return None
    try:
        orient = _normalize(getattr(wall, "Orientation", None))
        if orient is None:
            return None
        return XYZ(-orient.X, -orient.Y, -orient.Z)
    except Exception:
        return None


def orient_instance_on_wall(inst, wall=None, space=None, logger=None, sample_ft=0.75, pad_ft=0.01):
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
    doc = getattr(inst, "Document", None)

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

    wall_dir = _interior_direction(wall)
    if wall_dir is None:
        return

    if can_flip:
        try:
            dot = facing.X * wall_dir.X + facing.Y * wall_dir.Y + facing.Z * wall_dir.Z
        except Exception:
            dot = None
        if dot is not None and dot < 0.0:
            try:
                inst.FlipFacing()
                facing = _normalize(getattr(inst, "FacingOrientation", None)) or facing
            except Exception as ex:
                if logger:
                    try:
                        logger.debug(u"[ORIENT] flip (wall) failed: {}".format(ex))
                    except Exception:
                        pass

    # Recompute interior direction in case the family needs the opposite sign
    if facing is not None:
        try:
            dot = facing.X * wall_dir.X + facing.Y * wall_dir.Y + facing.Z * wall_dir.Z
            if dot is not None and dot < 0.0:
                wall_dir = XYZ(-wall_dir.X, -wall_dir.Y, -wall_dir.Z)
        except Exception:
            pass

    if origin is None or doc is None:
        return

    try:
        width = float(getattr(wall, "Width", 0.0) or 0.0)
    except Exception:
        width = 0.0

    desired_offset = max(0.0, width * 0.5 - float(pad_ft or 0.0))

    move_dir = wall_dir
    if facing is not None:
        try:
            dot_fd = facing.X * move_dir.X + facing.Y * move_dir.Y + facing.Z * move_dir.Z
            if dot_fd is not None and dot_fd < 0.0:
                move_dir = XYZ(-move_dir.X, -move_dir.Y, -move_dir.Z)
        except Exception:
            pass

    center_pt = None
    try:
        if isinstance(wall, Wall):
            loc = getattr(wall, "Location", None)
            curve = getattr(loc, "Curve", None)
            if curve is not None:
                proj = curve.Project(origin)
                if proj:
                    center_pt = proj.XYZPoint
    except Exception:
        center_pt = None

    if center_pt is None:
        center_pt = origin  # fallback

    try:
        current_vec = XYZ(origin.X - center_pt.X, origin.Y - center_pt.Y, origin.Z - center_pt.Z)
        current_offset = abs(current_vec.X * move_dir.X + current_vec.Y * move_dir.Y + current_vec.Z * move_dir.Z)
    except Exception:
        current_offset = 0.0

    delta = desired_offset - current_offset
    if abs(delta) < 1e-4:
        return

    try:
        shift = XYZ(move_dir.X * delta, move_dir.Y * delta, move_dir.Z * delta)
        ElementTransformUtils.MoveElement(doc, inst.Id, shift)
    except Exception as ex:
        if logger:
            try:
                logger.debug(u"[ORIENT] move failed: {}".format(ex))
            except Exception:
                pass
