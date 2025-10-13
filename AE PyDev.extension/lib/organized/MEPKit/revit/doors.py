# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/doors.py
from __future__ import absolute_import
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, LocationPoint, LocationCurve,
    XYZ, RevitLinkInstance
)
from organized.MEPKit.revit.params import get_param_value

_DOOR_META_CACHE = {}


def _door_location_xyz(door):
    loc = getattr(door, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point
    if isinstance(loc, LocationCurve):
        c = loc.Curve
        try:
            p0 = c.GetEndPoint(0); p1 = c.GetEndPoint(1)
            return XYZ(0.5*(p0.X+p1.X), 0.5*(p0.Y+p1.Y), 0.5*(p0.Z+p1.Z))
        except:
            pass
    bb = None
    try:
        bb = door.get_BoundingBox(None)
    except:
        bb = None
    if bb:
        return XYZ(0.5*(bb.Min.X+bb.Max.X), 0.5*(bb.Min.Y+bb.Max.Y), 0.5*(bb.Min.Z+bb.Max.Z))
    return None


def _store_door_meta(door, orient, width):
    if door is None:
        return
    key = getattr(door, "UniqueId", None)
    if not key:
        return
    ovec = None
    try:
        if orient is not None:
            mag = (orient.X ** 2 + orient.Y ** 2) ** 0.5
            if mag > 1e-6:
                ovec = XYZ(orient.X / mag, orient.Y / mag, 0.0)
    except Exception:
        ovec = None
    w = None
    try:
        if width is not None:
            w = float(width)
    except Exception:
        w = None
    _DOOR_META_CACHE[key] = (ovec, w)


def door_wall_meta(door):
    key = getattr(door, "UniqueId", None)
    if not key:
        return (None, None)
    return _DOOR_META_CACHE.get(key, (None, None))


def _get_link_transform(inst):
    try:
        return inst.GetTotalTransform()
    except Exception:
        try:
            return inst.GetTransform()
        except Exception:
            return None


def _distance_to_curve(point, curve):
    if curve is None or point is None:
        return None
    try:
        res = curve.Project(point)
        if res:
            return res.Distance
    except Exception:
        pass
    try:
        return curve.Distance(point)
    except Exception:
        return None


def door_points_on_wall(doc, wall, include_linked=False, link_tolerance_ft=3.0, boundary_curve=None):
    doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
    out = []
    curve = None
    try:
        if wall is not None:
            loc = getattr(wall, "Location", None)
            if isinstance(loc, LocationCurve):
                curve = loc.Curve
    except Exception:
        curve = None
    if curve is None and boundary_curve is not None:
        curve = boundary_curve

    tol = 3.0
    try:
        tol = float(link_tolerance_ft if link_tolerance_ft is not None else 3.0)
    except Exception:
        tol = 3.0

    for d in doors:
        try:
            host = getattr(d, "Host", None)
            p = _door_location_xyz(d)
            if p is None:
                continue
            if wall is not None and host is not None and host.Id == wall.Id:
                if host is not None:
                    orient = getattr(host, "Orientation", None)
                    width = getattr(host, "Width", None)
                    _store_door_meta(d, orient, width)
                out.append((d, p))
            elif wall is None and curve is not None:
                dist = _distance_to_curve(p, curve)
                if dist is not None and dist <= tol:
                    if host is not None:
                        orient = getattr(host, "Orientation", None)
                        width = getattr(host, "Width", None)
                        _store_door_meta(d, orient, width)
                    out.append((d, p))
        except Exception:
            pass

    if include_linked and curve is not None and tol is not None:
        try:
            for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
                ldoc = inst.GetLinkDocument()
                if ldoc is None:
                    continue
                tf = _get_link_transform(inst)
                try:
                    linked_doors = FilteredElementCollector(ldoc) \
                        .OfCategory(BuiltInCategory.OST_Doors) \
                        .WhereElementIsNotElementType()
                except Exception:
                    linked_doors = []
                for d in linked_doors:
                    p = _door_location_xyz(d)
                    if p is None:
                        continue
                    try:
                        if tf is not None:
                            p = tf.OfPoint(p)
                    except Exception:
                        pass
                    dist = _distance_to_curve(p, curve)
                    if dist is not None and dist <= tol:
                        host = getattr(d, "Host", None)
                        orient = None
                        width = None
                        if host is not None:
                            try:
                                orient = getattr(host, "Orientation", None)
                                if orient is not None and tf is not None:
                                    orient = tf.OfVector(orient)
                            except Exception:
                                orient = getattr(host, "Orientation", None)
                            try:
                                width = getattr(host, "Width", None)
                            except Exception:
                                width = None
                        _store_door_meta(d, orient, width)
                        out.append((d, p))
        except Exception:
            pass
    return out


def _door_dynamic_radius_ft(door_elem, base_radius_ft, edge_margin_ft):
    """Use door width when available to widen the keepout: max(base, width/2 + edge_margin)."""
    r = float(base_radius_ft or 0.0)
    try:
        w = get_param_value(door_elem, "Width")
        if w is not None:
            r = max(r, float(w)/2.0 + float(edge_margin_ft or 0.0))
    except:
        pass
    return r


def filter_points_by_doors(points, door_tuples, base_radius_ft, edge_margin_ft):
    if not points or not door_tuples: return points
    out = []
    for p in points:
        deny = False
        for (door, dp) in door_tuples:
            r = _door_dynamic_radius_ft(door, base_radius_ft, edge_margin_ft)
            dx, dy, dz = p.X-dp.X, p.Y-dp.Y, p.Z-dp.Z
            if (dx*dx + dy*dy + dz*dz) <= (r*r):
                deny = True; break
        if not deny:
            out.append(p)
    return out
