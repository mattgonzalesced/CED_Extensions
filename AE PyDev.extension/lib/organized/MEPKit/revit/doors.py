# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/doors.py
from __future__ import absolute_import
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, LocationPoint, LocationCurve,
    XYZ, RevitLinkInstance
)
from organized.MEPKit.revit.params import get_param_value

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

def _get_link_transform(inst):
    try:
        return inst.GetTotalTransform()
    except Exception:
        try:
            return inst.GetTransform()
        except Exception:
            return None

def door_points_on_wall(doc, wall, include_linked=False, link_tolerance_ft=3.0):
    doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
    out = []
    curve = None
    try:
        loc = getattr(wall, "Location", None)
        if isinstance(loc, LocationCurve):
            curve = loc.Curve
    except:
        curve = None
    for d in doors:
        try:
            if getattr(d, "Host", None) and d.Host.Id == wall.Id:
                p = _door_location_xyz(d)
                if p:
                    out.append((d, p))
        except:
            pass
    if include_linked and curve is not None and link_tolerance_ft is not None:
        try:
            tol = float(link_tolerance_ft)
        except:
            tol = 3.0
        try:
            for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
                ldoc = inst.GetLinkDocument()
                if ldoc is None:
                    continue
                tf = _get_link_transform(inst)
                try:
                    linked_doors = FilteredElementCollector(ldoc)\
                        .OfCategory(BuiltInCategory.OST_Doors)\
                        .WhereElementIsNotElementType()
                except:
                    linked_doors = []
                for d in linked_doors:
                    p = _door_location_xyz(d)
                    if p is None:
                        continue
                    try:
                        if tf is not None:
                            p = tf.OfPoint(p)
                    except:
                        pass
                    try:
                        proj = curve.Project(p)
                        if proj is None:
                            continue
                        if proj.Distance <= tol:
                            out.append((d, p))
                    except:
                        # fallback: simple XY distance to segment endpoints
                        try:
                            dist_vec = curve.Distance(p)
                            if dist_vec <= tol:
                                out.append((d, p))
                        except:
                            pass
        except:
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
