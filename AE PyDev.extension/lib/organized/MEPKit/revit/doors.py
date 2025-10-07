# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/doors.py
from __future__ import absolute_import
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, LocationPoint, LocationCurve, XYZ
)
from organized.MEPKit.revit.params import get_param_value

def door_points_on_wall(doc, wall):
    doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
    out = []
    for d in doors:
        try:
            if getattr(d, "Host", None) and d.Host.Id == wall.Id:
                loc = getattr(d, "Location", None)
                if isinstance(loc, LocationPoint):
                    out.append((d, loc.Point))
                elif isinstance(loc, LocationCurve):
                    c = loc.Curve
                    p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                    out.append((d, XYZ(0.5*(p0.X+p1.X), 0.5*(p0.Y+p1.Y), 0.5*(p0.Z+p1.Z))))
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