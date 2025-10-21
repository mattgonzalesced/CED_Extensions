# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/spaces.py
from __future__ import absolute_import
import math
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, SpatialElementBoundaryOptions,
    LocationPoint, LocationCurve, XYZ, Line, Wall
)
from organized.MEPKit.revit.params import get_param_value

def space_match_text(space):
    """Return a robust text blob for categorization: Name, Space/Room Name, Number, Dept, Function, Occupancy."""
    parts = []
    # direct props
    for attr in ("Name", "Number"):
        try:
            val = getattr(space, attr, None)
            if val: parts.append(val)
        except: pass
    # common parameters found on Spaces/Rooms
    for pname in ("Name", "Space Name", "Room Name", "Number", "Department", "Function", "Occupancy"):
        try:
            v = get_param_value(space, pname)
            if v: parts.append(str(v))
        except: pass
    # collapse to a single string
    blob = u" ".join([p for p in parts if p]).strip()
    return blob or u""

def collect_spaces_or_rooms(doc):
    spaces = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType())
    if spaces: return spaces
    return list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())

def space_name(space):
    return getattr(space, "Name", "") or ""

def boundary_loops(space):
    try:
        return space.GetBoundarySegments(SpatialElementBoundaryOptions()) or []
    except:
        return []

def segment_curve(seg):
    try: return seg.GetCurve()
    except: return None

def segment_host_wall(doc, seg):
    try:
        eid = seg.ElementId
        if eid and eid.IntegerValue > 0:
            el = doc.GetElement(eid)
            if isinstance(el, Wall): return el
    except:
        pass
    return None

def _seg_len(curve):
    if not curve: return 0.0
    try: return curve.ApproximateLength
    except:
        try: return curve.Length
        except: return 0.0

def _point_along(curve, d):
    L = _seg_len(curve)
    if L <= 1e-9: return curve.GetEndPoint(0)
    t = max(0.0, min(1.0, d/L))
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    return XYZ(p0.X + (p1.X-p0.X)*t, p0.Y + (p1.Y-p0.Y)*t, p0.Z + (p1.Z-p0.Z)*t)

def _inward_offset_xy(curve, mag=0.05):
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    vx, vy = (p1.X-p0.X), (p1.Y-p0.Y)
    L = math.hypot(vx, vy) or 1.0
    # rotate 90° in XY (approx inward); sign is heuristic
    nx, ny = -vy/L, vx/L
    return XYZ(nx*mag, ny*mag, 0.0)

def sample_points_on_segment(curve, first_ft, next_ft, corner_margin_ft, inset_ft):
    """Return list[XYZ] along a boundary segment, trimmed by corner margins, with inward inset."""
    L = _seg_len(curve)
    usable = L - 2.0*corner_margin_ft
    if usable <= 0.5: return []
    pts, d = [], corner_margin_ft + float(first_ft)
    inward = _inward_offset_xy(curve, inset_ft)
    end_limit = L - corner_margin_ft + 1e-6
    while d <= end_limit:
        p = _point_along(curve, d)
        p = XYZ(p.X + inward.X, p.Y + inward.Y, p.Z + inward.Z)
        pts.append(p)
        d += float(next_ft)
    return pts