# -*- coding: utf-8 -*-
# lib/geometry.py
import clr, math
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import XYZ


def _room_extents_xy(room, active_view):
    # Try view bbox first
    bb = room.get_BoundingBox(active_view)
    if bb:
        return bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y
    # Fallback: boundary loops extents
    opts = SpatialElementBoundaryOptions()
    loops = room.GetBoundarySegments(opts)
    if loops:
        xs, ys = [], []
        for loop in loops:
            for seg in loop:
                c = seg.GetCurve()
                try:
                    p, q = c.GetEndPoint(0), c.GetEndPoint(1)
                    xs.extend([p.X, q.X]); ys.extend([p.Y, q.Y])
                except: pass
        if xs and ys:
            return min(xs), min(ys), max(xs), max(ys)
    # Last resort: single point at room location
    try:
        lp = room.Location.Point
        return lp.X, lp.Y, lp.X, lp.Y
    except:
        return None

def propose_grid_points_from_rule(room, active_view, rule):
    try:
        spacing_cfg = (rule.get("spacing_ft") or {})
        target = float(spacing_cfg.get("target", 10.0))
        offset = float(rule.get("offset_ft", 2.0))
        align  = (rule.get("grid_align") or "center").lower()
    except Exception:
        return []

    ex = _room_extents_xy(room, active_view)
    if not ex or target <= 0:
        return []
    minx, miny, maxx, maxy = ex
    minx += offset; miny += offset; maxx -= offset; maxy -= offset
    if maxx <= minx or maxy <= miny:
        # ensure at least one point at center
        cx, cy = (minx+maxx)/2.0, (miny+maxy)/2.0
        return [XYZ(cx, cy, 0.0)]

    def _start(a, b):
        span = b - a
        return a if align == "edge" else a + min(target*0.5, span*0.5)

    xs, ys = [], []
    x = _start(minx, maxx)
    while x <= maxx + 1e-6:
        xs.append(x); x += target
    y = _start(miny, maxy)
    while y <= maxy + 1e-6:
        ys.append(y); y += target

    pts = [XYZ(x, y, 0.0) for x in xs for y in ys]
    if not pts:
        cx, cy = (minx+maxx)/2.0, (miny+maxy)/2.0
        pts = [XYZ(cx, cy, 0.0)]
    return pts