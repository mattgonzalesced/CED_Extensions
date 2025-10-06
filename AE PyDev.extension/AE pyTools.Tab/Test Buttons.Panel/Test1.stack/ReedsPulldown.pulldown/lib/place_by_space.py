# -*- coding: utf-8 -*-
# lib/place_by_space.py
# Utilities to drive placement by MEP Spaces (works for Rooms too).
# IronPython + Revit API friendly (defensive, null-safe).

import clr
clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, XYZ, Transaction,
    HostObjectUtils, FamilyPlacementType, SpatialElementBoundaryOptions,
    Plane, SketchPlane, BuiltInParameter, ElementTransformUtils
)
from Autodesk.Revit.DB.Structure import StructuralType
from rules_loader import deep_merge

# --- type guard (works even if Mechanical.Space isn't importable in this rev) ---
try:
    from Autodesk.Revit.DB.Mechanical import Space as _MepSpace
except:
    _MepSpace = type("Dummy", (), {})  # harmless fallback


# --- family plan size / clearance -------------------------------------------

def _param_double_in_feet(p):
    """Return parameter as feet (Double) if possible; parse simple strings like '600 mm' or '24 in'."""
    try:
        st = p.StorageType.ToString()
    except:
        st = None
    if st == "Double":
        try: return float(p.AsDouble() or 0.0)
        except: return 0.0
    # simple string parse fallback
    try:
        s = (p.AsString() or "").strip().lower()
        if not s:
            return 0.0
        # pull first number
        num = ""
        for ch in s:
            if ch.isdigit() or ch in ".-":
                num += ch
            elif num:
                break
        val = float(num) if num else 0.0
        if "mm" in s:
            return val / 304.8  # mm -> ft
        if "in" in s or "\"" in s:
            return val / 12.0   # inches -> ft
        return val  # assume feet
    except:
        return 0.0

def get_symbol_plan_halfspan_ft(sym, rule=None):
    """
    Estimate half the max plan dimension of a fixture FamilySymbol (in feet).
    Tries common params; fallback to rule['plan_halfspan_ft'] or 0.75 ft.
    """
    if not sym:
        return float((rule or {}).get("plan_halfspan_ft", 0.75))

    names_diam = ("Trim Diameter", "Face Diameter", "Diameter", "Nominal Diameter")
    names_w    = ("Nominal Width", "Width", "W")
    names_l    = ("Nominal Length", "Length", "L")

    dia = 0.0
    for n in names_diam:
        try:
            p = sym.LookupParameter(n)
            if p:
                dia = max(dia, _param_double_in_feet(p))
        except:
            pass

    w = l = 0.0
    for n in names_w:
        try:
            p = sym.LookupParameter(n)
            if p:
                w = max(w, _param_double_in_feet(p))
        except:
            pass
    for n in names_l:
        try:
            p = sym.LookupParameter(n)
            if p:
                l = max(l, _param_double_in_feet(p))
        except:
            pass

    span = 0.0
    if dia > 0.0:
        span = dia
    elif max(w, l) > 0.0:
        span = max(w, l)
    else:
        return float((rule or {}).get("plan_halfspan_ft", 0.75))

    return 0.5 * span  # half the largest plan dimension



# -----------------------
# Space/Room basic helpers
# -----------------------

def strict_space_filter(doc, spatial, points, min_clear_ft=0.0, require_ceiling_host=False):
    """
    Keep only points that are inside the Space and at least min_clear_ft from the boundary.
    If require_ceiling_host=True, also drop points without a ceiling host.
    """
    pts = list(points or [])
    if not pts:
        return []

    zt = _spatial_test_z(doc, spatial)
    polys = _space_boundary_polygons(spatial)
    kept = []
    dropped = 0

    for p in pts:
        # 1) must be inside
        if not spatial_point_contains(spatial, XYZ(p.X, p.Y, zt)):
            dropped += 1
            continue

        # 2) must respect minimum clearance from boundary (if polygons known)
        if min_clear_ft > 0.0 and polys:
            dmin, _ = _closest_boundary_info(polys, p)
            if dmin < (min_clear_ft - 1e-6):
                dropped += 1
                continue

        # 3) must have ceiling host if required
        if require_ceiling_host:
            host, _ = _find_ceiling_under_point(doc, spatial, p)
            if host is None:
                dropped += 1
                continue

        kept.append(p)

    if dropped:
        print("[CLIP] strict: dropped {} point(s) (outside/too near boundary)".format(dropped))
    return kept

def _ray_intersect_seg_2d(r0x, r0y, nx, ny, ax, ay, bx, by, eps=1e-9):
    """Return ray distance t>=0 where (r0 + t*n) meets segment ab, or None if no hit."""
    vx = bx - ax
    vy = by - ay
    det = ny*vx - nx*vy
    if abs(det) < eps:
        return None  # parallel
    dx = ax - r0x
    dy = ay - r0y
    # Cramer's rule
    t = (-dx*vy + dy*vx) / det
    if t < 0.0:
        return None
    # u in [0,1] for segment
    u = (nx*dy - ny*dx) / det
    if u < -eps or u > 1.0 + eps:
        return None
    return t

def enforce_midspan_clearance(doc, spatial, points, near_threshold_ft=0.25,
                              require_ceiling_host=False, min_center_clear_ft=0.0):
    """
    For points near the boundary, move them to midpoint between nearest boundary and the
    opposing boundary along the inward normal. If half the gap is smaller than
    min_center_clear_ft (e.g., fixture halfspan + margin), drop the point.
    """
    pts = list(points or [])
    if not pts:
        return []

    polys = _space_boundary_polygons(spatial)
    if not polys:
        return pts

    z_test = _spatial_test_z(doc, spatial)
    kept = []
    nudged = 0
    dropped = 0
    fallback = 0

    for p in pts:
        dmin, info = _closest_boundary_info(polys, p)
        if info is None or dmin >= (near_threshold_ft - 1e-6):
            kept.append(p); continue

        a, b, _, px, py = info
        tx = b.X - a.X; ty = b.Y - a.Y
        seg_len = (tx*tx + ty*ty) ** 0.5 or 1.0
        tx /= seg_len; ty /= seg_len
        n1 = (-ty, tx); n2 = (ty, -tx)

        test1 = XYZ(px + n1[0]*0.05, py + n1[1]*0.05, 0.0)
        test2 = XYZ(px + n2[0]*0.05, py + n2[1]*0.05, 0.0)
        inside1 = spatial_point_contains(spatial, XYZ(test1.X, test1.Y, z_test))
        inside2 = spatial_point_contains(spatial, XYZ(test2.X, test2.Y, z_test))
        if inside1 and not inside2:
            nx, ny = n1
        elif inside2 and not inside1:
            nx, ny = n2
        elif inside1 and inside2:
            d1, _ = _closest_boundary_info(polys, test1)
            d2, _ = _closest_boundary_info(polys, test2)
            nx, ny = (n1 if d1 >= d2 else n2)
        else:
            kept.append(XYZ(p.X + 0.1, p.Y + 0.1, p.Z)); fallback += 1; continue

        r0x = px + nx*0.01; r0y = py + ny*0.01

        best_t = None
        for poly in polys:
            n = len(poly)
            if n < 2: continue
            for i in range(n):
                ax = poly[i].X; ay = poly[i].Y
                bx = poly[(i+1)%n].X; by = poly[(i+1)%n].Y
                # skip same edge if it’s the one we touched
                if (abs(ax - a.X) < 1e-6 and abs(ay - a.Y) < 1e-6 and
                    abs(bx - b.X) < 1e-6 and abs(by - b.Y) < 1e-6):
                    continue
                t_hit = _ray_intersect_seg_2d(r0x, r0y, nx, ny, ax, ay, bx, by)
                if t_hit is None or t_hit < 1e-3: continue
                if best_t is None or t_hit < best_t:
                    best_t = t_hit

        if best_t is None:
            cand = XYZ(p.X + nx*0.5, p.Y + ny*0.5, p.Z)
            if spatial_point_contains(spatial, XYZ(cand.X, cand.Y, z_test)):
                if not require_ceiling_host or (_find_ceiling_under_point(doc, spatial, cand)[0] is not None):
                    kept.append(cand); fallback += 1; continue
            kept.append(p); fallback += 1; continue

        # ensure half-gap >= min_center_clear_ft (size-aware clearance)
        half_gap = 0.5 * best_t
        if min_center_clear_ft > 0.0 and half_gap < (min_center_clear_ft - 1e-6):
            dropped += 1
            continue

        mid = XYZ(r0x + half_gap*nx, r0y + half_gap*ny, p.Z)
        if not spatial_point_contains(spatial, XYZ(mid.X, mid.Y, z_test)):
            mid = XYZ(r0x + 0.45*best_t*nx, r0y + 0.45*best_t*ny, p.Z)
        if require_ceiling_host:
            host, _ = _find_ceiling_under_point(doc, spatial, mid)
            if host is None:
                mid = XYZ(r0x + 0.4*best_t*nx, r0y + 0.4*best_t*ny, p.Z)

        kept.append(mid); nudged += 1

    if (nudged + dropped + fallback) > 0:
        print("[EDGE] midspan: adjusted {}, dropped {}, fallback {}".format(nudged, dropped, fallback))
    return kept

def _dist_point_to_seg_2d(p, a, b):
    # p, a, b are XYZ (z ignored)
    vx, vy = (b.X - a.X), (b.Y - a.Y)
    wx, wy = (p.X - a.X), (p.Y - a.Y)
    seg2 = vx*vx + vy*vy
    t = 0.0 if seg2 <= 1e-12 else max(0.0, min(1.0, (wx*vx + wy*vy) / seg2))
    px, py = (a.X + t*vx), (a.Y + t*vy)
    dx, dy = (p.X - px), (p.Y - py)
    d = (dx*dx + dy*dy) ** 0.5
    return d, (a, b, t, px, py)

def _closest_boundary_info(polys, p):
    """Return (min_dist, (a,b,tx,px,py)) for the closest segment to point p."""
    best = (1e30, None)
    for poly in polys:
        n = len(poly)
        if n < 2:
            continue
        for i in range(n):
            a = poly[i]
            b = poly[(i + 1) % n]
            d, info = _dist_point_to_seg_2d(p, a, b)
            if d < best[0]:
                best = (d, info)
    return best

def enforce_boundary_clearance(doc, spatial, points, clearance_ft=2.0, require_ceiling_host=False):
    """
    Ensure every point is at least `clearance_ft` inside from the space boundary.
    If a point is too close, nudge it inward along the closest-segment normal.
    If nudging can't keep it inside, drop the point.
    """
    pts = list(points or [])
    if not pts:
        return []

    # get polygons (we'll rely on true loops if available)
    polys = _space_boundary_polygons(spatial)
    if not polys:
        return pts

    z_test = _spatial_test_z(doc, spatial)
    kept, nudged, dropped = [], 0, 0

    for p in pts:
        dmin, info = _closest_boundary_info(polys, p)
        if info is None:
            kept.append(p)
            continue

        if dmin >= (clearance_ft - 1e-6):
            # already safely inside
            kept.append(p)
            continue

        # Need to push inward by at least (clearance - dmin) + tiny epsilon
        a, b, t, px, py = info
        tx, ty = (b.X - a.X), (b.Y - a.Y)
        seg_len = (tx*tx + ty*ty) ** 0.5 or 1.0
        tx, ty = tx/seg_len, ty/seg_len

        # two candidate normals (perpendicular to segment)
        n1 = (-ty, tx)
        n2 = (ty, -tx)
        delta = (clearance_ft - dmin) + 0.05  # small buffer

        # try both directions and pick the one that stays inside
        candidates = [
            XYZ(p.X + n1[0]*delta, p.Y + n1[1]*delta, p.Z),
            XYZ(p.X + n2[0]*delta, p.Y + n2[1]*delta, p.Z),
        ]

        chosen = None
        for cand in candidates:
            # must be inside the space
            if not spatial_point_contains(spatial, XYZ(cand.X, cand.Y, z_test)):
                continue
            # recheck distance after move (avoid borderline ~epsilon)
            d2, _ = _closest_boundary_info(polys, cand)
            if d2 + 1e-6 >= clearance_ft:
                if require_ceiling_host:
                    host, _ = _find_ceiling_under_point(doc, spatial, cand)
                    if host is None:
                        continue
                chosen = cand
                break

        if chosen is not None:
            kept.append(chosen)
            nudged += 1
        else:
            # couldn't safely nudge; drop it
            dropped += 1

    if (nudged + dropped) > 0:
        print("[EDGE] clearance {:.2f}ft → nudged {}, dropped {}".format(clearance_ft, nudged, dropped))
    return kept


def is_space(spatial):
    try:
        return isinstance(spatial, _MepSpace)
    except:
        return False

def spatial_display_name(spatial):
    nm = ""; num = ""
    # try common params first
    for pname in ("Name", "ROOM_NAME", "SPACE_NAME"):
        try:
            p = spatial.LookupParameter(pname)
            if p:
                nm = (p.AsString() or nm)
        except:
            pass
    for pname in ("Number", "ROOM_NUMBER", "SPACE_NUMBER"):
        try:
            p = spatial.LookupParameter(pname)
            if p:
                num = (p.AsString() or num)
        except:
            pass
    return (nm or "Unnamed"), (num or "")

def get_space_level(doc, spatial):
    try:
        return doc.GetElement(spatial.LevelId)
    except:
        return None

# alias (lets existing code that calls get_room_level keep working)
get_room_level = get_space_level

def _spatial_test_z(doc, spatial):
    """Z height used for point-in-space tests (~3 ft above level)."""
    lvl = get_space_level(doc, spatial)
    base = (lvl.Elevation if lvl else 0.0)
    return base + 3.0

def spatial_point_contains(spatial, pt):
    """True if pt lies inside the Space/Room footprint."""
    # Rooms: IsPointInRoom; Spaces: IsPointInSpace
    try:
        return spatial.IsPointInRoom(pt)
    except:
        try:
            return spatial.IsPointInSpace(pt)
        except:
            return False

def get_target_spaces(doc, uidoc=None, view=None, only_current_level=False, prefer_selection=True):
    """Collect target MEP Spaces (honors selection & current-level if requested)."""
    # prefer selection (if any Space is selected)
    try:
        if prefer_selection and uidoc and uidoc.Selection:
            ids = list(uidoc.Selection.GetElementIds())
            if ids:
                chosen = []
                for eid in ids:
                    el = doc.GetElement(eid)
                    if el and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces):
                        chosen.append(el)
                if chosen:
                    return chosen
    except:
        pass

    col = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_MEPSpaces)\
        .WhereElementIsNotElementType()
    spaces = [s for s in col if getattr(s, "Area", 0.0) > 1e-6]

    if only_current_level and view is not None:
        try:
            gl = getattr(view, "GenLevel", None)
            vid = (gl.Id if gl else view.LevelId)
            spaces = [s for s in spaces if getattr(s, "LevelId", None) == vid]
        except:
            pass

    return spaces

# -----------------------
# Ceilings / hosting utils
# -----------------------

def _ceiling_underside_elev_ft(doc, ceiling):
    elev = 0.0
    try:
        lvl = doc.GetElement(ceiling.LevelId)
        elev = (lvl.Elevation if lvl else 0.0)
        p = ceiling.LookupParameter("Height Offset From Level")
        if p:
            elev += (p.AsDouble() or 0.0)
    except:
        pass
    return elev

def _find_ceiling_under_point(doc, spatial, pt_xy):
    """Find a ceiling 'over' this XY inside the same level. Returns (host, underside_z_ft) or (None, None)."""
    lvl = get_space_level(doc, spatial)
    cands = []
    for c in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType():
        try:
            if lvl and c.LevelId != lvl.Id:
                continue
        except:
            pass
        bb = c.get_BoundingBox(doc.ActiveView) or c.get_BoundingBox(None)
        if not bb:
            continue
        if (bb.Min.X - 1e-3) <= pt_xy.X <= (bb.Max.X + 1e-3) and (bb.Min.Y - 1e-3) <= pt_xy.Y <= (bb.Max.Y + 1e-3):
            cands.append(c)
    if not cands:
        return None, None
    cands.sort(key=lambda cc: _ceiling_underside_elev_ft(doc, cc))
    host = cands[-1]
    return host, _ceiling_underside_elev_ft(doc, host)

def _ensure_sketchplane_at_z(doc, view, z_ft):
    """Create a horizontal SketchPlane at Z and try to make the view use it."""
    plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, float(z_ft)))
    sp = SketchPlane.Create(doc, plane)
    try:
        view.SketchPlane = sp
        try:
            view.ShowActiveWorkPlane = True
        except:
            pass
    except:
        pass
    return sp

def _set_instance_offset_from_level(inst, offset_ft):
    """Try common params that control elevation/offset above level/workplane."""
    for bip in (
        BuiltInParameter.INSTANCE_ELEVATION_PARAM,
        BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
        BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
    ):
        try:
            p = inst.get_Parameter(bip)
            if p and p.StorageType.ToString() == "Double":
                p.Set(float(offset_ft))
                return True
        except:
            pass
    for pname in ("Elevation", "Offset", "Offset from Level", "Height Offset From Level"):
        try:
            p = inst.LookupParameter(pname)
            if p and p.StorageType.ToString() == "Double":
                p.Set(float(offset_ft))
                return True
        except:
            pass
    return False

def _raise_instance_to_z(doc, inst, z_target):
    """Last-resort nudge: move element in Z to desired height."""
    try:
        loc = inst.Location
        if hasattr(loc, "Point") and loc.Point:
            dz = float(z_target) - float(loc.Point.Z)
            if abs(dz) > 1e-3:
                ElementTransformUtils.MoveElement(doc, inst.Id, XYZ(0, 0, dz))
                return True
    except:
        pass
    return False

# -----------------------
# Grid planning (rule-driven)
# -----------------------

def _extents_from_spatial(spatial, active_view):
    # Try view bbox
    bb = spatial.get_BoundingBox(active_view)
    if bb:
        return bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y
    # Fallback: boundary loops extents
    try:
        opts = SpatialElementBoundaryOptions()
        loops = spatial.GetBoundarySegments(opts)
        xs, ys = [], []
        for loop in loops or []:
            for seg in loop:
                c = seg.GetCurve()
                p, q = c.GetEndPoint(0), c.GetEndPoint(1)
                xs.extend([p.X, q.X]); ys.extend([p.Y, q.Y])
        if xs and ys:
            return min(xs), min(ys), max(xs), max(ys)
    except:
        pass
    # Last resort: single point
    try:
        lp = spatial.Location.Point
        return lp.X, lp.Y, lp.X, lp.Y
    except:
        return None

def propose_grid_points_from_rule(spatial, active_view, rule):
    """Strict, JSON-driven grid: spacing_ft.target is the step; offset_ft trims edges; grid_align = 'center'|'edge'."""
    try:
        spacing_cfg = (rule.get("spacing_ft") or {})
        target = float(spacing_cfg.get("target", 10.0))
        offset = float(rule.get("offset_ft", 2.0))
        align  = (rule.get("grid_align") or "center").lower()
    except:
        return []

    ex = _extents_from_spatial(spatial, active_view)
    if not ex or target <= 0.0:
        return []
    minx, miny, maxx, maxy = ex
    minx += offset; miny += offset; maxx -= offset; maxy -= offset
    if maxx <= minx or maxy <= miny:
        cx, cy = (minx + maxx) * 0.5, (miny + maxy) * 0.5
        return [XYZ(cx, cy, 0.0)]

    def _start(a, b):
        span = b - a
        if align == "edge":
            return a
        return a + min(target * 0.5, span * 0.5)

    xs, ys = [], []
    x = _start(minx, maxx)
    while x <= maxx + 1e-6:
        xs.append(x); x += target
    y = _start(miny, maxy)
    while y <= maxy + 1e-6:
        ys.append(y); y += target

    pts = [XYZ(x, y, 0.0) for x in xs for y in ys]
    if not pts:
        cx, cy = (minx + maxx) * 0.5, (miny + maxy) * 0.5
        pts = [XYZ(cx, cy, 0.0)]
    return pts


def _space_boundary_polygons(spatial):
    """Return list of boundary loops as lists of XYZ (plan view), or [] if none."""
    try:
        opts = SpatialElementBoundaryOptions()
        loops = spatial.GetBoundarySegments(opts)
        polys = []
        for loop in loops or []:
            poly = []
            for seg in loop:
                c = seg.GetCurve()
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                if not poly:
                    poly.append(XYZ(p0.X, p0.Y, 0.0))
                poly.append(XYZ(p1.X, p1.Y, 0.0))
            if len(poly) >= 3:
                polys.append(poly)
        return polys
    except:
        return []

def _pip(point, poly):
    """Point-in-polygon (2D) using ray-cast; poly is list of XYZ in order."""
    x, y = point.X, point.Y
    inside = False
    n = len(poly)
    j = n - 1
    for i in range(n):
        xi, yi = poly[i].X, poly[i].Y
        xj, yj = poly[j].X, poly[j].Y
        cross = ((yi > y) != (yj > y)) and \
                (x < (xj - xi) * (y - yi) / ((yj - yi) or 1e-9) + xi)
        if cross:
            inside = not inside
        j = i
    return inside

def clip_points_to_space(doc, spatial, points, require_ceiling_host=False):
    """
    Keep only points inside the Space (not just bbox). Fallbacks:
      1) IsPointInSpace at safe Z
      2) boundary polygons test
      3) bbox if 1&2 unavailable
    """
    pts = list(points or [])
    if not pts:
        return []

    # 1) Try IsPointInSpace / IsPointInRoom (best)
    z_test = _spatial_test_z(doc, spatial)
    try:
        contained = [p for p in pts if spatial.IsPointInSpace(XYZ(p.X, p.Y, z_test))]
    except:
        try:
            contained = [p for p in pts if spatial.IsPointInRoom(XYZ(p.X, p.Y, z_test))]
        except:
            contained = None

    # 2) If IsPointInSpace failed or returned nothing, try boundary loops
    if contained is None or len(contained) == 0:
        polys = _space_boundary_polygons(spatial)
        if polys:
            contained = [p for p in pts if any(_pip(p, poly) for poly in polys)]
        else:
            contained = None

    # 3) Last resort: keep points within Space bbox (prevents 0 placements)
    if contained is None:
        bb = spatial.get_BoundingBox(doc.ActiveView)
        if bb:
            contained = [p for p in pts
                         if (bb.Min.X - 1e-6) <= p.X <= (bb.Max.X + 1e-6)
                         and (bb.Min.Y - 1e-6) <= p.Y <= (bb.Max.Y + 1e-6)]
        else:
            contained = []

    # Optional: require a ceiling host under each point
    if require_ceiling_host:
        kept = []
        for p in contained:
            host, _ = _find_ceiling_under_point(doc, spatial, p)
            if host is not None:
                kept.append(p)
        contained = kept

    return contained

# -----------------------
# Ceiling-aware instance placement
# -----------------------

NEG_Z = XYZ(0, 0, -1)

def place_light_instance(doc, sym, spatial, pt_plan, rule):
    """Prefer ceiling host; fallback to workplane at ceiling Z; else place at level + set offset/move."""
    # activate
    try:
        if sym and not sym.IsActive:
            sym.Activate()
    except:
        pass

    # FPT
    fpt = None
    try:
        fam = getattr(sym, "Family", None)
        if fam:
            fpt = fam.FamilyPlacementType
    except:
        fpt = None

    # target Z
    host, z_underside = _find_ceiling_under_point(doc, spatial, pt_plan)
    lvl = get_space_level(doc, spatial)
    lvl_elev = (lvl.Elevation if lvl else 0.0)
    mount_elev_ft = float(rule.get("mount_elev_ft", 9.0))
    z_target = (z_underside if z_underside is not None else (lvl_elev + mount_elev_ft))

    # 1) CeilingBased
    try:
        if fpt == FamilyPlacementType.CeilingBased and host is not None:
            return doc.Create.NewFamilyInstance(pt_plan, sym, host, StructuralType.NonStructural)
    except:
        pass

    # 2) FaceBased (host to underside)
    try:
        if fpt is not None and str(fpt).endswith("FaceBased") and host is not None:
            refs = HostObjectUtils.GetBottomFaces(host)
            if refs and len(refs) > 0:
                try:
                    return doc.Create.NewFamilyInstance(refs[0], pt_plan, NEG_Z, sym)
                except:
                    return doc.Create.NewFamilyInstance(refs[0], pt_plan, XYZ.BasisZ, sym)
    except:
        pass

    # 3) WorkPlaneBased: create SketchPlane at ceiling Z and place there; verify Z
    try:
        if fpt == FamilyPlacementType.WorkPlaneBased:
            sp = _ensure_sketchplane_at_z(doc, doc.ActiveView, z_target)
            doc.Regenerate()
            inst = None
            try:
                inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, 0.0), sym, sp, StructuralType.NonStructural)
            except:
                try:
                    inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, z_target), sym, doc.ActiveView)
                except:
                    inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, z_target), sym, sp)
            if inst:
                doc.Regenerate()
                if not _raise_instance_to_z(doc, inst, z_target):
                    _set_instance_offset_from_level(inst, z_target - lvl_elev)
                return inst
    except:
        pass

    # 4) Level/Other: place at level then set offset/move
    try:
        inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, lvl_elev), sym, lvl, StructuralType.NonStructural)
        if not _set_instance_offset_from_level(inst, z_target - lvl_elev):
            _raise_instance_to_z(doc, inst, z_target)
        return inst
    except:
        return None

# -----------------------
# Turn-key Space placer for lights
# -----------------------

def place_fixtures_in_space(doc, active_view, spatial, rule, symbol_picker, dry_run=True, verbose=True):
    """
    Place fixtures in a Space using a JSON rule.
    - symbol_picker: callable(doc, rule) -> FamilySymbol
    """
    name, number = spatial_display_name(spatial)

    # Let first candidate override category fields (spacing_ft, offset_ft, mount_elev_ft, etc.)
    cands = rule.get('fixture_candidates') or []
    eff = deep_merge(rule, cands[0]) if cands else dict(rule)

    # plan grid strictly from JSON
    pts = propose_grid_points_from_rule(spatial, active_view, eff)
    # basic inside clip
    pts = clip_points_to_space(doc, spatial, pts, require_ceiling_host=bool(eff.get("require_ceiling_host")))

    # pick symbol up-front (we need its size for clearance)
    sym = symbol_picker(doc, eff) if symbol_picker else None
    if not sym:
        if verbose: print("[WARN] No matching FamilySymbol for rule; skipping Space {}".format(number))
        return len(pts), []

    # compute size-aware required clearance
    halfspan = get_symbol_plan_halfspan_ft(sym, eff)
    margin = float(eff.get("edge_clear_margin_ft", 0.0))
    min_clear = max(float(eff.get("min_edge_clear_ft", 0.01)), halfspan + margin)
    if verbose:
        try:
            print("[DBG] footprint halfspan≈{:.2f} ft | min_clear≈{:.2f} ft".format(halfspan, min_clear))
        except:
            pass

    # edge policy: "drop" | "midspan" | "offset"
    edge_policy = (eff.get("edge_policy") or "drop").lower()

    # optional transforms BEFORE strict drop
    if edge_policy == "midspan":
        near_thr = float(eff.get("edge_near_threshold_ft", 0.25))  # ~3 in detection
        pts = enforce_midspan_clearance(doc, spatial, pts,
                                        near_threshold_ft=near_thr,
                                        require_ceiling_host=bool(eff.get("require_ceiling_host")),
                                        min_center_clear_ft=min_clear)
    elif edge_policy == "offset":
        # keep if you still want fixed-offset mode elsewhere; this doesn't use size, strict filter will.
        edge_clear = float(eff.get("edge_clear_ft", min_clear))
        pts = enforce_boundary_clearance(doc, spatial, pts,
                                         clearance_ft=edge_clear,
                                         require_ceiling_host=bool(eff.get("require_ceiling_host")))

    # FINAL strict guard: drop anything outside or violating size-aware clearance
    pts = strict_space_filter(doc, spatial, pts, min_clear_ft=min_clear,
                              require_ceiling_host=bool(eff.get("require_ceiling_host")))

    if verbose:
        print("[PLAN] {} {}: {} pts after clipping".format(name, number, len(pts)))
    if dry_run or not pts:
        return len(pts), []

    # Optional replace
    if eff.get("replace_existing"):
        victims, zt = [], _spatial_test_z(doc, spatial)
        bic = BuiltInCategory.OST_LightingFixtures  # change to .OST_Sprinklers for sprinkler runs
        for inst in FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType():
            lp = getattr(getattr(inst, "Location", None), "Point", None)
            if lp and spatial_point_contains(spatial, XYZ(lp.X, lp.Y, zt)):
                victims.append(inst)
        if victims:
            tdel = Transaction(doc, "Clear Space Fixtures")
            tdel.Start()
            try:
                for v in victims: doc.Delete(v.Id)
                tdel.Commit()
                if verbose: print("[PLACE] Removed {} existing instance(s)".format(len(victims)))
            except:
                tdel.RollBack()

    # Place
    placed = []
    t = Transaction(doc, "Place Fixtures (Space)")
    t.Start()

    zt = _spatial_test_z(doc, spatial)
    polys_cache = _space_boundary_polygons(spatial)
    min_clear = float(eff.get("min_edge_clear_ft", 0.01))

    try:
        for p in pts:
            if not spatial_point_contains(spatial, XYZ(p.X, p.Y, zt)):
                if verbose: print("[SKIP] outside after adjustments"); continue
            if polys_cache:
                dmin, _ = _closest_boundary_info(polys_cache, p)
                if dmin < (min_clear - 1e-6):
                    if verbose: print("[SKIP] violates clearance (d={:.3f}ft < {:.3f}ft)".format(dmin, min_clear))
                    continue
            inst = place_light_instance(doc, sym, spatial, p, eff)
            if inst: placed.append(inst)
        t.Commit()
        if verbose:
            print("[PLACE] Placed {} instances in {} {}".format(len(placed), name, number))
    except Exception as ex:
        if verbose:
            print("[ERROR] Placement failed in {} {}: {}".format(name, number, ex))
        try: t.RollBack()
        except: pass
    return len(pts), placed
