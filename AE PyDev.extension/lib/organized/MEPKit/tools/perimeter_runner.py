# -*- coding: utf-8 -*-
# lib/organized/MEPKit/electrical/perimeter_runner.py
from __future__ import absolute_import
from organized.MEPKit.core.log import get_logger
from organized.MEPKit.core.rules import (
    load_identify_rules, load_branch_rules, normalize_constraints,
    categorize_space_by_name, get_category_rule
)
from organized.MEPKit.revit.transactions import RunInTransaction
from organized.MEPKit.revit.spaces import collect_spaces_or_rooms, space_name, boundary_loops, segment_curve, segment_host_wall, sample_points_on_segment, space_match_text
from organized.MEPKit.revit.doors import door_points_on_wall, filter_points_by_doors
from organized.MEPKit.revit.placement import place_hosted, place_free
from organized.MEPKit.revit.symbols import resolve_or_load_symbol
from organized.MEPKit.revit.params import set_param_value  # optional for mounting height

from Autodesk.Revit.DB import (
    RevitLinkInstance, Opening, BuiltInCategory, FilteredElementCollector, XYZ, Wall, LocationCurve
)
from System import Double
from Autodesk.Revit.DB import HostObjectUtils, ShellLayerType, PlanarFace, Plane


# ---------Helpers to place recepts in small spaces-----------
def _curve_len(curve):
    try: return float(curve.ApproximateLength)
    except:
        try: return float(curve.Length)
        except: return 0.0

def _curve_point_at(curve, t01):
    # parameterized by curve parameter space (0..1 using IsProportion=True)
    try: return curve.Evaluate(float(t01), True)
    except: return None

def _midpoint(curve):
    return _curve_point_at(curve, 0.5)

def _two_points_on_one_curve(curve):
    # fallback for a single very short wall: place at ~1/3 and ~2/3
    return _curve_point_at(curve, 0.33), _curve_point_at(curve, 0.66)

def _unique_by_xy(pts, eps=1e-4):
    seen = set(); out = []
    for p in pts:
        k = (round(p.X/eps), round(p.Y/eps))
        if k in seen: continue
        seen.add(k); out.append(p)
    return out

def _refilter_with(base_constraints, overrides, wall_segments, doc, first_ft, next_ft, inset_ft, logger=None):
    """Re-sample all segments using the (possibly) relaxed constraints."""
    c = dict(base_constraints or {})
    c.update(overrides or {})
    avoid_corners_ft      = float(c.get('avoid_corners_ft', 0.0))
    avoid_doors_radius_ft = float(c.get('avoid_doors_radius_ft', 0.0))
    door_edge_margin_ft   = float(c.get('door_edge_margin_ft', 0.0))

    re_pts = []
    for curve, wall in wall_segments:
        if curve is None:
            continue
        pts = sample_points_on_segment(curve, first_ft, next_ft, avoid_corners_ft, inset_ft)
        if wall and (avoid_doors_radius_ft > 0.0 or door_edge_margin_ft > 0.0):
            doors = door_points_on_wall(doc, wall)
            pts = filter_points_by_doors(pts, doors, avoid_doors_radius_ft, door_edge_margin_ft)
        re_pts.extend(pts)
    return _unique_by_xy(re_pts)

def _host_for_point(p, wall_segments, tol=1e-3):
    # try to find the segment whose curve this point lies on
    for curve, wall in wall_segments:
        if curve is None:
            continue
        try:
            pr = curve.Project(p)
            if pr and pr.Distance <= tol:
                return wall
        except:
            pass
    # fallback: first wall if any
    return wall_segments[0][1] if wall_segments else None

def _xy_key(p, eps=1e-4):
    return (round(p.X/eps), round(p.Y/eps))

#---------------Wall collectors------------------

def _collect_walls(doc):
    try:
        return list(FilteredElementCollector(doc).OfClass(Wall))
    except:
        return []

def _nearest_wall_xy_distance(point_xyz, walls):
    """Return horizontal (XY) distance in feet from point to nearest wall location curve."""
    min_d = 1e30
    for w in walls:
        try:
            lc = w.Location
            if lc is None:
                continue
            crv = getattr(lc, "Curve", None)
            if crv is None:
                continue
            res = crv.Project(point_xyz)
            if res is None:
                continue
            # IntersectionResult has XYZPoint; fall back to Point if needed
            q = res.XYZPoint if hasattr(res, "XYZPoint") else getattr(res, "Point", None)
            if q is None:
                continue
            dx, dy = (point_xyz.X - q.X), (point_xyz.Y - q.Y)
            d = (dx*dx + dy*dy) ** 0.5
            if d < min_d:
                min_d = d
        except:
            pass
    return min_d

#-----------------Linked wall collectors-------------------

def _get_link_transform(link_inst):
    try:
        return link_inst.GetTotalTransform()  # newer API
    except:
        try:
            return link_inst.GetTransform()    # older API
        except:
            return None

def _collect_wall_curves_host(doc):
    curves = []
    try:
        for w in FilteredElementCollector(doc).OfClass(Wall):
            lc = w.Location
            if isinstance(lc, LocationCurve):
                crv = lc.Curve
                if crv: curves.append(crv)
    except:
        pass
    return curves

def _collect_wall_curves_linked(doc):
    curves = []
    try:
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            linkdoc = inst.GetLinkDocument()
            if linkdoc is None:
                continue
            tf = _get_link_transform(inst)
            for w in FilteredElementCollector(linkdoc).OfClass(Wall):
                lc = w.Location
                if isinstance(lc, LocationCurve):
                    crv = lc.Curve
                    if crv:
                        try:
                            if tf is not None:
                                crv = crv.CreateTransformed(tf)  # into host coords
                        except:
                            pass
                        curves.append(crv)
    except:
        pass
    return curves

def _nearest_wall_xy_distance(point_xyz, wall_curves):
    """Horizontal distance (ft) to nearest wall curve in host coords."""
    min_d = 1e30
    for crv in wall_curves:
        try:
            res = crv.Project(point_xyz)
            if res is None:
                continue
            q = res.XYZPoint if hasattr(res, "XYZPoint") else getattr(res, "Point", None)
            if q is None:
                continue
            dx, dy = (point_xyz.X - q.X), (point_xyz.Y - q.Y)
            d = (dx*dx + dy*dy) ** 0.5
            if d < min_d:
                min_d = d
        except:
            pass
    return min_d

#--------------Avoid those doors-------------




_linked_open_aabbs_cache = {}

def get_linked_open_aabbs(doc, pad_ft):
    key = round(float(pad_ft or 0.0), 3)
    hit = _linked_open_aabbs_cache.get(key)
    if hit is not None:
        return hit
    aabbs = _collect_linked_opening_aabbs(doc, pad_ft=key)
    _linked_open_aabbs_cache[key] = aabbs
    return aabbs

def _get_link_transform(link_inst):
    try:
        return link_inst.GetTotalTransform()
    except:
        try: return link_inst.GetTransform()
        except: return None

def _bbox_to_xy_aabb(bb, tf, pad_ft):
    """Return (xmin, ymin, xmax, ymax) in HOST coords, padded by pad_ft."""
    if bb is None:
        return None
    # 8 corners -> transform -> project to XY AABB
    pts = []
    for x in (bb.Min.X, bb.Max.X):
        for y in (bb.Min.Y, bb.Max.Y):
            for z in (bb.Min.Z, bb.Max.Z):
                p = XYZ(x, y, z)
                try:
                    if tf is not None: p = tf.OfPoint(p)
                except:
                    pass
                pts.append(p)
    xs = [p.X for p in pts]; ys = [p.Y for p in pts]
    return (min(xs)-pad_ft, min(ys)-pad_ft, max(xs)+pad_ft, max(ys)+pad_ft)

def _collect_linked_opening_aabbs(doc, pad_ft=2.0):
    """Doors + Openings from all links → XY AABBs in host coords (padded)."""
    aabbs = []
    try:
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = inst.GetLinkDocument()
            if ldoc is None:
                continue
            tf = _get_link_transform(inst)

            # Doors in link
            try:
                for el in FilteredElementCollector(ldoc)\
                        .OfCategory(BuiltInCategory.OST_Doors)\
                        .WhereElementIsNotElementType():
                    bb = el.get_BoundingBox(None)
                    a = _bbox_to_xy_aabb(bb, tf, pad_ft)
                    if a: aabbs.append(a)
            except:
                pass

            # Architectural openings / arches (Opening class)
            try:
                for el in FilteredElementCollector(ldoc).OfClass(Opening):
                    bb = el.get_BoundingBox(None)
                    a = _bbox_to_xy_aabb(bb, tf, pad_ft)
                    if a: aabbs.append(a)
            except:
                pass
    except:
        pass
    return aabbs

def _filter_points_by_linked_openings(pts, aabbs):
    """Remove points whose XY falls inside any linked opening AABB."""
    if not aabbs:
        return pts
    out = []
    for p in pts:
        inside = False
        for (xmin, ymin, xmax, ymax) in aabbs:
            try:
                if (xmin <= p.X <= xmax) and (ymin <= p.Y <= ymax):
                    inside = True
                    break
            except:
                pass
        if not inside:
            out.append(p)
    return out

#------------Avoid placing on thin boundaries------------

def _isfinite(x):
    try:
        return (not Double.IsNaN(x)) and (not Double.IsInfinity(x))
    except:
        # ultra-safe fallback
        try:
            import math
            return math.isfinite(float(x))
        except:
            return (x == x) and (abs(x) < 1e300)



def _seg_key_geom(seg, curve_fn):
    """Stable-ish key if ElementId is missing; uses rounded endpoints."""
    crv = curve_fn(seg)
    if not crv:
        return None
    p0 = crv.GetEndPoint(0); p1 = crv.GetEndPoint(1)
    return ("G",
            round(p0.X, 3), round(p0.Y, 3), round(p0.Z, 3),
            round(p1.X, 3), round(p1.Y, 3), round(p1.Z, 3))

def _seg_key(doc, seg, host_wall_fn, curve_fn):
    """Prefer a *non-wall* ElementId (separation line). Fall back to geometry key.
       Walls are allowed; we won't use their keys to exclude."""
    wall = host_wall_fn(doc, seg)
    if not wall:
        # non-wall: try element id (e.g., Space/Room Separation Line)
        try:
            eid = getattr(seg, "ElementId", None)
            if eid and eid.IntegerValue > 0:
                return ("E", eid.IntegerValue)
        except:
            pass
    # fallback
    return _seg_key_geom(seg, curve_fn)

def _collect_wall_curves_host(doc):
    curves = []
    try:
        for w in FilteredElementCollector(doc).OfClass(Wall):
            lc = w.Location
            if isinstance(lc, LocationCurve) and lc.Curve:
                curves.append(lc.Curve)
    except:
        pass
    return curves

def _get_link_transform(link_inst):
    try:
        return link_inst.GetTotalTransform()
    except:
        try: return link_inst.GetTransform()
        except: return None

def _collect_wall_curves_linked(doc):
    curves = []
    try:
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = inst.GetLinkDocument()
            if ldoc is None:
                continue
            tf = _get_link_transform(inst)
            for w in FilteredElementCollector(ldoc).OfClass(Wall):
                lc = w.Location
                if isinstance(lc, LocationCurve) and lc.Curve:
                    try:
                        crv = lc.Curve.CreateTransformed(tf) if tf else lc.Curve
                    except:
                        crv = lc.Curve
                    if crv:
                        curves.append(crv)
    except:
        pass
    return curves

def _nearest_parallel_offset_ft(curve, wall_curves, ang_tol_cos=0.98, max_search_ft=4.0):
    """Approximate 'thickness' as the smallest perpendicular distance to a nearby
       parallel wall curve (host or linked). Returns +inf if none are near."""
    if not curve:
        return float("inf")
    try:
        # tangent as vector from endpoints (fast enough)
        a = curve.GetEndPoint(0); b = curve.GetEndPoint(1)
        vx, vy = (b.X - a.X), (b.Y - a.Y)
        vlen = (vx*vx + vy*vy) ** 0.5
        if vlen < 1e-6:
            return float("inf")
        ux, uy = vx / vlen, vy / vlen  # unit tangent
        min_d = float("inf")
        mid = curve.Evaluate(0.5, True)
        for crv in wall_curves:
            try:
                a2 = crv.GetEndPoint(0); b2 = crv.GetEndPoint(1)
                wx, wy = (b2.X - a2.X), (b2.Y - a2.Y)
                wlen = (wx*wx + wy*wy) ** 0.5
                if wlen < 1e-6:
                    continue
                # parallel check in XY
                dot = (ux*wx + uy*wy) / wlen
                c = abs(dot) / 1.0  # since ux,uy is unit
                if c < ang_tol_cos:
                    continue  # directions differ too much
                proj = crv.Project(mid)
                if not proj:
                    continue
                q = getattr(proj, "XYZPoint", None) or getattr(proj, "Point", None)
                if not q:
                    continue
                dx, dy = (mid.X - q.X), (mid.Y - q.Y)
                d = (dx*dx + dy*dy) ** 0.5
                if d < min_d:
                    min_d = d
            except:
                pass
        return min_d if min_d <= max_search_ft else float("inf")
    except:
        return float("inf")

def estimate_boundary_thickness_ft(doc, seg, host_wall_fn, curve_fn,
                                   wall_curves_host, wall_curves_linked):
    """
    Heuristic 'thickness' for a boundary segment:
      - If hosted by a Wall: use Wall.Width (robust and fast).
      - Else if the boundary element is a Room/Space Separation Line: 0.0 (very thin).
      - Else: estimate by nearest parallel wall curve (host+linked) distance.
    """
    wall = host_wall_fn(doc, seg)
    if wall:
        try:
            w = float(wall.Width)  # in feet
            if w > 1e-6:
                return w
        except:
            pass

    # separation lines → zero “thickness”
    try:
        eid = getattr(seg, "ElementId", None)
        if eid and eid.IntegerValue > 0:
            el = doc.GetElement(eid)
            if el is not None and el.Category:
                bic = el.Category.Id.IntegerValue
                if bic in (int(BuiltInCategory.OST_SpaceSeparationLines),
                           int(BuiltInCategory.OST_RoomSeparationLines)):
                    return 0.0
    except:
        pass

    # fallback: look for a nearby parallel wall curve (host or linked)
    d_host  = _nearest_parallel_offset_ft(curve_fn(seg), wall_curves_host)
    d_link  = _nearest_parallel_offset_ft(curve_fn(seg), wall_curves_linked)
    d = min(d_host, d_link)
    return d if _isfinite(d) else float("inf")

def compute_thin_boundary_keys(doc, spaces, boundary_loops_fn, segment_host_wall, segment_curve,
                               per_space_factor=0.25,
                               min_abs_ft=0.14,
                               max_consider_ft=2.0,
                               include_host_walls=False,     # <-- default off
                               include_linked_walls=True,    # <-- default on
                               include_nonwalls=True,        # separation lines, etc.
                               logger=None):

    linked_refs = _collect_linked_wall_refs(doc)  # once

    thin_keys = set()
    for sp in spaces:
        loops = boundary_loops_fn(sp)
        if not loops:
            continue

        per_seg = []  # (key, t, is_host_wall, is_linked_wall)

        wall_t = []
        all_t  = []

        for loop in loops:
            for seg in loop or ():
                # gkey for skipping this segment geometry later
                gkey = _seg_key_geom(seg, segment_curve)
                if gkey is None:
                    continue

                # classify & measure thickness
                host_w = (segment_host_wall(doc, seg) is not None)
                crv = segment_curve(seg)
                is_l, linked_w = _segment_is_on_linked_wall(crv, linked_refs)

                t = estimate_boundary_lineband_ft(doc, seg, segment_host_wall, segment_curve, linked_refs)
                if (t != t) or (t == float('inf')) or (t > max_consider_ft):
                    t = max_consider_ft

                per_seg.append((gkey, t, host_w, is_l))
                all_t.append(t)
                if host_w or is_l:
                    wall_t.append(t)  # “wall-like” for median

        # --- threshold ---
        if per_space_factor <= 0.0:
            thresh = float(min_abs_ft)  # fixed floor only (e.g., 0.10 ft)
        else:
            src = wall_t if wall_t else all_t
            m = sorted(src)[len(src) // 2]
            thresh = max(min_abs_ft, per_space_factor * m)

        # --- add thin keys (use gkey) ---
        for gkey, t, host_w, link_w in per_seg:
            consider = ((host_w and include_host_walls) or
                        (link_w and include_linked_walls) or
                        ((not host_w and not link_w) and include_nonwalls))
            if not consider:
                continue
            if t < thresh:
                thin_keys.add(gkey)  # geometry key, not ("E", id)
                if logger:
                    logger.debug(u"[THIN] gkey={} t≈{:.3f}ft < thresh≈{:.3f}ft (host={}, linked={})"
                                 .format(gkey, t, thresh, host_w, link_w))

    if logger:
        logger.info("Thin boundary segments (to skip): {}".format(len(thin_keys)))
    return thin_keys

#-----------Avoid THIN BOUNDARIES 2----------

def _plane_from_face(face):
    # PlanarFace -> Plane (normal, origin)
    try:
        pl = face.GetSurface()  # Plane
        return pl  # has .Normal (XYZ) and .Origin (XYZ)
    except:
        return None

def _parallel_plane_gap_ft(pA, pB):
    # distance between two (roughly parallel) planes
    if pA is None or pB is None:
        return float('inf')
    nA, oA = pA.Normal, pA.Origin
    nB, oB = pB.Normal, pB.Origin
    # make normals co-directional
    dot = nA.X*nB.X + nA.Y*nB.Y + nA.Z*nB.Z
    if dot < 0.0:
        nB = -nB
    # gap = | nA · (oB - oA) |
    dx, dy, dz = (oB.X - oA.X), (oB.Y - oA.Y), (oB.Z - oA.Z)
    gap = abs(nA.X*dx + nA.Y*dy + nA.Z*dz)
    return gap

def _wall_lineband_thickness_ft(wall):
    faces = []
    try:
        for side in (ShellLayerType.Exterior, ShellLayerType.Interior):
            for rf in HostObjectUtils.GetSideFaces(wall, side) or []:
                f = wall.GetGeometryObjectFromReference(rf)
                if isinstance(f, PlanarFace):
                    pl = f.GetSurface()  # Plane
                    if pl: faces.append(pl)
    except:
        pass
    if len(faces) < 2:
        return float('inf')
    # max separation of (roughly) parallel planes
    best = 0.0
    for i in range(len(faces)):
        for j in range(i+1, len(faces)):
            nA, oA = faces[i].Normal, faces[i].Origin
            nB, oB = faces[j].Normal, faces[j].Origin
            dot = nA.X*nB.X + nA.Y*nB.Y + nA.Z*nB.Z
            if dot < 0.0:
                nB = -nB
            dx, dy, dz = (oB.X - oA.X), (oB.Y - oA.Y), (oB.Z - oA.Z)
            d = abs(nA.X*dx + nA.Y*dy + nA.Z*dz)
            if d > best:
                best = d
    return best if best > 0 else float('inf')

def estimate_boundary_lineband_ft(doc, seg, segment_host_wall, segment_curve,
                                  linked_wall_refs):
    """Return thickness of the drawn band:
       - host wall: face-to-face distance
       - linked wall: width (proxy for face-to-face)
       - separation line: 0.0
       - other: +inf
    """
    # host wall?
    wall = segment_host_wall(doc, seg)
    if wall:
        t = _wall_lineband_thickness_ft(wall)
        if t != float('inf'):
            return t

    # linked wall?
    crv = segment_curve(seg)
    is_l, w = _segment_is_on_linked_wall(crv, linked_wall_refs)
    if is_l and w > 0.0:
        return w

    # separation line?
    try:
        eid = getattr(seg, "ElementId", None)
        if eid and eid.IntegerValue > 0:
            el = doc.GetElement(eid)
            if el is not None and el.Category:
                bic = el.Category.Id.IntegerValue
                if bic in (int(BuiltInCategory.OST_SpaceSeparationLines),
                           int(BuiltInCategory.OST_RoomSeparationLines)):
                    return 0.0
    except:
        pass

    return float('inf')

#-----------------Check for THIN LINKED WALLS------------------


def _collect_linked_wall_refs(doc):
    """Return [(curve_in_host_coords, width_ft, wall_id_int), ...] for all linked walls."""
    refs = []
    try:
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = inst.GetLinkDocument()
            if not ldoc:
                continue
            # get transform into host coords
            try:
                tf = inst.GetTotalTransform()
            except:
                try: tf = inst.GetTransform()
                except: tf = None

            for w in FilteredElementCollector(ldoc).OfClass(Wall):
                lc = w.Location
                if not isinstance(lc, LocationCurve) or lc.Curve is None:
                    continue
                crv = lc.Curve
                try:
                    if tf: crv = crv.CreateTransformed(tf)
                except:
                    pass
                width = float(getattr(w, "Width", 0.0) or 0.0)
                refs.append((crv, width, w.Id.IntegerValue))
    except:
        pass
    return refs

def _segment_is_on_linked_wall(seg_curve, linked_refs, ang_tol_cos=0.98, dist_tol_ft=0.5):
    """
    Returns (True, width_ft) if seg_curve aligns with any linked wall curve:
      - near-parallel in XY (cos ≥ ang_tol_cos),
      - mid-point XY distance ≤ dist_tol_ft.
    """
    if not seg_curve:
        return (False, 0.0)
    try:
        a = seg_curve.GetEndPoint(0); b = seg_curve.GetEndPoint(1)
        vx, vy = (b.X - a.X), (b.Y - a.Y)
        vlen = (vx*vx + vy*vy) ** 0.5
        if vlen < 1e-6:
            return (False, 0.0)
        ux, uy = vx / vlen, vy / vlen
        mid = seg_curve.Evaluate(0.5, True)

        for crv, width, _ in linked_refs:
            try:
                a2 = crv.GetEndPoint(0); b2 = crv.GetEndPoint(1)
                wx, wy = (b2.X - a2.X), (b2.Y - a2.Y)
                wlen = (wx*wx + wy*wy) ** 0.5
                if wlen < 1e-6:
                    continue
                # parallel check in XY
                cosang = abs((ux*wx + uy*wy) / wlen)
                if cosang < ang_tol_cos:
                    continue
                # XY distance from mid to linked wall
                proj = crv.Project(mid)
                if not proj:
                    continue
                q = getattr(proj, "XYZPoint", None) or getattr(proj, "Point", None)
                if not q:
                    continue
                dx, dy = (mid.X - q.X), (mid.Y - q.Y)
                d = (dx*dx + dy*dy) ** 0.5
                if d <= dist_tol_ft:
                    return (True, float(width))
            except:
                pass
    except:
        pass
    return (False, 0.0)

#-----------------Main Function---------------------

@RunInTransaction("Electrical::PerimeterReceptsByRules")
def place_perimeter_recepts(doc, logger=None):
    log = logger or get_logger("MEPKit")

    id_rules = load_identify_rules()
    bc_rules = load_branch_rules()

    spaces = collect_spaces_or_rooms(doc)
    log.info("Spaces/Rooms found: {}".format(len(spaces)))
    if not spaces:
        return 0

    # NEW: host + linked wall curves (already in host coordinates)
    host_wall_curves = _collect_wall_curves_host(doc)
    linked_wall_curves = _collect_wall_curves_linked(doc)
    all_wall_curves = host_wall_curves + linked_wall_curves

    # NEW: collect linked doors & openings as XY AABBs (2 ft pad)
    linked_open_aabbs = _collect_linked_opening_aabbs(doc, pad_ft=2.0)
    log.info("Linked doors/openings (AABBs): {}".format(len(linked_open_aabbs)))

    # compute thin boundary keys once (uses walls from host + links under the hood)
    thin_keys = compute_thin_boundary_keys(
        doc, spaces,
        boundary_loops_fn=boundary_loops,
        segment_host_wall=segment_host_wall,
        segment_curve=segment_curve,
        per_space_factor=0.0,  # ← disables median logic
        min_abs_ft=0.20,  # ← fixed floor: 0.10 ft
        max_consider_ft=2.0,
        include_host_walls=False,
        include_linked_walls=True,
        include_nonwalls=True,
        logger=log
    )

    total = 0
    for sp in spaces:
        name = space_name(sp)
        match_text = space_match_text(sp)
        cat = categorize_space_by_name(match_text, id_rules)

        log.info(
            u"Space Id {} → name='{}' match_text='{}' → category [{}]".format(
                sp.Id.IntegerValue, name, match_text, cat
            )
        )

        cat_rule, general = get_category_rule(bc_rules, cat, fallback='Support')
        if not cat_rule:
            log.info("Skip space '{}' → category [{}] has no rule".format(name, cat))
            continue

        # spacing (first/next)
        spacing = cat_rule.get('wall_spacing_ft') or {}
        first_ft = float(spacing.get('first', spacing.get('next', 20.0)))
        next_ft  = float(spacing.get('next', first_ft))

        # mount height (feet)
        mh_in = cat_rule.get('mount_height_in', None)
        mh_ft = (float(mh_in) / 12.0) if mh_in is not None else None

        # constraints (normalize; DO NOT relax)
        gcon = normalize_constraints(general.get('placement_constraints', {}))
        ccon = normalize_constraints(cat_rule.get('placement_constraints', {}))
        avoid_corners_ft      = float(ccon.get('avoid_corners_ft', gcon.get('avoid_corners_ft', 2.0)))
        avoid_doors_radius_ft = float(ccon.get('avoid_doors_radius_ft', gcon.get('avoid_doors_radius_ft', 0.0)))
        door_edge_margin_ft   = float(ccon.get('door_edge_margin_ft', gcon.get('door_edge_margin_ft', 0.0)))
        avoid_linked_openings_ft = float(
            ccon.get('avoid_linked_openings_ft',
                     gcon.get('avoid_linked_openings_ft', 2.0))
        )

        # IMPORTANT: keep perimeter inset tiny & stable; do NOT use door snap tolerance here
        inset_ft = 0.05

        # symbol candidates (auto-load)
        sym = None
        for cand in (cat_rule.get('device_candidates') or []):
            fam  = cand.get('family')
            typ  = cand.get('type_catalog_name')
            path = cand.get('load_from')
            if fam:
                sym = resolve_or_load_symbol(doc, fam, typ, load_path=path, logger=log)
                if sym:
                    break
        if not sym:
            log.warning("No family symbol matched/loaded for space '{}' [{}]".format(name, cat))
            continue

        loops = boundary_loops(sp)
        if not loops:
            log.info("Space '{}' [{}] → no boundary loops (room-bounding?)".format(name, cat))
            continue

        placed_here = 0
        seg_count = 0
        pre_pts_total = 0
        post_pts_total = 0

        # Strict pass: apply corner/door rules; no “relax plan”
        for loop in loops:
            for seg in loop:
                seg_count += 1

                # NEW:
                gkey = _seg_key_geom(seg, segment_curve)
                if gkey and gkey in thin_keys:
                    continue



                curve = segment_curve(seg)
                if not curve:
                    continue

                # sample along the segment with corner inset
                pts = sample_points_on_segment(curve, first_ft, next_ft, avoid_corners_ft, inset_ft)
                pre_pts_total += len(pts)

                wall = segment_host_wall(doc, seg)

                # door filtering (if any)
                if wall and (avoid_doors_radius_ft > 0.0 or door_edge_margin_ft > 0.0):
                    doors = door_points_on_wall(doc, wall)
                    pts = filter_points_by_doors(pts, doors, avoid_doors_radius_ft, door_edge_margin_ft)


                # NEW: linked doors / arches filter (rule-driven buffer)
                if avoid_linked_openings_ft > 0.0:
                    linked_open_aabbs = get_linked_open_aabbs(doc, avoid_linked_openings_ft)
                    if linked_open_aabbs:
                        pts = _filter_points_by_linked_openings(pts, linked_open_aabbs)

                post_pts_total += len(pts)

                # place what survived filtering
                for p in pts:
                    try:
                        if wall:
                            inst = place_hosted(doc, wall, sym, p, mounting_height_ft=mh_ft, logger=log)
                        else:
                            inst = place_free(doc, sym, p, mounting_height_ft=mh_ft, logger=log)
                            # NEW: delete if this free-placed device is not near any wall (0.5 ft)
                            try:
                                d = _nearest_wall_xy_distance(p, all_wall_curves)
                                if d > 0.5:
                                    doc.Delete(inst.Id)
                                    placed_here -= 1
                                    log.info(u"Deleted (no wall within 0.5 ft): d≈{:.2f} ft".format(d))
                                    continue
                            except Exception as ex:
                                log.warning(u"Proximity check/delete failed: {}".format(ex))

                        if mh_ft is not None:
                            set_param_value(inst, "Mounting Height", mh_ft)

                        placed_here += 1
                    except Exception as ex:
                        log.warning(u"Placement failed at point → {}".format(ex))

        log.info("Space '{}' [{}] → loops={}, segs={}, pts pre/ post door = {}/{} → placed {}"
                 .format(name, cat, len(loops), seg_count, pre_pts_total, post_pts_total, placed_here))

        total += placed_here

    log.info("Total placed around perimeters: {}".format(total))
    return total