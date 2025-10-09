# -*- coding: utf-8 -*-
# lib/organized/MEPKit/electrical/perimeter_runner.py
from __future__ import absolute_import

import math

from organized.MEPKit.core.log import get_logger
from organized.MEPKit.core.rules import (
    load_identify_rules, load_branch_rules, normalize_constraints,
    categorize_space_by_name, get_category_rule
)
from organized.MEPKit.revit.transactions import RunInTransaction
from organized.MEPKit.revit.spaces import (
    collect_spaces_or_rooms, space_name, boundary_loops, segment_curve,
    segment_host_wall, sample_points_on_segment, space_match_text
)
from organized.MEPKit.revit.doors import door_points_on_wall, filter_points_by_doors
from organized.MEPKit.revit.placement import place_hosted, place_free
from organized.MEPKit.revit.symbols import resolve_or_load_symbol
from organized.MEPKit.revit.params import set_param_value  # optional for mounting height

from Autodesk.Revit.DB import (
    RevitLinkInstance, Opening, BuiltInCategory, FilteredElementCollector,
    XYZ, Wall, LocationCurve, SpatialElementBoundaryOptions, SpatialElementBoundaryLocation
)

# Toggle debug printing for shared-pair checks
PAIR_DIAG = True


# -------------------------- basic curve helpers --------------------------

def _seg_mid_xy(curve):
    try:
        p = curve.Evaluate(0.5, True)
    except Exception:
        p0 = curve.GetEndPoint(0); p1 = curve.GetEndPoint(1)
        p = XYZ((p0.X+p1.X)/2.0, (p0.Y+p1.Y)/2.0, (p0.Z+p1.Z)/2.0)
    return (p.X, p.Y)

def _seg_tan_xy(curve):
    try:
        der = curve.ComputeDerivatives(0.5, True)
        v = der.BasisX
    except Exception:
        p0 = curve.GetEndPoint(0); p1 = curve.GetEndPoint(1)
        v = XYZ(p1.X-p0.X, p1.Y-p0.Y, 0.0)
    mag = math.hypot(v.X, v.Y) or 1.0
    return (v.X/mag, v.Y/mag)  # unit tangent (XY)

def _space_sides_for_segment(space_id, curve, category_at_xy, probe_ft):
    """
    Sample a point ~probe_ft to each side of the segment normal (XY).
    Returns: (cat_left, cat_right, which_side_is_self: 'L'|'R'|None)
    """
    mx, my = _seg_mid_xy(curve)
    tx, ty = _seg_tan_xy(curve)
    nx, ny = -ty, tx  # perp to segment in XY

    a = (mx + nx*probe_ft, my + ny*probe_ft)  # “left” side
    b = (mx - nx*probe_ft, my - ny*probe_ft)  # “right” side

    sid_a, cat_a = category_at_xy(a)
    sid_b, cat_b = category_at_xy(b)

    which = None
    if sid_a == space_id:
        which = 'L'
    elif sid_b == space_id:
        which = 'R'

    return (cat_a or None, cat_b or None, which)

def _curve_point_at(curve, t01):
    try:
        return curve.Evaluate(float(t01), True)
    except:
        return None


# -------------------------- wall curve collectors --------------------------

def _get_link_transform(link_inst):
    try:
        return link_inst.GetTotalTransform()  # newer API
    except:
        try:
            return link_inst.GetTransform()    # older API
        except:
            return None

def _collect_host_wall_curves(doc):
    curves = []
    try:
        for w in FilteredElementCollector(doc).OfClass(Wall):
            lc = w.Location
            if isinstance(lc, LocationCurve) and lc.Curve is not None:
                curves.append(lc.Curve)
    except:
        pass
    return curves

def _collect_linked_wall_curves(doc):
    curves = []
    try:
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = inst.GetLinkDocument()
            if not ldoc:
                continue
            tf = _get_link_transform(inst)
            for w in FilteredElementCollector(ldoc).OfClass(Wall):
                lc = w.Location
                if isinstance(lc, LocationCurve) and lc.Curve is not None:
                    crv = lc.Curve
                    try:
                        if tf:
                            crv = crv.CreateTransformed(tf)
                    except:
                        pass
                    curves.append(crv)
    except:
        pass
    return curves

def _collect_all_wall_curves(doc):
    host = _collect_host_wall_curves(doc)
    linked = _collect_linked_wall_curves(doc)
    return host + linked, len(host), len(linked)


# -------------------------- linked doors / openings AABBs --------------------------

_linked_open_aabbs_cache = {}

def _bbox_to_xy_aabb(bb, tf, pad_ft):
    if bb is None:
        return None
    pts = []
    for x in (bb.Min.X, bb.Max.X):
        for y in (bb.Min.Y, bb.Max.Y):
            for z in (bb.Min.Z, bb.Max.Z):
                p = XYZ(x, y, z)
                try:
                    if tf is not None:
                        p = tf.OfPoint(p)
                except:
                    pass
                pts.append(p)
    xs = [p.X for p in pts]; ys = [p.Y for p in pts]
    return (min(xs)-pad_ft, min(ys)-pad_ft, max(xs)+pad_ft, max(ys)+pad_ft)

def _collect_linked_opening_aabbs(doc, pad_ft=2.0):
    aabbs = []
    try:
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = inst.GetLinkDocument()
            if ldoc is None:
                continue
            tf = _get_link_transform(inst)

            # Doors
            try:
                for el in FilteredElementCollector(ldoc)\
                        .OfCategory(BuiltInCategory.OST_Doors)\
                        .WhereElementIsNotElementType():
                    bb = el.get_BoundingBox(None)
                    a = _bbox_to_xy_aabb(bb, tf, pad_ft)
                    if a: aabbs.append(a)
            except:
                pass

            # Openings / arches
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

def get_linked_open_aabbs(doc, pad_ft):
    key = round(float(pad_ft or 0.0), 3)
    hit = _linked_open_aabbs_cache.get(key)
    if hit is not None:
        return hit
    aabbs = _collect_linked_opening_aabbs(doc, pad_ft=key)
    _linked_open_aabbs_cache[key] = aabbs
    return aabbs

def _filter_points_by_linked_openings(pts, aabbs):
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


# -------------------------- near-wall distance --------------------------

def _nearest_wall_xy_distance(p, curves):
    """2D distance (ft) from XYZ p to nearest wall curve in host coords; None if unknown."""
    if not curves:
        return None
    best = float('inf')
    for c in curves:
        try:
            pr = c.Project(p)
            q = getattr(pr, "XYZPoint", None) or getattr(pr, "Point", None)
            if q is None:
                q0 = c.GetEndPoint(0); q1 = c.GetEndPoint(1)
                d0 = math.hypot(p.X - q0.X, p.Y - q0.Y)
                d1 = math.hypot(p.X - q1.X, p.Y - q1.Y)
                d = min(d0, d1)
            else:
                d = math.hypot(p.X - q.X, p.Y - q.Y)
            if d < best:
                best = d
        except:
            try:
                q0 = c.GetEndPoint(0); q1 = c.GetEndPoint(1)
                d0 = math.hypot(p.X - q0.X, p.Y - q0.Y)
                d1 = math.hypot(p.X - q1.X, p.Y - q1.Y)
                d = min(d0, d1)
                if d < best:
                    best = d
            except:
                pass
    return None if best == float('inf') else best

def _nearest_curve_info_xy(p, curves):
    """Return (curve, dist_ft, (tx,ty)) for the nearest curve to XY of p; None if not found."""
    if not curves:
        return (None, None, None)
    best_curve, best_d, best_tan = (None, float('inf'), None)
    for c in curves:
        try:
            pr = c.Project(p)
            q = getattr(pr, "XYZPoint", None) or getattr(pr, "Point", None)
            if q is None:
                continue
            d = math.hypot(p.X - q.X, p.Y - q.Y)
            if d < best_d:
                # tangent at the projected parameter if possible
                try:
                    der = c.ComputeDerivatives(pr.Parameter, True)
                    v = der.BasisX
                except:
                    q0, q1 = c.GetEndPoint(0), c.GetEndPoint(1)
                    v = XYZ(q1.X - q0.X, q1.Y - q0.Y, 0.0)
                mag = math.hypot(v.X, v.Y) or 1.0
                best_curve, best_d, best_tan = c, d, (v.X / mag, v.Y / mag)
        except:
            pass
    if best_d == float('inf'):
        return (None, None, None)
    return (best_curve, best_d, best_tan)


# -------------------------- “point in space” locator --------------------------

def _ray_cast_point_in_poly(pt, poly_xy):
    x, y = pt
    inside = False
    j = len(poly_xy) - 1
    for i in range(len(poly_xy)):
        xi, yi = poly_xy[i]
        xj, yj = poly_xy[j]
        if ((yi > y) != (yj > y)):
            xint = xi + (y - yi) * (xj - xi) / (yj - yi)
            if xint >= x:
                inside = not inside
        j = i
    return inside

def _bbox_of_xy(poly_xy):
    xs = [p[0] for p in poly_xy]
    ys = [p[1] for p in poly_xy]
    return (min(xs), min(ys), max(xs), max(ys))

def _loop_to_xy(loop):
    """Convert Revit boundary loop (IList<BoundarySegment>) → {'xy':[(x,y)...],'perimeter_ft':float}."""
    pts = []
    perim = 0.0
    for bs in loop:
        crv = bs.GetCurve()
        p0 = crv.GetEndPoint(0); p1 = crv.GetEndPoint(1)
        if not pts:
            pts.append((p0.X, p0.Y))
        pts.append((p1.X, p1.Y))
        try:
            perim += crv.Length
        except:
            pass
    if len(pts) >= 2 and pts[0] == pts[-1]:
        pts = pts[:-1]
    return {"xy": pts, "perimeter_ft": perim}

def build_space_loops_by_id(doc, spaces, boundary_location="Finish"):
    """
    Returns: { space_id:int : [ {'xy':[(x,y)...], 'perimeter_ft':float}, ... ] }
    """
    opt = SpatialElementBoundaryOptions()
    if (boundary_location or "").lower().startswith("center"):
        opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center
    else:
        opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

    loops_map = {}
    for sp in spaces:
        sid = sp.Id.IntegerValue
        try:
            segloops = sp.GetBoundarySegments(opt)
        except:
            segloops = None
        if not segloops:
            continue
        loops = []
        for loop in segloops:
            try:
                loops.append(_loop_to_xy(loop))
            except:
                pass
        if loops:
            loops_map[sid] = loops
    return loops_map

class SpaceLocator(object):
    """Fast XY → (space_id, category) lookup built from space boundary loops."""

    __slots__ = ("_polys",)

    def __init__(self, polys):
        self._polys = polys

    @classmethod
    def from_space_loops(cls, spaces, space_loops_by_id, cat_by_spaceid):
        polys = []
        for sp in spaces:
            sid = sp.Id.IntegerValue
            loops = space_loops_by_id.get(sid) or []
            if not loops:
                continue
            if loops and isinstance(loops[0], dict):
                outer = max(loops, key=lambda lp: abs(lp.get("perimeter_ft", 0.0)))
                xy = outer.get("xy") or []
            else:
                outer = loops[0] if loops else []
                xy = outer or []
            if not xy:
                continue
            cat = (cat_by_spaceid.get(sid) or u"").strip().lower()
            bbox = _bbox_of_xy(xy)
            polys.append({'sid': sid, 'xy': xy, 'bbox': bbox, 'cat': cat})
        return cls(polys)

    def category_at_xy(self, pt):
        x, y = pt
        for item in self._polys:
            xmin, ymin, xmax, ymax = item['bbox']
            if x < xmin or x > xmax or y < ymin or y > ymax:
                continue
            try:
                if _ray_cast_point_in_poly((x, y), item['xy']):
                    return item['sid'], item['cat']
            except:
                pass
        return None, None


# -------------------------- pair-skip wiring --------------------------

def _segment_is_from_linked_wall(curve, linked_wall_curves, tol_ft=0.5):
    """Returns True if the boundary curve lies on any LINKED wall curve within tol_ft."""
    if not curve or not linked_wall_curves:
        return False
    try:
        pm = curve.Evaluate(0.5, True)
    except:
        p0 = curve.GetEndPoint(0); p1 = curve.GetEndPoint(1)
        pm = type(p0)((p0.X+p1.X)/2.0, (p0.Y+p1.Y)/2.0, (p0.Z+p1.Z)/2.0)
    for lc in linked_wall_curves:
        try:
            pr = lc.Project(pm)
            if pr and pr.Distance <= tol_ft:
                return True
        except:
            try:
                q0, q1 = lc.GetEndPoint(0), lc.GetEndPoint(1)
                d0 = math.hypot(pm.X - q0.X, pm.Y - q0.Y)
                d1 = math.hypot(pm.X - q1.X, pm.Y - q1.Y)
                if min(d0, d1) <= tol_ft:
                    return True
            except:
                pass
    return False

def _load_skip_pair_set(bc_rules):
    gen = (bc_rules or {}).get('general', {}) or {}
    rcfg = (bc_rules or {}).get('rule_config', {}) or {}
    raw = (gen.get('skip_shared_boundary_pairs') or []) + (rcfg.get('skip_shared_boundary_pairs') or [])
    s = set()
    for a, b in raw:
        s.add(tuple(sorted(((a or u'').strip().lower(), (b or u'').strip().lower()))))
    return s


# -------------------------- MAIN --------------------------

@RunInTransaction("Electrical::PerimeterReceptsByRules")
def place_perimeter_recepts(doc, logger=None):
    log = logger or get_logger("MEPKit")

    # Rules + spaces
    id_rules = load_identify_rules()
    bc_rules = load_branch_rules()
    spaces = collect_spaces_or_rooms(doc)

    # Skip-pair set (e.g., ("dairy/ produce cooler","sales floor"))
    skip_pair_set = _load_skip_pair_set(bc_rules)
    log.info(u"[PAIR] Skip shared pairs: {}".format(sorted(list(skip_pair_set))))

    # Global/general constraints (for near-wall cleanup + default pair params)
    gen = (bc_rules or {}).get('general', {}) or {}
    general_constraints = gen.get('placement_constraints', {}) or {}
    near_wall_ft = float(general_constraints.get('near_wall_threshold_ft', 0.5))
    default_pair_probe_ft = float(general_constraints.get('pair_probe_ft', 0.5))
    pair_linked_wall_tol_ft = float(general_constraints.get('pair_linked_wall_tol_ft', 1.5))
    log.info("Near-wall threshold: {:.2f} ft".format(near_wall_ft))

    log.info("Spaces/Rooms found: {}".format(len(spaces)))
    if not spaces:
        return 0

    # Wall curve caches
    all_wall_curves, host_n, link_n = _collect_all_wall_curves(doc)
    linked_wall_curves = _collect_linked_wall_curves(doc)
    log.info("Wall curve cache → host:{} linked:{}".format(host_n, link_n))
    if PAIR_DIAG:
        log.debug(u"[PAIRDBG] linked_wall_curves count: {}".format(len(linked_wall_curves or [])))

    # Linked doors/openings AABBs (default pad 2.0')
    linked_open_aabbs = _collect_linked_opening_aabbs(doc, pad_ft=2.0)
    log.info("Linked doors/openings (AABBs): {}".format(len(linked_open_aabbs)))

    # Categories for locator
    cat_by_spaceid = {}
    for sp in spaces:
        cat_by_spaceid[sp.Id.IntegerValue] = (categorize_space_by_name(space_match_text(sp), id_rules) or u"").strip().lower()

    space_loops_by_id = build_space_loops_by_id(doc, spaces, boundary_location="Finish")
    locator = SpaceLocator.from_space_loops(spaces, space_loops_by_id, cat_by_spaceid)
    category_at_xy = locator.category_at_xy

    total = 0

    for sp in spaces:
        space_id = sp.Id.IntegerValue
        name = space_name(sp)
        match_text = space_match_text(sp)
        cat = categorize_space_by_name(match_text, id_rules)

        log.info(
            u"Space Id {} → name='{}' match_text='{}' → category [{}]".format(
                space_id, name, match_text, cat
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

        # constraints (NO relax)
        gcon = normalize_constraints(general.get('placement_constraints', {}))
        ccon = normalize_constraints(cat_rule.get('placement_constraints', {}))

        avoid_corners_ft         = float(ccon.get('avoid_corners_ft', gcon.get('avoid_corners_ft', 2.0)))
        avoid_doors_radius_ft    = float(ccon.get('avoid_doors_radius_ft', gcon.get('avoid_doors_radius_ft', 0.0)))
        door_edge_margin_ft      = float(ccon.get('door_edge_margin_ft', gcon.get('door_edge_margin_ft', 0.0)))
        avoid_linked_openings_ft = float(ccon.get('avoid_linked_openings_ft',
                                                  gcon.get('avoid_linked_openings_ft', 2.0)))
        pair_probe_ft            = float(ccon.get('pair_probe_ft', gcon.get('pair_probe_ft', default_pair_probe_ft)))

        # tiny inset for sampling
        inset_ft = 0.05

        # resolve/load symbol
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

        for loop in loops:
            for seg in loop:
                seg_count += 1
                curve = segment_curve(seg)
                if curve is None:
                    continue

                # ----- NEW: shared-pair skip (only on LINKED wall boundary segments) -----
                if skip_pair_set and _segment_is_from_linked_wall(curve, linked_wall_curves, tol_ft=pair_linked_wall_tol_ft):
                    pair_hit = False
                    for scale in (1.0, 1.5, 2.0, 3.0):
                        L, R, self_side = _space_sides_for_segment(space_id, curve, category_at_xy,
                                                                   pair_probe_ft * scale)
                        if PAIR_DIAG:
                            try:
                                log.debug(u"[PAIRCHK] space={} probe≈{:.2f}ft L={} R={} self={}"
                                          .format(space_id, pair_probe_ft*scale, L, R, self_side))
                            except:
                                pass
                        if L and R:
                            pa = (L or u"").strip().lower()
                            pb = (R or u"").strip().lower()
                            if tuple(sorted((pa, pb))) in skip_pair_set:
                                if PAIR_DIAG:
                                    log.info(u"[PAIRHIT] space={} seg → {} | {} (probe≈{:.2f}ft) → SKIP"
                                             .format(space_id, pa, pb, pair_probe_ft*scale))
                                pair_hit = True
                                break
                    if pair_hit:
                        continue
                # -------------------------------------------------------------------------

                # sample points along segment
                pts = sample_points_on_segment(curve, first_ft, next_ft, avoid_corners_ft, inset_ft)
                pre_pts_total += len(pts)

                # “host” wall (for hosted placement + door filtering)
                wall = segment_host_wall(doc, seg)

                # remove near doors on host wall
                if wall and (avoid_doors_radius_ft > 0.0 or door_edge_margin_ft > 0.0):
                    doors = door_points_on_wall(doc, wall)
                    pts = filter_points_by_doors(pts, doors, avoid_doors_radius_ft, door_edge_margin_ft)

                # avoid linked door/arch openings (buffered AABBs)
                if avoid_linked_openings_ft > 0.0:
                    aabbs = get_linked_open_aabbs(doc, avoid_linked_openings_ft)
                    if aabbs:
                        pts = _filter_points_by_linked_openings(pts, aabbs)

                post_pts_total += len(pts)

                # place
                for p in pts:
                    try:
                        # --- point-level guard: skip points near a linked wall if categories form a skip pair ---
                        if skip_pair_set and linked_wall_curves:
                            lc, d_link, txy = _nearest_curve_info_xy(p, linked_wall_curves)
                            if lc and (d_link is not None) and (d_link <= pair_linked_wall_tol_ft) and txy:
                                tx, ty = txy
                                nx, ny = -ty, tx  # normal to the linked wall in XY
                                skip_this_point = False
                                for scale in (1.0, 1.5, 2.0, 3.0):
                                    off = pair_probe_ft * scale
                                    a = (p.X + nx * off, p.Y + ny * off)
                                    b = (p.X - nx * off, p.Y - ny * off)
                                    _sida, ca = category_at_xy(a)
                                    _sidb, cb = category_at_xy(b)
                                    if ca and cb:
                                        pa = (ca or u"").strip().lower()
                                        pb = (cb or u"").strip().lower()
                                        if tuple(sorted((pa, pb))) in skip_pair_set:
                                            if PAIR_DIAG:
                                                log.info(u"[PAIRHIT-PT] skip point near linked wall: {} | {} (d≈{:.2f}ft, off≈{:.2f}ft)"
                                                         .format(pa, pb, d_link, off))
                                            skip_this_point = True
                                            break
                                if skip_this_point:
                                    continue  # ← actually skip this candidate point
                        # ----------------------------------------------------------------------------------------

                        deleted = False
                        if wall:
                            inst = place_hosted(doc, wall, sym, p, mounting_height_ft=mh_ft, logger=log)
                        else:
                            inst = place_free(doc, sym, p, mounting_height_ft=mh_ft, logger=log)
                            # near-wall cleanup only for truly free instances
                            try:
                                if getattr(inst, "Host", None) is None:
                                    d = _nearest_wall_xy_distance(p, all_wall_curves)
                                    if d is not None and d > near_wall_ft:
                                        doc.Delete(inst.Id)
                                        deleted = True
                                        if PAIR_DIAG:
                                            log.info(u"Deleted (no wall within {:.2f} ft): d≈{:.2f} ft"
                                                     .format(near_wall_ft, d))
                            except Exception as ex:
                                log.warning("Near-wall cleanup error: {}".format(ex))

                        if deleted:
                            continue  # do not touch params / counters on deleted instance

                        if mh_ft is not None:
                            try:
                                set_param_value(inst, "Mounting Height", mh_ft)
                            except:
                                pass

                        placed_here += 1

                    except Exception as ex:
                        log.warning(u"Placement failed at point → {}".format(ex))

        log.info("Space '{}' [{}] → loops={}, segs={}, pts pre/ post door = {}/{} → placed {}"
                 .format(name, cat, len(loops), seg_count, pre_pts_total, post_pts_total, placed_here))

        total += placed_here

    log.info("Total placed around perimeters: {}".format(total))
    return total
