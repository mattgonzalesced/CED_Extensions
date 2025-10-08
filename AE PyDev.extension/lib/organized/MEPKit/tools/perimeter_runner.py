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
import math


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

def _nearest_wall_xy_distance(p, curves):
    """Return 2D distance (feet) from point p (XYZ) to nearest wall curve in host coords.
       Returns None if curves is empty or projection fails everywhere."""
    if not curves:
        return None
    best = float('inf')
    for c in curves:
        try:
            # Try curve.Project first (works on most Revit curves)
            pr = c.Project(p)
            q = getattr(pr, "XYZPoint", None) or getattr(pr, "Point", None)
            if q is None:
                # Fallback: closest endpoint
                q0 = c.GetEndPoint(0); q1 = c.GetEndPoint(1)
                d0 = math.hypot(p.X - q0.X, p.Y - q0.Y)
                d1 = math.hypot(p.X - q1.X, p.Y - q1.Y)
                d = min(d0, d1)
            else:
                d = math.hypot(p.X - q.X, p.Y - q.Y)
            if d < best:
                best = d
        except:
            # Last-resort fallback to endpoints
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

#---------------Collect hosted AND linked walls----------------

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
            # robust transform (works across Revit versions)
            try:
                tf = inst.GetTotalTransform()
            except:
                try: tf = inst.GetTransform()
                except: tf = None

            for w in FilteredElementCollector(ldoc).OfClass(Wall):
                lc = w.Location
                if isinstance(lc, LocationCurve) and lc.Curve is not None:
                    crv = lc.CCurve if hasattr(lc, "CCurve") else lc.Curve  # safety
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


#------------That window between diary and sales-----------

# --- read the pair rules from bc_rules.general ---
def _load_skip_pair_set(bc_rules):
    gen = (bc_rules or {}).get('general', {})
    pairs = gen.get('skip_shared_boundary_pairs', []) or []
    out = set()
    for pair in pairs:
        if isinstance(pair, (list, tuple)) and len(pair) == 2:
            a = (pair[0] or "").strip(); b = (pair[1] or "").strip()
            if a and b:
                out.add(tuple(sorted((a, b))))
    return out

# --- build an index: segment geometry -> [(space_id, category), ...] ---
def _build_shared_boundary_indexes(doc, spaces, id_rules):
    ix_g, ix_m = {}, {}
    for sp in spaces:
        cat = (categorize_space_by_name(space_match_text(sp), id_rules) or "").strip()
        for loop in (boundary_loops(sp) or []):
            for seg in loop or []:
                crv = segment_curve(seg)
                if not crv: continue
                gk = _seg_gkey2d(crv); mk = _seg_mkey2d(crv)
                if gk: ix_g.setdefault(gk, []).append((sp.Id.IntegerValue, cat))
                if mk: ix_m.setdefault(mk, []).append((sp.Id.IntegerValue, cat))
    return ix_g, ix_m


#------------------More to handle that window between Dairy and sales--------------------

def _pt2d_key(p, prec=3):
    r = lambda v: round(v, prec)
    return (r(p.X), r(p.Y))

def _seg_gkey2d(curve, prec=3):
    if not curve: return None
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    a, b = _pt2d_key(p0, prec), _pt2d_key(p1, prec)
    return ("G2",) + (a + b if a <= b else b + a)

def _seg_mkey2d(curve, prec=2):
    if not curve: return None
    try:
        pm = curve.Evaluate(0.5, True)
    except:
        p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
        pm = type(p0)((p0.X+p1.X)/2.0, (p0.Y+p1.Y)/2.0, (p0.Z+p1.Z)/2.0)
    return ("M2",) + _pt2d_key(pm, prec)

def _segment_is_from_linked_wall(curve, linked_wall_curves, tol_ft=0.2):
    """Returns True if the boundary curve sits on any LINKED wall curve within tol."""
    if not curve or not linked_wall_curves:
        return False
    try:
        pm = curve.Evaluate(0.5, True)
    except:
        p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
        pm = type(p0)((p0.X+p1.X)/2.0, (p0.Y+p1.Y)/2.0, (p0.Z+p1.Z)/2.0)
    for lc in linked_wall_curves:
        try:
            pr = lc.Project(pm)
            if pr and pr.Distance <= tol_ft:
                return True
        except:
            # fallback to endpoints if Project fails
            try:
                q0, q1 = lc.GetEndPoint(0), lc.GetEndPoint(1)
                # 2D distance
                import math
                d0 = math.hypot(pm.X - q0.X, pm.Y - q0.Y)
                d1 = math.hypot(pm.X - q1.X, pm.Y - q1.Y)
                if min(d0, d1) <= tol_ft:
                    return True
            except:
                pass
    return False


#-----------------Main Function---------------------

@RunInTransaction("Electrical::PerimeterReceptsByRules")
def place_perimeter_recepts(doc, logger=None):
    log = logger or get_logger("MEPKit")

    id_rules = load_identify_rules()
    bc_rules = load_branch_rules()
    spaces = collect_spaces_or_rooms(doc)

    # NEW: build shared-boundary indexes (geom & midpoint keys)
    shared_ix_g, shared_ix_m = _build_shared_boundary_indexes(doc, spaces, id_rules)


    skip_pair_set = _load_skip_pair_set(bc_rules)
    log.info(u"[PAIR] Skip shared pairs: {}".format(sorted(list(skip_pair_set))))

    # new at 1:19 10/8/25
    gen = (bc_rules or {}).get('general', {})
    constraints = gen.get('placement_constraints', {}) or {}
    near_wall_ft = float(constraints.get('near_wall_threshold_ft', 0.5))
    log.info("Near-wall threshold: {:.2f} ft".format(near_wall_ft))


    log.info("Spaces/Rooms found: {}".format(len(spaces)))
    if not spaces:
        return 0

    all_wall_curves, _host_n, _link_n = _collect_all_wall_curves(doc)
    linked_wall_curves = _collect_linked_wall_curves(doc)  # NEW dedicated list
    log.info("Wall curve cache → host:{} linked:{}".format(_host_n, _link_n))
    log.debug(u"[PAIRDBG] linked_wall_curves count: {}".format(len(linked_wall_curves or [])))

    # NEW: collect linked doors & openings as XY AABBs (2 ft pad)
    linked_open_aabbs = _collect_linked_opening_aabbs(doc, pad_ft=2.0)
    log.info("Linked doors/openings (AABBs): {}".format(len(linked_open_aabbs)))


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
        wall = None
        # Strict pass: apply corner/door rules; no “relax plan”
        for loop in loops:
            for seg in loop:
                seg_count += 1
                curve = segment_curve(seg)
                # ── NEW: only if this is a LINKED wall boundary, check the neighbor pair rule
                if skip_pair_set and wall is None:
                    # ensure this boundary lies on a linked wall curve
                    if _segment_is_from_linked_wall(curve, linked_wall_curves, tol_ft=0.2):
                        this_id = sp.Id.IntegerValue
                        this_cat = (cat or "").strip()
                        gk = _seg_gkey2d(curve)
                        mk = _seg_mkey2d(curve)

                        # Only get chatty if relevant categories or it is a linked boundary
                        is_linked_boundary = (wall is None) and _segment_is_from_linked_wall(curve, linked_wall_curves,
                                                                                             tol_ft=0.2)

                        log.info(u"[PAIR] space={} cat='{}' seg-mid={} gk={} mk={} linked={}".format(
                            sp.Id.IntegerValue, (cat or u"").strip(), _fmt_xy(pm), gk, mk, is_linked_boundary))
                        # find neighbors via either geom-key or midpoint-key
                        neigh = (shared_ix_g.get(gk) if gk in shared_ix_g else shared_ix_m.get(mk, []))
                        log.info(u"[PAIR] neighbors via key → {}".format(_dump_neighbors(neigh, this_id)))
                        # If the other side belongs to a category in the skip pair set, skip this segment.
                        blocked = False
                        for sid, ncat in (neigh or []):
                            if sid == this_id:
                                continue
                            pair = tuple(sorted((this_cat, (ncat or "").strip())))
                            if pair in skip_pair_set:
                                log.debug("[SKIP-PAIR] linked boundary '{}' | '{}' → skipping segment".format(this_cat,
                                                                                                              (
                                                                                                                          ncat or "").strip()))
                                blocked = True
                                break
                        if blocked:
                            continue  # skip this boundary segment entirely



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
                            # If the family is actually hosted to a wall face, skip the near-wall cleanup.
                            try:
                                if getattr(inst, "Host", None) is not None:
                                    # hosted: do not delete based on near-wall distance
                                    pass
                                else:
                                    d = _nearest_wall_xy_distance(p, all_wall_curves)
                                    # If we couldn’t measure (no curves), do NOT delete.
                                    if d is None:
                                        log.warning("Near-wall check skipped (no wall curves). Keeping instance.")
                                    else:
                                        if d > near_wall_ft:
                                            doc.Delete(inst.Id)
                                            placed_here -= 1
                                            log.info(
                                                u"Deleted (no wall within {:.2f} ft): d≈{:.2f} ft".format(near_wall_ft,
                                                                                                          d))
                            except Exception as ex:
                                log.warning("Near-wall cleanup error: {}".format(ex))

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