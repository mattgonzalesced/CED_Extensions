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

from Autodesk.Revit.DB import Wall, FilteredElementCollector


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


@RunInTransaction("Electrical::PerimeterReceptsByRules")
def place_perimeter_recepts(doc, logger=None):
    log = logger or get_logger("MEPKit")

    id_rules = load_identify_rules()
    bc_rules = load_branch_rules()

    spaces = collect_spaces_or_rooms(doc)
    log.info("Spaces/Rooms found: {}".format(len(spaces)))
    if not spaces:
        return 0

    all_walls = _collect_walls(doc)


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

                post_pts_total += len(pts)

                # place what survived filtering
                for p in pts:
                    try:
                        if wall:
                            inst = place_hosted(doc, wall, sym, p, mounting_height_ft=mh_ft, logger=log)
                        else:
                            inst = place_free(doc, sym, p, mounting_height_ft=mh_ft, logger=log)
                        if mh_ft is not None:
                            set_param_value(inst, "Mounting Height", mh_ft)
                        # NEW: delete if not near any wall (0.5 ft threshold)
                        try:
                            d = _nearest_wall_xy_distance(p, all_walls)
                            if d > 0.5:
                                doc.Delete(inst.Id)
                                placed_here -= 1
                                log.info(u"Deleted receptacle (> 0.5 ft from wall): d≈{:.2f} ft".format(d))
                        except Exception as ex:
                            log.warning(u"Proximity check/delete failed: {}".format(ex))

                        placed_here += 1
                    except Exception as ex:
                        log.warning(u"Placement failed at point → {}".format(ex))

        log.info("Space '{}' [{}] → loops={}, segs={}, pts pre/ post door = {}/{} → placed {}"
                 .format(name, cat, len(loops), seg_count, pre_pts_total, post_pts_total, placed_here))

        total += placed_here

    log.info("Total placed around perimeters: {}".format(total))
    return total