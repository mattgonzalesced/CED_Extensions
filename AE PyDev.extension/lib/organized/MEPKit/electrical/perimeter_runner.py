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


#-------------Ensure minimum recepts get placed---------
def ensure_min_points_for_space(all_samples_xy, post_points_xy, wall_segments, min_count,
                                constraints, logger, doc, first_ft, next_ft, inset_ft):
    """
    all_samples_xy: list[XYZ] before door/corner filtering
    post_points_xy: list[XYZ] after filtering (what you currently place)
    wall_segments:  list[(curve, host_wall or None)] that you sampled from
    constraints:    dict with avoid_corners_ft, avoid_doors_radius_ft, etc.
    """
    pts = list(post_points_xy)
    if len(pts) >= min_count:
        return pts[:min_count]

    need = min_count - len(pts)
    log = logger

    # --- Step 1: relax corner & door margins progressively
    ac = float(constraints.get('avoid_corners_ft', 2.0))
    ad = float(constraints.get('avoid_doors_radius_ft', 2.0))
    relax_plan = [
        dict(avoid_corners_ft=max(ac * 0.5, 0.0), avoid_doors_radius_ft=ad),   # halve corner margin
        dict(avoid_corners_ft=0.0,               avoid_doors_radius_ft=ad),   # no corner margin
        dict(avoid_corners_ft=0.0,               avoid_doors_radius_ft=max(ad * 0.5, 0.0)),  # halve door
        dict(avoid_corners_ft=0.0,               avoid_doors_radius_ft=0.0),  # no door margin
    ]
    # Re-filter from all_samples with relaxed margins
    for i, r in enumerate(relax_plan, 1):
        if len(pts) >= min_count: break
        cand = _refilter_with(constraints, r, wall_segments, doc, first_ft, next_ft, inset_ft, logger=log)
        # keep only new points we don't already have
        merged = _unique_by_xy(pts + cand)
        if len(merged) > len(pts):
            if log: log.info(u"Relax step {} → +{} pts (corners={}, doors={})"
                             .format(i, len(merged)-len(pts), r['avoid_corners_ft'], r['avoid_doors_radius_ft']))
            pts = merged
    if len(pts) >= min_count:
        return pts[:min_count]

    # --- Step 2: last-resort geometric fallback on segments
    # pick midpoints of the two longest wall segments (if available)
    segs = [(c, w) for (c, w) in wall_segments if c is not None]
    segs.sort(key=lambda cw: _curve_len(cw[0]), reverse=True)

    fallback_pts = []
    if len(segs) >= 2:
        p1 = _midpoint(segs[0][0]); p2 = _midpoint(segs[1][0])
        if p1: fallback_pts.append(p1)
        if p2: fallback_pts.append(p2)
    elif len(segs) == 1:
        a, b = _two_points_on_one_curve(segs[0][0])
        if a: fallback_pts.append(a)
        if b: fallback_pts.append(b)

    if fallback_pts:
        if log: log.warning(u"Fallback midpoint placement: +{} pts".format(len(fallback_pts)))
        pts = _unique_by_xy(pts + fallback_pts)

    return pts[:min_count]

@RunInTransaction("Electrical::PerimeterReceptsByRules")
def place_perimeter_recepts(doc, logger=None):
    log = logger or get_logger("MEPKit")

    id_rules = load_identify_rules()
    bc_rules = load_branch_rules()

    spaces = collect_spaces_or_rooms(doc)
    log.info("Spaces/Rooms found: {}".format(len(spaces)))
    if not spaces:
        return 0

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

        # mount height
        mh_in = cat_rule.get('mount_height_in', None)
        mh_ft = (float(mh_in)/12.0) if mh_in is not None else None

        # constraints (normalize)
        gcon = normalize_constraints(general.get('placement_constraints', {}))
        ccon = normalize_constraints(cat_rule.get('placement_constraints', {}))
        avoid_corners_ft      = float(ccon.get('avoid_corners_ft', gcon.get('avoid_corners_ft', 2.0)))
        avoid_doors_radius_ft = float(ccon.get('avoid_doors_radius_ft', gcon.get('avoid_doors_radius_ft', 0.0)))
        door_edge_margin_ft   = float(ccon.get('door_edge_margin_ft', gcon.get('door_edge_margin_ft', 0.0)))

        # IMPORTANT: keep perimeter inset tiny & stable; do NOT use door snap tolerance here
        inset_ft = 0.05

        # symbol candidates (now with auto-load)
        sym = None
        for cand in (cat_rule.get('device_candidates') or []):
            fam = cand.get('family'); typ = cand.get('type_catalog_name'); path = cand.get('load_from')
            if fam:
                sym = resolve_or_load_symbol(doc, fam, typ, load_path=path, logger=log)
                if sym: break
        if not sym:
            log.warning("No family symbol matched/loaded for space '{}' [{}]".format(name, cat))
            continue

        loops = boundary_loops(sp)
        if not loops:
            log.info("Space '{}' [{}] → no boundary loops (room-bounding?)".format(name, cat))
            continue

        # --- NEW: per-space accumulators ---
        placed_pts_xy = []  # points that actually placed during the first pass
        all_samples_xy = []  # pre-door filter samples (aggregated)
        post_points_xy = []  # post-door samples (aggregated)
        wall_segments = []  # [(curve, wall)]

        placed_here = 0
        seg_count = 0
        pre_pts_total = 0
        post_pts_total = 0

        for loop in loops:
            for seg in loop:
                seg_count += 1
                curve = segment_curve(seg)
                if not curve:
                    continue

                pts = sample_points_on_segment(curve, first_ft, next_ft, avoid_corners_ft, inset_ft)
                pre_pts = len(pts)
                all_samples_xy.extend(pts)
                wall = segment_host_wall(doc, seg)
                wall_segments.append((curve, wall))

                if wall and (avoid_doors_radius_ft > 0.0 or door_edge_margin_ft > 0.0):
                    doors = door_points_on_wall(doc, wall)
                    pts = filter_points_by_doors(pts, doors, avoid_doors_radius_ft, door_edge_margin_ft)

                post_pts = len(pts)
                post_points_xy.extend(pts)
                pre_pts_total += pre_pts
                post_pts_total += post_pts

                for p in pts:
                    try:
                        inst = place_hosted(doc, wall, sym, p, mounting_height_ft=mh_ft, logger=log) if wall else place_free(doc, sym, p, mounting_height_ft=mh_ft, logger=log)
                        if mh_ft is not None:
                            set_param_value(inst, "Mounting Height", mh_ft)
                        placed_here += 1
                        placed_pts_xy.append(p)
                    except Exception as ex:
                        log.warning(u"Placement failed at point → {}".format(ex))

        log.info("Space '{}' [{}] → loops={}, segs={}, pts pre/ post door = {}/{} → placed {}"
                 .format(name, cat, len(loops), seg_count, pre_pts_total, post_pts_total, placed_here))

        # --- NEW: enforce minimum per space ---
        # pull from category first, else general default (e.g., 2)
        min_ps = int(cat_rule.get('min_per_space', general.get('min_per_space_default', 2)))
        constraints = {
            'avoid_corners_ft': avoid_corners_ft,
            'avoid_doors_radius_ft': avoid_doors_radius_ft,
            'door_edge_margin_ft': door_edge_margin_ft,
        }

        if placed_here < min_ps:
            final_pts = ensure_min_points_for_space(
                all_samples_xy=all_samples_xy,
                post_points_xy=post_points_xy,
                wall_segments=wall_segments,
                min_count=min_ps,
                constraints=constraints,
                logger=log,
                # extra context for resampling:
                doc=doc, first_ft=first_ft, next_ft=next_ft, inset_ft=inset_ft
            )

            # place any missing points (dedupe vs. what we already placed)
            have = set(_xy_key(p) for p in placed_pts_xy)
            need = [p for p in final_pts if _xy_key(p) not in have]
            added = 0
            for p in need:
                try:
                    wall = _host_for_point(p, wall_segments)
                    inst = place_hosted(doc, wall, sym, p, mounting_height_ft=mh_ft, logger=log) if wall \
                        else place_free(doc, sym, p, mounting_height_ft=mh_ft, logger=log)
                    if mh_ft is not None:
                        set_param_value(inst, "Mounting Height", mh_ft)
                    placed_here += 1
                    added += 1
                except Exception as ex:
                    log.warning(u"Placement failed at fallback point → {}".format(ex))

            if added > 0:
                log.info(u"Min-per-space enforced → required={}, had={}, added={}".format(min_ps, len(have), added))

        # tally
        total += placed_here

    log.info("Total placed around perimeters: {}".format(total))
    return total