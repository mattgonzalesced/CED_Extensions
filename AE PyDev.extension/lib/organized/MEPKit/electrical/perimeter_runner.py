# -*- coding: utf-8 -*-
# lib/organized/MEPKit/electrical/perimeter_runner.py
from __future__ import absolute_import
from organized.MEPKit.core.log import get_logger
from organized.MEPKit.core.rules import (
    load_identify_rules, load_branch_rules, normalize_constraints,
    categorize_space_by_name, get_category_rule
)
from organized.MEPKit.revit.transactions import RunInTransaction
from organized.MEPKit.revit.spaces import collect_spaces_or_rooms, space_name, boundary_loops, segment_curve, segment_host_wall, sample_points_on_segment
from organized.MEPKit.revit.doors import door_points_on_wall, filter_points_by_doors
from organized.MEPKit.revit.symbols import resolve_or_load_symbol, place_hosted, place_free
from organized.MEPKit.revit.params import set_param_value  # optional for mounting height

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
        cat = categorize_space_by_name(name, id_rules)
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

                wall = segment_host_wall(doc, seg)
                if wall and (avoid_doors_radius_ft > 0.0 or door_edge_margin_ft > 0.0):
                    doors = door_points_on_wall(doc, wall)
                    pts = filter_points_by_doors(pts, doors, avoid_doors_radius_ft, door_edge_margin_ft)

                post_pts = len(pts)
                pre_pts_total += pre_pts
                post_pts_total += post_pts

                for p in pts:
                    try:
                        inst = place_hosted(doc, wall, sym, p) if wall else place_free(doc, sym, p)
                        if mh_ft is not None:
                            set_param_value(inst, "Mounting Height", mh_ft)
                        placed_here += 1
                    except:
                        # keep going on errors
                        pass

        log.info("Space '{}' [{}] → loops={}, segs={}, pts pre/ post door = {}/{} → placed {}"
                 .format(name, cat, len(loops), seg_count, pre_pts_total, post_pts_total, placed_here))

        total += placed_here

    log.info("Total placed around perimeters: {}".format(total))
    return total