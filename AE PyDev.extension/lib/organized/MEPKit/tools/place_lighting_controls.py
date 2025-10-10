# -*- coding: utf-8 -*-
# lib/organized/MEPKit/tools/place_lighting_controls.py
from __future__ import absolute_import

import math
import os

from Autodesk.Revit.DB import XYZ, LocationPoint

from organized.MEPKit.revit.appdoc import get_doc
from organized.MEPKit.revit.transactions import RunInTransaction
from organized.MEPKit.core.log import get_logger
from organized.MEPKit.core.rules import load_identify_rules, categorize_space_by_name
from organized.MEPKit.core.io_utils import read_json
from organized.MEPKit.core.paths import rules_dir
from organized.MEPKit.revit.spaces import (
    collect_spaces_or_rooms,
    space_match_text,
    boundary_loops,
    segment_curve,
    segment_host_wall,
)
from organized.MEPKit.revit.placement import place_free, place_hosted
from organized.MEPKit.revit.symbols import resolve_or_load_symbol
from organized.MEPKit.revit.doors import door_points_on_wall
from organized.MEPKit.electrical.selection import active_phase


def _load_lighting_rules():
    path = os.path.join(rules_dir('electrical'), 'lighting_controls.json')
    return read_json(path, default={}) or {}


def _space_center_point(space):
    loc = getattr(space, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point
    bb = space.get_BoundingBox(None)
    if bb:
        return XYZ(
            (bb.Min.X + bb.Max.X) * 0.5,
            (bb.Min.Y + bb.Max.Y) * 0.5,
            (bb.Min.Z + bb.Max.Z) * 0.5,
        )
    return None


def _category_rules(rules, category, fallback="Support"):
    lighting = rules.get('lighting_controls') or {}
    sensor_rules = lighting.get('lighting_controls_rules_by_sensors') or {}
    switch_rules = lighting.get('lighting_controls_rules_by_switches') or {}
    sensor_rule = sensor_rules.get(category) or {}
    switch_rule = switch_rules.get(category) or {}
    if not sensor_rule and fallback:
        sensor_rule = sensor_rules.get(fallback) or {}
    if not switch_rule and fallback:
        switch_rule = switch_rules.get(fallback) or {}
    general = lighting.get('general') or {}
    return sensor_rule, switch_rule, general


_SYMBOL_CACHE = {}


def _resolve_candidate_symbol(doc, candidate, logger):
    if not candidate:
        return None
    fam = candidate.get('family')
    typ = candidate.get('type_catalog_name')
    load_path = candidate.get('load_from')
    key = (fam or u"", typ or u"", load_path or u"")
    hit = _SYMBOL_CACHE.get(key)
    if hit:
        return hit
    path = load_path
    if path and not os.path.exists(path):
        if logger:
            logger.warning(u"[LOAD] Family path missing -> {} (continuing without load)".format(path))
        path = None
    sym = resolve_or_load_symbol(doc, fam, typ, load_path=path, logger=logger)
    if sym:
        _SYMBOL_CACHE[key] = sym
    return sym


def _resolve_first_available_symbol(doc, candidates, logger):
    for cand in (candidates or []):
        sym = _resolve_candidate_symbol(doc, cand, logger)
        if sym:
            return sym
    return None


def _unit_xy_from_curve(curve):
    try:
        der = curve.ComputeDerivatives(0.5, True)
        vec = der.BasisX
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            vec = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
        except Exception:
            return (1.0, 0.0)
    mag = math.hypot(vec.X, vec.Y)
    if mag < 1e-6:
        return (1.0, 0.0)
    return (vec.X / mag, vec.Y / mag)


def _space_inward_normal(space, door_point, dir_xy, probe_ft=0.5):
    normals = [(-dir_xy[1], dir_xy[0]), (dir_xy[1], -dir_xy[0])]
    for nx, ny in normals:
        probe = XYZ(door_point.X + nx * probe_ft,
                    door_point.Y + ny * probe_ft,
                    door_point.Z)
        try:
            if space.IsPointInSpace(probe):
                return (nx, ny)
        except Exception:
            continue
    return None


def _door_belongs_to_space(door, space, phase):
    if not door or not space:
        return False
    if phase is None:
        return True
    for attr in ("ToRoom", "FromRoom"):
        try:
            accessor = getattr(door, attr, None)
            if accessor:
                room = accessor[phase]
                if room and room.Id == space.Id:
                    return True
        except Exception:
            continue
    return False


def _switch_point_for_door(space, wall, door_point, near_door_ft):
    if not wall or not door_point:
        return None
    loc = getattr(wall, "Location", None)
    curve = getattr(loc, "Curve", None)
    if curve is None:
        return None
    dir_xy = _unit_xy_from_curve(curve)
    inward_xy = _space_inward_normal(space, door_point, dir_xy)
    if inward_xy is None:
        return None

    for sign in (1.0, -1.0):
        dx = dir_xy[0] * near_door_ft * sign
        dy = dir_xy[1] * near_door_ft * sign
        candidate = XYZ(door_point.X + dx, door_point.Y + dy, door_point.Z)
        probe = XYZ(candidate.X + inward_xy[0] * 0.5,
                    candidate.Y + inward_xy[1] * 0.5,
                    candidate.Z)
        try:
            if space.IsPointInSpace(probe):
                return candidate
        except Exception:
            continue

    fallback = XYZ(door_point.X + inward_xy[0] * 0.25,
                   door_point.Y + inward_xy[1] * 0.25,
                   door_point.Z)
    try:
        if space.IsPointInSpace(fallback):
            return fallback
    except Exception:
        pass
    return None


def _space_level(doc, space):
    try:
        lvl_id = getattr(space, "LevelId", None)
        if lvl_id:
            return doc.GetElement(lvl_id)
    except Exception:
        pass
    return None


def _place_switch(doc, symbol, wall, point, mounting_height_ft, logger, level=None):
    if not symbol or not point:
        return None
    try:
        return place_hosted(doc, wall, symbol, point, mounting_height_ft=mounting_height_ft, logger=logger)
    except Exception as ex:
        if logger:
            logger.warning(u"[HOST] Wall placement failed ({}) -> trying free placement".format(ex))
        try:
            return place_free(doc, symbol, point, level=level, mounting_height_ft=mounting_height_ft, logger=logger)
        except Exception as inner:
            if logger:
                logger.error(u"[PLACE] Switch placement failed: {}".format(inner))
    return None


@RunInTransaction("Electrical::PlaceLightingControls")
def place_lighting_controls(doc, logger=None):
    log = logger or get_logger("LightingControls")
    rules = _load_lighting_rules()
    id_rules = load_identify_rules()
    spaces = collect_spaces_or_rooms(doc)
    phase = active_phase(doc)

    log.info(u"Spaces/Rooms found: {}".format(len(spaces)))

    occ_count = 0
    switch_count = 0

    for space in spaces:
        match_text = space_match_text(space)
        category = categorize_space_by_name(match_text, id_rules)
        sensor_rule, switch_rule, general = _category_rules(rules, category)
        sensor_candidates = (sensor_rule.get('device_candidates') or [])
        switch_candidates = (switch_rule.get('device_candidates') or [])

        if not sensor_candidates:
            log.debug(u"No occupancy device candidates for '{}' [{}]".format(
                getattr(space, "Name", u"<unnamed>"), category))
        if not switch_candidates:
            log.debug(u"No switch device candidates for '{}' [{}]".format(
                getattr(space, "Name", u"<unnamed>"), category))

        occ_symbol = _resolve_first_available_symbol(doc, sensor_candidates, log)
        switch_symbol = _resolve_first_available_symbol(doc, switch_candidates, log)

        level = _space_level(doc, space)

        center_pt = _space_center_point(space)
        if occ_symbol and center_pt:
            try:
                place_free(doc, occ_symbol, center_pt, level=level, mounting_height_ft=None, logger=log)
                occ_count += 1
            except Exception as ex:
                log.warning(u"[PLACE] Occupancy placement failed for '{}' -> {}".format(
                    getattr(space, "Name", u"<unnamed>"), ex))
        else:
            log.debug(u"No occupancy symbol or center point for '{}'".format(
                getattr(space, "Name", u"<unnamed>")))

        if not switch_symbol:
            log.debug(u"No switch symbol resolved for '{}'".format(
                getattr(space, "Name", u"<unnamed>")))
            continue

        constraints = switch_rule.get('placement_constraints', {}) or {}
        general_constraints = (general.get('placement_constraints') or {}) if general else {}
        near_door_ft = float(constraints.get('place_near_door_ft',
                                             general_constraints.get('place_near_door_ft', 2.0)))
        mounting_height_ft = 4.0

        loops = boundary_loops(space) or []
        processed_doors = set()

        for loop in loops:
            for seg in loop:
                curve = segment_curve(seg)
                wall = segment_host_wall(doc, seg)
                if wall is None or curve is None:
                    continue
                for door, door_point in door_points_on_wall(doc, wall):
                    door_id = door.Id.IntegerValue if door else None
                    if door_id in processed_doors:
                        continue
                    if not _door_belongs_to_space(door, space, phase):
                        continue

                    point = _switch_point_for_door(space, wall, door_point, near_door_ft)
                    if not point:
                        log.debug(u"[SKIP] No valid switch point near door {} for space '{}'".format(
                            door_id, getattr(space, "Name", u"<unnamed>")))
                        continue

                    inst = _place_switch(doc, switch_symbol, wall, point,
                                         mounting_height_ft=mounting_height_ft,
                                         logger=log, level=level)
                    if inst:
                        processed_doors.add(door_id)
                        switch_count += 1
                    else:
                        log.warning(u"[PLACE] Switch placement failed for door {} in space '{}'".format(
                            door_id, getattr(space, "Name", u"<unnamed>")))

    log.info(u"Occupancy-style devices placed: {}".format(occ_count))
    log.info(u"Door switches placed: {}".format(switch_count))
    return {"occupancy": occ_count, "switches": switch_count}
