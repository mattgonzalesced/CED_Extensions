# -*- coding: utf-8 -*-
from __future__ import absolute_import

import math
import os

from Autodesk.Revit.DB import (
    XYZ,
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    RevitLinkInstance,
)

try:
    basestring
except NameError:
    basestring = (str,)

from organized.MEPKit.core.io_utils import read_json
from organized.MEPKit.core.log import get_logger
from organized.MEPKit.core.paths import rules_dir
from organized.MEPKit.core.rules import (
    categorize_space_by_name,
    load_identify_rules,
)
from organized.MEPKit.revit.appdoc import get_uidoc
from organized.MEPKit.revit.placement import place_free, place_hosted
from organized.MEPKit.revit.symbols import resolve_or_load_symbol
from organized.MEPKit.revit.spaces import (
    boundary_loops,
    collect_spaces_or_rooms,
    segment_curve,
    space_match_text,
    space_name,
)
from organized.MEPKit.revit.transactions import RunInTransaction


def _load_emergency_rules():
    path = os.path.join(rules_dir('electrical'), 'emergency_lighting.json')
    data = read_json(path, default={}) or {}
    return data.get('emergency_lighting') or {}


_SYMBOL_CACHE = {}
_SYMBOL_FAIL = object()


def _resolve_candidate_symbol(doc, candidate, logger):
    if not candidate:
        return None
    family = (candidate.get('family') or u"").strip()
    type_name = (candidate.get('type_catalog_name') or u"").strip()
    load_path = (candidate.get('load_from') or u"").strip()
    key = (family, type_name, load_path)
    cached = _SYMBOL_CACHE.get(key)
    if cached is _SYMBOL_FAIL:
        return None
    if cached:
        return cached

    attempts = []
    if family or type_name:
        attempts.append((family, type_name or None))
    attempts.append((family, None))

    for fam_name, typ_name in attempts:
        sym = resolve_or_load_symbol(doc, fam_name, typ_name, load_path=load_path or None, logger=logger)
        if sym:
            if typ_name is None and type_name:
                logger.warning(
                    u"Emergency lighting rule requested type '{}', but only '{}' was found; "
                    u"verify the type catalog or update the rules JSON."
                    .format(type_name, getattr(sym, "Name", u"<unnamed>"))
                )
            _SYMBOL_CACHE[key] = sym
            return sym
        # if a specific type failed, continue to generic attempt

    _SYMBOL_CACHE[key] = _SYMBOL_FAIL
    return None


def _resolve_symbol_for_rule(doc, rule_candidates, general_candidates, logger):
    for bucket in (rule_candidates, general_candidates):
        for candidate in bucket or []:
            sym = _resolve_candidate_symbol(doc, candidate, logger)
            if sym:
                return sym
    sym = _fallback_emergency_symbol(doc, logger)
    if sym:
        return sym
    return None


_FALLBACK_SYMBOL = None


def _fallback_emergency_symbol(doc, logger):
    global _FALLBACK_SYMBOL
    if _FALLBACK_SYMBOL is not None:
        return _FALLBACK_SYMBOL
    hints = ('emergency', 'egress', 'backup', 'night', 'exit')
    col = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_LightingFixtures).WhereElementIsElementType()
    best = None
    for sym in col:
        try:
            fam = getattr(sym, "Family", None)
            fam_name = getattr(fam, "Name", u"") or u""
        except Exception:
            fam_name = u""
        sym_name = getattr(sym, "Name", u"") or u""
        text = (fam_name + u" " + sym_name).strip().lower()
        if best is None and (fam_name or sym_name):
            best = sym
        if text and any(h in text for h in hints):
            _FALLBACK_SYMBOL = sym
            if logger:
                logger.warning(u"Using existing symbol '{}' for emergency lighting placements."
                               .format(sym_name or fam_name))
            return _FALLBACK_SYMBOL
    if best is not None:
        _FALLBACK_SYMBOL = best
        if logger:
            logger.warning(u"No emergency-labelled symbol found; using '{}' as fallback. Update rules JSON to "
                           u"target a specific family."
                           .format(getattr(best, "Name", u"<unnamed>")))
        return _FALLBACK_SYMBOL
    return None


def _coerce_positive_float(*values):
    for value in values:
        try:
            if value is None:
                continue
            if isinstance(value, basestring) and not value.strip():
                continue
            num = float(value)
            if num > 0.0:
                return num
        except Exception:
            continue
    return None


def _space_level(doc, space):
    try:
        lvl_id = getattr(space, "LevelId", None)
        if lvl_id:
            return doc.GetElement(lvl_id)
    except Exception:
        pass
    return None


def _space_test_z(doc, space):
    lvl = _space_level(doc, space)
    base = getattr(lvl, "Elevation", None)
    if base is None:
        try:
            loc = getattr(space, "Location", None)
            if loc and hasattr(loc, "Point"):
                base = loc.Point.Z
        except Exception:
            base = 0.0
    return float(base or 0.0) + 3.0


def _outer_polygon(space):
    loops = boundary_loops(space) or []
    best = None
    best_len = 0.0
    for loop in loops:
        pts = []
        total = 0.0
        for seg in loop:
            curve = segment_curve(seg)
            if not curve:
                continue
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            if not pts:
                pts.append(XYZ(p0.X, p0.Y, 0.0))
            pts.append(XYZ(p1.X, p1.Y, 0.0))
            try:
                total += float(curve.Length)
            except Exception:
                try:
                    total += float(getattr(curve, "ApproximateLength", 0.0))
                except Exception:
                    pass
        if pts and len(pts) >= 3:
            if abs(pts[0].X - pts[-1].X) < 1e-6 and abs(pts[0].Y - pts[-1].Y) < 1e-6:
                pts = pts[:-1]
        if pts and len(pts) >= 3 and total > best_len:
            best = pts
            best_len = total
    return best or []


def _corner_index(points):
    best_idx = None
    best_score = None
    for idx, pt in enumerate(points):
        score = pt.X + pt.Y
        if best_score is None or score < best_score:
            best_idx = idx
            best_score = score
    return best_idx


def _vector_xy(origin, target):
    return XYZ(target.X - origin.X, target.Y - origin.Y, 0.0)


def _normalize_xy(vec):
    mag = math.hypot(vec.X, vec.Y)
    if mag < 1e-6:
        return None
    return XYZ(vec.X / mag, vec.Y / mag, 0.0)


def _dot_xy(a, b):
    return (a.X * b.X) + (a.Y * b.Y)


def _point_in_polygon(point, polygon):
    x = point.X
    y = point.Y
    inside = False
    j = len(polygon) - 1
    for i in range(len(polygon)):
        xi = polygon[i].X
        yi = polygon[i].Y
        xj = polygon[j].X
        yj = polygon[j].Y
        intersects = ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / ((yj - yi) or 1e-9) + xi)
        if intersects:
            inside = not inside
        j = i
    return inside


def _point_within_space(doc, space, point, polygon, test_z):
    try:
        if space.IsPointInSpace(XYZ(point.X, point.Y, test_z)):
            return True
    except Exception:
        try:
            if space.IsPointInRoom(XYZ(point.X, point.Y, test_z)):
                return True
        except Exception:
            pass
    if polygon:
        return _point_in_polygon(point, polygon)
    return False


def _build_grid_points(doc, space, polygon, spacing, start_offset_ft):
    if len(polygon) < 3 or spacing <= 0.0:
        return []
    idx = _corner_index(polygon)
    if idx is None:
        return []
    corner = polygon[idx]
    prev_pt = polygon[(idx - 1) % len(polygon)]
    next_pt = polygon[(idx + 1) % len(polygon)]

    axis1 = _normalize_xy(_vector_xy(corner, next_pt))
    axis2 = _normalize_xy(_vector_xy(corner, prev_pt))
    if not axis1 or not axis2:
        return []
    if abs(_dot_xy(axis1, axis2)) > 0.999:
        axis2 = XYZ(-axis1.Y, axis1.X, 0.0)

    bisector = _normalize_xy(XYZ(axis1.X + axis2.X, axis1.Y + axis2.Y, 0.0))
    if bisector is None:
        axis2 = XYZ(-axis2.X, -axis2.Y, 0.0)
        bisector = _normalize_xy(XYZ(axis1.X + axis2.X, axis1.Y + axis2.Y, 0.0))
    if bisector is None:
        bisector = axis1

    test_z = _space_test_z(doc, space)
    start = XYZ(corner.X + bisector.X * start_offset_ft,
                corner.Y + bisector.Y * start_offset_ft,
                0.0)

    if not _point_within_space(doc, space, start, polygon, test_z):
        axis2 = XYZ(-axis2.X, -axis2.Y, 0.0)
        bisector = _normalize_xy(XYZ(axis1.X + axis2.X, axis1.Y + axis2.Y, 0.0)) or axis1
        start = XYZ(corner.X + bisector.X * start_offset_ft,
                    corner.Y + bisector.Y * start_offset_ft,
                    0.0)

    if not _point_within_space(doc, space, start, polygon, test_z):
        start = XYZ(corner.X + axis1.X * start_offset_ft,
                    corner.Y + axis1.Y * start_offset_ft,
                    0.0)

    if not _point_within_space(doc, space, start, polygon, test_z):
        return []

    offset1 = _dot_xy(axis1, _vector_xy(corner, start))
    offset2 = _dot_xy(axis2, _vector_xy(corner, start))
    if offset1 < 0.0:
        offset1 += math.ceil(-offset1 / spacing) * spacing
    if offset2 < 0.0:
        offset2 += math.ceil(-offset2 / spacing) * spacing

    max1 = max(_dot_xy(axis1, _vector_xy(corner, pt)) for pt in polygon)
    max2 = max(_dot_xy(axis2, _vector_xy(corner, pt)) for pt in polygon)

    pts = []
    seen = set()
    limit = 512
    i = 0
    while offset1 + i * spacing <= max1 + 1e-3 and i < limit:
        dist1 = offset1 + i * spacing
        j = 0
        while offset2 + j * spacing <= max2 + 1e-3 and j < limit:
            dist2 = offset2 + j * spacing
            candidate = XYZ(
                corner.X + axis1.X * dist1 + axis2.X * dist2,
                corner.Y + axis1.Y * dist1 + axis2.Y * dist2,
                0.0
            )
            key = (round(candidate.X, 4), round(candidate.Y, 4))
            if key not in seen and _point_within_space(doc, space, candidate, polygon, test_z):
                pts.append(candidate)
                seen.add(key)
            j += 1
        i += 1
    return pts


def _ceiling_underside_elev(doc, ceiling):
    elev = 0.0
    try:
        lvl = doc.GetElement(ceiling.LevelId)
        if lvl:
            elev = float(getattr(lvl, "Elevation", 0.0) or 0.0)
    except Exception:
        pass
    try:
        bip = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
        if bip and bip.HasValue:
            elev += float(bip.AsDouble() or 0.0)
    except Exception:
        pass
    try:
        p = ceiling.LookupParameter("Height Offset From Level")
        if p and p.HasValue:
            elev += float(p.AsDouble() or 0.0)
    except Exception:
        pass
    return elev


_HOST_CEILINGS = None


def _host_ceiling_elements(doc):
    global _HOST_CEILINGS
    if _HOST_CEILINGS is None:
        _HOST_CEILINGS = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Ceilings)
            .WhereElementIsNotElementType()
        )
    return _HOST_CEILINGS


def _find_host_ceiling(doc, space, point_xy):
    ceilings = _host_ceiling_elements(doc)
    level = _space_level(doc, space)
    level_id = getattr(level, "Id", None)
    tol = 0.1

    for restrict_level in (True, False):
        best = None
        best_elev = None
        for ceiling in ceilings:
            try:
                if restrict_level and level_id and getattr(ceiling, "LevelId", None) != level_id:
                    continue
            except Exception:
                pass
            bb = ceiling.get_BoundingBox(None)
            if not bb and hasattr(doc, "ActiveView"):
                try:
                    bb = ceiling.get_BoundingBox(doc.ActiveView)
                except Exception:
                    bb = None
            if not bb:
                continue
            if (bb.Min.X - tol) <= point_xy.X <= (bb.Max.X + tol) and \
               (bb.Min.Y - tol) <= point_xy.Y <= (bb.Max.Y + tol):
                elev = _ceiling_underside_elev(doc, ceiling)
                if best is None or elev > best_elev:
                    best = ceiling
                    best_elev = elev
        if best:
            return best, best_elev
    return None, None


_LINK_CEILING_SOURCES = None


def _prepare_link_ceiling_sources(doc):
    global _LINK_CEILING_SOURCES
    if _LINK_CEILING_SOURCES is not None:
        return _LINK_CEILING_SOURCES
    sources = []
    for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = inst.GetLinkDocument()
        if not link_doc:
            continue
        try:
            tf = inst.GetTotalTransform()
        except Exception:
            tf = inst.GetTransform()
        inv = tf.Inverse
        ceilings = []
        for ceiling in FilteredElementCollector(link_doc)\
                .OfCategory(BuiltInCategory.OST_Ceilings)\
                .WhereElementIsNotElementType():
            bb = ceiling.get_BoundingBox(None)
            if not bb:
                continue
            ceilings.append((ceiling, bb))
        if ceilings:
            sources.append((tf, inv, link_doc, ceilings))
    _LINK_CEILING_SOURCES = sources
    return sources


def _find_linked_ceiling(doc, point_xy):
    sources = _prepare_link_ceiling_sources(doc)
    tol = 0.1
    best_elev = None
    for tf, inv, link_doc, ceilings in sources:
        local_xy = inv.OfPoint(XYZ(point_xy.X, point_xy.Y, 0.0))
        for ceiling, bb in ceilings:
            if (bb.Min.X - tol) <= local_xy.X <= (bb.Max.X + tol) and \
               (bb.Min.Y - tol) <= local_xy.Y <= (bb.Max.Y + tol):
                elev_local = _ceiling_underside_elev(link_doc, ceiling)
                world_point = tf.OfPoint(XYZ(local_xy.X, local_xy.Y, elev_local))
                if best_elev is None or world_point.Z > best_elev:
                    best_elev = world_point.Z
    if best_elev is not None:
        return None, best_elev
    return None, None


def _ceiling_under_point(doc, space, point_xy):
    host, elev = _find_host_ceiling(doc, space, point_xy)
    if host or elev is not None:
        return host, elev
    return _find_linked_ceiling(doc, point_xy)


def _target_spaces(doc):
    uidoc = None
    try:
        uidoc = get_uidoc()
    except Exception:
        uidoc = None
    selected = []
    if uidoc:
        try:
            ids = list(uidoc.Selection.GetElementIds())
            for eid in ids:
                el = doc.GetElement(eid)
                if not el or not getattr(el, "Category", None):
                    continue
                cat_id = el.Category.Id.IntegerValue
                if cat_id == BuiltInCategory.OST_MEPSpaces.value__ or \
                   cat_id == BuiltInCategory.OST_Rooms.value__:
                    selected.append(el)
        except Exception:
            selected = []
    if selected:
        return selected
    return collect_spaces_or_rooms(doc)


@RunInTransaction("Electrical::PlaceEmergencyLighting")
def place_emergency_lighting(doc, logger=None):
    log = logger or get_logger("EmergencyLighting")
    rules_root = _load_emergency_rules()
    general = rules_root.get('general') or {}
    per_category = rules_root.get('rules_by_category') or {}

    identify_rules = load_identify_rules()
    spaces = _target_spaces(doc)
    if not spaces:
        log.warning("No spaces or rooms found to process.")
        return 0

    total = 0
    default_spacing = _coerce_positive_float(general.get('spacing_ft_default'), 30.0) or 30.0
    default_offset = _coerce_positive_float(general.get('start_offset_ft'), 10.0) or 10.0
    general_candidates = general.get('device_candidates') or []

    for space in spaces:
        match_text = space_match_text(space)
        category = categorize_space_by_name(match_text, identify_rules)
        rule = per_category.get(category) or per_category.get('Support') or {}

        spacing = _coerce_positive_float(rule.get('spacing_ft'), default_spacing, 30.0) or default_spacing
        start_offset = _coerce_positive_float(rule.get('start_offset_ft'), default_offset, 10.0) or default_offset
        symbol = _resolve_symbol_for_rule(doc, rule.get('device_candidates'), general_candidates, log)
        if not symbol:
            log.warning(u"No emergency lighting family resolved for category '{}'; skipping space '{}'."
                        .format(category, space_name(space)))
            continue

        polygon = _outer_polygon(space)
        if len(polygon) < 3:
            log.warning(u"Space '{}' has no valid boundary polygon; skipping.".format(space_name(space)))
            continue

        points = _build_grid_points(doc, space, polygon, spacing, start_offset)
        if not points:
            log.debug(u"Space '{}' produced no candidate grid points.".format(space_name(space)))
            continue

        level = _space_level(doc, space)
        level_elev = float(getattr(level, "Elevation", 0.0) or 0.0) if level else 0.0
        placed_here = 0

        for pt in points:
            host, underside = _ceiling_under_point(doc, space, pt)
            if host is None and underside is None:
                log.debug(u"No ceiling host found for point ({:.2f}, {:.2f}) in space '{}'."
                          .format(pt.X, pt.Y, space_name(space)))
                continue
            target_z = underside if underside is not None else (level_elev + default_offset)
            mount_height = target_z - level_elev
            point_xyz = XYZ(pt.X, pt.Y, target_z)
            try:
                if host is not None:
                    place_hosted(doc, host, symbol, point_xyz, mounting_height_ft=mount_height, logger=log)
                else:
                    place_free(doc, symbol, point_xyz, level=level, mounting_height_ft=mount_height, logger=log)
                placed_here += 1
            except Exception as ex:
                log.warning(u"Failed to place emergency light in space '{}': {}"
                            .format(space_name(space), ex))

        if placed_here:
            total += placed_here
            log.info(u"Space '{}' [{}]: placed {} emergency light(s)."
                     .format(space_name(space), category, placed_here))

    log.info(u"Total emergency lights placed: {}".format(total))
    return total
