# -*- coding: utf-8 -*-
__title__ = "Name Piping Systems"
__doc__ = "Place tags/text notes on pipes based on connected pipe traversal."

import math

from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()


PIPE_CATEGORY_IDS = set()
for _name in ("OST_PipeCurves", "OST_FlexPipe", "OST_PipePlaceholder", "OST_FabricationPipework"):
    try:
        _bic = getattr(DB.BuiltInCategory, _name)
    except Exception:
        _bic = None
    if _bic is not None:
        try:
            PIPE_CATEGORY_IDS.add(int(_bic))
        except Exception:
            pass

FITTING_CATEGORY_IDS = set()
for _name in ("OST_PipeFitting",):
    try:
        _bic = getattr(DB.BuiltInCategory, _name)
    except Exception:
        _bic = None
    if _bic is not None:
        try:
            FITTING_CATEGORY_IDS.add(int(_bic))
        except Exception:
            pass

class _PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):  # noqa: N802
        return _is_pipe_like(elem)

    def AllowReference(self, reference, position):  # noqa: N802
        return False



def _is_pipe_like(elem):
    if elem is None:
        return False
    try:
        cat = elem.Category
    except Exception:
        cat = None
    if cat is None:
        return False
    try:
        if cat.Id.IntegerValue in PIPE_CATEGORY_IDS:
            return True
    except Exception:
        pass
    try:
        if _is_pipe_fitting(elem):
            return False
    except Exception:
        pass
    try:
        loc = elem.Location
    except Exception:
        loc = None
    if loc is not None and hasattr(loc, "Curve") and loc.Curve is not None:
        return True
    return False


def _is_pipe_fitting(elem):
    if elem is None:
        return False
    try:
        cat = elem.Category
    except Exception:
        cat = None
    if cat is None:
        return False
    try:
        return cat.Id.IntegerValue in FITTING_CATEGORY_IDS
    except Exception:
        return False


BALL_VALVE_KEYWORDS = ("BALL VALVE", "BALLVALVE", "BALL-VALVE", "BALL_VALVE")
BALL_VALVE_FAMILY = "GENERIC ANNOTATIONS"
BALL_VALVE_TYPE = "BALL_VALVE"
BALL_VALVE_MAX_OFFSET_FT = 2.0 / 12.0

_MECH_EQUIP_VIEW_CACHE = {}
_ROOF_RUNNER_BOXES = {}
_MECH_EQUIP_BOXES = {}
ROOF_RUNNER_FAMILY = "CED-R-ROOF RUNNER"
ROOF_RUNNER_TYPE = "XXX"


def _mechanical_equipment_ids_in_view(view):
    if view is None:
        return None
    try:
        vid = view.Id.IntegerValue
    except Exception:
        return None
    cached = _MECH_EQUIP_VIEW_CACHE.get(vid)
    if cached is not None:
        return cached
    try:
        mec_cat = DB.BuiltInCategory.OST_MechanicalEquipment
    except Exception:
        return None
    try:
        collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(mec_cat)
    except Exception:
        return None
    ids = set()
    try:
        for elem in collector.WhereElementIsNotElementType():
            try:
                ids.add(elem.Id.IntegerValue)
            except Exception:
                continue
    except Exception:
        pass
    _MECH_EQUIP_VIEW_CACHE[vid] = ids
    return ids


def _mechanical_equipment_bboxes(view):
    if view is None:
        return []
    try:
        vid = view.Id.IntegerValue
    except Exception:
        return []
    cached = _MECH_EQUIP_BOXES.get(vid)
    if cached is not None:
        return cached
    boxes = []
    try:
        mec_cat = DB.BuiltInCategory.OST_MechanicalEquipment
    except Exception:
        return []
    try:
        collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(mec_cat)
    except Exception:
        return []
    for elem in collector.WhereElementIsNotElementType():
        try:
            sym = elem.Symbol
        except Exception:
            sym = None
        try:
            fam_name = sym.Family.Name if sym and sym.Family else ""
        except Exception:
            fam_name = ""
        if ROOF_RUNNER_FAMILY.upper() in (fam_name or "").upper():
            continue
        bbox = None
        try:
            bbox = elem.get_BoundingBox(view)
        except Exception:
            bbox = None
        if not bbox:
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
        if not bbox:
            continue
        if bbox.Min is None or bbox.Max is None:
            continue
        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            elem_id = None
        boxes.append((elem_id, bbox.Min, bbox.Max, elem))
    _MECH_EQUIP_BOXES[vid] = boxes
    return boxes


def _equipment_label_from_intersection(pipe, view):
    if pipe is None or view is None:
        return None, None
    curve = _pipe_curve(pipe)
    if curve is None:
        return None, None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        mid = curve.Evaluate(0.5, True)
    except Exception:
        return None, None
    best_label = None
    best_id = None
    best_dist = None
    for elem_id, min_pt, max_pt, elem in _mechanical_equipment_bboxes(view):
        if not _segment_intersects_rect_xy(p0, p1, min_pt, max_pt):
            continue
        label = _format_identity_mark(_get_leaf_identity_value(elem))
        if not label:
            label = "XXXX"
        center = (min_pt + max_pt) * 0.5
        try:
            dist = mid.DistanceTo(center)
        except Exception:
            dist = 0.0
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best_label = label
            best_id = elem_id
    return best_label, best_id


def _roof_runner_bboxes(view):
    if view is None:
        return []
    try:
        vid = view.Id.IntegerValue
    except Exception:
        return []
    cached = _ROOF_RUNNER_BOXES.get(vid)
    if cached is not None:
        return cached
    boxes = []
    try:
        mec_cat = DB.BuiltInCategory.OST_MechanicalEquipment
    except Exception:
        return []
    try:
        collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(mec_cat)
    except Exception:
        return []
    for elem in collector.WhereElementIsNotElementType():
        try:
            sym = elem.Symbol
        except Exception:
            sym = None
        try:
            fam_name = sym.Family.Name if sym and sym.Family else ""
        except Exception:
            fam_name = ""
        try:
            type_name = sym.Name if sym else ""
        except Exception:
            type_name = ""
        try:
            elem_name = elem.Name or ""
        except Exception:
            elem_name = ""
        fam_upper = (fam_name or "").upper()
        type_upper = (type_name or "").upper()
        elem_upper = (elem_name or "").upper()
        if ROOF_RUNNER_FAMILY.upper() not in fam_upper:
            continue
        if ROOF_RUNNER_TYPE.upper() not in type_upper and ROOF_RUNNER_TYPE.upper() not in elem_upper:
            continue
        bbox = None
        try:
            bbox = elem.get_BoundingBox(view)
        except Exception:
            bbox = None
        if not bbox:
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
        if not bbox:
            continue
        if bbox.Min is None or bbox.Max is None:
            continue
        boxes.append((bbox.Min, bbox.Max))
    _ROOF_RUNNER_BOXES[vid] = boxes
    if not boxes:
        logger.info("Roof runner: no '{}' '{}' instances found in view {}".format(
            ROOF_RUNNER_FAMILY,
            ROOF_RUNNER_TYPE,
            vid,
        ))
    return boxes


def _segment_intersects_rect_xy(p0, p1, min_pt, max_pt):
    if p0 is None or p1 is None or min_pt is None or max_pt is None:
        return False
    minx, maxx = min_pt.X, max_pt.X
    miny, maxy = min_pt.Y, max_pt.Y
    x0, y0 = p0.X, p0.Y
    x1, y1 = p1.X, p1.Y

    if (minx <= x0 <= maxx and miny <= y0 <= maxy) or (minx <= x1 <= maxx and miny <= y1 <= maxy):
        return True

    dx = x1 - x0
    dy = y1 - y0
    p = [-dx, dx, -dy, dy]
    q = [x0 - minx, maxx - x0, y0 - miny, maxy - y0]
    u1, u2 = 0.0, 1.0
    for pi, qi in zip(p, q):
        if abs(pi) < 1e-12:
            if qi < 0:
                return False
        else:
            t = qi / pi
            if pi < 0:
                if t > u2:
                    return False
                if t > u1:
                    u1 = t
            else:
                if t < u1:
                    return False
                if t < u2:
                    u2 = t
    return True


def _pipe_hits_roof_runner(pipe, view):
    curve = _pipe_curve(pipe)
    if curve is None or view is None:
        return False
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return False
    for min_pt, max_pt in _roof_runner_bboxes(view):
        if _segment_intersects_rect_xy(p0, p1, min_pt, max_pt):
            return True
    return False



def _pipe_curve(pipe):
    try:
        loc = pipe.Location
    except Exception:
        loc = None
    if loc is None or not hasattr(loc, "Curve"):
        return None
    return loc.Curve


def _process_label(num, max_decimals=3):
    s = str(num) if num is not None else ""
    letter = ""
    if s and s[-1].isalpha():
        letter = s[-1]
        s = s[:-1]
    if "." in s:
        whole, frac = s.split(".", 1)
        if max_decimals is not None and len(frac) > max_decimals:
            frac = frac[:max_decimals]
        decimals = len(frac)
    else:
        whole, frac, decimals = s, "", 0
    if not whole:
        whole = "0"
    try:
        whole_int = int(whole)
    except Exception:
        whole_int = 0
    if decimals == 0:
        inc_num = str(whole_int + 1) + letter
        inc_digit = whole + "1" + letter
        return inc_num, inc_digit
    try:
        frac_int = int(frac) if frac else 0
    except Exception:
        frac_int = 0
    frac_int += 1
    if frac_int >= 10 ** decimals:
        whole_int += 1
        frac_int = 0
    inc_frac = str(frac_int).zfill(decimals)
    inc_num = "{}.{}{}".format(whole_int, inc_frac, letter)
    if max_decimals is not None and decimals >= max_decimals:
        inc_digit = inc_num
    else:
        inc_digit = "{}.{}1{}".format(whole, frac, letter)
    return inc_num, inc_digit


def _normalize_label_for_process(label, max_decimals=3):
    s = str(label) if label is not None else ""
    if "." in s:
        if max_decimals is None:
            return s
        head, tail = s.split(".", 1)
        letter = ""
        if tail and tail[-1].isalpha():
            letter = tail[-1]
            tail = tail[:-1]
        if len(tail) > max_decimals:
            tail = tail[:max_decimals]
        return "{}.{}{}".format(head, tail, letter)
    if not s:
        return "0.0"
    letter = ""
    if s[-1].isalpha():
        letter = s[-1]
        s = s[:-1]
    return "{}.0{}".format(s, letter)


def _element_point(elem, view=None):
    if elem is None:
        return None
    try:
        loc = elem.Location
    except Exception:
        loc = None
    if loc is not None:
        try:
            if hasattr(loc, "Point") and loc.Point:
                return loc.Point
        except Exception:
            pass
        try:
            if hasattr(loc, "Curve") and loc.Curve:
                return loc.Curve.Evaluate(0.5, True)
        except Exception:
            pass
    bbox = None
    try:
        bbox = elem.get_BoundingBox(view)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _ball_valve_annotations(view):
    elements = []
    categories = []
    for name in ("OST_GenericAnnotation", "OST_DetailComponents", "OST_DetailItems"):
        try:
            categories.append(getattr(DB.BuiltInCategory, name))
        except Exception:
            continue
    for cat in categories:
        try:
            collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(cat)
        except Exception:
            collector = DB.FilteredElementCollector(doc).OfCategory(cat)
        for elem in collector.WhereElementIsNotElementType():
            try:
                name = elem.Name or ""
            except Exception:
                name = ""
            try:
                sym = elem.Symbol
            except Exception:
                sym = None
            try:
                fam_name = sym.Family.Name if sym and sym.Family else ""
            except Exception:
                fam_name = ""
            try:
                type_name = sym.Name if sym else ""
            except Exception:
                type_name = ""
            fam_upper = (fam_name or "").upper()
            type_upper = (type_name or "").upper()
            full = "{} {} {}".format(name, fam_name, type_name).upper()
            if fam_upper == BALL_VALVE_FAMILY and type_upper == BALL_VALVE_TYPE:
                pass
            elif not any(k in full for k in BALL_VALVE_KEYWORDS):
                continue
            pt = _element_point(elem, view)
            if pt is None:
                continue
            elements.append((elem, pt))
    return elements


def _distance_point_to_curve(curve, pt):
    if curve is None or pt is None:
        return None
    try:
        proj = curve.Project(pt)
    except Exception:
        proj = None
    if proj is None:
        return None
    try:
        proj_pt = proj.XYZPoint
    except Exception:
        return None
    try:
        return pt.DistanceTo(proj_pt)
    except Exception:
        return None


def _distance_point_to_curve_xy(curve, pt):
    if curve is None or pt is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return _distance_point_to_curve(curve, pt)
    if p0 is None or p1 is None:
        return _distance_point_to_curve(curve, pt)
    try:
        vx = p1.X - p0.X
        vy = p1.Y - p0.Y
    except Exception:
        return _distance_point_to_curve(curve, pt)
    denom = vx * vx + vy * vy
    if denom <= 1e-12:
        try:
            dx = pt.X - p0.X
            dy = pt.Y - p0.Y
            return (dx * dx + dy * dy) ** 0.5
        except Exception:
            return _distance_point_to_curve(curve, pt)
    try:
        wx = pt.X - p0.X
        wy = pt.Y - p0.Y
    except Exception:
        return _distance_point_to_curve(curve, pt)
    # Use the XY segment (pipe length) so "in line" requires projection on the pipe.
    t = (wx * vx + wy * vy) / denom
    if t < 0.0 or t > 1.0:
        return None
    projx = p0.X + t * vx
    projy = p0.Y + t * vy
    dx = pt.X - projx
    dy = pt.Y - projy
    return (dx * dx + dy * dy) ** 0.5


def _nearest_ball_valve_on_pipe(pipe, valves, max_offset):
    if pipe is None or not valves:
        return None
    curve = _pipe_curve(pipe)
    if curve is None:
        return None
    best = None
    best_dist = None
    for elem, pt in valves:
        dist = _distance_point_to_curve_xy(curve, pt)
        if dist is None:
            continue
        if max_offset is not None and dist > max_offset:
            continue
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best = (elem, pt)
    return best



def _distance_along_curve(curve, from_pt, to_pt):
    if curve is None or from_pt is None or to_pt is None:
        return 0.0
    try:
        proj_from = curve.Project(from_pt)
        proj_to = curve.Project(to_pt)
    except Exception:
        return 0.0
    if proj_from is None or proj_to is None:
        return 0.0
    try:
        param0 = curve.GetEndParameter(0)
        param1 = curve.GetEndParameter(1)
        p_from = proj_from.Parameter
        p_to = proj_to.Parameter
    except Exception:
        return 0.0
    if abs(param1 - param0) < 1e-9:
        return 0.0
    try:
        frac_from = (p_from - param0) / (param1 - param0)
        frac_to = (p_to - param0) / (param1 - param0)
        length = curve.Length
    except Exception:
        return 0.0
    return abs(frac_to - frac_from) * length


def _connection_point_to_prev(pipe, prev_id):
    if pipe is None or prev_id is None:
        return None
    fitting_ids = _fittings_connecting(pipe, prev_id)
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None:
                continue
            if _is_pipe_like(owner) and owner.Id.IntegerValue == prev_id:
                try:
                    return conn.Origin
                except Exception:
                    return None
            if _is_pipe_fitting(owner) and owner.Id.IntegerValue in fitting_ids:
                try:
                    return conn.Origin
                except Exception:
                    return None
    return None


def _oriented_pipe_direction(pipe, prev_pipe):
    dir_xy = _pipe_direction_xy(pipe)
    if dir_xy is None:
        return None
    if prev_pipe is None:
        return dir_xy
    curr_pt = _element_point(pipe)
    prev_pt = _element_point(prev_pipe)
    if curr_pt is None or prev_pt is None:
        return dir_xy
    try:
        vec = prev_pt - curr_pt
    except Exception:
        return dir_xy
    try:
        if dir_xy.DotProduct(vec) > 0:
            dir_xy = DB.XYZ(-dir_xy.X, -dir_xy.Y, -dir_xy.Z)
    except Exception:
        pass
    return dir_xy


def _equipment_label_near_point_with_id(point, direction, view=None):
    if point is None:
        return None, None
    try:
        mec_cat = DB.BuiltInCategory.OST_MechanicalEquipment
    except Exception:
        return None, None
    try:
        if view is not None:
            collector = DB.FilteredElementCollector(doc, view.Id)
        else:
            collector = DB.FilteredElementCollector(doc)
    except Exception:
        return None, None
    best_label = None
    best_id = None
    best_metric = None
    for elem in collector.OfCategory(mec_cat).WhereElementIsNotElementType():
        label = _get_leaf_identity_value(elem)
        if not label:
            continue
        label = _format_identity_mark(label)
        if not label:
            continue
        pt = _element_point(elem, view)
        if pt is None:
            continue
        try:
            vec = pt - point
        except Exception:
            continue
        try:
            dist = vec.GetLength()
        except Exception:
            continue
        metric = (dist, 0.0)
        if best_metric is None or metric < best_metric:
            best_metric = metric
            best_label = label
            try:
                best_id = elem.Id.IntegerValue
            except Exception:
                best_id = None
    return best_label, best_id


def _equipment_label_near_point(point, direction, view=None):
    label, _ = _equipment_label_near_point_with_id(point, direction, view)
    return label


def _equipment_label_from_valve_bbox_with_id(point, view=None):
    if point is None:
        return None, None
    try:
        mec_cat = DB.BuiltInCategory.OST_MechanicalEquipment
    except Exception:
        return None, None
    try:
        if view is not None:
            collector = DB.FilteredElementCollector(doc, view.Id)
        else:
            collector = DB.FilteredElementCollector(doc)
    except Exception:
        return None, None
    for elem in collector.OfCategory(mec_cat).WhereElementIsNotElementType():
        label = _get_leaf_identity_value(elem)
        if not label:
            continue
        label = _format_identity_mark(label)
        if not label:
            continue
        bbox = None
        try:
            bbox = elem.get_BoundingBox(view)
        except Exception:
            bbox = None
        if not bbox:
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
        if not bbox:
            continue
        min_pt = bbox.Min
        max_pt = bbox.Max
        if min_pt is None or max_pt is None:
            continue
        if (min_pt.X <= point.X <= max_pt.X) and (min_pt.Y <= point.Y <= max_pt.Y):
            try:
                return label, elem.Id.IntegerValue
            except Exception:
                return label, None
    return None, None


def _equipment_label_from_valve_bbox(point, view=None):
    label, _ = _equipment_label_from_valve_bbox_with_id(point, view)
    return label




def _has_dot_comment(pipe):
    if pipe is None:
        return False
    try:
        param = pipe.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    except Exception:
        param = None
    if param:
        try:
            value = param.AsString() or param.AsValueString()
        except Exception:
            value = None
    else:
        try:
            value = pipe.LookupParameter("Comments").AsString()
        except Exception:
            value = None
    if not value:
        return False
    return "DOT" in value.upper()


def _element_center(elem, view=None):
    if elem is None:
        return None
    bbox = None
    try:
        bbox = elem.get_BoundingBox(view)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _equipment_label_for_terminal_with_id(pipe, view=None):
    if pipe is None:
        return None, None
    mech_ids = _mechanical_equipment_ids_in_view(view) if view is not None else None
    best_label = None
    best_id = None
    best_dist = None
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        has_pipe = False
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None:
                continue
            if _is_pipe_like(owner):
                has_pipe = True
                break
        if has_pipe:
            continue
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or _is_pipe_like(owner):
                continue
            if mech_ids is not None:
                try:
                    if owner.Id.IntegerValue not in mech_ids:
                        continue
                except Exception:
                    continue
            mark = _get_leaf_identity_value(owner)
            if not mark:
                continue
            label = _format_identity_mark(mark)
            if not label:
                continue
            try:
                ref_origin = ref.Origin
            except Exception:
                ref_origin = None
            if ref_origin is None:
                ref_origin = _element_center(owner)
            if ref_origin is None:
                dist = 0.0
            else:
                try:
                    dist = conn.Origin.DistanceTo(ref_origin)
                except Exception:
                    dist = 0.0
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best_label = label
                try:
                    best_id = owner.Id.IntegerValue
                except Exception:
                    best_id = None
    if best_label:
        return best_label, best_id
    return _equipment_label_for_pipe_with_id(pipe, view)


def _equipment_label_for_terminal(pipe, view=None):
    label, _ = _equipment_label_for_terminal_with_id(pipe, view)
    return label


def _is_inline_fitting(fitting):
    if fitting is None:
        return False
    try:
        ccount = len(_get_connectors(fitting))
    except Exception:
        ccount = None
    if ccount is None:
        return False
    return ccount == 2


def _pipes_through_fitting(start_fitting, max_depth=4):
    pipes = {}
    if start_fitting is None:
        return []
    visited = set()
    stack = [(start_fitting, 0)]
    while stack:
        fitting, depth = stack.pop()
        if fitting is None:
            continue
        fid = fitting.Id.IntegerValue
        if fid in visited:
            continue
        visited.add(fid)
        for conn in _get_connectors(fitting):
            if conn is None:
                continue
            try:
                refs = conn.AllRefs
            except Exception:
                refs = []
            for ref in refs:
                try:
                    owner = ref.Owner
                except Exception:
                    owner = None
                if owner is None:
                    continue
                if _is_pipe_like(owner):
                    pipes[owner.Id.IntegerValue] = owner
                elif _is_pipe_fitting(owner):
                    if depth + 1 > max_depth:
                        continue
                    if _is_inline_fitting(owner):
                        stack.append((owner, depth + 1))
    return list(pipes.values())


def _fitting_connected_pipes(fitting):
    if fitting is None:
        return []
    return _pipes_through_fitting(fitting)


def _fittings_connecting(pipe, prev_id):
    if pipe is None or prev_id is None:
        return set()
    fittings = set()
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or not _is_pipe_fitting(owner):
                continue
            pipes = _fitting_connected_pipes(owner)
            for p in pipes:
                if p.Id.IntegerValue == prev_id:
                    fittings.add(owner.Id.IntegerValue)
                    break
    return fittings


def _pipe_neighbors(pipe, prev_id=None):
    neighbors = {}
    ignore_fittings = _fittings_connecting(pipe, prev_id)
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or owner.Id == pipe.Id:
                continue
            if _is_pipe_like(owner):
                pid = owner.Id.IntegerValue
                if pid not in neighbors:
                    neighbors[pid] = {
                        "pipe": owner,
                        "fitting_id": None,
                        "fitting": None,
                        "is_branch": False,
                    }
                continue
            if _is_pipe_fitting(owner):
                fid = owner.Id.IntegerValue
                if fid in ignore_fittings:
                    continue
                pipes = _fitting_connected_pipes(owner)
                pipe_count = len(pipes)
                is_branch = pipe_count >= 3
                for p in pipes:
                    if p.Id == pipe.Id:
                        continue
                    pid = p.Id.IntegerValue
                    entry = neighbors.get(pid)
                    if entry is None:
                        neighbors[pid] = {
                            "pipe": p,
                            "fitting_id": fid,
                            "fitting": owner,
                            "is_branch": is_branch,
                        }
                    else:
                        if is_branch:
                            entry["is_branch"] = True
                            entry["fitting_id"] = fid
                            entry["fitting"] = owner
    return list(neighbors.values())


def _get_connectors(elem):
    if elem is None:
        return []
    cm = None
    try:
        cm = elem.ConnectorManager
    except Exception:
        cm = None
    if cm is None:
        try:
            mep = elem.MEPModel
            if mep:
                cm = mep.ConnectorManager
        except Exception:
            cm = None
    if cm is None:
        return []
    try:
        return list(cm.Connectors)
    except Exception:
        return []


def _connected_pipes(pipe):
    neighbors = {}
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or owner.Id == pipe.Id:
                continue
            if _is_pipe_like(owner):
                neighbors[owner.Id.IntegerValue] = owner
                continue
            for oconn in _get_connectors(owner):
                if oconn is None:
                    continue
                try:
                    orefs = oconn.AllRefs
                except Exception:
                    orefs = []
                for oref in orefs:
                    try:
                        oowner = oref.Owner
                    except Exception:
                        oowner = None
                    if oowner is None or oowner.Id == pipe.Id:
                        continue
                    if _is_pipe_like(oowner):
                        neighbors[oowner.Id.IntegerValue] = oowner
    return list(neighbors.values())


def _choose_trunk(pipe, candidates):
    if not candidates:
        return None
    dir_curr = _pipe_direction_xy(pipe)
    if dir_curr is None:
        return sorted(candidates, key=lambda x: x.Id.IntegerValue)[0]
    best = None
    best_score = -1.0
    for cand in candidates:
        dir_c = _pipe_direction_xy(cand)
        if dir_c is None:
            score = -1.0
        else:
            try:
                score = abs(dir_curr.DotProduct(dir_c))
            except Exception:
                score = -1.0
        if score > best_score:
            best = cand
            best_score = score
        elif score == best_score and best is not None:
            if cand.Id.IntegerValue < best.Id.IntegerValue:
                best = cand
    return best


def _get_identity_mark(elem):
    if elem is None:
        return None
    try:
        param = elem.LookupParameter("Identity Mark")
    except Exception:
        param = None
    if not param or not param.HasValue:
        return None
    try:
        value = param.AsString() or ""
    except Exception:
        try:
            value = param.AsValueString() or ""
        except Exception:
            value = ""
    value = value.strip()
    return value or None


def _get_system_first(elem):
    if elem is None:
        return None
    try:
        param = elem.LookupParameter("System First")
    except Exception:
        param = None
    if not param or not param.HasValue:
        return None
    try:
        value = param.AsString() or ""
    except Exception:
        try:
            value = param.AsValueString() or ""
        except Exception:
            value = ""
    value = value.strip()
    return value or None


def _get_leaf_identity_value(elem):
    value = _get_identity_mark(elem)
    if value:
        return value
    value = _get_system_first(elem)
    if value:
        return value
    return "XXXX"


def _format_identity_mark(mark):
    if not mark:
        return None
    cleaned = mark.strip()
    if not cleaned:
        return None
    if cleaned[-1].isalpha():
        cleaned = cleaned[:-1]
    idx = None
    for i, ch in enumerate(cleaned):
        if ch.isdigit():
            idx = i
            break
    if idx is None:
        return cleaned
    prefix = cleaned[:idx]
    digits = cleaned[idx:]
    if len(prefix) > 2:
        prefix = prefix[:-2]
        if not prefix:
            prefix = cleaned[:idx]
    return prefix + digits


def _equipment_label_for_pipe_with_id(pipe, view=None):
    mech_ids = _mechanical_equipment_ids_in_view(view) if view is not None else None
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or owner.Id == pipe.Id:
                continue
            if _is_pipe_like(owner):
                continue
            if mech_ids is not None:
                try:
                    if owner.Id.IntegerValue not in mech_ids:
                        continue
                except Exception:
                    continue
            mark = _get_leaf_identity_value(owner)
            if mark:
                try:
                    return _format_identity_mark(mark), owner.Id.IntegerValue
                except Exception:
                    return _format_identity_mark(mark), None
    return None, None


def _equipment_label_for_pipe(pipe, view=None):
    label, _ = _equipment_label_for_pipe_with_id(pipe, view)
    return label



def _pick_text_type():
    types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    if not types:
        return None
    for t in types:
        try:
            name = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            font = t.get_Parameter(DB.BuiltInParameter.TEXT_FONT).AsString() or ""
            if ("3/32" in name) and ("Arial" in font):
                return t
        except Exception:
            continue
    return types[0]



def _pipe_tag_types():
    tag_types = []
    for cat_name in ("OST_PipeTags", "OST_MultiCategoryTags"):
        try:
            cat = getattr(DB.BuiltInCategory, cat_name)
        except Exception:
            cat = None
        if cat is None:
            continue
        tag_types.extend(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilySymbol)
            .OfCategory(cat)
            .ToElements()
        )
    return tag_types



def _tag_type_label(tag_type):
    try:
        fam_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
    except Exception:
        fam_name = None
    try:
        type_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except Exception:
        type_name = None
    try:
        cat_name = tag_type.Category.Name if tag_type.Category else "Tag"
    except Exception:
        cat_name = "Tag"
    return "[{}] {} : {}".format(cat_name, fam_name or "?", type_name or "?")



def _pick_pipe_tag_type():
    tag_types = _pipe_tag_types()
    if not tag_types:
        return None
    options = [_tag_type_label(t) for t in tag_types]
    picked = forms.SelectFromList.show(
        options,
        multiselect=False,
        title="Select Pipe Tag Type",
    )
    if not picked:
        return None
    try:
        idx = options.index(picked)
    except Exception:
        return None
    return tag_types[idx]



def _place_text_notes(label_map, pipe_map, text_type, view):
    if not label_map:
        return 0
    count = 0
    with revit.Transaction("Name Piping Systems - Text Notes"):
        for pid, label in label_map.items():
            pipe = pipe_map.get(pid)
            if pipe is None:
                continue
            curve = _pipe_curve(pipe)
            if curve is None:
                continue
            try:
                pt = curve.Evaluate(0.5, True)
            except Exception:
                continue
            try:
                opts = DB.TextNoteOptions()
                opts.TypeId = text_type.Id
                try:
                    opts.HorizontalAlignment = DB.HorizontalTextAlignment.Center
                except Exception:
                    pass
                try:
                    opts.VerticalAlignment = DB.VerticalTextAlignment.Middle
                except Exception:
                    pass
                DB.TextNote.Create(doc, view.Id, pt, label, opts)
                count += 1
            except Exception:
                continue
    return count



def _pipe_direction_xy(pipe):
    curve = _pipe_curve(pipe)
    if curve is None:
        return None
    try:
        deriv = curve.ComputeDerivatives(0.5, True)
        tangent = deriv.BasisX
    except Exception:
        tangent = None
    if tangent is None:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            tangent = p1 - p0
        except Exception:
            return None
    try:
        vec = DB.XYZ(tangent.X, tangent.Y, 0.0)
        if abs(vec.X) + abs(vec.Y) < 1e-9:
            return None
        return vec.Normalize()
    except Exception:
        return None



def _rotate_element_about_z(elem_id, origin, angle):
    if elem_id is None or origin is None:
        return
    if abs(angle) < 1e-7:
        return
    axis = DB.Line.CreateUnbound(origin, DB.XYZ(0, 0, 1))
    DB.ElementTransformUtils.RotateElement(doc, elem_id, axis, angle)



def _place_pipe_tags(label_map, pipe_map, tag_type, view):
    if not label_map:
        return 0
    count = 0
    with revit.Transaction("Name Piping Systems - Tags"):
        if not tag_type.IsActive:
            tag_type.Activate()
            doc.Regenerate()
        for pid in label_map.keys():
            pipe = pipe_map.get(pid)
            if pipe is None:
                continue
            curve = _pipe_curve(pipe)
            if curve is None:
                continue
            try:
                pt = curve.Evaluate(0.5, True)
            except Exception:
                continue
            try:
                reference = DB.Reference(pipe)
                tag = DB.IndependentTag.Create(
                    doc,
                    tag_type.Id,
                    view.Id,
                    reference,
                    False,
                    DB.TagOrientation.Horizontal,
                    pt,
                )
                if tag:
                    dir_xy = _pipe_direction_xy(pipe)
                    if dir_xy is not None:
                        angle = math.atan2(dir_xy.Y, dir_xy.X)
                        try:
                            _rotate_element_about_z(tag.Id, tag.TagHeadPosition, angle)
                        except Exception:
                            _rotate_element_about_z(tag.Id, pt, angle)
                    count += 1
            except Exception:
                continue
    return count


def _set_identity_marks(label_map, pipe_map):
    if not label_map:
        return 0
    count = 0
    with revit.Transaction("Name Piping Systems - Identity Mark"):
        for pid, label in label_map.items():
            pipe = pipe_map.get(pid)
            if pipe is None:
                continue
            try:
                param = pipe.LookupParameter("Identity Mark")
            except Exception:
                param = None
            if not param or param.IsReadOnly:
                continue
            try:
                param.Set(label)
                count += 1
            except Exception:
                continue
    return count


def _log_pipe_connections(label_map, pipe_map):
    if not label_map:
        return
    logger.info("Name Piping Systems - Pipe connection list (labeled pipes)")
    for pid in sorted(label_map.keys()):
        pipe = pipe_map.get(pid)
        if pipe is None:
            continue
        neighbors = _connected_pipes(pipe)
        neighbor_ids = [n.Id.IntegerValue for n in neighbors]
        logger.info("pipe {} label {} -> connected {}".format(pid, label_map.get(pid), neighbor_ids))


def _fitting_kind(fitting):
    if fitting is None:
        return "Fitting"
    try:
        pt = fitting.MEPModel.PartType
        if pt:
            return str(pt)
    except Exception:
        pass
    name = ""
    try:
        name = "{} {}".format(fitting.Symbol.Family.Name, fitting.Symbol.Name)
    except Exception:
        try:
            name = fitting.Name
        except Exception:
            name = ""
    lname = (name or "").lower()
    if "tee" in lname:
        return "Tee"
    if "elbow" in lname:
        return "Elbow"
    if "cross" in lname:
        return "Cross"
    if "union" in lname:
        return "Union"
    try:
        ccount = len(_get_connectors(fitting))
        if ccount >= 3:
            return "Tee"
        if ccount == 2:
            return "Elbow"
    except Exception:
        pass
    return "Fitting"


def _record_fitting_connections(pipe, fitting_map, fitting_objs):
    if pipe is None:
        return
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or owner.Id == pipe.Id:
                continue
            if _is_pipe_fitting(owner):
                fid = owner.Id.IntegerValue
                if fid not in fitting_map:
                    fitting_map[fid] = set()
                fitting_map[fid].add(pipe.Id.IntegerValue)
                fitting_objs[fid] = owner


def _pipe_direct_connection_summary(pipe, fitting_objs):
    if pipe is None:
        return [], []
    pipe_ids = set()
    fitting_ids = set()
    for conn in _get_connectors(pipe):
        if conn is None:
            continue
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []
        for ref in refs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or owner.Id == pipe.Id:
                continue
            if _is_pipe_like(owner):
                pipe_ids.add(owner.Id.IntegerValue)
            elif _is_pipe_fitting(owner):
                fitting_ids.add(owner.Id.IntegerValue)
                fitting_objs[owner.Id.IntegerValue] = owner
    return sorted(pipe_ids), sorted(fitting_ids)


def _log_traversal(root_label, order_list, label_map, pipe_map, fitting_map, fitting_objs):
    logger.info("Name Piping Systems - Traversal order for root {}".format(root_label))
    for idx, pid in enumerate(order_list, 1):
        pipe = pipe_map.get(pid)
        if pipe is None:
            continue
        pipe_ids, fitting_ids = _pipe_direct_connection_summary(pipe, fitting_objs)
        fittings = []
        for fid in fitting_ids:
            kind = _fitting_kind(fitting_objs.get(fid))
            fittings.append("{}({})".format(fid, kind))
        logger.info(
            "step {} pipe {} label {} -> pipes {} fittings {}".format(
                idx,
                pid,
                label_map.get(pid),
                pipe_ids,
                fittings,
            )
        )
    if fitting_map:
        logger.info("Name Piping Systems - Fitting connections for root {}".format(root_label))
        for fid in sorted(fitting_map.keys()):
            fitting = fitting_objs.get(fid)
            kind = _fitting_kind(fitting)
            pipes = sorted(fitting_map[fid])
            logger.info("fitting {} ({}) -> pipes {}".format(fid, kind, pipes))


def _traverse_and_label(start_pipe, start_label, label_map, pipe_map, valves, view, used_valve_ids):
    visited = set()
    order_list = []
    fitting_map = {}
    fitting_objs = {}

    def _find_branch_leaf_marker(branch_pipe, prev_id):
        if branch_pipe is None:
            return None, None, None
        curr = branch_pipe
        back = prev_id
        while curr is not None:
            found = _nearest_ball_valve_on_pipe(
                curr,
                valves,
                BALL_VALVE_MAX_OFFSET_FT,
            )
            if found:
                try:
                    valve_id = found[0].Id.IntegerValue
                except Exception:
                    valve_id = None
                if valve_id is None or valve_id not in used_valve_ids:
                    return curr, "valve", found
            equip_label, equip_id = _equipment_label_from_intersection(curr, view)
            if equip_label:
                return curr, "equip", (equip_label, equip_id)
            neighbors = _pipe_neighbors(curr, back)
            if not neighbors or len(neighbors) != 1:
                break
            back = curr.Id.IntegerValue
            curr = neighbors[0]["pipe"]
        return None, None, None

    def _walk_leaf_branch(start_pipe, prev_id, branch_label, marker_pipe, marker_type, marker_info):
        if start_pipe is None:
            return False
        curr = start_pipe
        back = prev_id
        while curr is not None:
            cid = curr.Id.IntegerValue
            if cid in visited:
                return False
            visited.add(cid)
            order_list.append(cid)
            _record_fitting_connections(curr, fitting_map, fitting_objs)
            label_map[cid] = branch_label
            pipe_map[cid] = curr
            if _pipe_hits_roof_runner(curr, view):
                return True

            if marker_pipe is not None and cid == marker_pipe.Id.IntegerValue and marker_type == "equip":
                equip_label, equip_id = marker_info
                if equip_label:
                    label_map[cid] = equip_label
                    pipe_map[cid] = curr
                    logger.info(
                        "Leaf label: pipe {} -> {} (mech {})".format(
                            cid,
                            equip_label,
                            equip_id if equip_id is not None else "None",
                        )
                    )
                return True

            if marker_pipe is not None and cid == marker_pipe.Id.IntegerValue and marker_type == "valve" and marker_info:
                prev_pipe = None
                if back is not None:
                    try:
                        prev_pipe = doc.GetElement(DB.ElementId(back))
                    except Exception:
                        prev_pipe = None
                direction = _oriented_pipe_direction(curr, prev_pipe)
                valve_elem, valve_point = valve_info
                try:
                    valve_id = valve_elem.Id.IntegerValue
                except Exception:
                    valve_id = None
                if valve_id is None or valve_id not in used_valve_ids:
                    label, mech_id = _equipment_label_from_valve_bbox_with_id(valve_point, view)
                    if not label:
                        label, mech_id = _equipment_label_near_point_with_id(
                            valve_point,
                            direction,
                            view,
                        )
                    if not label:
                        label, mech_id = _equipment_label_for_terminal_with_id(curr, view)
                    if label:
                        label_map[cid] = label
                        pipe_map[cid] = curr
                        logger.info(
                            "Leaf label: pipe {} -> {} (mech {})".format(
                                cid,
                                label,
                                mech_id if mech_id is not None else "None",
                            )
                        )
                        if valve_id is not None:
                            used_valve_ids.add(valve_id)
                return True

            neighbors = _pipe_neighbors(curr, back)
            if not neighbors or len(neighbors) != 1:
                return False
            back = cid
            curr = neighbors[0]["pipe"]
        return False

    def walk_trunk(curr, label, prev_id):
        while curr is not None:
            cid = curr.Id.IntegerValue
            if cid in visited:
                return
            visited.add(cid)
            order_list.append(cid)
            _record_fitting_connections(curr, fitting_map, fitting_objs)
            label_map[cid] = label
            pipe_map[cid] = curr
            if _pipe_hits_roof_runner(curr, view):
                return
            neighbors = _pipe_neighbors(curr, prev_id)
            branch_groups = {}
            linear_neighbors = []
            for n in neighbors:
                if n["is_branch"] and n["fitting_id"] is not None:
                    branch_groups.setdefault(n["fitting_id"], []).append(n["pipe"])
                else:
                    linear_neighbors.append(n["pipe"])
            if branch_groups:
                group = None
                for pipes in branch_groups.values():
                    if group is None or len(pipes) > len(group):
                        group = pipes
                branch_pipes = group or []
                trunk_candidates = []
                leaf_candidates = []
                leaf_valves = {}
                for n in branch_pipes:
                    marker_pipe, marker_type, marker_info = _find_branch_leaf_marker(n, cid)
                    if marker_pipe is not None and marker_info:
                        leaf_candidates.append(n)
                        leaf_valves[n.Id.IntegerValue] = (marker_pipe, marker_type, marker_info)
                    else:
                        trunk_candidates.append(n)

                proc_label = _normalize_label_for_process(label)
                next_num, next_digit = _process_label(proc_label)
                if next_digit == next_num:
                    next_digit = _process_label(_normalize_label_for_process(next_digit))[0]

                if len(trunk_candidates) >= 2:
                    trunk_branch = _choose_trunk(curr, trunk_candidates)
                    if trunk_branch is None:
                        trunk_branch = trunk_candidates[0]
                    for n in sorted(trunk_candidates, key=lambda x: x.Id.IntegerValue):
                        if n.Id.IntegerValue == trunk_branch.Id.IntegerValue:
                            continue
                        walk_trunk(n, next_digit, cid)
                        next_digit = _process_label(_normalize_label_for_process(next_digit))[0]
                    walk_trunk(trunk_branch, next_num, cid)
                    for leaf in sorted(leaf_candidates, key=lambda x: x.Id.IntegerValue):
                        marker_pipe, marker_type, marker_info = leaf_valves.get(leaf.Id.IntegerValue, (None, None, None))
                        _walk_leaf_branch(leaf, cid, next_digit, marker_pipe, marker_type, marker_info)
                        next_digit = _process_label(_normalize_label_for_process(next_digit))[0]
                elif len(trunk_candidates) == 1:
                    trunk_branch = trunk_candidates[0]
                    walk_trunk(trunk_branch, next_num, cid)
                    for leaf in sorted(leaf_candidates, key=lambda x: x.Id.IntegerValue):
                        marker_pipe, marker_type, marker_info = leaf_valves.get(leaf.Id.IntegerValue, (None, None, None))
                        _walk_leaf_branch(leaf, cid, next_digit, marker_pipe, marker_type, marker_info)
                        next_digit = _process_label(_normalize_label_for_process(next_digit))[0]
                else:
                    first_leaf = None
                    for leaf in sorted(leaf_candidates, key=lambda x: x.Id.IntegerValue):
                        if first_leaf is None:
                            first_leaf = leaf
                            branch_label = next_num
                        else:
                            branch_label = next_digit
                        marker_pipe, marker_type, marker_info = leaf_valves.get(leaf.Id.IntegerValue, (None, None, None))
                        _walk_leaf_branch(leaf, cid, branch_label, marker_pipe, marker_type, marker_info)
                        if first_leaf is not None and leaf.Id.IntegerValue != first_leaf.Id.IntegerValue:
                            next_digit = _process_label(_normalize_label_for_process(next_digit))[0]
                for n in linear_neighbors:
                    walk_trunk(n, label, cid)
                return

            if not neighbors:
                return
            if len(linear_neighbors) == 1:
                prev_id = cid
                curr = linear_neighbors[0]
                continue
            if len(linear_neighbors) > 1:
                for n in linear_neighbors:
                    walk_trunk(n, label, cid)
                return

    walk_trunk(start_pipe, start_label, None)
    return order_list, fitting_map, fitting_objs



def _prompt_start_pipes():
    try:
        forms.alert(
            "Select the start pipes for this rack in order (2-5).",
            title="Select Start Pipes",
            warn_icon=False,
        )
    except Exception:
        pass
    start_pipes = []
    selected_ids = []
    while True:
        if len(start_pipes) >= 5:
            break
        prompt = "Select start pipe #{} (ESC to finish)".format(len(start_pipes) + 1)
        try:
            with forms.WarningBar(title=prompt):
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    _PipeSelectionFilter(),
                    prompt,
                )
        except Exception:
            break
        try:
            elem = doc.GetElement(ref.ElementId)
        except Exception:
            elem = None
        if not _is_pipe_like(elem):
            continue
        elem_id = elem.Id.IntegerValue
        if elem_id in selected_ids:
            continue
        start_pipes.append(elem)
        selected_ids.append(elem_id)
        try:
            uidoc.Selection.SetElementIds(List[DB.ElementId]([p.Id for p in start_pipes]))
        except Exception:
            pass

    if len(start_pipes) < 2 or len(start_pipes) > 5:
        forms.alert(
            "Select 2 to 5 start pipes. Detected {} pipe(s).".format(len(start_pipes)),
            exitscript=True,
        )
    return start_pipes



def main():
    active_view = revit.active_view
    if active_view.IsTemplate:
        forms.alert("Active view is a template. Open a working view first.", exitscript=True)

    start_pipes = _prompt_start_pipes()

    output_mode = None
    while not output_mode:
        output_mode = forms.CommandSwitchWindow.show(
            ["Tags", "Text Notes", "Both"],
            message="Place labels as tags, text notes, or both?",
        )

    tag_type = None
    text_type = None
    if output_mode in ("Tags", "Both"):
        tag_type = _pick_pipe_tag_type()
        if tag_type is None:
            forms.alert("No pipe tag types found in this project.", exitscript=True)
    if output_mode in ("Text Notes", "Both"):
        text_type = _pick_text_type()
        if text_type is None:
            forms.alert("No text note types found in this project.", exitscript=True)

    valves = _ball_valve_annotations(active_view)

    label_map = {}
    pipe_map = {}
    used_valve_ids = set()
    for idx, start_pipe in enumerate(start_pipes, 1):
        pid = start_pipe.Id.IntegerValue
        if pid in label_map:
            continue
        start_label = "{}A".format(idx)
        local_label_map = {}
        local_pipe_map = {}
        order_list, fitting_map, fitting_objs = _traverse_and_label(
            start_pipe,
            start_label,
            local_label_map,
            local_pipe_map,
            valves,
            active_view,
            used_valve_ids,
        )
        _log_traversal(
            start_label,
            order_list,
            local_label_map,
            local_pipe_map,
            fitting_map,
            fitting_objs,
        )
        for lpid, lbl in local_label_map.items():
            if lpid in label_map:
                continue
            label_map[lpid] = lbl
            pipe_map[lpid] = local_pipe_map.get(lpid)

    if not label_map:
        forms.alert("No pipes were labeled from the selected start pipes.", exitscript=True)

    marks_set = _set_identity_marks(label_map, pipe_map)
    tag_count = 0
    text_count = 0
    if output_mode in ("Tags", "Both"):
        tag_count = _place_pipe_tags(label_map, pipe_map, tag_type, active_view)
    if output_mode in ("Text Notes", "Both"):
        text_count = _place_text_notes(label_map, pipe_map, text_type, active_view)

    forms.alert(
        "Applied {} label(s). Tags: {}. Text notes: {}.".format(
            marks_set,
            tag_count,
            text_count,
        )
    )


if __name__ == "__main__":
    main()
