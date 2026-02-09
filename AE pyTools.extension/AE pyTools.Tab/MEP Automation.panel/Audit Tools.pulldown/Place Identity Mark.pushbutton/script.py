# -*- coding: utf-8 -*-
"""
Place Identity Mark
-------------------
Copy Identity Mark from refrigerated cases in workset CED REFRIGERATED CASES
onto mechanical control devices hosted on or located within/above the case.
"""

from collections import defaultdict

from pyrevit import revit, forms, script, DB
from System.Collections.Generic import List

TITLE = "Place Identity Mark"
XY_TOLERANCE_FT = 0.1
ORANGE_RGB = (255, 128, 0)
ZOOM_TIGHTEN = 0.6
TARGET_CONTROLLER_FAMILY = "RCD-U_General Refrigeration Device_CED"
PANEL_VALUE = "RA, RB, RC, RD"
CIRCUIT_SUFFIX = "_CASECONTROLLER"

output = script.get_output()
output.close_others()


def _normalize(text):
    return (text or "").strip().lower()


def _get_identity_param(elem):
    if not elem:
        return None
    try:
        param = elem.LookupParameter("Identity Mark")
    except Exception:
        param = None
    if param:
        return param
    try:
        return elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    except Exception:
        return None


def _get_identity_value(elem):
    param = _get_identity_param(elem)
    if not param:
        return None
    try:
        value = param.AsString()
        if value is not None:
            return value
    except Exception:
        pass
    try:
        value = param.AsValueString()
        if value is not None:
            return value
    except Exception:
        pass
    return None


def _set_identity_value(elem, value):
    param = _get_identity_param(elem)
    if not param or param.IsReadOnly:
        return False
    try:
        param.Set(value or "")
        return True
    except Exception:
        return False


def _set_param_text(elem, name, value):
    if not elem or not name:
        return None
    try:
        param = elem.LookupParameter(name)
    except Exception:
        param = None
    if not param:
        return None
    if param.IsReadOnly:
        return False
    try:
        param.Set(value or "")
        return True
    except Exception:
        return False


def _get_bbox(elem):
    if not elem:
        return None
    bbox = None
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox:
        return bbox
    try:
        bbox = elem.get_BoundingBox(revit.active_view)
    except Exception:
        bbox = None
    return bbox


def _get_solid_fill_pattern_id(doc):
    for fpe in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp and fp.IsSolidFill:
                return fpe.Id
        except Exception:
            pass
    return DB.ElementId.InvalidElementId


def _build_orange_overrides(doc):
    ogs = DB.OverrideGraphicSettings()
    r, g, b = ORANGE_RGB
    col = DB.Color(bytearray([r])[0], bytearray([g])[0], bytearray([b])[0])
    try:
        ogs.SetProjectionLineColor(col)
    except Exception:
        pass
    try:
        ogs.SetCutLineColor(col)
    except Exception:
        pass
    solid_id = _get_solid_fill_pattern_id(doc)
    if solid_id and solid_id != DB.ElementId.InvalidElementId:
        try:
            ogs.SetSurfaceForegroundPatternId(solid_id)
            ogs.SetSurfaceForegroundPatternColor(col)
        except Exception:
            pass
        try:
            ogs.SetCutForegroundPatternId(solid_id)
            ogs.SetCutForegroundPatternColor(col)
        except Exception:
            pass
    return ogs


def _apply_overrides(view, element_ids, ogs):
    if view is None:
        return
    for elem_id in element_ids:
        try:
            view.SetElementOverrides(DB.ElementId(int(elem_id)), ogs)
        except Exception:
            continue


def _union_bbox(current_min, current_max, bbox):
    if not bbox:
        return current_min, current_max
    if current_min is None:
        return bbox.Min, bbox.Max
    return (
        DB.XYZ(
            min(current_min.X, bbox.Min.X),
            min(current_min.Y, bbox.Min.Y),
            min(current_min.Z, bbox.Min.Z),
        ),
        DB.XYZ(
            max(current_max.X, bbox.Max.X),
            max(current_max.Y, bbox.Max.Y),
            max(current_max.Z, bbox.Max.Z),
        ),
    )


def _shrink_bbox(min_point, max_point, factor):
    if min_point is None or max_point is None:
        return min_point, max_point
    factor = max(0.1, min(1.0, float(factor)))
    cx = (min_point.X + max_point.X) * 0.5
    cy = (min_point.Y + max_point.Y) * 0.5
    cz = (min_point.Z + max_point.Z) * 0.5
    dx = (max_point.X - min_point.X) * 0.5 * factor
    dy = (max_point.Y - min_point.Y) * 0.5 * factor
    dz = (max_point.Z - min_point.Z) * 0.5 * factor
    return (
        DB.XYZ(cx - dx, cy - dy, cz - dz),
        DB.XYZ(cx + dx, cy + dy, cz + dz),
    )


def _bbox_center(bbox):
    if not bbox:
        return None
    try:
        return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        return None


def _get_point(elem):
    if not elem:
        return None
    location = getattr(elem, "Location", None)
    if location is not None:
        point = getattr(location, "Point", None)
        if point:
            return point
        curve = getattr(location, "Curve", None)
        if curve:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                pass
    bbox = _get_bbox(elem)
    if bbox:
        return _bbox_center(bbox)
    return None


def _controller_label(elem):
    if elem is None:
        return ""
    try:
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            return u"{} : {}".format(fam_name, type_name)
        if fam_name:
            return fam_name
        if type_name:
            return type_name
    except Exception:
        pass
    try:
        name = getattr(elem, "Name", None)
        if name:
            return str(name)
    except Exception:
        pass
    return ""


def _controller_family(elem):
    if elem is None:
        return ""
    try:
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        if fam_name:
            return str(fam_name)
    except Exception:
        pass
    return ""


def _strip_trailing_letter(value):
    if not value:
        return ""
    text = str(value).strip()
    if len(text) >= 2 and text[-1].isalpha():
        return text[:-1]
    return text


def _point_in_bbox_xy(point, bbox, xy_tol):
    if not point or not bbox:
        return False
    min_x = bbox.Min.X - xy_tol
    min_y = bbox.Min.Y - xy_tol
    max_x = bbox.Max.X + xy_tol
    max_y = bbox.Max.Y + xy_tol
    return (
        min_x <= point.X <= max_x
        and min_y <= point.Y <= max_y
    )


def _distance(point_a, point_b):
    if not point_a or not point_b:
        return None
    try:
        return point_a.DistanceTo(point_b)
    except Exception:
        try:
            dx = point_a.X - point_b.X
            dy = point_a.Y - point_b.Y
            dz = point_a.Z - point_b.Z
            return (dx * dx + dy * dy + dz * dz) ** 0.5
        except Exception:
            return None


def _collect_cases(doc):
    cases = []
    blank_ids = []
    if not doc:
        return cases, blank_ids
    collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
    collector = collector.WhereElementIsNotElementType()
    for elem in collector:
        mark = _get_identity_value(elem)
        if not (mark or "").strip():
            blank_ids.append(elem.Id.IntegerValue)
        bbox = _get_bbox(elem)
        center = _bbox_center(bbox) if bbox else _get_point(elem)
        cases.append({
            "element": elem,
            "id": elem.Id.IntegerValue,
            "mark": mark,
            "bbox": bbox,
            "center": center,
        })
    return cases, blank_ids


def _collect_controllers(doc):
    if not doc:
        return []
    collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalControlDevices)
    collector = collector.WhereElementIsNotElementType()
    return list(collector)


def _match_case_for_controller(doc, controller, cases_by_id, cases_with_bbox):
    if controller is None:
        return None

    point = _get_point(controller)
    if point is None:
        return None

    best_case = None
    best_dist = None
    for case in cases_with_bbox:
        bbox = case.get("bbox")
        if not bbox:
            continue
        if not _point_in_bbox_xy(point, bbox, XY_TOLERANCE_FT):
            continue
        center = case.get("center")
        dist = _distance(point, center) if center else None
        if best_dist is None or (dist is not None and dist < best_dist):
            best_dist = dist
            best_case = case
    return best_case


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    cases, blank_case_ids = _collect_cases(doc)
    if not cases:
        forms.alert("No Mechanical Equipment found in the host model.", title=TITLE)
        return

    cases_by_id = {case["id"]: case for case in cases}
    cases_with_bbox = [case for case in cases if case.get("bbox")]

    controllers = _collect_controllers(doc)
    if not controllers:
        forms.alert("No Mechanical Control Devices found.", title=TITLE)
        return

    stats = defaultdict(int)
    updated_ids = []
    unmatched_ids = []
    active_view = revit.active_view

    if blank_case_ids:
        output.print_md("### Place Identity Mark - Missing Case Identity Marks")
        output.print_md("Cases with blank Identity Mark: `{}`".format(len(blank_case_ids)))
        output.print_md("**Case ElementIds (blank Identity Mark)**")
        for elem_id in blank_case_ids:
            output.print_md("- `{}`".format(elem_id))
        forms.alert(
            "Found {} refrigerated case(s) with blank Identity Mark.\n"
            "See the pyRevit output panel for the ElementId list, then fix and rerun."
            .format(len(blank_case_ids)),
            title=TITLE,
        )

    tx = DB.Transaction(doc, "Place Identity Mark")
    tx.Start()
    try:
        for controller in controllers:
            stats["controllers_total"] += 1
            fam_name = _controller_family(controller)
            if fam_name != TARGET_CONTROLLER_FAMILY:
                stats["controllers_skipped_type"] += 1
                continue
            case = _match_case_for_controller(doc, controller, cases_by_id, cases_with_bbox)
            if not case:
                stats["controllers_unmatched"] += 1
                unmatched_ids.append(controller.Id.IntegerValue)
                continue

            case_mark = case.get("mark") or ""
            base_mark = _strip_trailing_letter(case_mark)
            updates_applied = False
            readonly_hit = False

            current_identity = _get_identity_value(controller)
            if current_identity != case_mark:
                if _set_identity_value(controller, case_mark):
                    updates_applied = True
                else:
                    readonly_hit = True

            panel_result = _set_param_text(controller, "CKT_Panel_CEDT", PANEL_VALUE)
            if panel_result is True:
                updates_applied = True
            elif panel_result is False:
                readonly_hit = True

            circuit_value = "{}{}".format(base_mark, CIRCUIT_SUFFIX) if base_mark else CIRCUIT_SUFFIX
            circuit_result = _set_param_text(controller, "CKT_Circuit Number_CEDT", circuit_value)
            if circuit_result is True:
                updates_applied = True
            elif circuit_result is False:
                readonly_hit = True

            load_result = _set_param_text(controller, "CKT_Load Name_CEDT", base_mark)
            if load_result is True:
                updates_applied = True
            elif load_result is False:
                readonly_hit = True

            if readonly_hit:
                stats["controllers_readonly"] += 1
                continue

            if not updates_applied:
                stats["controllers_already"] += 1
                continue

            stats["controllers_updated"] += 1
            updated_ids.append(controller.Id.IntegerValue)

        if active_view and unmatched_ids:
            ogs = _build_orange_overrides(doc)
            _apply_overrides(active_view, unmatched_ids, ogs)

        tx.Commit()
    except Exception as ex:
        tx.RollBack()
        forms.alert("Failed to update Identity Mark values:\n\n{}".format(ex), title=TITLE)
        return

    output.print_md("### Place Identity Mark")
    output.print_md("Cases found: `{}`".format(len(cases)))
    output.print_md("Controllers checked: `{}`".format(stats["controllers_total"]))
    output.print_md("Controllers skipped (type mismatch): `{}`".format(stats["controllers_skipped_type"]))
    output.print_md("Controllers updated: `{}`".format(stats["controllers_updated"]))
    output.print_md("Controllers already correct: `{}`".format(stats["controllers_already"]))
    output.print_md("Controllers unmatched: `{}`".format(stats["controllers_unmatched"]))
    output.print_md("Controllers read-only: `{}`".format(stats["controllers_readonly"]))

    if updated_ids:
        output.print_md("")
        output.print_md("**Updated Controller ElementIds**")
        for elem_id in updated_ids:
            output.print_md("- `{}`".format(elem_id))

    forms.alert(
        "Updated {} controller(s).\nUnmatched: {}\nAlready correct: {}".format(
            stats["controllers_updated"],
            stats["controllers_unmatched"],
            stats["controllers_already"],
        ),
        title=TITLE,
    )

    if unmatched_ids:
        show_unmatched = forms.alert(
            "Show and zoom to the {} unmatched controller(s)?".format(len(unmatched_ids)),
            title=TITLE,
            yes=True,
            no=True,
        )
        if show_unmatched:
            try:
                uidoc = revit.uidoc
                elem_ids = [DB.ElementId(int(eid)) for eid in unmatched_ids]
                if uidoc:
                    id_list = List[DB.ElementId](elem_ids)
                    uidoc.Selection.SetElementIds(id_list)
                    uidoc.ShowElements(id_list)
                    active_view = doc.ActiveView
                    ui_view = None
                    for candidate in uidoc.GetOpenUIViews():
                        if candidate.ViewId == active_view.Id:
                            ui_view = candidate
                            break
                    if ui_view:
                        min_point = None
                        max_point = None
                        for elem_id in elem_ids:
                            try:
                                elem = doc.GetElement(elem_id)
                            except Exception:
                                elem = None
                            if not elem:
                                continue
                            try:
                                bbox = elem.get_BoundingBox(active_view)
                            except Exception:
                                bbox = None
                            if not bbox:
                                try:
                                    bbox = elem.get_BoundingBox(None)
                                except Exception:
                                    bbox = None
                            min_point, max_point = _union_bbox(min_point, max_point, bbox)
                        if min_point and max_point:
                            zoom_min, zoom_max = _shrink_bbox(min_point, max_point, ZOOM_TIGHTEN)
                            ui_view.ZoomAndCenterRectangle(zoom_min, zoom_max)
            except Exception:
                pass


if __name__ == "__main__":
    main()
