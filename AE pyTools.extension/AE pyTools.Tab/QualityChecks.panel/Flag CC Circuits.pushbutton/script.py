# -*- coding: utf-8 -*-
"""
Flag CC Circuits
----------------
Compare refrigerated-case Identity Mark values (Mechanical Equipment)
to circuit load names on panels RA/RB/RC/RD.

For circuit load names, "CASE POWER -" is removed before comparison.
Reports mismatches in both directions.
"""

__title__ = "Flag CC\nCircuits"
__doc__ = (
    "Flags mismatches between refrigerated-case Identity Mark values and "
    "RA/RB/RC/RD circuit load names."
)

from pyrevit import DB, forms, revit, script

TITLE = "Flag CC Circuits"
TARGET_PANELS = set(["RA", "RB", "RC", "RD"])
CASE_POWER_PREFIX = "CASE POWER -"
TARGET_CONTROLLER_FAMILY = "RCD-U_General Refrigeration Device_CED"
XY_TOLERANCE_FT = 0.1


def _param_text(param):
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


def _normalize(text):
    return " ".join((text or "").strip().upper().split())


def _strip_trailing_letter(value):
    text = (value or "").strip()
    if len(text) >= 2 and text[-1].isalpha():
        return text[:-1].strip()
    return text


def _strip_case_power_prefix(load_name):
    text = (load_name or "").strip()
    if not text:
        return ""
    upper = text.upper()
    prefix = CASE_POWER_PREFIX.upper()
    if upper.startswith(prefix):
        return text[len(CASE_POWER_PREFIX):].strip()
    return text


def _identity_value(elem):
    if not elem:
        return None
    param = None
    try:
        param = elem.LookupParameter("Identity Mark")
    except Exception:
        param = None
    if not param:
        try:
            param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        except Exception:
            param = None
    return _param_text(param)


def _controller_family(elem):
    if elem is None:
        return ""
    try:
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        return (fam_name or "").strip()
    except Exception:
        return ""


def _get_bbox(elem):
    if elem is None:
        return None
    try:
        return elem.get_BoundingBox(None)
    except Exception:
        return None


def _bbox_center(bbox):
    if bbox is None:
        return None
    try:
        return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        return None


def _get_point(elem):
    if elem is None:
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
    return _bbox_center(_get_bbox(elem))


def _point_in_bbox_xy(point, bbox, xy_tol=0.0):
    if point is None or bbox is None:
        return False
    return (
        (bbox.Min.X - xy_tol) <= point.X <= (bbox.Max.X + xy_tol)
        and (bbox.Min.Y - xy_tol) <= point.Y <= (bbox.Max.Y + xy_tol)
    )


def _collect_controller_points(doc):
    points = []
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_MechanicalControlDevices)
        .WhereElementIsNotElementType()
        .WherePasses(option_filter)
    )
    for elem in collector:
        if _controller_family(elem) != TARGET_CONTROLLER_FAMILY:
            continue
        point = _get_point(elem)
        if point:
            points.append(point)
    return points


def _collect_case_marks(doc, controller_points):
    marks = {}
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
        .WhereElementIsNotElementType()
        .WherePasses(option_filter)
    )
    for elem in collector:
        bbox = _get_bbox(elem)
        if bbox is None:
            continue
        has_controller = any(_point_in_bbox_xy(pt, bbox, XY_TOLERANCE_FT) for pt in controller_points)
        if not has_controller:
            continue

        raw_mark = (_identity_value(elem) or "").strip()
        if not raw_mark:
            continue
        base_mark = _strip_trailing_letter(raw_mark)
        norm_mark = _normalize(base_mark)
        if not norm_mark:
            continue
        if norm_mark not in marks:
            marks[norm_mark] = base_mark
    return marks


def _collect_panel_load_names(doc):
    loads = {}
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit)
        .WhereElementIsNotElementType()
        .WherePasses(option_filter)
    )
    for ckt in collector:
        panel_param = None
        load_param = None
        try:
            panel_param = ckt.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        except Exception:
            panel_param = None
        try:
            load_param = ckt.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
        except Exception:
            load_param = None

        panel_name = _normalize(_param_text(panel_param))
        if panel_name not in TARGET_PANELS:
            continue

        raw_load = (_param_text(load_param) or "").strip()
        if not raw_load:
            try:
                raw_load = (ckt.LoadName or "").strip()
            except Exception:
                raw_load = ""
        if not raw_load:
            continue

        cleaned_load = _strip_case_power_prefix(raw_load)
        norm_load = _normalize(cleaned_load)
        if not norm_load:
            continue
        if norm_load not in loads:
            loads[norm_load] = cleaned_load
    return loads


def _format_value_lines(values):
    if not values:
        return ["  (none)"]
    return ["  - {}".format(value) for value in values]


def run_check(doc):
    controller_points = _collect_controller_points(doc)
    if not controller_points:
        forms.alert(
            "No refrigeration case controllers were found for family:\n{}\n\n"
            "The check uses case equipment only when a controller point is inside the case XY bounds.".format(
                TARGET_CONTROLLER_FAMILY
            ),
            title=TITLE,
        )
        return [], []

    case_marks = _collect_case_marks(doc, controller_points)
    panel_loads = _collect_panel_load_names(doc)

    case_keys = set(case_marks.keys())
    load_keys = set(panel_loads.keys())

    case_only = sorted([case_marks[key] for key in (case_keys - load_keys)], key=lambda x: x.upper())
    load_only = sorted([panel_loads[key] for key in (load_keys - case_keys)], key=lambda x: x.upper())

    output = script.get_output()
    output.print_md("# Flag CC Circuits")
    output.print_md(
        "Refrigeration controllers found (family `{}`): `{}`".format(
            TARGET_CONTROLLER_FAMILY, len(controller_points)
        )
    )
    output.print_md("Case Identity Marks found: `{}`".format(len(case_marks)))
    output.print_md("RA/RB/RC/RD load names found: `{}`".format(len(panel_loads)))

    output.print_md("## Case Identity Marks Missing in Circuits")
    if case_only:
        output.print_table([[value] for value in case_only], columns=["Identity Mark"])
    else:
        output.print_md("None.")

    output.print_md("## Circuit Load Names Missing in Cases")
    if load_only:
        output.print_table([[value] for value in load_only], columns=["Circuit Load Name (prefix removed)"])
    else:
        output.print_md("None.")

    if not case_only and not load_only:
        forms.alert(
            "No mismatches found.\n\n"
            "All unique case Identity Marks match RA/RB/RC/RD circuit load names\n"
            "after removing the \"{}\" prefix.".format(CASE_POWER_PREFIX),
            title=TITLE,
        )
        return case_only, load_only

    lines = []
    lines.append("Case Identity Marks not in RA/RB/RC/RD circuits:")
    lines.extend(_format_value_lines(case_only))
    lines.append("")
    lines.append("RA/RB/RC/RD circuit load names not in case Identity Marks:")
    lines.extend(_format_value_lines(load_only))

    forms.alert("\n".join(lines), title=TITLE)
    return case_only, load_only


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    if getattr(doc, "IsFamilyDocument", False):
        forms.alert("This check requires a project document.", title=TITLE)
        return
    run_check(doc)


if __name__ == "__main__":
    main()
