# -*- coding: utf-8 -*-
"""
Mechanical: Insulation on cold duct / chilled water / refrigeration.
MCP_Answers: Cold systems by name (CHW, CHILLED WATER, REFRIG, SUCTION, LIQUID LINE). Min 1.0" CHW, 1.5" refrig.
"""

import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402

COLD_DUCT_KEYWORDS = ("CHW", "CHILLED WATER", "COLD AIR", "REFRIG")
COLD_PIPE_KEYWORDS = ("CHW", "CHILLED WATER", "REFRIG", "SUCTION", "LIQUID LINE")
MIN_CHW_IN = 1.0
MIN_REFRIG_IN = 1.5
INSUL_PARAM_NAMES = ("CED_InsulationThickness", "Insulation Thickness", "Thickness")


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _name_contains_any(name, keywords):
    if not name:
        return False
    u = name.upper()
    return any(k.upper() in u for k in keywords)


def _get_insulation_thickness_inches(elem):
    if elem is None:
        return None
    for pname in INSUL_PARAM_NAMES:
        try:
            for param in elem.GetOrderedParameters():
                if param is None:
                    continue
                if getattr(param, "Definition", None) and param.Definition.Name == pname:
                    if param.HasValue and param.StorageType == DB.StorageType.Double:
                        val = param.AsDouble()
                        if val is not None:
                            return val * 12.0  # ft to in if in feet
                    break
        except Exception:
            continue
        try:
            param = elem.LookupParameter(pname)
            if param is not None and param.HasValue:
                val = param.AsDouble()
                if val is not None:
                    return val * 12.0
        except Exception:
            continue
    return None


def _system_name(doc, elem):
    try:
        if hasattr(elem, "MEPSystem"):
            sys_obj = elem.MEPSystem
            if sys_obj is not None and hasattr(sys_obj, "Name"):
                return sys_obj.Name or ""
    except Exception:
        pass
    return ""


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    min_chw = options.get("min_chw_in", MIN_CHW_IN)
    min_refrig = options.get("min_refrig_in", MIN_REFRIG_IN)
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    hits = []

    # Duct
    for bic in (DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_FlexDuctCurves):
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(opt_filter)
            ):
                sys_name = _system_name(doc, elem)
                if not _name_contains_any(sys_name, COLD_DUCT_KEYWORDS):
                    continue
                thick = _get_insulation_thickness_inches(elem)
                if thick is None:
                    hits.append({
                        "a_id": elem.Id, "a_label": "Duct (no insulation param)",
                        "b_id": None, "b_label": None,
                        "distance_ft": None, "min_required_in": min_chw, "actual_in": None,
                    })
                elif thick < min_chw:
                    hits.append({
                        "a_id": elem.Id, "a_label": "Duct (insul {:.2f}\")".format(thick),
                        "b_id": None, "b_label": None,
                        "distance_ft": None, "min_required_in": min_chw, "actual_in": thick,
                    })
        except Exception:
            continue

    # Pipe
    for bic in (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_FlexPipeCurves):
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(opt_filter)
            ):
                sys_name = _system_name(doc, elem)
                if not _name_contains_any(sys_name, COLD_PIPE_KEYWORDS):
                    continue
                thick = _get_insulation_thickness_inches(elem)
                is_refrig = any(k in sys_name.upper() for k in ("REFRIG", "SUCTION", "LIQUID LINE"))
                req = min_refrig if is_refrig else min_chw
                if thick is None:
                    hits.append({
                        "a_id": elem.Id, "a_label": "Pipe (no insulation param)",
                        "b_id": None, "b_label": None,
                        "distance_ft": None, "min_required_in": req, "actual_in": None,
                    })
                elif thick < req:
                    hits.append({
                        "a_id": elem.Id, "a_label": "Pipe (insul {:.2f}\")".format(thick),
                        "b_id": None, "b_label": None,
                        "distance_ft": None, "min_required_in": req, "actual_in": thick,
                    })
        except Exception:
            continue

    return hits


def run_check(doc=None, show_ui=True, show_empty=False, options=None):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    options = options or {}
    results = collect_hits(doc, options=options)
    if show_ui:
        # Use a simple table: element, min required, actual
        output = __import__("pyrevit", fromlist=["script"]).script.get_output()
        output.set_width(1000)
        output.print_md("# Cold System Insulation Check")
        output.print_md("CHW/cold duct and pipe: min {:.1f}\"; refrigeration: min {:.1f}\"".format(
            options.get("min_chw_in", MIN_CHW_IN), options.get("min_refrig_in", MIN_REFRIG_IN)))
        if not results:
            if show_empty:
                output.print_md("No issues found.")
            return results
        for h in results:
            output.print_md("- {} (min {:.1f}\") {}".format(
                output.linkify(h.get("a_id")),
                h.get("min_required_in", 0),
                "missing" if h.get("actual_in") is None else "actual {:.2f}\"".format(h.get("actual_in")),
            ))
        forms.alert("Found {} cold system insulation issue(s). See output panel.".format(len(results)), title="Cold Insulation")
    return results


__all__ = ["collect_hits", "run_check"]
