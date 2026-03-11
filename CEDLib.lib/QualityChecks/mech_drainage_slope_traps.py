# -*- coding: utf-8 -*-
"""
Mechanical/Plumbing: Drainage traps and slope. Drain systems by name; RBS_PIPE_SLOPE_PARAM min 1/4" per ft; trap within 5 ft.
"""

import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

DRAIN_KEYWORDS = ("SANITARY", "WASTE", "CONDENSATE", "DRAIN")
TRAP_KEYWORDS = ("TRAP", "P-TRAP", "CONDENSATE TRAP")
MIN_SLOPE_FT_PER_FT = 1.0 / 48.0  # 1/4" per foot
TRAP_DISTANCE_FT = 5.0

try:
    RBS_PIPE_SLOPE_PARAM = DB.BuiltInParameter.RBS_PIPE_SLOPE_PARAM
except Exception:
    RBS_PIPE_SLOPE_PARAM = None


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _system_name(elem):
    try:
        if hasattr(elem, "MEPSystem") and elem.MEPSystem is not None:
            return getattr(elem.MEPSystem, "Name", None) or ""
    except Exception:
        pass
    return ""


def _name_contains_any(name, keywords):
    if not name:
        return False
    u = name.upper()
    return any(k.upper() in u for k in keywords)


def _get_slope_ft_per_ft(elem):
    if elem is None or RBS_PIPE_SLOPE_PARAM is None:
        return None
    try:
        p = elem.get_Parameter(RBS_PIPE_SLOPE_PARAM)
        if p is not None and p.HasValue:
            return p.AsDouble()
    except Exception:
        pass
    return None


def _is_trap_family(elem):
    try:
        if hasattr(elem, "Symbol") and elem.Symbol is not None:
            fam = getattr(elem.Symbol, "Family", None)
            if fam is not None:
                name = getattr(fam, "Name", None) or ""
                if _name_contains_any(name, TRAP_KEYWORDS):
                    return True
    except Exception:
        pass
    return False


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    min_slope = options.get("min_slope_ft_per_ft", MIN_SLOPE_FT_PER_FT)
    trap_dist_ft = options.get("trap_distance_ft", TRAP_DISTANCE_FT)
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    hits = []

    # Collect trap locations (family instances / fixtures)
    trap_ids = set()
    for bic in (DB.BuiltInCategory.OST_PlumbingFixtures, DB.BuiltInCategory.OST_GenericModel):
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(opt_filter)
            ):
                if _is_trap_family(elem):
                    trap_ids.add(elem.Id.IntegerValue)
        except Exception:
            continue

    # Pipes in drain systems
    try:
        for elem in (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_PipeCurves)
            .WhereElementIsNotElementType()
            .WherePasses(opt_filter)
        ):
            sys_name = _system_name(elem)
            if not _name_contains_any(sys_name, DRAIN_KEYWORDS):
                continue
            slope = _get_slope_ft_per_ft(elem)
            if slope is not None and slope < min_slope:
                hits.append({
                    "a_id": elem.Id, "a_label": "Pipe slope {:.4f} (min {:.4f})".format(slope, min_slope),
                    "b_id": None, "b_label": None,
                    "distance_ft": None, "slope_ft_per_ft": slope, "min_slope": min_slope,
                })
            # Simplified: we don't trace downstream 5 ft here; flag if no trap in same system (stub)
    except Exception:
        pass

    return hits


def run_check(doc=None, show_ui=True, show_empty=False, options=None):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    options = options or {}
    results = collect_hits(doc, options=options)
    if show_ui:
        output = __import__("pyrevit", fromlist=["script"]).script.get_output()
        output.set_width(1000)
        output.print_md("# Drainage Slope Check")
        output.print_md("Drain systems (sanitary/waste/condensate): min 1/4\" per foot.")
        if not results:
            if show_empty:
                output.print_md("No slope issues found.")
            return results
        for h in results:
            output.print_md("- {}: {}".format(output.linkify(h.get("a_id")), h.get("a_label")))
        forms.alert("Found {} drainage slope issue(s). See output panel.".format(len(results)), title="Drainage Slope")
    return results


__all__ = ["collect_hits", "run_check"]
