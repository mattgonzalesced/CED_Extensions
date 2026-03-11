# -*- coding: utf-8 -*-
"""
Electrical: Conduit vs hot/cold pipe clearance. MCP_Answers: conduit 6" from hot, 3" from cold. 3D bbox.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402

THRESHOLD_HOT_IN = 6.0
THRESHOLD_COLD_IN = 3.0
HOT_PIPE_KEYWORDS = ("HW", "HOT WATER", "STEAM", "HEATING", "GAS")
COLD_PIPE_KEYWORDS = ("CHW", "CHILLED", "REFRIG", "COLD WATER")


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _sys_name(elem):
    try:
        if hasattr(elem, "MEPSystem") and elem.MEPSystem is not None:
            return getattr(elem.MEPSystem, "Name", None) or ""
    except Exception:
        pass
    return ""


def _name_has_any(name, keywords):
    if not name:
        return False
    u = name.upper()
    return any(k.upper() in u for k in keywords)


def _label(elem):
    if elem is None:
        return "<missing>"
    try:
        sym = getattr(elem, "Symbol", None)
        fam = getattr(sym, "Family", None) if sym else None
        fn = getattr(fam, "Name", None) if fam else None
        tn = getattr(sym, "Name", None) if sym else None
        if fn and tn:
            return "{} : {}".format(fn, tn)
    except Exception:
        pass
    return getattr(elem, "Name", None) or "<element>"


def _bbox_dist(bbox_a, bbox_b):
    if bbox_a is None or bbox_b is None:
        return None
    def ax(a_min, a_max, b_min, b_max):
        if a_max < b_min:
            return b_min - a_max
        if b_max < a_min:
            return a_min - b_max
        return 0.0
    dx = ax(bbox_a.Min.X, bbox_a.Max.X, bbox_b.Min.X, bbox_b.Max.X)
    dy = ax(bbox_a.Min.Y, bbox_a.Max.Y, bbox_b.Min.Y, bbox_b.Max.Y)
    dz = ax(bbox_a.Min.Z, bbox_a.Max.Z, bbox_b.Min.Z, bbox_b.Max.Z)
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    th_hot_ft = (options.get("threshold_hot_in") or THRESHOLD_HOT_IN) / 12.0
    th_cold_ft = (options.get("threshold_cold_in") or THRESHOLD_COLD_IN) / 12.0
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    conduits = []
    hot_pipes = []
    cold_pipes = []

    for bic in (DB.BuiltInCategory.OST_Conduit, DB.BuiltInCategory.OST_ConduitFitting):
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(opt_filter)
            ):
                try:
                    bbox = elem.get_BoundingBox(None)
                except Exception:
                    bbox = None
                if bbox is not None:
                    conduits.append({"id": elem.Id, "label": _label(elem), "bbox": bbox})
        except Exception:
            continue

    for bic in (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_FlexPipeCurves):
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(opt_filter)
            ):
                name = _sys_name(elem)
                try:
                    bbox = elem.get_BoundingBox(None)
                except Exception:
                    bbox = None
                if bbox is None:
                    continue
                item = {"id": elem.Id, "label": _label(elem), "bbox": bbox}
                if _name_has_any(name, HOT_PIPE_KEYWORDS):
                    hot_pipes.append(item)
                if _name_has_any(name, COLD_PIPE_KEYWORDS):
                    cold_pipes.append(item)
        except Exception:
            continue

    hits = []
    for c in conduits:
        cb = c.get("bbox")
        if cb is None:
            continue
        for p in hot_pipes:
            d = _bbox_dist(cb, p.get("bbox"))
            if d is not None and d <= th_hot_ft:
                hits.append({
                    "a_id": c.get("id"), "a_label": c.get("label"),
                    "b_id": p.get("id"), "b_label": p.get("label"),
                    "distance_ft": d,
                })
        for p in cold_pipes:
            d = _bbox_dist(cb, p.get("bbox"))
            if d is not None and d <= th_cold_ft:
                hits.append({
                    "a_id": c.get("id"), "a_label": c.get("label"),
                    "b_id": p.get("id"), "b_label": p.get("label"),
                    "distance_ft": d,
                })
    return hits


def run_check(doc=None, show_ui=True, show_empty=False, options=None):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    options = options or {}
    results = collect_hits(doc, options=options)
    if show_ui:
        hits = [
            {"a_id": h.get("a_id"), "a_label": h.get("a_label"), "b_id": h.get("b_id"), "b_label": h.get("b_label"), "distance_ft": h.get("distance_ft")}
            for h in results
        ]
        report_proximity_hits(
            title="Conduit vs Pipe Clearance (Electrical)",
            subtitle="Conduit vs hot 6\", vs cold/chilled 3\"",
            hits=hits,
            columns=["Conduit ID", "Conduit", "Pipe ID", "Pipe", "Distance (in)"],
            show_empty=show_empty,
        )
        if hits:
            forms.alert("Found {} conduit/pipe clearance issue(s). See output panel.".format(len(hits)), title="Conduit vs Pipe")
    return results


__all__ = ["collect_hits", "run_check"]
