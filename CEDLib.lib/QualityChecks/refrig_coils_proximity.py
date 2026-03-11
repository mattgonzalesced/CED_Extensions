# -*- coding: utf-8 -*-
"""
Refrigeration: Coils vs heat sources (24") and coils vs sprinklers (18"). 3D bbox.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402

THRESHOLD_COIL_HEAT_IN = 24.0
THRESHOLD_COIL_SPRINKLER_IN = 18.0
COIL_KEYWORDS = ("CED-R-KRACK", "KRACK", "COIL", "EVAPORATOR", "CONDENSER", "REFRIG", "RACK")
HEAT_SOURCE_KEYWORDS = ("UNIT HEATER", "UH", "HW COIL", "HEATING COIL")
HEAT_PIPE_KEYWORDS = ("HW", "HEATING", "STEAM")


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _family_name(elem):
    try:
        sym = getattr(elem, "Symbol", None)
        fam = getattr(sym, "Family", None) if sym else None
        return getattr(fam, "Name", None) or ""
    except Exception:
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
        fn = _family_name(elem)
        sym = getattr(elem, "Symbol", None)
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


def _sys_name(elem):
    try:
        if hasattr(elem, "MEPSystem") and elem.MEPSystem is not None:
            return getattr(elem.MEPSystem, "Name", None) or ""
    except Exception:
        pass
    return ""


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    th_heat_ft = (options.get("threshold_coil_heat_in") or THRESHOLD_COIL_HEAT_IN) / 12.0
    th_sprink_ft = (options.get("threshold_coil_sprinkler_in") or THRESHOLD_COIL_SPRINKLER_IN) / 12.0
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    coils = []
    heat_sources = []
    sprinklers = []

    for bic in (DB.BuiltInCategory.OST_MechanicalEquipment, DB.BuiltInCategory.OST_SpecialityEquipment):
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(opt_filter)
            ):
                name = _family_name(elem)
                try:
                    bbox = elem.get_BoundingBox(None)
                except Exception:
                    bbox = None
                if bbox is None:
                    continue
                if _name_has_any(name, COIL_KEYWORDS):
                    coils.append({"id": elem.Id, "label": _label(elem), "bbox": bbox})
                if _name_has_any(name, HEAT_SOURCE_KEYWORDS):
                    heat_sources.append({"id": elem.Id, "label": _label(elem), "bbox": bbox})
        except Exception:
            continue

    try:
        for elem in (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_PipeCurves)
            .WhereElementIsNotElementType()
            .WherePasses(opt_filter)
        ):
            if _name_has_any(_sys_name(elem), HEAT_PIPE_KEYWORDS):
                try:
                    bbox = elem.get_BoundingBox(None)
                except Exception:
                    bbox = None
                if bbox is not None:
                    heat_sources.append({"id": elem.Id, "label": _label(elem), "bbox": bbox})
    except Exception:
        pass

    try:
        for elem in (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_Sprinklers)
            .WhereElementIsNotElementType()
            .WherePasses(opt_filter)
        ):
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
            if bbox is not None:
                sprinklers.append({"id": elem.Id, "label": _label(elem), "bbox": bbox})
    except Exception:
        pass

    hits = []
    for c in coils:
        cb = c.get("bbox")
        if cb is None:
            continue
        for h in heat_sources:
            d = _bbox_dist(cb, h.get("bbox"))
            if d is not None and d <= th_heat_ft:
                hits.append({
                    "a_id": c.get("id"), "a_label": c.get("label"),
                    "b_id": h.get("id"), "b_label": h.get("label"),
                    "distance_ft": d,
                })
        for s in sprinklers:
            d = _bbox_dist(cb, s.get("bbox"))
            if d is not None and d <= th_sprink_ft:
                hits.append({
                    "a_id": c.get("id"), "a_label": c.get("label"),
                    "b_id": s.get("id"), "b_label": s.get("label"),
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
            title="Refrigeration Coils Proximity",
            subtitle="Coils vs heat sources 24\", vs sprinklers 18\"",
            hits=hits,
            columns=["Coil ID", "Coil", "Other ID", "Other", "Distance (in)"],
            show_empty=show_empty,
        )
        if hits:
            forms.alert("Found {} refrigeration coil proximity issue(s). See output panel.".format(len(hits)), title="Coils Proximity")
    return results


__all__ = ["collect_hits", "run_check"]
