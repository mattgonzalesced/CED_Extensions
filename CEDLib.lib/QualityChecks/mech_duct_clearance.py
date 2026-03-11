# -*- coding: utf-8 -*-
"""
Mechanical: Duct clearances vs structure, other duct, pipe, cable tray.
MCP_Answers: 2" duct-structure, 1.5" duct-duct, duct-pipe, duct-cable tray. 3D bbox distance.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402

THRESHOLD_DUCT_STRUCTURE_IN = 2.0
THRESHOLD_DUCT_DUCT_IN = 1.5
THRESHOLD_DUCT_PIPE_IN = 1.5
THRESHOLD_DUCT_TRAY_IN = 1.5

DUCT_BICS = (DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_FlexDuctCurves)
STRUCTURE_BICS = (
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_StructuralFoundation,
)
PIPE_BICS = (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_FlexPipeCurves)
TRAY_BICS = (DB.BuiltInCategory.OST_CableTray, DB.BuiltInCategory.OST_CableTrayFitting)


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _elem_label(elem):
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
    try:
        return getattr(elem, "Name", None) or "<element>"
    except Exception:
        return "<element>"


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


def _collect_by_bics(doc, bics, option_filter):
    out = []
    for bic in bics:
        try:
            for elem in (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(option_filter)
            ):
                try:
                    bbox = elem.get_BoundingBox(None)
                except Exception:
                    bbox = None
                if bbox is None:
                    continue
                out.append({"id": elem.Id, "label": _elem_label(elem), "bbox": bbox})
        except Exception:
            continue
    return out


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    ducts = _collect_by_bics(doc, DUCT_BICS, opt_filter)
    structure = _collect_by_bics(doc, STRUCTURE_BICS, opt_filter)
    pipes = _collect_by_bics(doc, PIPE_BICS, opt_filter)
    trays = _collect_by_bics(doc, TRAY_BICS, opt_filter)

    th_duct_struct_ft = (options.get("threshold_duct_structure_in") or THRESHOLD_DUCT_STRUCTURE_IN) / 12.0
    th_duct_duct_ft = (options.get("threshold_duct_duct_in") or THRESHOLD_DUCT_DUCT_IN) / 12.0
    th_duct_pipe_ft = (options.get("threshold_duct_pipe_in") or THRESHOLD_DUCT_PIPE_IN) / 12.0
    th_duct_tray_ft = (options.get("threshold_duct_tray_in") or THRESHOLD_DUCT_TRAY_IN) / 12.0

    hits = []

    for d in ducts:
        db = d.get("bbox")
        if db is None:
            continue
        for s in structure:
            dist = _bbox_dist(db, s.get("bbox"))
            if dist is not None and dist <= th_duct_struct_ft:
                hits.append({
                    "a_id": d.get("id"), "a_label": d.get("label"),
                    "b_id": s.get("id"), "b_label": s.get("label"),
                    "distance_ft": dist, "pair_type": "Duct vs Structure",
                })
        for o in ducts:
            if o.get("id") == d.get("id"):
                continue
            dist = _bbox_dist(db, o.get("bbox"))
            if dist is not None and dist <= th_duct_duct_ft:
                hits.append({
                    "a_id": d.get("id"), "a_label": d.get("label"),
                    "b_id": o.get("id"), "b_label": o.get("label"),
                    "distance_ft": dist, "pair_type": "Duct vs Duct",
                })
        for p in pipes:
            dist = _bbox_dist(db, p.get("bbox"))
            if dist is not None and dist <= th_duct_pipe_ft:
                hits.append({
                    "a_id": d.get("id"), "a_label": d.get("label"),
                    "b_id": p.get("id"), "b_label": p.get("label"),
                    "distance_ft": dist, "pair_type": "Duct vs Pipe",
                })
        for t in trays:
            dist = _bbox_dist(db, t.get("bbox"))
            if dist is not None and dist <= th_duct_tray_ft:
                hits.append({
                    "a_id": d.get("id"), "a_label": d.get("label"),
                    "b_id": t.get("id"), "b_label": t.get("label"),
                    "distance_ft": dist, "pair_type": "Duct vs Cable Tray",
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
            title="Duct Clearance (Mechanical)",
            subtitle="Duct vs structure (2\"), duct/pipe/tray (1.5\")",
            hits=hits,
            columns=["Duct ID", "Duct", "Other ID", "Other", "Distance (in)"],
            show_empty=show_empty,
        )
        if hits:
            forms.alert(
                "Found {} duct clearance issue(s). See output panel.".format(len(hits)),
                title="Duct Clearance",
            )
    return results


__all__ = ["collect_hits", "run_check"]
