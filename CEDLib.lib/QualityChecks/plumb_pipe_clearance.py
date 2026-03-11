# -*- coding: utf-8 -*-
"""
Plumbing: Pipe vs structure/electrical/duct clearance. MCP_Answers: pipe-structure 2", pipe-duct 1.5", pipe-electrical 3".
"""

import math
import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402

THRESHOLD_PIPE_STRUCTURE_IN = 2.0
THRESHOLD_PIPE_DUCT_IN = 1.5
THRESHOLD_PIPE_ELECTRICAL_IN = 3.0

PIPE_BICS = (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_FlexPipeCurves)
STRUCTURE_BICS = (DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_StructuralColumns)
DUCT_BICS = (DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_FlexDuctCurves)
ELECTRICAL_BICS = (
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_Conduit,
    DB.BuiltInCategory.OST_CableTray,
    DB.BuiltInCategory.OST_CableTrayFitting,
)


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


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


def _collect_by_bics(doc, bics, opt_filter):
    out = []
    for bic in bics:
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
                    out.append({"id": elem.Id, "label": _label(elem), "bbox": bbox})
        except Exception:
            continue
    return out


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    pipes = _collect_by_bics(doc, PIPE_BICS, opt_filter)
    structure = _collect_by_bics(doc, STRUCTURE_BICS, opt_filter)
    ducts = _collect_by_bics(doc, DUCT_BICS, opt_filter)
    electrical = _collect_by_bics(doc, ELECTRICAL_BICS, opt_filter)

    th_struct_ft = (options.get("threshold_pipe_structure_in") or THRESHOLD_PIPE_STRUCTURE_IN) / 12.0
    th_duct_ft = (options.get("threshold_pipe_duct_in") or THRESHOLD_PIPE_DUCT_IN) / 12.0
    th_elec_ft = (options.get("threshold_pipe_electrical_in") or THRESHOLD_PIPE_ELECTRICAL_IN) / 12.0

    hits = []
    for p in pipes:
        pb = p.get("bbox")
        if pb is None:
            continue
        for s in structure:
            d = _bbox_dist(pb, s.get("bbox"))
            if d is not None and d <= th_struct_ft:
                hits.append({"a_id": p.get("id"), "a_label": p.get("label"), "b_id": s.get("id"), "b_label": s.get("label"), "distance_ft": d})
        for d in ducts:
            dist = _bbox_dist(pb, d.get("bbox"))
            if dist is not None and dist <= th_duct_ft:
                hits.append({"a_id": p.get("id"), "a_label": p.get("label"), "b_id": d.get("id"), "b_label": d.get("label"), "distance_ft": dist})
        for e in electrical:
            dist = _bbox_dist(pb, e.get("bbox"))
            if dist is not None and dist <= th_elec_ft:
                hits.append({"a_id": p.get("id"), "a_label": p.get("label"), "b_id": e.get("id"), "b_label": e.get("label"), "distance_ft": dist})
    return hits


def run_check(doc=None, show_ui=True, show_empty=False, options=None):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    options = options or {}
    results = collect_hits(doc, options=options)
    if show_ui:
        hits = [{"a_id": h.get("a_id"), "a_label": h.get("a_label"), "b_id": h.get("b_id"), "b_label": h.get("b_label"), "distance_ft": h.get("distance_ft")} for h in results]
        report_proximity_hits(
            title="Pipe Clearance (Plumbing)",
            subtitle="Pipe vs structure 2\", duct 1.5\", electrical 3\"",
            hits=hits,
            columns=["Pipe ID", "Pipe", "Other ID", "Other", "Distance (in)"],
            show_empty=show_empty,
        )
        if hits:
            forms.alert("Found {} pipe clearance issue(s). See output panel.".format(len(hits)), title="Pipe Clearance")
    return results


__all__ = ["collect_hits", "run_check"]
