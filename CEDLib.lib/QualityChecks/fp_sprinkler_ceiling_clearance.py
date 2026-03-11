# -*- coding: utf-8 -*-
"""
Fire Protection: Sprinkler clearance to ceiling/soffit. MCP_Answers: min 1" below ceiling, max 12" below.
"""

import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

MIN_DEFLECTOR_BELOW_CEILING_FT = 1.0 / 12.0   # 1 inch
MAX_DEFLECTOR_BELOW_CEILING_FT = 12.0 / 12.0  # 12 inches


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


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    min_ft = (options.get("min_deflector_below_ceiling_in") or 1.0) / 12.0
    max_ft = (options.get("max_deflector_below_ceiling_in") or 12.0) / 12.0
    opt_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    hits = []

    # Ceilings: get top face or level
    ceiling_z_by_bbox = {}
    try:
        for elem in (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_Ceilings)
            .WhereElementIsNotElementType()
            .WherePasses(opt_filter)
        ):
            try:
                bbox = elem.get_BoundingBox(None)
                if bbox is not None:
                    ceiling_z_by_bbox[elem.Id.IntegerValue] = (bbox.Min.Z + bbox.Max.Z) * 0.5
            except Exception:
                continue
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
            if bbox is None:
                continue
            # Deflector typically at bottom of sprinkler; use bbox min Z as proxy
            deflector_z = bbox.Min.Z
            # Find nearest ceiling (simplified: use max Z of any ceiling as reference for that area)
            ceiling_z = None
            for cid, cz in ceiling_z_by_bbox.items():
                if ceiling_z is None or cz > deflector_z:
                    ceiling_z = cz
            if ceiling_z is not None:
                below_ft = ceiling_z - deflector_z
                if below_ft < min_ft:
                    hits.append({
                        "a_id": elem.Id, "a_label": _label(elem),
                        "b_id": None, "b_label": "Ceiling",
                        "distance_ft": below_ft, "issue": "too close (min 1\" below)",
                    })
                elif below_ft > max_ft:
                    hits.append({
                        "a_id": elem.Id, "a_label": _label(elem),
                        "b_id": None, "b_label": "Ceiling",
                        "distance_ft": below_ft, "issue": "too far (max 12\" below)",
                    })
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
        output.print_md("# Sprinkler Clearance to Ceiling")
        output.print_md("Deflector should be 1\"–12\" below ceiling.")
        if not results:
            if show_empty:
                output.print_md("No issues found.")
            return results
        for h in results:
            dist_in = (h.get("distance_ft") or 0) * 12.0
            output.print_md("- {}: {:.2f}\" below — {}".format(
                output.linkify(h.get("a_id")), dist_in, h.get("issue", "")))
        forms.alert("Found {} sprinkler/ceiling clearance issue(s). See output panel.".format(len(results)), title="Sprinkler Ceiling")
    return results


__all__ = ["collect_hits", "run_check"]
