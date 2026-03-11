# -*- coding: utf-8 -*-
"""
Equipment access clearance check: required clear space in front of equipment vs walls/duct/pipe/etc.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402


THRESHOLD_INCHES_DEFAULT = 36.0


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _family_type_label(elem):
    if elem is None:
        return "<missing>"
    label = None
    try:
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            label = "{} : {}".format(fam_name, type_name)
    except Exception:
        label = None
    if not label:
        try:
            label = getattr(elem, "Name", None)
        except Exception:
            label = None
    return label or "<element>"


def _bbox_distance(bbox_a, bbox_b):
    if bbox_a is None or bbox_b is None:
        return None

    def axis_distance(a_min, a_max, b_min, b_max):
        if a_max < b_min:
            return b_min - a_max
        if b_max < a_min:
            return a_min - b_max
        return 0.0

    dx = axis_distance(bbox_a.Min.X, bbox_a.Max.X, bbox_b.Min.X, bbox_b.Max.X)
    dy = axis_distance(bbox_a.Min.Y, bbox_a.Max.Y, bbox_b.Min.Y, bbox_b.Max.Y)
    dz = axis_distance(bbox_a.Min.Z, bbox_a.Max.Z, bbox_b.Min.Z, bbox_b.Max.Z)
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def _collect_equipment(doc, option_filter):
    bics = (
        DB.BuiltInCategory.OST_MechanicalEquipment,
        DB.BuiltInCategory.OST_ElectricalEquipment,
    )
    items = []
    for bic in bics:
        try:
            collector = (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(option_filter)
            )
        except Exception:
            continue
        for elem in collector:
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
            if bbox is None:
                continue
            items.append(
                {
                    "id": elem.Id,
                    "label": _family_type_label(elem),
                    "bbox": bbox,
                }
            )
    return items


def _collect_obstructions(doc, option_filter):
    bics = (
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_FlexDuctCurves,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_CableTray,
    )
    items = []
    for bic in bics:
        try:
            collector = (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .WherePasses(option_filter)
            )
        except Exception:
            continue
        for elem in collector:
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
            if bbox is None:
                continue
            items.append(
                {
                    "id": elem.Id,
                    "label": _family_type_label(elem),
                    "bbox": bbox,
                }
            )
    return items


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    threshold_inches = options.get("threshold_inches", THRESHOLD_INCHES_DEFAULT)
    threshold_feet = float(threshold_inches) / 12.0

    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    equipment = _collect_equipment(doc, option_filter)
    obstructions = _collect_obstructions(doc, option_filter)
    if not equipment or not obstructions:
        return []

    hits = []
    for equip in equipment:
        equip_bbox = equip.get("bbox")
        if equip_bbox is None:
            continue
        for obs in obstructions:
            dist_ft = _bbox_distance(equip_bbox, obs.get("bbox"))
            if dist_ft is None:
                continue
            if dist_ft <= threshold_feet:
                hits.append(
                    {
                        "equipment_id": equip.get("id"),
                        "equipment_label": equip.get("label"),
                        "obstruction_id": obs.get("id"),
                        "obstruction_label": obs.get("label"),
                        "distance_ft": dist_ft,
                    }
                )
    return hits


def run_check(doc=None, show_ui=True, show_empty=False, options=None):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    options = options or {}
    threshold_inches = options.get("threshold_inches", THRESHOLD_INCHES_DEFAULT)
    results = collect_hits(doc, options=options)
    if show_ui:
        hits = [
            {
                "a_id": item.get("equipment_id"),
                "a_label": item.get("equipment_label"),
                "b_id": item.get("obstruction_id"),
                "b_label": item.get("obstruction_label"),
                "distance_ft": item.get("distance_ft"),
            }
            for item in results or []
        ]
        report_proximity_hits(
            title="Equipment Access Clearance",
            subtitle="Equipment with obstructions within {:.0f} inches".format(
                threshold_inches
            ),
            hits=hits,
            columns=[
                "Equipment ID",
                "Equipment Family : Type",
                "Obstruction ID",
                "Obstruction Family : Type",
                "Distance (in)",
            ],
            show_empty=show_empty,
        )
        if hits:
            forms.alert(
                "Found {} equipment instance(s) with obstructions within {:.0f} inches.\n\n"
                "See the output panel for details.".format(len(hits), threshold_inches),
                title="Equipment Access Clearance",
            )
    return results


__all__ = [
    "collect_hits",
    "run_check",
]

