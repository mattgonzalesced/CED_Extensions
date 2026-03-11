# -*- coding: utf-8 -*-
"""
Proximity check for lighting fixtures near sprinklers.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402


THRESHOLD_INCHES_DEFAULT = 18.0


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


def _collect_lights(doc, option_filter):
    lights = []
    try:
        collector = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_LightingFixtures)
            .WhereElementIsNotElementType()
            .WherePasses(option_filter)
        )
    except Exception:
        return lights
    for elem in collector:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        lights.append(
            {
                "id": elem.Id,
                "label": _family_type_label(elem),
                "bbox": bbox,
            }
        )
    return lights


def _collect_sprinklers(doc, option_filter):
    sprinklers = []
    try:
        collector = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_Sprinklers)
            .WhereElementIsNotElementType()
            .WherePasses(option_filter)
        )
    except Exception:
        return sprinklers
    for elem in collector:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        sprinklers.append(
            {
                "id": elem.Id,
                "label": _family_type_label(elem),
                "bbox": bbox,
            }
        )
    return sprinklers


def collect_hits(doc, options=None):
    if doc is None:
        return []
    options = options or {}
    threshold_inches = options.get("threshold_inches", THRESHOLD_INCHES_DEFAULT)
    threshold_feet = float(threshold_inches) / 12.0

    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    lights = _collect_lights(doc, option_filter)
    sprinklers = _collect_sprinklers(doc, option_filter)
    if not lights or not sprinklers:
        return []

    hits = []
    for light in lights:
        light_bbox = light.get("bbox")
        if light_bbox is None:
            continue
        for spr in sprinklers:
            dist_ft = _bbox_distance(light_bbox, spr.get("bbox"))
            if dist_ft is None:
                continue
            if dist_ft <= threshold_feet:
                hits.append(
                    {
                        "light_id": light.get("id"),
                        "light_label": light.get("label"),
                        "sprinkler_id": spr.get("id"),
                        "sprinkler_label": spr.get("label"),
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
                "a_id": item.get("light_id"),
                "a_label": item.get("light_label"),
                "b_id": item.get("sprinkler_id"),
                "b_label": item.get("sprinkler_label"),
                "distance_ft": item.get("distance_ft"),
            }
            for item in results or []
        ]
        report_proximity_hits(
            title="Proximity Check: Lights-Sprinklers",
            subtitle="Lighting fixtures within {:.0f} inches of sprinklers".format(
                threshold_inches
            ),
            hits=hits,
            columns=[
                "Lighting ID",
                "Lighting Family : Type",
                "Sprinkler ID",
                "Sprinkler Family : Type",
                "Distance (in)",
            ],
            show_empty=show_empty,
        )
        if hits:
            forms.alert(
                "Found {} lighting fixture(s) within {:.0f} inches of sprinklers.\n\n"
                "See the output panel for details.".format(len(hits), threshold_inches),
                title="Proximity Check: Lights-Sprinklers",
            )
    return results


__all__ = [
    "collect_hits",
    "run_check",
]

