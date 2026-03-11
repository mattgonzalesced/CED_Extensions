# -*- coding: utf-8 -*-
"""
Proximity check for lighting fixtures near CED-R-KRACK coils.
Runs after sync when enabled.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit, script

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage import ExtensibleStorage  # noqa: E402
from QualityChecks.quality_check_core import report_proximity_hits  # noqa: E402

SETTING_KEY = "proximity_lights_coils_check"
FAMILY_PREFIX = "CED-R-KRACK"
THRESHOLD_INCHES = 18.0
THRESHOLD_FEET = THRESHOLD_INCHES / 12.0


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def _coerce_bool(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return bool(default)
    if isinstance(value, (int, float)):
        return value != 0
    try:
        text = str(value).strip().lower()
    except Exception:
        return bool(default)
    if text in ("1", "true", "yes", "y", "on", "enabled"):
        return True
    if text in ("0", "false", "no", "n", "off", "disabled", ""):
        return False
    return bool(default)


def get_setting(default=True, doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return bool(default)
    value = ExtensibleStorage.get_user_setting(doc, SETTING_KEY, default=None)
    if value is None:
        return bool(default)
    return _coerce_bool(value, default=default)


def set_setting(value, doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return False
    return ExtensibleStorage.set_user_setting(doc, SETTING_KEY, _coerce_bool(value, default=False))


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


def _collect_coils(doc, option_filter):
    coils = []
    try:
        collector = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilyInstance)
            .WhereElementIsNotElementType()
            .WherePasses(option_filter)
        )
    except Exception:
        return coils
    prefix = FAMILY_PREFIX.upper()
    for elem in collector:
        try:
            symbol = getattr(elem, "Symbol", None)
            family = getattr(symbol, "Family", None) if symbol else None
            fam_name = getattr(family, "Name", None) if family else None
        except Exception:
            fam_name = None
        if not fam_name:
            continue
        if not fam_name.upper().startswith(prefix):
            continue
        bbox = None
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        coils.append({
            "id": elem.Id,
            "label": _family_type_label(elem),
            "bbox": bbox,
        })
    return coils


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
        bbox = None
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        lights.append({
            "id": elem.Id,
            "label": _family_type_label(elem),
            "bbox": bbox,
        })
    return lights


def collect_hits(doc):
    if doc is None:
        return []
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    coils = _collect_coils(doc, option_filter)
    lights = _collect_lights(doc, option_filter)
    if not coils or not lights:
        return []
    hits = []
    for light in lights:
        light_bbox = light.get("bbox")
        if light_bbox is None:
            continue
        for coil in coils:
            dist_ft = _bbox_distance(light_bbox, coil.get("bbox"))
            if dist_ft is None:
                continue
            if dist_ft <= THRESHOLD_FEET:
                hits.append({
                    "light_id": light.get("id"),
                    "light_label": light.get("label"),
                    "coil_id": coil.get("id"),
                    "coil_label": coil.get("label"),
                    "distance_ft": dist_ft,
                })
    return hits


def _report_results(results, show_empty=False):
    hits = [
        {
            "a_id": item.get("light_id"),
            "a_label": item.get("light_label"),
            "b_id": item.get("coil_id"),
            "b_label": item.get("coil_label"),
            "distance_ft": item.get("distance_ft"),
        }
        for item in results or []
    ]
    report_proximity_hits(
        title="Proximity Check: Lights-Coils",
        subtitle="Lighting fixtures within {:.0f} inches of CED-R-KRACK coils".format(
            THRESHOLD_INCHES
        ),
        hits=hits,
        columns=[
            "Lighting ID",
            "Lighting Family : Type",
            "Coil ID",
            "Coil Family : Type",
            "Distance (in)",
        ],
        show_empty=show_empty,
    )
    if hits:
        forms.alert(
            "Found {} lighting fixture(s) within {:.0f} inches of CED-R-KRACK coils.\n\n"
            "See the output panel for details.".format(len(hits), THRESHOLD_INCHES),
            title="Proximity Check: Lights-Coils",
        )


def run_check(doc=None, show_ui=True, show_empty=False):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    results = collect_hits(doc)
    if show_ui:
        _report_results(results, show_empty=show_empty)
    return results


def run_sync_check(doc):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return
    if not get_setting(default=True, doc=doc):
        return
    run_check(doc, show_ui=True, show_empty=False)


__all__ = [
    "get_setting",
    "set_setting",
    "collect_hits",
    "run_check",
    "run_sync_check",
]
