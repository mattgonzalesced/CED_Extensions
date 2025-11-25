# -*- coding: utf-8 -*-
"""Reassign the parent equipment definition for a linked child."""

import math
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import XYZ

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from profile_schema import load_data as load_profile_data, save_data as save_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402
from LogicClasses.linked_equipment import (  # noqa: E402
    compute_offsets_from_points,
    ensure_relations,
    find_equipment_by_id,
    remove_child,
    set_parent,
    upsert_child,
)

TITLE = "Change Parent"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")


# --------------------------------------------------------------------------- #
# YAML helpers
# --------------------------------------------------------------------------- #
def _pick_profile_data_path():
    cached = get_cached_yaml_path()
    if cached and os.path.exists(cached):
        return cached
    default_path = os.path.join(LIB_ROOT, "profileData.yaml")
    init_dir = os.path.dirname(default_path)
    path = forms.pick_file(
        file_ext="yaml",
        title="Select profileData YAML file",
        init_dir=init_dir,
    )
    if path:
        set_cached_yaml_path(path)
    return path


def _load_profile_store(data_path):
    return load_profile_data(data_path)


def _save_profile_store(data_path, data):
    try:
        save_profile_data(data_path, data)
        return True
    except Exception:
        forms.alert("Failed to save profileData.yaml.", title=TITLE)
        return False


# --------------------------------------------------------------------------- #
# Element_Linker parsing utilities
# --------------------------------------------------------------------------- #
def _get_element_linker_payload(elem):
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        value = None
        try:
            value = param.AsString()
        except Exception:
            value = None
        if not value:
            try:
                value = param.AsValueString()
            except Exception:
                value = None
        if value and value.strip():
            return value.strip()
    return None


def _parse_xyz_string(text):
    if not text:
        return None
    parts = [token.strip() for token in text.split(",")]
    if len(parts) != 3:
        return None
    try:
        return XYZ(float(parts[0]), float(parts[1]), float(parts[2]))
    except Exception:
        return None


def _parse_float(value, default=0.0):
    if value is None:
        return default
    try:
        return float(value)
    except Exception:
        try:
            return float(str(value).strip())
        except Exception:
            return default


def _parse_element_linker_payload(payload_text):
    if not payload_text:
        return {}
    entries = {}
    for raw_line in payload_text.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, _, remainder = line.partition(":")
        entries[key.strip()] = remainder.strip()
    return {
        "led_id": entries.get("Linked Element Definition ID", "").strip(),
        "set_id": entries.get("Set Definition ID", "").strip(),
        "location": _parse_xyz_string(entries.get("Location XYZ (ft)")),
        "rotation_deg": _parse_float(entries.get("Rotation (deg)"), 0.0),
        "facing": _parse_xyz_string(entries.get("FacingOrientation")),
        "raw": entries,
    }


def _find_led_entry(data, led_id, set_id=None):
    needle = (led_id or "").strip().lower()
    target_set = (set_id or "").strip().lower()
    if not needle:
        return None
    for eq_def in data.get("equipment_definitions") or []:
        for linked_set in eq_def.get("linked_sets") or []:
            set_id_value = (linked_set.get("id") or "").strip().lower()
            if target_set and target_set != set_id_value:
                continue
            for led_entry in linked_set.get("linked_element_definitions") or []:
                led_value = (led_entry.get("id") or "").strip().lower()
                if led_value == needle:
                    return eq_def, linked_set, led_entry
    return None


# --------------------------------------------------------------------------- #
# Geometry helpers
# --------------------------------------------------------------------------- #
def _get_point(elem):
    if not elem:
        return None
    location = getattr(elem, "Location", None)
    if not location:
        return None
    try:
        point = location.Point
        if point:
            return point
    except Exception:
        pass
    try:
        curve = location.Curve
        if curve:
            return curve.GetEndPoint(0)
    except Exception:
        pass
    return None


def _get_rotation_degrees(elem):
    if not elem:
        return None
    try:
        location = elem.Location
    except Exception:
        location = None
    if hasattr(location, "Rotation"):
        try:
            return float(location.Rotation * 180.0 / math.pi)
        except Exception:
            pass
    try:
        facing = getattr(elem, "FacingOrientation", None)
        if facing:
            angle = XYZ.BasisX.AngleTo(facing)
            cross = XYZ.BasisX.CrossProduct(facing)
            if cross.Z < 0:
                angle = -angle
            return float(angle * 180.0 / math.pi)
    except Exception:
        pass
    return None


def _inches_to_feet(value):
    try:
        return float(value) / 12.0
    except Exception:
        return 0.0


def _rotate_xy(vec, angle_deg):
    if not isinstance(vec, XYZ):
        return XYZ(0, 0, 0)
    try:
        angle_rad = math.radians(float(angle_deg or 0.0))
    except Exception:
        angle_rad = 0.0
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)
    return XYZ(
        vec.X * cos_a - vec.Y * sin_a,
        vec.X * sin_a + vec.Y * cos_a,
        vec.Z,
    )


def _average_points(points):
    valid = [pt for pt in points if isinstance(pt, XYZ)]
    if not valid:
        return None
    count = float(len(valid))
    return XYZ(
        sum(pt.X for pt in valid) / count,
        sum(pt.Y for pt in valid) / count,
        sum(pt.Z for pt in valid) / count,
    )


def _resolve_equipment_info(elem, data):
    payload_text = _get_element_linker_payload(elem)
    if not payload_text:
        return None
    payload = _parse_element_linker_payload(payload_text)
    entry = _find_led_entry(data, payload.get("led_id"), payload.get("set_id"))
    if not entry:
        return None
    eq_def, linked_set, led_entry = entry
    payload_point = payload.get("location")
    payload_rotation = payload.get("rotation_deg") or 0.0
    live_point = _get_point(elem)
    live_rotation = _get_rotation_degrees(elem)
    element_point = live_point or payload_point
    rotation = live_rotation if live_rotation is not None else payload_rotation
    offsets = led_entry.get("offsets") or []
    if isinstance(offsets, list):
        offsets = offsets[0] if offsets else {}
    if not isinstance(offsets, dict):
        offsets = {}
    local_vec = XYZ(
        _inches_to_feet(offsets.get("x_inches") or 0.0),
        _inches_to_feet(offsets.get("y_inches") or 0.0),
        _inches_to_feet(offsets.get("z_inches") or 0.0),
    )
    base_point = element_point - _rotate_xy(local_vec, rotation) if element_point else None
    eq_id = (eq_def.get("id") or "").strip()
    eq_name = (eq_def.get("name") or eq_id or "").strip()
    led_id = (led_entry.get("id") or "").strip()
    return {
        "eq_def": eq_def,
        "linked_set": linked_set,
        "led_entry": led_entry,
        "eq_id": eq_id,
        "eq_name": eq_name,
        "element_point": element_point,
        "payload_point": payload_point,
        "rotation_deg": rotation,
        "payload_rotation": payload_rotation,
        "base_point": base_point,
        "led_id": led_id,
    }


# --------------------------------------------------------------------------- #
# Selection + validation logic
# --------------------------------------------------------------------------- #
def _collect_child_selection(selection, data):
    selection = list(selection or [])
    if not selection:
        forms.alert("No child elements were selected.", title=TITLE)
        return None

    eq_map = {}
    for elem in selection:
        info = _resolve_equipment_info(elem, data)
        if not info:
            continue
        eq_id = info.get("eq_id")
        if not eq_id:
            continue
        entry = eq_map.setdefault(
            eq_id,
            {
                "eq_def": info["eq_def"],
                "eq_name": info.get("eq_name") or eq_id,
                "points": [],
                "led_ids": set(),
                "rotation_deg": info.get("rotation_deg"),
                "fallback_point": info.get("element_point") or info.get("payload_point"),
            },
        )
        point = info.get("element_point") or info.get("payload_point")
        if point:
            entry["points"].append(point)
        if info.get("rotation_deg") is not None and entry.get("rotation_deg") is None:
            entry["rotation_deg"] = info.get("rotation_deg")
        led_id = info.get("led_id")
        if led_id:
            entry["led_ids"].add(led_id)

    if not eq_map:
        forms.alert("Selected elements do not belong to tracked equipment definitions.", title=TITLE)
        return None
    if len(eq_map) != 1:
        forms.alert("Select elements from exactly one equipment definition before changing its parent.", title=TITLE)
        return None

    child_id, info = list(eq_map.items())[0]
    eq_def = info.get("eq_def")
    expected = 0
    if eq_def:
        for linked_set in eq_def.get("linked_sets") or []:
            expected += len(linked_set.get("linked_element_definitions") or [])
    actual = len(info.get("led_ids") or [])
    if expected and actual != expected:
        forms.alert(
            "You selected {} element(s) for '{}', but the equipment definition expects {}. "
            "Select every element in the equipment definition before changing its parent.".format(
                actual,
                info.get("eq_name") or child_id,
                expected,
            ),
            title=TITLE,
        )
        return None

    points = info.get("points") or []
    if not points and info.get("fallback_point"):
        points = [info.get("fallback_point")]
    avg_point = _average_points(points)
    if avg_point is None:
        forms.alert("Could not determine a reference point for '{}'. Try moving it slightly and retry.".format(
            info.get("eq_name") or child_id,
        ), title=TITLE)
        return None

    info.update({
        "eq_id": child_id,
        "avg_point": avg_point,
    })
    if info.get("rotation_deg") is None:
        info["rotation_deg"] = 0.0
    return info


def _remove_child_from_parent(data, parent_id, child_id):
    if not parent_id:
        return
    parent_eq = find_equipment_by_id(data, parent_id)
    if parent_eq:
        remove_child(parent_eq, child_id)


# --------------------------------------------------------------------------- #
# Main entry point
# --------------------------------------------------------------------------- #
def main():
    data_path = _pick_profile_data_path()
    if not data_path:
        return
    data = _load_profile_store(data_path)

    selection = list(revit.get_selection().elements)
    if not selection:
        try:
            selection = list(revit.pick_elements(message="Select child equipment elements"))
        except Exception:
            selection = []
    child_info = _collect_child_selection(selection, data)
    if not child_info:
        return

    child_eq = child_info["eq_def"]
    child_id = child_info["eq_id"]
    child_name = child_info.get("eq_name") or child_id
    child_point = child_info.get("avg_point")
    child_rotation = child_info.get("rotation_deg") or 0.0

    relations = ensure_relations(child_eq)
    parent_entry = relations.get("parent") or {}
    existing_parent_id = (parent_entry.get("equipment_id") or "").strip()
    if not existing_parent_id:
        forms.alert("'{}' does not have a parent to change.".format(child_name), title=TITLE)
        return

    choice = forms.alert(
        "Change the parent of '{}'? \n\nSelect YES to pick a new parent element, "
        "NO to remove the parent link, or CANCEL to abort.".format(child_name),
        title=TITLE,
        yes=True,
        no=True,
        cancel=True,
    )
    if choice is None:
        return

    if choice is False:
        _remove_child_from_parent(data, existing_parent_id, child_id)
        set_parent(child_eq, None)
        if _save_profile_store(data_path, data):
            forms.alert("Removed the parent relationship for '{}'.".format(child_name), title=TITLE)
        return

    forms.alert("Select the new parent element for '{}'.".format(child_name), title=TITLE)
    try:
        parent_elem = revit.pick_element(message="Select new parent element")
    except Exception:
        parent_elem = None
    if not parent_elem:
        return

    parent_info = _resolve_equipment_info(parent_elem, data)
    if not parent_info:
        forms.alert("Selected parent element is missing Element_Linker metadata.", title=TITLE)
        return
    parent_eq = parent_info["eq_def"]
    parent_eq_id = parent_info.get("eq_id")
    parent_led_id = parent_info.get("led_id")
    parent_name = parent_info.get("eq_name") or parent_eq_id
    if not parent_eq_id or parent_eq_id == child_id:
        forms.alert("Select a different equipment definition to act as the parent.", title=TITLE)
        return

    parent_anchor_point = parent_info.get("element_point") or parent_info.get("payload_point")
    parent_base_point = parent_info.get("base_point")
    parent_rotation = parent_info.get("rotation_deg") or 0.0
    if parent_anchor_point is None or parent_base_point is None:
        forms.alert("Could not determine the parent element's midpoint.", title=TITLE)
        return
    if child_point is None:
        forms.alert("Unable to determine the child's reference point. Try reloading the equipment.", title=TITLE)
        return

    offsets = compute_offsets_from_points(parent_anchor_point, parent_rotation, child_point, child_rotation)
    anchor_offsets = compute_offsets_from_points(parent_base_point, parent_rotation, parent_anchor_point, parent_rotation)

    if existing_parent_id and existing_parent_id != parent_eq_id:
        _remove_child_from_parent(data, existing_parent_id, child_id)

    set_parent(child_eq, parent_eq_id, offsets, parent_led_id)
    upsert_child(parent_eq, child_id, offsets, anchor_offsets, parent_led_id)

    if _save_profile_store(data_path, data):
        forms.alert("Parent of '{}' updated to '{}'.".format(child_name, parent_name), title=TITLE)


if __name__ == "__main__":
    main()
