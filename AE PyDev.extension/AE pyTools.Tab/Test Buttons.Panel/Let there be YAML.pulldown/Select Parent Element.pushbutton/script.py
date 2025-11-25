# -*- coding: utf-8 -*-
"""Link child equipment definitions to a selected parent element."""

import math
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import XYZ

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from profile_schema import load_data as load_profile_data, save_data as save_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402
from LogicClasses.linked_equipment import compute_offsets_from_points, get_parent_id, set_parent, upsert_child  # noqa: E402

TITLE = "Select Parent Element"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")


def _pick_profile_data_path():
    cached = get_cached_yaml_path()
    if cached and os.path.exists(cached):
        return cached
    path = forms.pick_file(file_ext="yaml", title="Select profileData YAML file", init_dir=os.path.dirname(os.path.join(LIB_ROOT, "profileData.yaml")))
    if path:
        set_cached_yaml_path(path)
    return path


def _load_profile_store(data_path):
    return load_profile_data(data_path)


def _get_element_linker_payload(elem):
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if param:
            try:
                value = param.AsString() or param.AsValueString()
            except Exception:
                value = None
            if value:
                return value
    return None


def _parse_payload(text):
    payload = {}
    if not text:
        return payload
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip().lower()
        value = value.strip()
        payload[key] = value
    loc_token = payload.get("location xyz (ft)")
    if loc_token:
        try:
            parts = [float(token.strip()) for token in loc_token.split(",")]
            if len(parts) == 3:
                payload["location"] = XYZ(parts[0], parts[1], parts[2])
        except Exception:
            payload["location"] = None
    rot_token = payload.get("rotation (deg)")
    if rot_token is not None:
        try:
            payload["rotation_deg"] = float(rot_token)
        except Exception:
            payload["rotation_deg"] = 0.0
    led_token = payload.get("linked element definition id")
    if led_token:
        payload["led_id"] = led_token.strip()
    return payload


def _get_point(elem):
    if elem is None:
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


def _get_rotation(elem):
    if elem is None:
        return 0.0
    try:
        location = elem.Location
    except Exception:
        location = None
    if hasattr(location, "Rotation"):
        try:
            return float(location.Rotation * 180.0 / 3.141592653589793)
        except Exception:
            pass
    try:
        facing = getattr(elem, "FacingOrientation", None)
        if facing:
            angle = XYZ.BasisX.AngleTo(facing)
            cross = XYZ.BasisX.CrossProduct(facing)
            if cross.Z < 0:
                angle = -angle
            return float(angle * 180.0 / 3.141592653589793)
    except Exception:
        pass
    return 0.0


def _feet_to_inches(value):
    try:
        return float(value) * 12.0
    except Exception:
        return 0.0


def _inches_to_feet(value):
    try:
        return float(value) / 12.0
    except Exception:
        return 0.0


def _rotate_xy(vec, angle_deg):
    if vec is None:
        return XYZ(0, 0, 0)
    try:
        ang = math.radians(float(angle_deg))
    except Exception:
        ang = 0.0
    cos_a = math.cos(ang)
    sin_a = math.sin(ang)
    x = vec.X * cos_a - vec.Y * sin_a
    y = vec.X * sin_a + vec.Y * cos_a
    return XYZ(x, y, vec.Z)


def _average_points(points):
    if not points:
        return None
    count = float(len(points))
    sum_x = sum(p.X for p in points)
    sum_y = sum(p.Y for p in points)
    sum_z = sum(p.Z for p in points)
    return XYZ(sum_x / count, sum_y / count, sum_z / count)


def _find_equipment_by_led(data, led_id):
    target = (led_id or "").strip().lower()
    if not target:
        return None
    for eq_def in data.get("equipment_definitions") or []:
        for linked_set in eq_def.get("linked_sets") or []:
            for led_entry in linked_set.get("linked_element_definitions") or []:
                if (led_entry.get("id") or "").strip().lower() == target:
                    return eq_def, linked_set, led_entry
    return None


def _resolve_equipment_info(elem, data):
    payload_text = _get_element_linker_payload(elem)
    if not payload_text:
        return None
    payload = _parse_payload(payload_text)
    entry = _find_equipment_by_led(data, payload.get("led_id"))
    if not entry:
        return None
    eq_def, linked_set, led_entry = entry
    payload_point = payload.get("location")
    payload_rotation = payload.get("rotation_deg") or 0.0
    live_point = _get_point(elem)
    live_rotation = _get_rotation(elem)
    element_point = live_point or payload_point
    rotation = live_rotation if live_point else payload_rotation
    # approximate equipment base by subtracting stored offsets
    offsets = (led_entry.get("offsets") or [])
    if isinstance(offsets, list) and offsets:
        offsets = offsets[0]
    if not isinstance(offsets, dict):
        offsets = {}
    local_vec = XYZ(
        _inches_to_feet(offsets.get("x_inches") or 0.0),
        _inches_to_feet(offsets.get("y_inches") or 0.0),
        _inches_to_feet(offsets.get("z_inches") or 0.0),
    )
    base_point = element_point - _rotate_xy(local_vec, rotation)
    eq_id = (eq_def.get("id") or "").strip()
    eq_name = (eq_def.get("name") or eq_id or "").strip()
    led_id = (led_entry.get("id") or "").strip() if isinstance(led_entry, dict) else None
    return {
        "eq_def": eq_def,
        "linked_set": linked_set,
        "led_entry": led_entry,
        "eq_id": eq_id,
        "eq_name": eq_name,
        "base_point": base_point,
        "element_point": element_point,
        "rotation_deg": rotation,
        "payload_point": payload_point,
        "payload_rotation": payload_rotation,
        "led_id": led_id,
    }


def _summarize(children_updates, parent_name):
    if not children_updates:
        forms.alert("No relationships were updated.", title=TITLE)
        return
    lines = ["Linked {} child(ren) to '{}':".format(len(children_updates), parent_name)]
    for entry in children_updates:
        lines.append(" - {}".format(entry))
    forms.alert("\n".join(lines), title=TITLE)


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    data_path = _pick_profile_data_path()
    if not data_path:
        return
    data = _load_profile_store(data_path)

    selection = list(revit.get_selection().elements)
    if not selection:
        try:
            selection = list(revit.pick_elements(message="Select child element(s)"))
        except Exception:
            selection = []
    if not selection:
        forms.alert("No child elements were selected.", title=TITLE)
        return

    child_map = {}
    for elem in selection:
        info = _resolve_equipment_info(elem, data)
        if not info:
            continue
        eq_id = info.get("eq_id")
        if not eq_id:
            continue
        entry = child_map.setdefault(eq_id, {
            "eq_def": info["eq_def"],
            "eq_name": info["eq_name"],
            "rotation_deg": info.get("rotation_deg"),
            "fallback_point": info.get("element_point"),
            "points": [],
        })
        element_point = _get_point(elem) or info.get("element_point")
        if element_point:
            entry["points"].append(element_point)
        if entry.get("rotation_deg") is None:
            entry["rotation_deg"] = info.get("rotation_deg")
        if not entry.get("fallback_point"):
            entry["fallback_point"] = info.get("element_point")
    if not child_map:
        forms.alert("Selected elements do not belong to tracked equipment definitions.", title=TITLE)
        return
    if len(child_map) != 1:
        forms.alert("Select elements from exactly one equipment definition when establishing a parent link.", title=TITLE)
        return

    child_id, info = list(child_map.items())[0]
    eq_def_child = info.get("eq_def")
    expected_count = 0
    if eq_def_child:
        for linked_set in eq_def_child.get("linked_sets") or []:
            expected_count += len(linked_set.get("linked_element_definitions") or [])
    actual_count = len(info.get("points") or [])
    if actual_count == 0 and info.get("fallback_point"):
        actual_count = 1
    if expected_count and actual_count != expected_count:
        forms.alert(
            "You selected {} element(s) for '{}', but the equipment definition expects {}. Select all elements before choosing a parent.".format(
                actual_count,
                info.get("eq_name") or child_id,
                expected_count,
            ),
            title=TITLE,
        )
        return
    child_map = {child_id: info}

    child_entry = list(child_map.items())[0][1]
    child_eq = child_entry.get("eq_def")
    existing_parent = get_parent_id(child_eq) if child_eq else None
    if existing_parent:
        forms.alert("Child '{}' already has a parent ({}). Use 'Change Parent' to reassign.".format(
            child_entry.get("eq_name") or child_id,
            existing_parent,
        ), title=TITLE)
        return

    forms.alert("Select the parent element for the chosen child equipment.", title=TITLE)
    try:
        parent_elem = revit.pick_element(message="Select parent element")
    except Exception:
        parent_elem = None
    if not parent_elem:
        return
    parent_info = _resolve_equipment_info(parent_elem, data)
    if not parent_info:
        forms.alert("Parent element is missing Element_Linker metadata.", title=TITLE)
        return
    parent_eq = parent_info["eq_def"]
    parent_eq_id = parent_info["eq_id"]
    parent_led_id = parent_info.get("led_id")
    parent_name = parent_info["eq_name"] or parent_eq_id or "(unknown)"
    parent_anchor_point = parent_info.get("element_point") or parent_info.get("base_point")
    parent_base_point = parent_info.get("base_point")
    parent_rotation = parent_info["rotation_deg"]
    if parent_anchor_point is None or parent_base_point is None:
        forms.alert("Could not determine the parent element's base point.", title=TITLE)
        return

    updates = []
    modified = False
    for child_id, info in sorted(child_map.items()):
        if child_id == parent_eq_id:
            continue
        child_eq = info.get("eq_def")
        points = info.get("points") or []
        if not points and info.get("fallback_point"):
            points = [info.get("fallback_point")]
        child_point = _average_points(points)
        if child_point is None:
            continue
        child_rotation = info.get("rotation_deg") or 0.0
        offsets = compute_offsets_from_points(parent_anchor_point, parent_rotation, child_point, child_rotation)
        anchor_offsets = compute_offsets_from_points(parent_base_point, parent_rotation, parent_anchor_point, parent_rotation)
        upsert_child(parent_eq, child_id, offsets, anchor_offsets, parent_led_id)
        set_parent(child_eq, parent_eq_id, offsets, parent_led_id)
        updates.append(info["eq_name"] or child_id)
        modified = True

    if not modified:
        forms.alert("No child relationships were updated.", title=TITLE)
        return

    try:
        save_profile_data(data_path, data)
    except Exception as exc:
        forms.alert("Failed to save profileData.yaml:\n\n{}".format(exc), title=TITLE)
        return

    _summarize(updates, parent_name)


if __name__ == "__main__":
    main()
