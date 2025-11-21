# -*- coding: utf-8 -*-
"""
Set New Rotation Offset
-----------------------
Allows rotating a placed element around a picked pivot while previewing the
change. Once satisfied, writes the resulting rotation offset into
CEDLib.lib/profileData.yaml so future placements inherit the adjustment.
"""

import datetime
import io
import math
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementTransformUtils,
    Group,
    GroupType,
    Line,
    Transaction,
    XYZ,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from profile_schema import load_data as load_profile_data, save_data as save_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")

try:
    basestring
except NameError:
    basestring = str


# --------------------------------------------------------------------------- #
# YAML helpers (shared with Set New XYZ Offset)
# --------------------------------------------------------------------------- #


def _parse_scalar(token):
    token = (token or "").strip()
    if not token:
        return ""
    if token in ("{}",):
        return {}
    if token in ("[]",):
        return []
    if token.startswith('"') and token.endswith('"'):
        return token[1:-1]
    lowered = token.lower()
    if lowered in ("true", "false"):
        return lowered == "true"
    if lowered == "null":
        return None
    try:
        if "." in token:
            return float(token)
        return int(token)
    except Exception:
        return token


def _simple_yaml_parse(text):
    lines = text.splitlines()

    def parse_block(start_idx, base_indent):
        idx = start_idx
        result = None
        while idx < len(lines):
            raw_line = lines[idx]
            stripped = raw_line.strip()
            if not stripped or stripped.startswith("#"):
                idx += 1
                continue
            indent = len(raw_line) - len(raw_line.lstrip(" "))
            if indent < base_indent:
                break
            if stripped.startswith("-"):
                if result is None:
                    result = []
                elif not isinstance(result, list):
                    break
                remainder = stripped[1:].strip()
                if remainder:
                    result.append(_parse_scalar(remainder))
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result.append(value)
            else:
                if result is None:
                    result = {}
                elif isinstance(result, list):
                    break
                key, _, remainder = stripped.partition(":")
                key = key.strip().strip('"')
                remainder = remainder.strip()
                if remainder:
                    result[key] = _parse_scalar(remainder)
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result[key] = value
        if result is None:
            result = {}
        return result, idx

    parsed, _ = parse_block(0, 0)
    return parsed if isinstance(parsed, dict) else {}


def _pick_profile_data_path():
    cached = get_cached_yaml_path()
    if cached and os.path.exists(cached):
        return cached
    path = forms.pick_file(
        file_ext="yaml",
        title="Select profileData YAML file",
    )
    if path:
        set_cached_yaml_path(path)
    return path


def _load_profile_store(data_path):
    data = load_profile_data(data_path)
    if data.get("equipment_definitions"):
        return data
    try:
        with io.open(data_path, "r", encoding="utf-8") as handle:
            fallback = _simple_yaml_parse(handle.read())
        if fallback.get("equipment_definitions"):
            return fallback
    except Exception:
        pass
    return data


def _log_entry(payload, log_path):
    try:
        with io.open(log_path, "a", encoding="utf-8") as handle:
            handle.write("{} :: {}\n".format(datetime.datetime.utcnow().isoformat() + "Z", payload))
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Element helpers
# --------------------------------------------------------------------------- #


def _get_point(elem):
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    if hasattr(loc, "Point") and loc.Point:
        return loc.Point
    if hasattr(loc, "Curve") and loc.Curve:
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            return None
    return None


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


def _label_variants(value):
    base = (value or "").strip()
    variants = []
    if base:
        variants.append(base)
        if "#" in base:
            variants.append(base.split("#")[0].strip())
    return [v for v in variants if v]


def _build_label(elem):
    fam_name = None
    type_name = None

    if isinstance(elem, (Group, GroupType)):
        target = elem
        try:
            if hasattr(elem, "GroupType") and elem.GroupType:
                target = elem.GroupType
        except Exception:
            pass
        try:
            fam_name = getattr(target, "Name", None)
            type_name = getattr(target, "Name", None)
        except Exception:
            fam_name = None
            type_name = None
        if not fam_name:
            try:
                fam_name = getattr(elem, "Name", None)
                type_name = fam_name
            except Exception:
                pass
    else:
        try:
            sym = getattr(elem, "Symbol", None) or getattr(elem, "GroupType", None)
            if sym is None and hasattr(elem, "GetTypeId"):
                try:
                    sym = elem.Document.GetElement(elem.GetTypeId())
                except Exception:
                    sym = None
            if sym:
                fam = getattr(sym, "Family", None)
                fam_name = getattr(fam, "Name", None) if fam else None
                type_name = getattr(sym, "Name", None)
                if not type_name:
                    try:
                        tparam = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if tparam:
                            type_name = tparam.AsString()
                    except Exception:
                        pass
        except Exception:
            fam_name = None
            type_name = None

    if fam_name and type_name:
        return u"{} : {}".format(fam_name, type_name)
    if type_name:
        return type_name
    if fam_name:
        return fam_name
    return ""


def _resolve_target_element():
    selection = None
    try:
        selection = revit.get_selection()
    except Exception:
        selection = None
    picked = None
    if selection:
        try:
            elems = list(selection.elements)
        except Exception:
            elems = []
        if elems:
            picked = elems[0]
    if picked:
        return picked
    try:
        return revit.pick_element(message="Select equipment to update rotation offset")
    except Exception:
        return None


def _ensure_offset_entry(led):
    offsets = led.get("offsets")
    if isinstance(offsets, list) and offsets:
        entry = offsets[0]
    elif isinstance(offsets, dict):
        entry = offsets
        led["offsets"] = [entry]
    else:
        entry = {}
        led["offsets"] = [entry]
    entry.setdefault("x_inches", 0.0)
    entry.setdefault("y_inches", 0.0)
    entry.setdefault("z_inches", 0.0)
    entry.setdefault("rotation_deg", 0.0)
    return entry


def _offset_xyz(offset_entry):
    return XYZ(
        _inches_to_feet(offset_entry.get("x_inches")),
        _inches_to_feet(offset_entry.get("y_inches")),
        _inches_to_feet(offset_entry.get("z_inches")),
    )


def _rotate_xy(vec, degrees):
    if vec is None:
        return None
    angle_deg = degrees or 0.0
    try:
        angle_rad = math.radians(angle_deg)
    except Exception:
        angle_rad = 0.0
    if abs(angle_rad) < 1e-9:
        return XYZ(vec.X, vec.Y, vec.Z)
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)
    x = vec.X * cos_a - vec.Y * sin_a
    y = vec.X * sin_a + vec.Y * cos_a
    return XYZ(x, y, vec.Z)


def _rotate_xy(vec, degrees):
    if vec is None:
        return None
    angle_deg = degrees or 0.0
    try:
        angle_rad = math.radians(angle_deg)
    except Exception:
        angle_rad = 0.0
    if abs(angle_rad) < 1e-9:
        return XYZ(vec.X, vec.Y, vec.Z)
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)
    x = vec.X * cos_a - vec.Y * sin_a
    y = vec.X * sin_a + vec.Y * cos_a
    return XYZ(x, y, vec.Z)


def _find_led_matches(data, label):
    label_variants = _label_variants(label)
    label_variants_lower = [v.lower() for v in label_variants]
    matches = []
    for eq in data.get("equipment_definitions") or []:
        linked_sets = eq.get("linked_sets") or []
        for set_entry in linked_sets:
            leds = set_entry.get("linked_element_definitions") or []
            for led in leds:
                led_label = (led.get("label") or "").strip()
                if not led_label:
                    continue
                led_variants_lower = [v.lower() for v in _label_variants(led_label)]
                if any(v in led_variants_lower for v in label_variants_lower):
                    matches.append((eq, set_entry, led))
    return matches


def _choose_match(matches):
    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]
    options = []
    mapping = {}
    for idx, (eq, set_entry, led) in enumerate(matches):
        desc = u"{} ({}) :: {} [{}]".format(
            eq.get("name") or eq.get("id") or "<Unnamed>",
            set_entry.get("name") or set_entry.get("id") or "Set",
            led.get("label") or "<Label?>",
            led.get("id") or "LED",
        )
        options.append(desc)
        mapping[desc] = idx
    chosen = forms.SelectFromList.show(options, title="Select equipment definition to update", multiselect=False)
    if not chosen:
        return None
    choice = chosen if isinstance(chosen, basestring) else chosen[0]
    idx = mapping.get(choice)
    if idx is None:
        return None
    return matches[idx]


def _auto_pivot(elem, fallback_point):
    doc = getattr(elem, "Document", None) or revit.doc
    bbox = None
    try:
        view = getattr(doc, "ActiveView", None)
    except Exception:
        view = None
    try:
        if elem is not None:
            bbox = elem.get_BoundingBox(view)
    except Exception:
        bbox = None
    if bbox is None:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
    if bbox:
        pivot = XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z)
        return pivot
    return fallback_point


def _rotate_element(elem, pivot_point, delta_degrees, description):
    doc = getattr(elem, "Document", None) or revit.doc
    if doc is None:
        return False
    try:
        if abs(delta_degrees) < 1e-6:
            return True
    except Exception:
        return True
    axis = Line.CreateBound(pivot_point, pivot_point + XYZ(0, 0, 1))
    radians = math.radians(delta_degrees)
    t = Transaction(doc, description)
    try:
        t.Start()
        ElementTransformUtils.RotateElement(doc, elem.Id, axis, radians)
        t.Commit()
        return True
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return False


def _interactive_rotation(elem, pivot_point, step_degrees):
    """
    Allows the user to rotate CCW repeatedly. Returns (total_rotation, success_flag).
    """
    total = 0.0
    step = abs(step_degrees) if step_degrees not in (None, 0) else 45.0
    while True:
        msg = "Current preview rotation: {:.3f}°\nClick 'Rotate CCW' to rotate by {:.3f}°.".format(total, step)
        choice = forms.CommandSwitchWindow.show(
            ["Rotate CCW", "Reset Preview", "Finish", "Cancel"],
            message=msg,
            title="Rotate Equipment Preview",
        )
        if not choice or choice == "Cancel":
            if abs(total) > 1e-6:
                _rotate_element(elem, pivot_point, -total, "Cancel rotation preview")
            return 0.0, False
        if choice == "Rotate CCW":
            if _rotate_element(elem, pivot_point, step, "Rotate CCW preview"):
                total += step
            else:
                forms.alert("Failed to rotate the element. It may be pinned or constrained.", title="Set New Rotation Offset")
        elif choice == "Reset Preview":
            if abs(total) > 1e-6:
                if _rotate_element(elem, pivot_point, -total, "Reset rotation preview"):
                    total = 0.0
                else:
                    forms.alert("Failed to reset the preview rotation.", title="Set New Rotation Offset")
        elif choice == "Finish":
            return total, True


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #


def main():
    elem = _resolve_target_element()
    if not elem:
        forms.alert("No element selected.", title="Set New Rotation Offset")
        return

    label = _build_label(elem)
    if not label:
        forms.alert("Could not determine the label for the selected element.", title="Set New Rotation Offset")
        return

    elem_point = _get_point(elem)
    if not elem_point:
        forms.alert("Unable to read the element location.", title="Set New Rotation Offset")
        return

    data_path = _pick_profile_data_path()
    if not data_path:
        return
    data = _load_profile_store(data_path)

    matches = _find_led_matches(data, label)
    if not matches:
        forms.alert(
            "No equipment definition entry was found for label '{}'. Use 'Add YAML Profiles' first.".format(label),
            title="Set New Rotation Offset",
        )
        return

    eq_entry = _choose_match(matches)
    if not eq_entry:
        return
    eq_def, set_entry, led_entry = eq_entry
    offset_entry = _ensure_offset_entry(led_entry)
    offset_vec = _offset_xyz(offset_entry)
    existing_rotation = float(offset_entry.get("rotation_deg") or 0.0)
    rotated_offset = _rotate_xy(offset_vec, existing_rotation)

    pivot_point = _auto_pivot(elem, elem_point)
    if not pivot_point:
        forms.alert("Unable to determine a pivot point for this element.", title="Set New Rotation Offset")
        return

    step_degrees = 45.0
    base_point = elem_point - rotated_offset

    total_rotation, success = _interactive_rotation(elem, pivot_point, step_degrees)
    if not success:
        return

    new_rotation = round(existing_rotation + total_rotation, 6)
    offset_entry["rotation_deg"] = new_rotation

    elem_point_after = _get_point(elem)
    if elem_point_after:
        new_offset_world = elem_point_after - base_point
        new_offset_local = _rotate_xy(new_offset_world, -new_rotation)
        offset_entry["x_inches"] = round(_feet_to_inches(new_offset_local.X), 6)
        offset_entry["y_inches"] = round(_feet_to_inches(new_offset_local.Y), 6)
        offset_entry["z_inches"] = round(_feet_to_inches(new_offset_local.Z), 6)

    try:
        save_profile_data(data_path, data)
    except Exception as ex:
        forms.alert("Failed to update profileData.yaml:\n\n{}".format(ex), title="Set New Rotation Offset")
        return

    log_path = os.path.join(os.path.dirname(data_path), "profileData.log")
    _log_entry(
        {
            "action": "set_rotation_offset",
            "label": label,
            "equipment": eq_def.get("name") or eq_def.get("id"),
            "led_id": led_entry.get("id"),
            "rotation_deg": new_rotation,
            "delta_deg": total_rotation,
            "user": os.getenv("USERNAME") or os.getenv("USER") or "unknown",
        },
        log_path,
    )

    forms.alert(
        "Updated rotation offset for '{}'\nRotation offset (degrees): {:.3f}".format(label, new_rotation),
        title="Set New Rotation Offset",
    )


if __name__ == "__main__":
    main()
