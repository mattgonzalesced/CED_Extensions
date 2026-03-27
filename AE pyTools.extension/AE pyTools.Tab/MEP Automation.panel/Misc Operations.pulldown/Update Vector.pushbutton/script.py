# -*- coding: utf-8 -*-
"""
Update Vector
-------------
Updates selected elements by recalculating their offset and rotation
relative to the parent element stored in Element_Linker metadata,
and writing those values back to active YAML Extensible Storage.
"""

import math
import os
import re
import sys

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    RevitLinkInstance,
    Transaction,
    XYZ,
)

output = script.get_output()
output.close_others()

TITLE = "Update Vector"
LOG = script.get_logger()
LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
Z_MOVE_TOL_FT = 1.0 / 256.0

INLINE_PAYLOAD_PATTERN = re.compile(
    r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|"
    r"Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
    r"Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation)\s*:\s*",
    re.IGNORECASE,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402


def _to_int(value, default=None):
    if value in (None, ""):
        return default
    try:
        return int(value)
    except Exception:
        try:
            return int(float(value))
        except Exception:
            return default


def _to_float(value, default=None):
    if value in (None, ""):
        return default
    try:
        return float(value)
    except Exception:
        return default


def _normalize_angle(deg):
    value = float(deg or 0.0)
    while value <= -180.0:
        value += 360.0
    while value > 180.0:
        value -= 360.0
    return value


def _rotate_xy(vec, deg):
    if vec is None:
        return None
    angle = math.radians(float(deg or 0.0))
    cos_a = math.cos(angle)
    sin_a = math.sin(angle)
    return XYZ(
        vec.X * cos_a - vec.Y * sin_a,
        vec.X * sin_a + vec.Y * cos_a,
        vec.Z,
    )


def _feet_to_inches(ft):
    try:
        return float(ft) * 12.0
    except Exception:
        return 0.0


def _level_relative_z_inches(elem, world_point):
    if elem is None:
        return 0.0
    doc = getattr(elem, "Document", None)
    level_elem = None

    level_id = getattr(elem, "LevelId", None)
    if level_id and doc:
        try:
            level_elem = doc.GetElement(level_id)
        except Exception:
            level_elem = None

    if not level_elem:
        level_param_names = (
            "INSTANCE_REFERENCE_LEVEL_PARAM",
            "FAMILY_LEVEL_PARAM",
            "INSTANCE_LEVEL_PARAM",
            "SCHEDULE_LEVEL_PARAM",
        )
        for param_name in level_param_names:
            bip = getattr(BuiltInParameter, param_name, None)
            if not bip:
                continue
            try:
                param = elem.get_Parameter(bip)
            except Exception:
                param = None
            if not param:
                continue
            try:
                elem_id = param.AsElementId()
            except Exception:
                elem_id = None
            if elem_id and doc:
                try:
                    level_elem = doc.GetElement(elem_id)
                except Exception:
                    level_elem = None
                if level_elem:
                    break

    level_elev = 0.0
    if level_elem:
        try:
            level_elev = float(getattr(level_elem, "Elevation", 0.0) or 0.0)
        except Exception:
            level_elev = 0.0

    world_z = 0.0
    if world_point is not None:
        try:
            world_z = float(getattr(world_point, "Z", 0.0) or 0.0)
        except Exception:
            world_z = 0.0

    return _feet_to_inches(world_z - level_elev)


def _element_id_value(elem_id, default=None):
    if elem_id is None:
        return default
    for attr in ("Value", "IntegerValue"):
        try:
            raw = getattr(elem_id, attr)
        except Exception:
            raw = None
        if raw is None:
            continue
        try:
            return int(raw)
        except Exception:
            continue
    return default


def _get_point(elem):
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None

    point = getattr(loc, "Point", None)
    if point is not None:
        return point

    curve = getattr(loc, "Curve", None)
    if curve is not None:
        try:
            return curve.Evaluate(0.5, True)
        except Exception:
            return None

    return None


def _vector_angle_deg(vec):
    if vec is None:
        return None
    try:
        x_val = float(vec.X)
        y_val = float(vec.Y)
    except Exception:
        return None
    if abs(x_val) < 1e-9 and abs(y_val) < 1e-9:
        return None
    try:
        return math.degrees(math.atan2(y_val, x_val))
    except Exception:
        return None


def _get_orientation_vector(elem):
    for attr in ("HandOrientation", "FacingOrientation"):
        vec = getattr(elem, attr, None)
        if vec is None:
            continue
        try:
            x_val = float(vec.X)
            y_val = float(vec.Y)
        except Exception:
            continue
        if abs(x_val) < 1e-9 and abs(y_val) < 1e-9:
            continue
        return XYZ(vec.X, vec.Y, vec.Z)

    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            angle = float(loc.Rotation)
            return XYZ(math.cos(angle), math.sin(angle), 0.0)
        except Exception:
            pass

    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        for basis_name in ("BasisX", "BasisY"):
            basis = getattr(transform, basis_name, None)
            if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
                return XYZ(basis.X, basis.Y, basis.Z)

    return None


def _get_rotation_deg(elem, link_transform=None):
    vec = _get_orientation_vector(elem)
    if vec is not None and link_transform is not None:
        try:
            vec = link_transform.OfVector(vec)
        except Exception:
            pass
    angle = _vector_angle_deg(vec)
    if angle is not None:
        return angle
    return 0.0


def _get_linker_value(elem):
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue

        try:
            raw = param.AsString()
        except Exception:
            raw = None
        if not raw:
            try:
                raw = param.AsValueString()
            except Exception:
                raw = None

        if raw and str(raw).strip():
            return str(raw).strip()

    return ""


def _set_linker_value(elem, text):
    wrote = False
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(text)
            wrote = True
        except Exception:
            continue
    return wrote


def _parse_xyz(text):
    if not text:
        return None
    parts = [p.strip() for p in str(text).split(",")]
    if len(parts) != 3:
        return None
    try:
        return XYZ(float(parts[0]), float(parts[1]), float(parts[2]))
    except Exception:
        return None


def _format_xyz(point):
    if point is None:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(point.X, point.Y, point.Z)


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _name_variants(elem):
    names = set()
    if elem is None:
        return names
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            names.add(raw_name)
    except Exception:
        pass

    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if family_name and type_name:
            names.add(u"{} : {}".format(family_name, type_name))
        if family_name:
            names.add(family_name)
        if type_name:
            names.add(type_name)
    elif isinstance(elem, Group):
        group_type = getattr(elem, "GroupType", None)
        group_name = getattr(group_type, "Name", None) if group_type else None
        if group_name:
            names.add(group_name)

    return {_normalize_name(name) for name in names if _normalize_name(name)}


def _doc_key(doc):
    if doc is None:
        return None
    try:
        return doc.PathName or doc.Title
    except Exception:
        return None


def _get_link_transform(link_inst):
    if link_inst is None:
        return None
    try:
        return link_inst.GetTotalTransform()
    except Exception:
        try:
            return link_inst.GetTransform()
        except Exception:
            return None


def _combine_transform(parent_transform, child_transform):
    if parent_transform is None:
        return child_transform
    if child_transform is None:
        return parent_transform
    try:
        return parent_transform.Multiply(child_transform)
    except Exception:
        return None


def _walk_link_documents(doc, parent_transform, doc_chain):
    if doc is None:
        return
    key = _doc_key(doc)
    if key and key in doc_chain:
        return
    next_chain = set(doc_chain or set())
    if key:
        next_chain.add(key)

    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        transform = _combine_transform(parent_transform, _get_link_transform(link_inst))
        yield link_doc, transform
        for nested in _walk_link_documents(link_doc, transform, next_chain):
            yield nested


def _iter_link_documents(doc):
    for link_doc, transform in _walk_link_documents(doc, None, set()):
        yield link_doc, transform


def _transform_point(transform, point):
    if transform is None or point is None:
        return point
    try:
        return transform.OfPoint(point)
    except Exception:
        return point


def _collect_parent_candidates_by_host_name(doc, host_name):
    target = _normalize_name(host_name)
    if not target:
        return []

    candidates = []

    def _collect_from_doc(scan_doc, link_transform=None):
        if scan_doc is None:
            return
        for cls in (FamilyInstance, Group):
            try:
                collector = FilteredElementCollector(scan_doc).OfClass(cls).WhereElementIsNotElementType()
            except Exception:
                continue
            for elem in collector:
                variants = _name_variants(elem)
                if not variants or target not in variants:
                    continue
                point = _get_point(elem)
                if point is None:
                    continue
                host_point = _transform_point(link_transform, point)
                candidates.append({
                    "element": elem,
                    "point": host_point,
                    "rotation_deg": _get_rotation_deg(elem, link_transform=link_transform),
                    "parent_id": _element_id_value(getattr(elem, "Id", None), None),
                    "is_linked": bool(link_transform),
                })

    _collect_from_doc(doc, link_transform=None)
    for link_doc, link_transform in _iter_link_documents(doc):
        _collect_from_doc(link_doc, link_transform=link_transform)
    return candidates


def _choose_nearest_candidate(candidates, reference_point, preferred_parent_id=None):
    if not candidates:
        return None
    best = None
    best_key = None
    for cand in candidates:
        point = cand.get("point")
        id_match = 1 if (preferred_parent_id is not None and cand.get("parent_id") == preferred_parent_id) else 0
        if reference_point is not None and point is not None:
            try:
                dist = point.DistanceTo(reference_point)
            except Exception:
                dist = 1e99
        else:
            dist = 1e99
        key = (id_match, -dist)
        if best_key is None or key > best_key:
            best = cand
            best_key = key
    return best


def _resolve_parent_context(doc, parsed_payload, child_point):
    parent_id = parsed_payload.get("parent_id")
    host_name = parsed_payload.get("host_name")
    stored_parent_point = parsed_payload.get("parent_location")
    stored_parent_rot = parsed_payload.get("parent_rotation_deg")

    if parent_id is not None:
        try:
            parent_elem = doc.GetElement(ElementId(int(parent_id)))
        except Exception:
            parent_elem = None
        if parent_elem is not None:
            parent_point = _get_point(parent_elem)
            if parent_point is not None:
                return {
                    "parent_id": parent_id,
                    "parent_point": parent_point,
                    "parent_rotation_deg": _get_rotation_deg(parent_elem),
                    "resolution": "parent_id",
                }

    if host_name:
        reference_point = stored_parent_point or child_point
        candidates = _collect_parent_candidates_by_host_name(doc, host_name)
        chosen = _choose_nearest_candidate(candidates, reference_point, preferred_parent_id=parent_id)
        if chosen is not None:
            resolution = "host_name_linked" if chosen.get("is_linked") else "host_name"
            return {
                "parent_id": chosen.get("parent_id"),
                "parent_point": chosen.get("point"),
                "parent_rotation_deg": chosen.get("rotation_deg") or 0.0,
                "resolution": resolution,
            }

    if stored_parent_point is not None:
        return {
            "parent_id": parent_id,
            "parent_point": stored_parent_point,
            "parent_rotation_deg": float(stored_parent_rot or 0.0),
            "resolution": "payload_parent",
        }

    return None


def _parse_linker_payload(payload_text):
    entries = {}
    if not payload_text:
        return {
            "led_id": "",
            "set_id": "",
            "parent_id": None,
            "host_name": "",
            "parent_location": None,
            "parent_rotation_deg": 0.0,
            "entries": entries,
        }

    if "\n" in payload_text:
        for raw_line in payload_text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, val = line.partition(":")
            entries[key.strip()] = val.strip()
    else:
        matches = list(INLINE_PAYLOAD_PATTERN.finditer(payload_text))
        for idx, match in enumerate(matches):
            key = match.group(1).strip()
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(payload_text)
            val = payload_text[start:end].strip().strip(",")
            entries[key] = val.strip()

    parent_id = _to_int(entries.get("Parent ElementId"))
    if parent_id is None:
        parent_id = _to_int(entries.get("Parent Element ID"))

    return {
        "led_id": (entries.get("Linked Element Definition ID") or "").strip(),
        "set_id": (entries.get("Set Definition ID") or "").strip(),
        "parent_id": parent_id,
        "host_name": (entries.get("Host Name") or "").strip(),
        "location": _parse_xyz(entries.get("Location XYZ (ft)")),
        "parent_location": _parse_xyz(entries.get("Parent_location")),
        "parent_rotation_deg": _to_float(entries.get("Parent Rotation (deg)"), 0.0),
        "entries": entries,
    }


def _build_linker_payload(parsed, child_elem, child_pt, child_rot, parent_pt, parent_rot, parent_id):
    entries = dict((parsed or {}).get("entries") or {})

    entries["Linked Element Definition ID"] = (parsed.get("led_id") or entries.get("Linked Element Definition ID") or "").strip()
    entries["Set Definition ID"] = (parsed.get("set_id") or entries.get("Set Definition ID") or "").strip()
    entries["Parent_location"] = _format_xyz(parent_pt)
    entries["Location XYZ (ft)"] = _format_xyz(child_pt)
    entries["Rotation (deg)"] = "{:.6f}".format(float(child_rot or 0.0))
    entries["Parent Rotation (deg)"] = "{:.6f}".format(float(parent_rot or 0.0))
    entries["Parent ElementId"] = str(parent_id if parent_id is not None else "")
    entries["LevelId"] = str(_element_id_value(getattr(child_elem, "LevelId", None), "") or "")
    entries["ElementId"] = str(_element_id_value(getattr(child_elem, "Id", None), "") or "")
    entries["FacingOrientation"] = _format_xyz(getattr(child_elem, "FacingOrientation", None))

    ordered_keys = [
        "Linked Element Definition ID",
        "Set Definition ID",
        "Host Name",
        "Parent_location",
        "Location XYZ (ft)",
        "Rotation (deg)",
        "Parent Rotation (deg)",
        "Parent ElementId",
        "LevelId",
        "ElementId",
        "FacingOrientation",
    ]

    lines = []
    used = set()

    for key in ordered_keys:
        if key in entries:
            lines.append("{}: {}".format(key, entries.get(key, "")))
            used.add(key)

    for key, val in entries.items():
        if key in used:
            continue
        lines.append("{}: {}".format(key, val if val is not None else ""))

    return "\n".join(lines).strip()


def _find_linked_def(data, led_id, set_id):
    target_led = (led_id or "").strip().lower()
    target_set = (set_id or "").strip().lower()
    if not target_led:
        return None

    fallback = None
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            current_set = (linked_set.get("id") or "").strip().lower()
            for led in linked_set.get("linked_element_definitions") or []:
                current_led = (led.get("id") or led.get("led_id") or "").strip().lower()
                if current_led != target_led:
                    continue
                if target_set and target_set == current_set:
                    return led
                if fallback is None:
                    fallback = led
    return fallback


def _ensure_offsets(led_entry):
    offsets = led_entry.setdefault("offsets", [])
    if not isinstance(offsets, list):
        offsets = [{}]
        led_entry["offsets"] = offsets
    if not offsets:
        offsets.append({})

    first = offsets[0]
    if not isinstance(first, dict):
        first = {}
        offsets[0] = first

    first.setdefault("x_inches", 0.0)
    first.setdefault("y_inches", 0.0)
    first.setdefault("z_inches", 0.0)
    first.setdefault("rotation_deg", 0.0)
    return first


def _selected_elements(doc):
    uidoc = getattr(revit, "uidoc", None)
    if uidoc is None:
        return []
    try:
        ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        ids = []

    elems = []
    for elem_id in ids:
        try:
            elem = doc.GetElement(elem_id)
        except Exception:
            elem = None
        if elem is not None:
            elems.append(elem)
    return elems


def _apply_payload_updates(doc, updates):
    if not updates:
        return 0

    count = 0
    tx = Transaction(doc, "Update Vector Element_Linker")
    try:
        tx.Start()
        for elem_id, payload_text in updates.items():
            if elem_id is None:
                continue
            try:
                elem = doc.GetElement(ElementId(int(elem_id)))
            except Exception:
                elem = None
            if elem is None:
                continue
            if _set_linker_value(elem, payload_text):
                count += 1
        tx.Commit()
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
        raise
    return count


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return
    if getattr(doc, "IsFamilyDocument", False):
        forms.alert("Run this in a project document.", title=TITLE)
        return

    elements = _selected_elements(doc)
    if not elements:
        forms.alert("Select one or more placed elements and run again.", title=TITLE)
        return

    try:
        data_path, data = load_active_yaml_data()
    except Exception as exc:
        forms.alert("Failed to load active YAML.\n\n{}".format(exc), title=TITLE)
        return

    yaml_label = get_yaml_display_name(data_path)

    stats = {
        "selected": len(elements),
        "yaml_updated": 0,
        "payload_updated": 0,
        "skip_no_linker": 0,
        "skip_no_led": 0,
        "skip_no_parent_context": 0,
        "skip_no_point": 0,
        "skip_led_missing": 0,
        "resolved_by_parent_id": 0,
        "resolved_by_host_name": 0,
        "resolved_by_host_name_linked": 0,
        "resolved_by_payload_parent": 0,
    }

    payload_updates = {}

    for elem in elements:
        linker_text = _get_linker_value(elem)
        if not linker_text:
            stats["skip_no_linker"] += 1
            continue

        parsed = _parse_linker_payload(linker_text)
        led_id = parsed.get("led_id")
        set_id = parsed.get("set_id")

        if not led_id:
            stats["skip_no_led"] += 1
            continue

        child_pt = _get_point(elem)
        if child_pt is None:
            stats["skip_no_point"] += 1
            continue

        parent_ctx = _resolve_parent_context(doc, parsed, child_pt)
        if not parent_ctx:
            stats["skip_no_parent_context"] += 1
            continue

        parent_pt = parent_ctx.get("parent_point")
        if parent_pt is None:
            stats["skip_no_parent_context"] += 1
            continue

        resolution = parent_ctx.get("resolution")
        if resolution == "parent_id":
            stats["resolved_by_parent_id"] += 1
        elif resolution == "host_name":
            stats["resolved_by_host_name"] += 1
        elif resolution == "host_name_linked":
            stats["resolved_by_host_name_linked"] += 1
        elif resolution == "payload_parent":
            stats["resolved_by_payload_parent"] += 1

        parent_id = parent_ctx.get("parent_id")
        child_rot = _get_rotation_deg(elem)
        parent_rot = float(parent_ctx.get("parent_rotation_deg") or 0.0)

        world_delta = child_pt - parent_pt
        local_delta = _rotate_xy(world_delta, -parent_rot)
        rel_rot = _normalize_angle(child_rot - parent_rot)

        led_entry = _find_linked_def(data, led_id, set_id)
        if led_entry is None:
            stats["skip_led_missing"] += 1
            continue

        offsets = _ensure_offsets(led_entry)
        previous_z_inches = offsets.get("z_inches")
        offsets["x_inches"] = round(_feet_to_inches(local_delta.X), 6)
        offsets["y_inches"] = round(_feet_to_inches(local_delta.Y), 6)
        payload_loc = parsed.get("location")
        moved_vertically = False
        if payload_loc is not None:
            try:
                moved_vertically = abs(float(child_pt.Z) - float(payload_loc.Z)) > Z_MOVE_TOL_FT
            except Exception:
                moved_vertically = False

        if moved_vertically or previous_z_inches in (None, ""):
            offsets["z_inches"] = round(_level_relative_z_inches(elem, child_pt), 6)
        else:
            try:
                offsets["z_inches"] = round(float(previous_z_inches), 6)
            except Exception:
                offsets["z_inches"] = round(_level_relative_z_inches(elem, child_pt), 6)
        offsets["rotation_deg"] = round(rel_rot, 6)

        stats["yaml_updated"] += 1

        payload_updates[_element_id_value(elem.Id)] = _build_linker_payload(
            parsed,
            elem,
            child_pt,
            child_rot,
            parent_pt,
            parent_rot,
            parent_id,
        )

        LOG.info(
            "[Update Vector] elem=%s parent=%s mode=%s led=%s set=%s offsets=(%.3f, %.3f, %.3f) rot=%.3f",
            _element_id_value(elem.Id),
            parent_id,
            resolution,
            led_id,
            set_id or "",
            offsets["x_inches"],
            offsets["y_inches"],
            offsets["z_inches"],
            offsets["rotation_deg"],
        )

    if stats["yaml_updated"] > 0:
        try:
            save_active_yaml_data(
                None,
                data,
                "Update Vector",
                "Updated selected offsets/rotation relative to parent elements",
            )
        except Exception as exc:
            forms.alert("Failed saving to {}.\n\n{}".format(yaml_label, exc), title=TITLE)
            return

    try:
        stats["payload_updated"] = _apply_payload_updates(doc, payload_updates)
    except Exception as exc:
        forms.alert("YAML saved, but failed to write Element_Linker payloads.\n\n{}".format(exc), title=TITLE)
        return

    lines = [
        "YAML source: {}".format(yaml_label),
        "Selected: {}".format(stats["selected"]),
        "YAML offsets updated: {}".format(stats["yaml_updated"]),
        "Element_Linker payloads updated: {}".format(stats["payload_updated"]),
        "",
        "Skipped (no Element_Linker): {}".format(stats["skip_no_linker"]),
        "Skipped (missing LED id): {}".format(stats["skip_no_led"]),
        "Skipped (no parent context from id/host/payload): {}".format(stats["skip_no_parent_context"]),
        "Skipped (missing location point): {}".format(stats["skip_no_point"]),
        "Skipped (LED not found in YAML): {}".format(stats["skip_led_missing"]),
        "",
        "Resolved parent by Parent ElementId: {}".format(stats["resolved_by_parent_id"]),
        "Resolved parent by Host Name (active model): {}".format(stats["resolved_by_host_name"]),
        "Resolved parent by Host Name (linked model): {}".format(stats["resolved_by_host_name_linked"]),
        "Resolved parent by payload Parent_location: {}".format(stats["resolved_by_payload_parent"]),
    ]

    forms.alert("\n".join(lines), title=TITLE)


if __name__ == "__main__":
    main()
