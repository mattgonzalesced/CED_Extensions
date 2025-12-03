# -*- coding: utf-8 -*-
"""Attach selected elements to the linked equipment definition of a parent element."""

import math
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    Group,
    IndependentTag,
    RevitLinkInstance,
    Transaction,
    Transform,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ObjectType

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from profile_schema import get_type_set, load_data as load_profile_data, next_led_id, save_data as save_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402
from LogicClasses.linked_equipment import compute_offsets_from_points, find_equipment_by_name  # noqa: E402

TITLE = "Select Parent Element"
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", ELEMENT_LINKER_PARAM_NAME)


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


def _linked_element_from_reference(doc, reference):
    if doc is None or reference is None:
        return None, None
    linked_id = getattr(reference, "LinkedElementId", None)
    if not isinstance(linked_id, ElementId):
        return None, None
    if linked_id == ElementId.InvalidElementId:
        return None, None
    try:
        host_elem = doc.GetElement(reference.ElementId)
    except Exception:
        host_elem = None
    if not isinstance(host_elem, RevitLinkInstance):
        return None, None
    try:
        link_doc = host_elem.GetLinkDocument()
    except Exception:
        link_doc = None
    if link_doc is None:
        return None, None
    try:
        transform = host_elem.GetTransform()
        if not isinstance(transform, Transform):
            transform = None
    except Exception:
        transform = None
    try:
        return link_doc.GetElement(linked_id), transform
    except Exception:
        return None, transform


def _pick_linked_parent_element(message):
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return None, None
    try:
        reference = uidoc.Selection.PickObject(ObjectType.LinkedElement, message)
    except Exception:
        return None, None
    return _linked_element_from_reference(doc, reference)


def _pick_parent_element(message):
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return None, None
    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, message)
    except Exception:
        return None, None
    parent_elem, transform = _linked_element_from_reference(doc, reference)
    if parent_elem:
        return parent_elem, transform
    try:
        elem = doc.GetElement(reference.ElementId)
    except Exception:
        elem = None
    if isinstance(elem, RevitLinkInstance):
        forms.alert("Selected a linked model. After this message, pick the specific parent element inside that link.", title=TITLE)
        linked_elem, linked_transform = _pick_linked_parent_element(message + " (linked model)")
        return linked_elem, linked_transform
    return elem, None


def _candidate_equipment_names(elem):
    names = []
    if elem is None:
        return names
    try:
        if hasattr(elem, "Name") and elem.Name:
            names.append(elem.Name.strip())
    except Exception:
        pass
    try:
        sym = getattr(elem, "Symbol", None) or getattr(elem, "GroupType", None)
    except Exception:
        sym = None
    if sym:
        try:
            if getattr(sym, "Name", None):
                names.append(sym.Name.strip())
        except Exception:
            pass
        try:
            fam = getattr(sym, "Family", None)
            if fam and getattr(fam, "Name", None):
                names.append(fam.Name.strip())
        except Exception:
            pass
        try:
            fam_name = getattr(sym, "FamilyName", None)
            if fam_name:
                names.append(fam_name.strip())
        except Exception:
            pass
    fam_label, _ = _build_label_info(elem)
    if fam_label:
        names.append(fam_label)
    uniq = []
    seen = set()
    for name in names:
        if not name:
            continue
        norm = name.strip()
        if not norm or norm.lower() in seen:
            continue
        seen.add(norm.lower())
        uniq.append(norm)
    return uniq


def _find_equipment_by_names(elem, data):
    candidates = _candidate_equipment_names(elem)
    for name in candidates:
        eq_def = find_equipment_by_name(data, name)
        if eq_def:
            return eq_def, name, candidates
    return None, None, candidates


def _collect_params(elem):
    try:
        cat = getattr(elem, "Category", None)
        cat_name = getattr(cat, "Name", "") if cat else ""
    except Exception:
        cat_name = ""
    cat_l = (cat_name or "").lower()
    is_electrical = ("electrical" in cat_l) or ("lighting" in cat_l) or ("data" in cat_l)

    base_targets = {
        "dev-Group ID": ["dev-Group ID", "dev_Group ID"],
    }
    electrical_targets = {
        "CKT_Rating_CED": ["CKT_Rating_CED"],
        "CKT_Panel_CEDT": ["CKT_Panel_CED", "CKT_Panel_CEDT"],
        "CKT_Schedule Notes_CEDT": ["CKT_Schedule Notes_CED", "CKT_Schedule Notes_CEDT"],
        "CKT_Circuit Number_CEDT": ["CKT_Circuit Number_CED", "CKT_Circuit Number_CEDT"],
        "CKT_Load Name_CEDT": ["CKT_Load Name_CED", "CKT_Load Name_CEDT"],
    }

    targets = dict(base_targets)
    if is_electrical:
        targets.update(electrical_targets)

    found = {k: "" for k in targets.keys()}
    for param in getattr(elem, "Parameters", []) or []:
        try:
            name = param.Definition.Name
        except Exception:
            continue
        target_key = None
        for out_key, aliases in targets.items():
            if name in aliases:
                target_key = out_key
                break
        if not target_key:
            continue
        try:
            storage = param.StorageType.ToString()
        except Exception:
            storage = ""
        try:
            if storage == "String":
                found[target_key] = param.AsString() or ""
            elif storage == "Double":
                found[target_key] = param.AsDouble()
            elif storage == "Integer":
                found[target_key] = param.AsInteger()
            else:
                found[target_key] = param.AsValueString() or ""
        except Exception:
            continue

    if "dev-Group ID" not in found:
        found["dev-Group ID"] = ""
    if is_electrical:
        return found
    if any(value for key, value in found.items() if key != "dev-Group ID" and value):
        return found
    return {"dev-Group ID": found.get("dev-Group ID", "")}


def _get_level_element_id(elem):
    try:
        lvl = getattr(elem, "LevelId", None)
        if lvl and getattr(lvl, "IntegerValue", -1) > 0:
            return lvl.IntegerValue
    except Exception:
        pass
    level_params = (
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_LEVEL_PARAM,
    )
    for bip in level_params:
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if not param:
            continue
        try:
            eid = param.AsElementId()
            if eid and getattr(eid, "IntegerValue", -1) > 0:
                return eid.IntegerValue
        except Exception:
            continue
    return None


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)


def _build_element_linker_payload(led_id, set_id, elem, host_point):
    point = host_point or _get_point(elem)
    rotation_deg = _get_rotation(elem)
    level_id = _get_level_element_id(elem)
    try:
        elem_id = elem.Id.IntegerValue
    except Exception:
        elem_id = ""
    facing = getattr(elem, "FacingOrientation", None)
    lines = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Location XYZ (ft): {}".format(_format_xyz(point)),
        "Rotation (deg): {:.6f}".format(rotation_deg),
        "LevelId: {}".format(level_id if level_id is not None else ""),
        "ElementId: {}".format(elem_id),
        "FacingOrientation: {}".format(_format_xyz(facing)),
    ]
    return "\n".join(lines).strip()


def _set_element_linker_parameter(elem, value):
    if not elem or value is None:
        return False
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(value)
            return True
        except Exception:
            continue
    return False


def _collect_hosted_tags(elem, host_point):
    doc = getattr(elem, "Document", None)
    if doc is None or host_point is None:
        return []
    try:
        deps = list(elem.GetDependentElements(None))
    except Exception:
        deps = []
    tags = []
    for dep_id in deps:
        try:
            tag = doc.GetElement(dep_id)
        except Exception:
            tag = None
        if not tag or not isinstance(tag, IndependentTag):
            continue
        try:
            tag_pt = tag.TagHeadPosition
        except Exception:
            tag_pt = None
        tag_symbol = None
        try:
            tag_symbol = doc.GetElement(tag.GetTypeId())
        except Exception:
            tag_symbol = None
        fam_name = None
        type_name = None
        category_name = None
        if tag_symbol:
            try:
                fam_name = getattr(tag_symbol, "FamilyName", None)
                if not fam_name:
                    fam = getattr(tag_symbol, "Family", None)
                    fam_name = getattr(fam, "Name", None) if fam else None
            except Exception:
                fam_name = None
            try:
                type_name = getattr(tag_symbol, "Name", None)
                if not type_name and hasattr(tag_symbol, "get_Parameter"):
                    sparam = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if sparam:
                        type_name = sparam.AsString()
            except Exception:
                type_name = None
            try:
                cat = getattr(tag_symbol, "Category", None)
                category_name = getattr(cat, "Name", None) if cat else None
            except Exception:
                category_name = None
        if not category_name:
            try:
                cat = getattr(tag, "Category", None)
                category_name = getattr(cat, "Name", None) if cat else None
            except Exception:
                category_name = None
        if not fam_name:
            try:
                sym = getattr(tag, "Symbol", None)
                fam = getattr(sym, "Family", None) if sym else None
                fam_name = getattr(fam, "Name", None) if fam else fam_name
            except Exception:
                pass
        if not type_name:
            try:
                tag_type = getattr(tag, "TagType", None)
                type_name = getattr(tag_type, "Name", None)
            except Exception:
                pass
        if not fam_name or not type_name:
            continue
        offsets = {
            "x_inches": 0.0,
            "y_inches": 0.0,
            "z_inches": 0.0,
            "rotation_deg": 0.0,
        }
        if tag_pt:
            delta = tag_pt - host_point
            offsets["x_inches"] = _feet_to_inches(delta.X)
            offsets["y_inches"] = _feet_to_inches(delta.Y)
        tags.append({
            "family_name": fam_name,
            "type_name": type_name,
            "category_name": category_name,
            "parameters": {},
            "offsets": offsets,
        })
    return tags


def _get_category_name(elem):
    try:
        cat = elem.Category
        if cat:
            return cat.Name or ""
    except Exception:
        pass
    return ""


def _build_label_info(elem):
    fam_name = None
    type_name = None
    is_group = isinstance(elem, Group)
    if isinstance(elem, Group):
        try:
            fam_name = elem.Name
            type_name = elem.Name
        except Exception:
            pass
    else:
        try:
            sym = getattr(elem, "Symbol", None) or getattr(elem, "GroupType", None)
            if sym:
                fam = getattr(sym, "Family", None)
                fam_name = getattr(fam, "Name", None) if fam else None
                type_name = getattr(sym, "Name", None)
                if not type_name and hasattr(sym, "get_Parameter"):
                    tparam = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if tparam:
                        type_name = tparam.AsString()
        except Exception:
            pass
    if not fam_name and hasattr(elem, "Name"):
        try:
            fam_name = elem.Name
        except Exception:
            pass
    if not type_name:
        type_name = fam_name
    if fam_name and type_name:
        label = "{} : {}".format(fam_name, type_name)
    elif type_name:
        label = type_name
    elif fam_name:
        label = fam_name
    else:
        label = "Unnamed"
    return label, is_group


def _build_child_entries(elements):
    entries = []
    for elem in elements:
        if elem is None:
            continue
        point = _get_point(elem)
        if point is None:
            continue
        label, is_group = _build_label_info(elem)
        if not label:
            continue
        entry = {
            "element": elem,
            "point": point,
            "rotation_deg": _get_rotation(elem),
            "label": label,
            "is_group": is_group,
            "category": _get_category_name(elem),
            "parameters": _collect_params(elem),
            "tags": _collect_hosted_tags(elem, point),
        }
        entries.append(entry)
    return entries


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
    try:
        hand = getattr(elem, "HandOrientation", None)
        if hand:
            projected = XYZ(hand.X, hand.Y, 0.0)
            if abs(projected.X) > 1e-9 or abs(projected.Y) > 1e-9:
                angle = math.atan2(projected.Y, projected.X)
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


def _transform_point(point, transform):
    if point is None or transform is None:
        return point
    try:
        return transform.OfPoint(point)
    except Exception:
        return point


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


def _resolve_equipment_info(elem, data, include_reason=False, allow_name_lookup=False, link_transform=None):
    payload_text = _get_element_linker_payload(elem)
    if not payload_text:
        if allow_name_lookup:
            eq_def, matched_name, candidates = _find_equipment_by_names(elem, data)
            if eq_def:
                element_point = _transform_point(_get_point(elem), link_transform)
                if element_point is None:
                    reason = {"code": "missing-location"}
                    return (None, reason) if include_reason else None
                rotation = _get_rotation(elem)
                eq_id = (eq_def.get("id") or "").strip()
                eq_name = (eq_def.get("name") or matched_name or eq_id or "").strip()
                linked_set = get_type_set(eq_def)
                result = {
                    "eq_def": eq_def,
                    "linked_set": linked_set,
                    "led_entry": None,
                    "eq_id": eq_id,
                    "eq_name": eq_name,
                    "base_point": element_point,
                    "element_point": element_point,
                    "rotation_deg": rotation,
                    "payload_point": None,
                    "payload_rotation": None,
                    "led_id": None,
                }
                if include_reason:
                    return result, None
                return result
            reason = {"code": "missing-equipment-by-name", "candidates": candidates}
            return (None, reason) if include_reason else None
        return (None, {"code": "missing-metadata"}) if include_reason else None
    payload = _parse_payload(payload_text)
    entry = _find_equipment_by_led(data, payload.get("led_id"))
    if not entry:
        return (None, {"code": "missing-equipment"}) if include_reason else None
    eq_def, linked_set, led_entry = entry
    payload_point = payload.get("location")
    payload_rotation = payload.get("rotation_deg") or 0.0
    live_point = _transform_point(_get_point(elem), link_transform)
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
    result = {
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
    if include_reason:
        return result, None
    return result


def _summarize(children_updates, parent_name):
    if not children_updates:
        forms.alert("No entries were added to the parent equipment definition.", title=TITLE)
        return
    lines = ["Added {} type(s) to '{}':".format(len(children_updates), parent_name)]
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
            selection = list(revit.pick_elements(message="Select element(s) to attach to the parent"))
        except Exception:
            selection = []
    if not selection:
        forms.alert("No child elements were selected.", title=TITLE)
        return

    child_entries = _build_child_entries(selection)
    if not child_entries:
        forms.alert("Selected elements could not be processed.", title=TITLE)
        return

    forms.alert("Select the parent element that owns the linked equipment definition.", title=TITLE)
    parent_elem, parent_transform = _pick_parent_element("Select parent element")
    if not parent_elem:
        return

    parent_info, parent_error = _resolve_equipment_info(parent_elem, data, include_reason=True, allow_name_lookup=True, link_transform=parent_transform)
    if not parent_info:
        code = (parent_error or {}).get("code") if isinstance(parent_error, dict) else parent_error
        if code in ("missing-equipment", "missing-equipment-by-name"):
            forms.alert("No equipment definition with that name exists yet, please create one with Add Yaml Profiles", title=TITLE)
        elif code == "missing-location":
            forms.alert("Could not determine the parent element's location.", title=TITLE)
        else:
            forms.alert("Parent element is missing Element_Linker metadata.", title=TITLE)
        return

    parent_eq = parent_info["eq_def"]
    parent_name = parent_info["eq_name"] or parent_info.get("eq_id") or "(unknown)"
    parent_origin_point = parent_info.get("element_point") or parent_info.get("base_point")
    parent_rotation = parent_info.get("rotation_deg") or 0.0
    if parent_origin_point is None:
        forms.alert("Could not determine the parent element's location.", title=TITLE)
        return

    linked_set = parent_info.get("linked_set") or get_type_set(parent_eq)
    if not linked_set:
        forms.alert("Parent equipment definition is missing a linked set to host new types.", title=TITLE)
        return
    led_list = linked_set.setdefault("linked_element_definitions", [])
    set_id = linked_set.get("id") or (parent_eq.get("id") or "")

    metadata_updates = []
    labels_added = []
    for child in child_entries:
        offsets = compute_offsets_from_points(parent_origin_point, parent_rotation, child["point"], child["rotation_deg"])
        led_id = next_led_id(linked_set, parent_eq)
        params = dict(child.get("parameters") or {})
        payload = _build_element_linker_payload(led_id, set_id, child["element"], child["point"])
        params[ELEMENT_LINKER_PARAM_NAME] = payload
        led_entry = {
            "id": led_id,
            "label": child["label"],
            "category": child["category"],
            "is_group": child["is_group"],
            "offsets": [offsets],
            "parameters": params,
            "tags": child["tags"],
        }
        led_list.append(led_entry)
        metadata_updates.append((child["element"], payload))
        labels_added.append(child["label"])

    if not labels_added:
        forms.alert("No new linked element definitions were created.", title=TITLE)
        return

    if metadata_updates and doc:
        t = Transaction(doc, "Store Element Linker metadata")
        try:
            t.Start()
            for element, payload in metadata_updates:
                _set_element_linker_parameter(element, payload)
            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass

    try:
        save_profile_data(data_path, data)
    except Exception as exc:
        forms.alert("Failed to save profileData.yaml:\n\n{}".format(exc), title=TITLE)
        return

    _summarize(labels_added, parent_name)


if __name__ == "__main__":
    main()
