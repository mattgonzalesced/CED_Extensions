# -*- coding: utf-8 -*-
"""Attach selected elements to the linked equipment definition stored in Extensible Storage."""

import math
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Group,
    IndependentTag,
    RevitLinkInstance,
    Transaction,
    TransactionGroup,
    Transform,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ObjectType

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.profile_schema import get_type_set, next_led_id, ensure_equipment_definition  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from LogicClasses.linked_equipment import compute_offsets_from_points, find_equipment_by_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Select Parent Element"
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", ELEMENT_LINKER_PARAM_NAME)
LEVEL_PARAM_NAMES = (
    "SCHEDULE_LEVEL_PARAM",
    "INSTANCE_REFERENCE_LEVEL_PARAM",
    "FAMILY_LEVEL_PARAM",
    "INSTANCE_LEVEL_PARAM",
)


def _iter_level_bips():
    for name in LEVEL_PARAM_NAMES:
        try:
            yield getattr(BuiltInParameter, name)
        except AttributeError:
            continue


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
        names.insert(0, fam_label)
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
        "Number of Poles_CED": ["Number of Poles_CED", "Number of Poles_CEDT"],
        "Apparent Load Input_CED": ["Apparent Load Input_CED", "Apparent Load Input_CEDT"],
        "Voltage_CED": ["Voltage_CED", "Voltage_CEDT"],
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
    if not found.get("Voltage_CED"):
        found["Voltage_CED"] = 120
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
    for bip in _iter_level_bips():
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


def _normalize_angle(angle_deg):
    try:
        value = float(angle_deg)
    except Exception:
        value = 0.0
    while value > 180.0:
        value -= 360.0
    while value <= -180.0:
        value += 360.0
    return value


def _build_element_linker_payload(led_id, set_id, elem, host_point, rotation_override=None, parent_rotation_deg=None):
    point = host_point or _get_point(elem)
    rotation_deg = _get_rotation(elem) if rotation_override is None else rotation_override
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
        "Parent Rotation (deg): {}".format("{:.6f}".format(parent_rotation_deg) if parent_rotation_deg is not None else ""),
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


def _next_eq_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        eq_id = (eq.get("id") or "").strip()
        if eq_id.upper().startswith("EQ-"):
            try:
                num = int(eq_id.split("-")[-1])
                if num > max_id:
                    max_id = num
            except Exception:
                continue
    return max_id + 1


def _max_set_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        for set_entry in eq.get("linked_sets") or []:
            sid = (set_entry.get("id") or "").strip()
            if not sid:
                continue
            upper = sid.upper()
            if upper.startswith("SET-"):
                try:
                    num = int(upper.split("-")[-1])
                    if num > max_id:
                        max_id = num
                except Exception:
                    continue
    return max_id


def _allocate_new_set_id(existing_ids, seed):
    counter = max(seed, 1)
    while True:
        candidate = "SET-{:03d}".format(counter)
        upper = candidate.upper()
        if upper not in existing_ids:
            existing_ids.add(upper)
            return candidate, counter + 1
        counter += 1


def _renumber_led_ids(set_entry, new_set_id):
    replacements = {}
    led_list = set_entry.get("linked_element_definitions") or []
    counter = 0
    for led in led_list:
        if not isinstance(led, dict):
            continue
        old_id = led.get("id") or ""
        if led.get("is_parent_anchor"):
            new_id = "{}-LED-000".format(new_set_id)
        else:
            counter += 1
            new_id = "{}-LED-{:03d}".format(new_set_id, counter)
        if old_id:
            replacements[old_id] = new_id
        led["id"] = new_id
    return replacements


def _update_linker_payload_ids(doc, replacements):
    if doc is None or not replacements:
        return 0
    changed = 0
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    txn = Transaction(doc, "Repair Element_Linker IDs")
    try:
        txn.Start()
        for elem in collector:
            payload = _get_element_linker_payload(elem)
            if not payload:
                continue
            new_payload = payload
            for old, new in replacements.items():
                if old and old in new_payload:
                    new_payload = new_payload.replace(old, new)
            if new_payload != payload:
                if _set_element_linker_parameter(elem, new_payload):
                    changed += 1
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass
        raise
    return changed


def _repair_duplicate_set_ids(doc, data):
    eq_defs = data.get("equipment_definitions") or []
    if not eq_defs:
        return False, {}
    used_ids = set()
    replacements = {}
    seed = max(_max_set_number(data), _next_eq_number(data))
    next_idx = seed + 1
    for eq in eq_defs:
        linked_sets = eq.get("linked_sets") or []
        if not linked_sets:
            continue
        set_entry = linked_sets[0]
        set_id = (set_entry.get("id") or "").strip()
        norm = set_id.upper()
        if norm and norm not in used_ids:
            used_ids.add(norm)
            continue
        new_set_id, next_idx = _allocate_new_set_id(used_ids, next_idx)
        if set_id:
            replacements[set_id] = new_set_id
        replacements.update(_renumber_led_ids(set_entry, new_set_id))
        set_entry["id"] = new_set_id
        set_entry["name"] = "{} Types".format(eq.get("name") or "Types")
    if not replacements:
        return False, {}
    if doc:
        _update_linker_payload_ids(doc, replacements)
    return True, replacements


def _assign_unique_ids_to_new_definition(eq_def, data):
    next_idx = _next_eq_number(data)
    eq_def["id"] = "EQ-{:03d}".format(next_idx)
    set_id = "SET-{:03d}".format(next_idx)
    for linked_set in eq_def.get("linked_sets") or []:
        linked_set["id"] = set_id
        linked_set["name"] = "{} Types".format(eq_def.get("name") or "Types")
    return set_id


def _find_equipment_definition_by_name(data, eq_name):
    target = (eq_name or "").strip().lower()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        current = (eq.get("name") or eq.get("id") or "").strip().lower()
        if current == target:
            return eq
    return None


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
    parent_rot = payload.get("parent rotation (deg)")
    if parent_rot is not None:
        try:
            payload["parent_rotation_deg"] = float(parent_rot)
        except Exception:
            payload["parent_rotation_deg"] = None
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


def _orientation_vector(elem):
    if elem is None:
        return None
    try:
        location = getattr(elem, "Location", None)
    except Exception:
        location = None
    if location is not None and hasattr(location, "Rotation"):
        try:
            ang = float(location.Rotation)
            return XYZ(math.cos(ang), math.sin(ang), 0.0)
        except Exception:
            pass
    try:
        facing = getattr(elem, "FacingOrientation", None)
        if facing and (abs(facing.X) > 1e-9 or abs(facing.Y) > 1e-9):
            return XYZ(facing.X, facing.Y, 0.0)
    except Exception:
        pass
    try:
        hand = getattr(elem, "HandOrientation", None)
        if hand and (abs(hand.X) > 1e-9 or abs(hand.Y) > 1e-9):
            return XYZ(hand.X, hand.Y, 0.0)
    except Exception:
        pass
    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        basis = getattr(transform, "BasisX", None)
        if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
            return XYZ(basis.X, basis.Y, 0.0)
        basis = getattr(transform, "BasisY", None)
        if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
            return XYZ(basis.X, basis.Y, 0.0)
    return None


def _get_rotation(elem):
    vec = _orientation_vector(elem)
    if vec is None:
        return 0.0
    try:
        return float(math.degrees(math.atan2(vec.Y, vec.X)))
    except Exception:
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
        for name in level_param_names:
            bip = getattr(BuiltInParameter, name, None)
            if not bip:
                continue
            try:
                param = elem.get_Parameter(bip)
            except Exception:
                param = None
            if not param:
                continue
            try:
                eid = param.AsElementId()
            except Exception:
                eid = None
            if eid and doc:
                try:
                    level_elem = doc.GetElement(eid)
                except Exception:
                    level_elem = None
                if level_elem:
                    break
    level_elev = 0.0
    if level_elem:
        try:
            level_elev = getattr(level_elem, "Elevation", 0.0) or 0.0
        except Exception:
            level_elev = 0.0
    world_z = world_point.Z if world_point else 0.0
    relative_ft = world_z - level_elev
    return _feet_to_inches(relative_ft)


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


def _transform_rotation(rotation_deg, transform):
    if transform is None:
        return rotation_deg
    try:
        ang = math.radians(float(rotation_deg or 0.0))
    except Exception:
        ang = 0.0
    vec = XYZ(math.cos(ang), math.sin(ang), 0.0)
    try:
        world_vec = transform.OfVector(vec)
    except Exception:
        return rotation_deg
    try:
        return math.degrees(math.atan2(world_vec.Y, world_vec.X))
    except Exception:
        return rotation_deg


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
                rotation = _transform_rotation(_get_rotation(elem), link_transform)
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
                    "link_transform": link_transform,
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
    payload_parent_rotation = payload.get("parent_rotation_deg")
    live_point = _transform_point(_get_point(elem), link_transform)
    live_rotation = _get_rotation(elem)
    element_point = live_point or payload_point
    rotation = live_rotation if live_point else payload_rotation
    rotation = _transform_rotation(rotation, link_transform)
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
    stored_rot_offset = float((offsets.get("rotation_deg") or 0.0))
    if payload_parent_rotation is not None:
        parent_rotation = payload_parent_rotation
    else:
        parent_rotation = rotation - stored_rot_offset
    base_point = element_point - _rotate_xy(local_vec, parent_rotation)
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
        "rotation_deg": parent_rotation,
        "payload_point": payload_point,
        "payload_rotation": payload_rotation,
        "parent_rotation": parent_rotation,
        "led_id": led_id,
        "link_transform": link_transform,
    }
    if include_reason:
        return result, None
    return result


def _seed_parent_equipment_definition(parent_elem, data, link_transform=None):
    candidates = _candidate_equipment_names(parent_elem)
    default_name = candidates[0] if candidates else ""
    eq_name = forms.ask_for_string(
        prompt="Enter a name for the new equipment definition",
        title=TITLE,
        default=default_name,
    )
    if not eq_name:
        forms.alert("Equipment definition creation canceled.", title=TITLE)
        return None
    label, is_group = _build_label_info(parent_elem)
    sample_entry = {
        "label": label,
        "category_name": _get_category_name(parent_elem),
        "is_group": is_group,
    }
    existing = _find_equipment_definition_by_name(data, eq_name)
    if existing:
        overwrite = forms.alert(
            "An equipment definition named '{}' already exists.\n"
            "Selecting parent again will overwrite its contents.\n"
            "Do you want to continue?".format(eq_name),
            title=TITLE,
            ok=False,
            yes=True,
            no=True,
        )
        if not overwrite:
            forms.alert("Select Parent Element canceled.", title=TITLE)
            return None
    eq_def = ensure_equipment_definition(data, eq_name, sample_entry)
    newly_created = existing is None
    if newly_created:
        _assign_unique_ids_to_new_definition(eq_def, data)
    linked_set = get_type_set(eq_def)
    led_list = linked_set.setdefault("linked_element_definitions", [])
    set_id = linked_set.get("id") or (eq_def.get("id") or "SET")

    parent_point = _transform_point(_get_point(parent_elem), link_transform)
    if parent_point is None:
        forms.alert("Could not determine the parent element's location.", title=TITLE)
        return None
    parent_rotation = _transform_rotation(_get_rotation(parent_elem), link_transform)
    led_id = "{}-LED-000".format(set_id)
    params = dict(_collect_params(parent_elem) or {})
    payload = _build_element_linker_payload(led_id, set_id, parent_elem, parent_point, parent_rotation, parent_rotation)
    params[ELEMENT_LINKER_PARAM_NAME] = payload
    led_entry = {
        "id": led_id,
        "label": label,
        "category": _get_category_name(parent_elem),
        "is_group": is_group,
        "offsets": [{
            "x_inches": 0.0,
            "y_inches": 0.0,
            "z_inches": 0.0,
            "rotation_deg": 0.0,
        }],
        "parameters": params,
        "tags": _collect_hosted_tags(parent_elem, parent_point),
        "is_parent_anchor": True,
    }
    led_list.append(led_entry)

    doc = revit.doc
    if doc:
        txn = Transaction(doc, "Select Parent Element: Initialize Parent Definition")
        try:
            txn.Start()
            _set_element_linker_parameter(parent_elem, payload)
            txn.Commit()
        except Exception:
            try:
                txn.RollBack()
            except Exception:
                pass
    try:
        save_active_yaml_data(
            None,
            data,
            "Select Parent Element",
            "Created equipment definition '{}'".format(eq_name),
        )
    except Exception:
        pass
    info = {
        "eq_def": eq_def,
        "linked_set": linked_set,
        "led_entry": led_entry,
        "eq_id": eq_def.get("id") or "",
        "eq_name": eq_name,
        "base_point": parent_point,
        "element_point": parent_point,
        "rotation_deg": parent_rotation,
        "payload_point": parent_point,
        "payload_rotation": parent_rotation,
        "led_id": led_id,
        "link_transform": link_transform,
    }
    return info


def _prune_anchor_only_definitions(data):
    changed = False
    eq_defs = data.get("equipment_definitions") or []
    survivors = []
    for entry in eq_defs:
        linked_sets = entry.get("linked_sets") or []
        has_real = False
        for linked_set in linked_sets:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                if not led.get("is_parent_anchor"):
                    has_real = True
                    break
            if has_real:
                break
        if has_real:
            survivors.append(entry)
        else:
            changed = True
    if changed:
        data["equipment_definitions"] = survivors
    return changed


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

    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    repaired, replacements = _repair_duplicate_set_ids(doc, data)
    if repaired:
        try:
            save_active_yaml_data(
                None,
                data,
                "Repair Equipment Definition IDs",
                "Reassigned duplicate set identifiers",
            )
        except Exception as repair_exc:
            forms.alert("Failed to repair duplicate IDs:\n\n{}".format(repair_exc), title=TITLE)
            return
    yaml_label = get_yaml_display_name(data_path)

    trans_group = TransactionGroup(doc, "Select Parent Element")
    trans_group.Start()
    success = False
    try:
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

        parent_definition_seeded = False
        parent_info, parent_error = _resolve_equipment_info(parent_elem, data, include_reason=True, allow_name_lookup=True, link_transform=parent_transform)
        if not parent_info:
            code = (parent_error or {}).get("code") if isinstance(parent_error, dict) else parent_error
            if code in ("missing-equipment", "missing-equipment-by-name"):
                seed_result = _seed_parent_equipment_definition(parent_elem, data, parent_transform)
                if not seed_result:
                    return
                parent_definition_seeded = True
                parent_info = seed_result
                parent_info, parent_error = _resolve_equipment_info(
                    parent_elem,
                    data,
                    include_reason=True,
                    allow_name_lookup=True,
                    link_transform=parent_transform,
                )
                if not parent_info:
                    code = (parent_error or {}).get("code") if isinstance(parent_error, dict) else parent_error
                    if code == "missing-location":
                        forms.alert("Could not determine the parent element's location.", title=TITLE)
                    else:
                        forms.alert("Parent element is missing Element_Linker metadata.", title=TITLE)
                    return
            elif code == "missing-location":
                forms.alert("Could not determine the parent element's location.", title=TITLE)
                return
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

        parent_transform = parent_info.get("link_transform")
        if parent_transform:
            parent_doc = getattr(parent_elem, "Document", None)
            for entry in child_entries:
                child_elem = entry.get("element")
                if parent_doc is not None and getattr(child_elem, "Document", None) is parent_doc:
                    entry["point"] = _transform_point(entry.get("point"), parent_transform)
                    entry["rotation_deg"] = _transform_rotation(entry.get("rotation_deg"), parent_transform)

        linked_set = parent_info.get("linked_set") or get_type_set(parent_eq)
        if not linked_set:
            forms.alert("Parent equipment definition is missing a linked set to host new types.", title=TITLE)
            return
        led_list = linked_set.setdefault("linked_element_definitions", [])
        set_id = linked_set.get("id") or (parent_eq.get("id") or "")

        metadata_updates = []
        labels_added = []
        for child in child_entries:
            child_rotation = child.get("rotation_deg")
            offsets = compute_offsets_from_points(parent_origin_point, parent_rotation, child["point"], child_rotation)
            offsets["z_inches"] = _level_relative_z_inches(child["element"], child["point"])
            if child.get("is_group"):
                rel_rot = _normalize_angle((child_rotation or 0.0) - (parent_rotation or 0.0))
                offsets["rotation_deg"] = rel_rot
            led_id = next_led_id(linked_set, parent_eq)
            params = dict(child.get("parameters") or {})
            payload = _build_element_linker_payload(
                led_id,
                set_id,
                child["element"],
                child["point"],
                child_rotation,
                parent_rotation,
            )
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
            if parent_definition_seeded:
                _remove_seeded_parent_definition(parent_info, data, parent_elem)
            forms.alert("No new linked element definitions were created.", title=TITLE)
            return

        if metadata_updates and doc:
            txn_name = "Select Parent Element: Store Element Linker metadata ({})".format(len(metadata_updates))
            t = Transaction(doc, txn_name)
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

        _prune_anchor_only_definitions(data)
        try:
            save_active_yaml_data(
                None,
                data,
                "Select Parent Element",
                "Added {} type(s) to {}".format(len(labels_added), parent_name),
            )
        except Exception as exc:
            forms.alert("Failed to save {}:\n\n{}".format(yaml_label, exc), title=TITLE)
            return

        _summarize(labels_added, parent_name)
        success = True
    finally:
        try:
            if success:
                trans_group.Assimilate()
            else:
                trans_group.RollBack()
        except Exception:
            pass


if __name__ == "__main__":
    main()
def _remove_seeded_parent_definition(parent_info, data, parent_elem):
    eq_def = parent_info.get("eq_def")
    if not eq_def:
        return
    eq_id = eq_def.get("id")
    defs = data.get("equipment_definitions") or []
    data["equipment_definitions"] = [entry for entry in defs if entry is not eq_def]
    doc = getattr(revit, "doc", None)
    if doc and parent_elem:
        txn = Transaction(doc, "Select Parent Element: Clear Placeholder Definition")
        try:
            txn.Start()
            _set_element_linker_parameter(parent_elem, "")
            txn.Commit()
        except Exception:
            try:
                txn.RollBack()
            except Exception:
                pass
