# -*- coding: utf-8 -*-
"""Add Equipment to Profiles
-------------------------
Append equipment elements to an existing YAML profile by selecting the linked
parent element first and then selecting the equipment to add.
"""

import math
import os
import sys

from pyrevit import revit, forms, script
output = script.get_output()
output.close_others()
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
    UnitUtils,
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

TITLE = "Add Equipment to Profiles"
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", ELEMENT_LINKER_PARAM_NAME)
LEVEL_PARAM_NAMES = (
    "SCHEDULE_LEVEL_PARAM",
    "INSTANCE_REFERENCE_LEVEL_PARAM",
    "FAMILY_LEVEL_PARAM",
    "INSTANCE_LEVEL_PARAM",
)




def _element_id_value(elem_id, default=None):
    if elem_id is None:
        return default
    for attr in ("Value", "IntegerValue"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return int(value)
        except Exception:
            try:
                return value
            except Exception:
                continue
    return default
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
        "Comments": ["Comments", "Comment", "Comments_CED", "Comments_CEDT"],
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
                internal = param.AsDouble()
                unit_id = None
                get_unit = getattr(param, "GetUnitTypeId", None)
                if callable(get_unit):
                    try:
                        unit_id = get_unit()
                    except Exception:
                        unit_id = None
                if unit_id is not None:
                    try:
                        found[target_key] = UnitUtils.ConvertFromInternalUnits(float(internal), unit_id)
                    except Exception:
                        found[target_key] = internal
                else:
                    found[target_key] = internal
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
        if lvl and _element_id_value(lvl, -1) > 0:
            return _element_id_value(lvl)
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
            if eid and _element_id_value(eid, -1) > 0:
                return _element_id_value(eid)
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


def _build_element_linker_payload(led_id, set_id, elem, host_point, rotation_override=None, parent_rotation_deg=None, parent_elem_id=None):
    point = host_point or _get_point(elem)
    rotation_deg = _get_rotation(elem) if rotation_override is None else rotation_override
    level_id = _get_level_element_id(elem)
    try:
        elem_id = _element_id_value(elem.Id)
    except Exception:
        elem_id = ""
    facing = getattr(elem, "FacingOrientation", None)
    lines = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Location XYZ (ft): {}".format(_format_xyz(point)),
        "Rotation (deg): {:.6f}".format(rotation_deg),
        "Parent Rotation (deg): {}".format("{:.6f}".format(parent_rotation_deg) if parent_rotation_deg is not None else ""),
        "Parent ElementId: {}".format(parent_elem_id if parent_elem_id is not None else ""),
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


def _tag_signature(name):
    return " ".join((name or "").strip().split()).lower()


def _is_tag_like(elem):
    if isinstance(elem, IndependentTag):
        return True
    try:
        _ = elem.TagHeadPosition
    except Exception:
        return False
    return True


def _normalize_keynote_family(value):
    if not value:
        return ""
    text = str(value)
    if ":" in text:
        text = text.split(":", 1)[0]
    return "".join([ch for ch in text.lower() if ch.isalnum()])


def _is_ga_keynote_symbol(family_name):
    return _normalize_keynote_family(family_name) == "gakeynotesymbolced"


def _annotation_family_type(elem):
    fam_name = None
    type_name = None
    try:
        symbol = getattr(elem, "Symbol", None)
        if symbol:
            fam = getattr(symbol, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else getattr(symbol, "FamilyName", None)
            type_name = getattr(symbol, "Name", None)
            if not type_name and hasattr(symbol, "get_Parameter"):
                param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    type_name = param.AsString()
    except Exception:
        pass
    if fam_name and type_name:
        return fam_name, type_name
    label = getattr(elem, "Name", None)
    return label or "", type_name or ""


def _is_ga_keynote_symbol_element(elem):
    if elem is None:
        return False
    try:
        cat = getattr(elem, "Category", None)
        cat_name = getattr(cat, "Name", None) if cat else ""
    except Exception:
        cat_name = ""
    if "generic annotation" not in (cat_name or "").lower():
        return False
    fam_name, _ = _annotation_family_type(elem)
    return _is_ga_keynote_symbol(fam_name)


def _collect_element_parameters(elem, include_read_only=True):
    results = {}
    if elem is None:
        return results
    for param in getattr(elem, "Parameters", []) or []:
        try:
            definition = getattr(param, "Definition", None)
            name = getattr(definition, "Name", None)
        except Exception:
            name = None
        if not name:
            continue
        if not include_read_only:
            try:
                if param.IsReadOnly:
                    continue
            except Exception:
                pass
        try:
            storage = param.StorageType.ToString()
        except Exception:
            storage = ""
        try:
            if storage == "String":
                value = param.AsString()
                if value is None:
                    try:
                        value = param.AsValueString()
                    except Exception:
                        value = None
            elif storage == "Integer":
                value = param.AsInteger()
            elif storage == "Double":
                value = param.AsDouble()
            elif storage == "ElementId":
                elem_id = param.AsElementId()
                value = _element_id_value(elem_id) if elem_id else None
                if value is None:
                    try:
                        value = param.AsValueString()
                    except Exception:
                        value = None
            else:
                value = param.AsValueString()
        except Exception:
            value = None
        if value is None:
            continue
        results[name] = value
    return results


def _normalize_keynote_params(params):
    if not isinstance(params, dict):
        return {}
    out = dict(params)
    if out.get("Keynote Value") not in (None, ""):
        out.pop("Key Value", None)
        out.pop("Keynote", None)
    elif out.get("Key Value") not in (None, ""):
        out["Keynote Value"] = out.pop("Key Value")
        out.pop("Keynote", None)
    elif out.get("Keynote") not in (None, ""):
        out["Keynote Value"] = out.pop("Keynote")
    if out.get("Keynote Description") not in (None, ""):
        out.pop("Keynote Text", None)
    elif out.get("Keynote Text") not in (None, ""):
        out["Keynote Description"] = out.pop("Keynote Text")
    return out


def _collect_keynote_parameters(annotation_elem):
    merged = {}
    type_elem = None
    try:
        sym = getattr(annotation_elem, "Symbol", None)
        if sym:
            type_elem = sym
    except Exception:
        type_elem = None
    if type_elem is None:
        try:
            doc = getattr(annotation_elem, "Document", None)
            type_id = annotation_elem.GetTypeId()
            if doc and type_id:
                type_elem = doc.GetElement(type_id)
        except Exception:
            type_elem = None
    if type_elem is not None:
        merged.update(_collect_element_parameters(type_elem, include_read_only=True))
    merged.update(_collect_element_parameters(annotation_elem, include_read_only=True))

    params = {}
    for key in ("Keynote Value", "Key Value", "Keynote"):
        value = merged.get(key)
        if value not in (None, ""):
            params["Keynote Value"] = value
            break
    for key in ("Keynote Description", "Keynote Text"):
        value = merged.get(key)
        if value not in (None, ""):
            params["Keynote Description"] = value
            break
    return _normalize_keynote_params(params)


def _keynote_entry_key(keynote_entry):
    if not isinstance(keynote_entry, dict):
        return None
    family = keynote_entry.get("family_name") or keynote_entry.get("family") or ""
    type_name = keynote_entry.get("type_name") or keynote_entry.get("type") or ""
    category = keynote_entry.get("category_name") or keynote_entry.get("category") or ""
    params = _normalize_keynote_params(keynote_entry.get("parameters") or {})
    key_value = params.get("Keynote Value")
    key_desc = params.get("Keynote Description")
    return (
        _tag_signature(family),
        _tag_signature(type_name),
        _tag_signature(category),
        "" if key_value in (None, "") else str(key_value).strip(),
        "" if key_desc in (None, "") else str(key_desc).strip(),
    )


def _build_keynote_entry(annotation_elem, host_point):
    if annotation_elem is None or host_point is None:
        return None
    if not _is_ga_keynote_symbol_element(annotation_elem):
        return None
    ann_point = _get_point(annotation_elem)
    if ann_point is None:
        return None
    fam_name, type_name = _annotation_family_type(annotation_elem)
    if not fam_name:
        return None
    offsets = {
        "x_inches": _feet_to_inches(ann_point.X - host_point.X),
        "y_inches": _feet_to_inches(ann_point.Y - host_point.Y),
        "z_inches": _feet_to_inches(ann_point.Z - host_point.Z),
        "rotation_deg": 0.0,
    }
    category_name = _get_category_name(annotation_elem) or "Generic Annotations"
    return {
        "family_name": fam_name,
        "type_name": type_name or "",
        "category_name": category_name,
        "parameters": _collect_keynote_parameters(annotation_elem),
        "offsets": offsets,
    }


def _tag_host_element_id(tag):
    if tag is None:
        return None
    try:
        getter = getattr(tag, "GetTaggedLocalElementIds", None)
        if callable(getter):
            ids = list(getter() or [])
            if ids:
                return ids[0]
    except Exception:
        pass
    for attr in ("TaggedLocalElementId", "TaggedElementId"):
        try:
            value = getattr(tag, attr, None)
        except Exception:
            value = None
        if value:
            return value
    return None


def _tag_entry_key(tag_entry):
    if not tag_entry:
        return None
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
        type_name = tag_entry.get("type_name") or tag_entry.get("type") or ""
        category = tag_entry.get("category_name") or tag_entry.get("category") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
        type_name = getattr(tag_entry, "type_name", None) or getattr(tag_entry, "type", None) or ""
        category = getattr(tag_entry, "category_name", None) or getattr(tag_entry, "category", None) or ""
    return (_tag_signature(family), _tag_signature(type_name), _tag_signature(category))


def _tag_offsets_near(tag_a, tag_b, pos_tol=0.05, rot_tol=0.5):
    def _extract(entry):
        offsets = {}
        if isinstance(entry, dict):
            offsets = entry.get("offsets") or {}
        else:
            offsets = getattr(entry, "offsets", None) or {}
        try:
            x_val = float(offsets.get("x_inches", 0.0) or 0.0)
        except Exception:
            x_val = 0.0
        try:
            y_val = float(offsets.get("y_inches", 0.0) or 0.0)
        except Exception:
            y_val = 0.0
        try:
            z_val = float(offsets.get("z_inches", 0.0) or 0.0)
        except Exception:
            z_val = 0.0
        try:
            rot_val = float(offsets.get("rotation_deg", 0.0) or 0.0)
        except Exception:
            rot_val = 0.0
        return (x_val, y_val, z_val, rot_val)

    ax, ay, az, ar = _extract(tag_a)
    bx, by, bz, br = _extract(tag_b)
    return (
        abs(ax - bx) <= pos_tol
        and abs(ay - by) <= pos_tol
        and abs(az - bz) <= pos_tol
        and abs(ar - br) <= rot_tol
    )


def _point_offsets_dict(point, origin):
    if point is None or origin is None:
        return None
    return {
        "x_inches": _feet_to_inches(point.X - origin.X),
        "y_inches": _feet_to_inches(point.Y - origin.Y),
        "z_inches": _feet_to_inches(point.Z - origin.Z),
    }


def _build_independent_tag_entry(tag, host_point):
    if tag is None or host_point is None:
        return None
    try:
        tag_point = tag.TagHeadPosition
    except Exception:
        tag_point = None
    if tag_point is None:
        return None
    doc = getattr(tag, "Document", None)
    tag_symbol = None
    if doc is not None:
        try:
            tag_symbol = doc.GetElement(tag.GetTypeId())
        except Exception:
            tag_symbol = None
    fam_name = None
    type_name = None
    category_name = None
    if tag_symbol:
        try:
            fam = getattr(tag_symbol, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else getattr(tag_symbol, "FamilyName", None)
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
            type_name = None
    if not category_name:
        try:
            cat = getattr(tag, "Category", None)
            category_name = getattr(cat, "Name", None) if cat else None
        except Exception:
            category_name = None
    if not (fam_name and type_name):
        return None
    offsets = {
        "x_inches": _feet_to_inches((tag_point.X - host_point.X)),
        "y_inches": _feet_to_inches((tag_point.Y - host_point.Y)),
        "z_inches": _feet_to_inches((tag_point.Z - host_point.Z)),
        "rotation_deg": 0.0,
    }
    leader_elbow = None
    leader_end = None
    try:
        elbow_point = getattr(tag, "LeaderElbow", None)
    except Exception:
        elbow_point = None
    if elbow_point:
        leader_elbow = _point_offsets_dict(elbow_point, host_point)
    try:
        end_point = getattr(tag, "LeaderEnd", None)
    except Exception:
        end_point = None
    if end_point:
        leader_end = _point_offsets_dict(end_point, host_point)
    return {
        "family_name": fam_name,
        "type_name": type_name,
        "category_name": category_name,
        "parameters": {},
        "offsets": offsets,
        "leader_elbow": leader_elbow,
        "leader_end": leader_end,
    }


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
        if tag is None or not _is_tag_like(tag):
            continue
        entry = _build_independent_tag_entry(tag, host_point)
        if entry:
            tags.append(entry)
    return tags


def _collect_hosted_keynotes(elem, host_point):
    doc = getattr(elem, "Document", None)
    if doc is None or host_point is None:
        return []
    try:
        deps = list(elem.GetDependentElements(None))
    except Exception:
        deps = []
    keynotes = []
    for dep_id in deps:
        try:
            dep_elem = doc.GetElement(dep_id)
        except Exception:
            dep_elem = None
        entry = _build_keynote_entry(dep_elem, host_point)
        if entry:
            keynotes.append(entry)
    return keynotes


def _find_closest_entry_by_point(entries, point):
    if not entries or point is None:
        return None
    closest_idx = None
    closest_dist = None
    for idx, entry in enumerate(entries):
        host_point = entry.get("point")
        if host_point is None:
            continue
        try:
            dist = host_point.DistanceTo(point)
        except Exception:
            try:
                dx = host_point.X - point.X
                dy = host_point.Y - point.Y
                dz = host_point.Z - point.Z
                dist = math.sqrt(dx * dx + dy * dy + dz * dz)
            except Exception:
                continue
        if closest_idx is None or dist < closest_dist:
            closest_idx = idx
            closest_dist = dist
    return closest_idx


def _assign_selected_tags(child_entries, tag_elems):
    if not child_entries or not tag_elems:
        return
    host_index = {}
    for idx, entry in enumerate(child_entries):
        elem = entry.get("element")
        if elem is None:
            continue
        try:
            elem_id = _element_id_value(elem.Id)
        except Exception:
            elem_id = None
        if elem_id is not None and elem_id not in host_index:
            host_index[elem_id] = idx
    for tag in tag_elems:
        target_idx = None
        host_id = _tag_host_element_id(tag)
        host_id_val = None
        if host_id is not None:
            try:
                host_id_val = _element_id_value(host_id)
            except Exception:
                try:
                    host_id_val = int(host_id)
                except Exception:
                    host_id_val = None
            if host_id_val is not None:
                target_idx = host_index.get(host_id_val)
                if target_idx is None:
                    continue
        if target_idx is None:
            try:
                tag_point = tag.TagHeadPosition
            except Exception:
                tag_point = None
            if tag_point is None:
                continue
            target_idx = _find_closest_entry_by_point(child_entries, tag_point)
        if target_idx is None:
            continue
        host_point = child_entries[target_idx].get("point")
        if host_point is None:
            continue
        entry = _build_independent_tag_entry(tag, host_point)
        if not entry:
            continue
        tags = child_entries[target_idx].setdefault("tags", [])
        entry_key = _tag_entry_key(entry)
        if entry_key:
            existing = False
            for existing_tag in tags:
                if _tag_entry_key(existing_tag) != entry_key:
                    continue
                if _tag_offsets_near(existing_tag, entry):
                    existing = True
                    break
            if existing:
                continue
        tags.append(entry)


def _assign_selected_keynotes(child_entries, keynote_elems):
    if not child_entries or not keynote_elems:
        return
    for keynote_elem in keynote_elems:
        if keynote_elem is None:
            continue
        keynote_point = _get_point(keynote_elem)
        if keynote_point is None:
            continue
        target_idx = _find_closest_entry_by_point(child_entries, keynote_point)
        if target_idx is None:
            continue
        host_point = child_entries[target_idx].get("point")
        if host_point is None:
            continue
        entry = _build_keynote_entry(keynote_elem, host_point)
        if not entry:
            continue
        keynotes = child_entries[target_idx].setdefault("keynotes", [])
        entry_key = _keynote_entry_key(entry)
        if entry_key:
            exists = False
            for existing in keynotes:
                if _keynote_entry_key(existing) != entry_key:
                    continue
                if _tag_offsets_near(existing, entry):
                    exists = True
                    break
            if exists:
                continue
        keynotes.append(entry)


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
            "local_point": point,
            "rotation_deg": _get_rotation(elem),
            "label": label,
            "is_group": is_group,
            "category": _get_category_name(elem),
            "parameters": _collect_params(elem),
            "tags": _collect_hosted_tags(elem, point),
            "keynotes": _collect_hosted_keynotes(elem, point),
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
    parent_elem = payload.get("parent elementid")
    if parent_elem:
        try:
            payload["parent_element_id"] = int(parent_elem)
        except Exception:
            payload["parent_element_id"] = None
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
    direct = _instance_elevation_inches(elem)
    if direct is not None:
        return direct
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


def _instance_elevation_inches(elem):
    param_names = (
        "INSTANCE_ELEV_PARAM",
        "INSTANCE_ELEVATION_PARAM",
        "INSTANCE_FREE_HOST_OFFSET_PARAM",
        "INSTANCE_SILL_HEIGHT_PARAM",
        "SILL_HEIGHT_PARAM",
        "HEAD_HEIGHT_PARAM",
    )
    for name in param_names:
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
            raw = param.AsDouble()
        except Exception:
            raw = None
        if raw is None:
            continue
        return _feet_to_inches(raw)
    return None


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
    try:
        parent_elem_id = _element_id_value(parent_elem.Id)
    except Exception:
        parent_elem_id = None
    led_id = "{}-LED-000".format(set_id)
    params = dict(_collect_params(parent_elem) or {})
    payload = _build_element_linker_payload(
        led_id,
        set_id,
        parent_elem,
        parent_point,
        parent_rotation,
        parent_rotation,
        parent_elem_id,
    )
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
        "keynotes": _collect_hosted_keynotes(parent_elem, parent_point),
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
        forms.alert("No entries were added to the profile.", title=TITLE)
        return
    lines = ["Added {} item(s) to '{}':".format(len(children_updates), parent_name)]
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
    yaml_label = get_yaml_display_name(data_path)

    forms.alert("Select the linked parent element that already has a profile.", title=TITLE)
    parent_elem, parent_transform = _pick_parent_element("Select parent element")
    if not parent_elem:
        return

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
            forms.alert(
                "The selected element has no profile in {}. Operation canceled.".format(yaml_label),
                title=TITLE,
            )
        return

    parent_eq = parent_info["eq_def"]
    parent_name = parent_info["eq_name"] or parent_info.get("eq_id") or "(unknown)"
    parent_origin_point = parent_info.get("element_point") or parent_info.get("base_point")
    parent_rotation = parent_info.get("rotation_deg") or 0.0
    try:
        parent_elem_id = _element_id_value(parent_elem.Id)
    except Exception:
        parent_elem_id = None
    if parent_origin_point is None:
        forms.alert("Could not determine the parent element's location.", title=TITLE)
        return

    forms.alert(
        "Select the elements to add to the profile.\nYou can also select tags/keynotes to capture with those elements.",
        title=TITLE,
    )
    try:
        selection = list(revit.pick_elements(message="Select equipment, tags, and keynotes to add to '{}'".format(parent_name)))
    except Exception:
        selection = []
    if not selection:
        forms.alert("No elements were selected.", title=TITLE)
        return

    tag_elems = []
    keynote_elems = []
    equipment_selection = []
    for elem in selection:
        if elem is None:
            continue
        if _is_tag_like(elem):
            tag_elems.append(elem)
        elif _is_ga_keynote_symbol_element(elem):
            keynote_elems.append(elem)
        else:
            equipment_selection.append(elem)
    if not equipment_selection:
        forms.alert("Select at least one equipment element. Tags are optional.", title=TITLE)
        return

    child_entries = _build_child_entries(equipment_selection)
    if not child_entries:
        forms.alert("Selected elements could not be processed.", title=TITLE)
        return
    _assign_selected_tags(child_entries, tag_elems)
    _assign_selected_keynotes(child_entries, keynote_elems)

    parent_doc = getattr(parent_elem, "Document", None)
    parent_transform = parent_info.get("link_transform")
    if parent_transform:
        for entry in child_entries:
            child_elem = entry.get("element")
            if parent_doc is not None and getattr(child_elem, "Document", None) is parent_doc:
                if entry.get("local_point") is None:
                    entry["local_point"] = entry.get("point")
                entry["point"] = _transform_point(entry.get("point"), parent_transform)
                entry["rotation_deg"] = _transform_rotation(entry.get("rotation_deg"), parent_transform)

    linked_set = parent_info.get("linked_set") or get_type_set(parent_eq)
    if not linked_set:
        forms.alert("Parent profile is missing a linked set to host new items.", title=TITLE)
        return
    led_list = linked_set.setdefault("linked_element_definitions", [])
    set_id = linked_set.get("id") or (parent_eq.get("id") or "")

    trans_group = TransactionGroup(doc, TITLE)
    trans_group.Start()
    success = False
    try:
        metadata_updates = []
        labels_added = []
        for child in child_entries:
            child_rotation = child.get("rotation_deg")
            offsets = compute_offsets_from_points(parent_origin_point, parent_rotation, child["point"], child_rotation)
            z_point = child.get("local_point") or child.get("point")
            offsets["z_inches"] = _level_relative_z_inches(child["element"], z_point)
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
                parent_elem_id,
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
                "keynotes": child.get("keynotes") or [],
            }
            led_list.append(led_entry)
            metadata_updates.append((child["element"], payload))
            labels_added.append(child["label"])

        if not labels_added:
            forms.alert("No new linked element definitions were created.", title=TITLE)
            return

        if metadata_updates and doc:
            txn_name = "Add Equipment to Profiles: Store Element Linker metadata ({})".format(len(metadata_updates))
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

        try:
            save_active_yaml_data(
                None,
                data,
                TITLE,
                "Added {} item(s) to {}".format(len(labels_added), parent_name),
            )
        except Exception as exc:
            forms.alert("Failed to save {}\n\n{}".format(yaml_label, exc), title=TITLE)
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
