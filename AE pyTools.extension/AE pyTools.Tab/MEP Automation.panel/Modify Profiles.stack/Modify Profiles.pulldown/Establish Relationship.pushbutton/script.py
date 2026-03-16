# -*- coding: utf-8 -*-
"""Create one independent parent/children relationship profile and parameter flag mappings."""

import math
import os
import sys
import imp

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import BuiltInParameter, Group, Transaction, TransactionGroup

output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.linked_equipment import compute_offsets_from_points  # noqa: E402
from LogicClasses.profile_schema import (  # noqa: E402
    ensure_equipment_definition,
    get_type_set,
    next_led_id,
)
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str

TITLE = "Establish Relationship"
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", ELEMENT_LINKER_PARAM_NAME)
LEVEL_PARAM_NAMES = (
    "SCHEDULE_LEVEL_PARAM",
    "INSTANCE_REFERENCE_LEVEL_PARAM",
    "FAMILY_LEVEL_PARAM",
    "INSTANCE_LEVEL_PARAM",
)

FLAG_BYPARENT = "BYPARENT"
FLAG_BYSIBLING = "BYSIBLING"
FLAG_STATIC = "static"




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
def _load_active_yaml_data_resilient(doc):
    try:
        return load_active_yaml_data(doc)
    except Exception as exc:
        raise RuntimeError(
            "Could not load active YAML from project Extensible Storage.\n\n{}".format(exc)
        )


def _save_active_yaml_data_resilient(doc, data, action, description):
    save_active_yaml_data(doc, data, action, description)


def _iter_level_bips():
    for name in LEVEL_PARAM_NAMES:
        try:
            yield getattr(BuiltInParameter, name)
        except AttributeError:
            continue


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
            return (math.cos(ang), math.sin(ang))
        except Exception:
            pass
    try:
        facing = getattr(elem, "FacingOrientation", None)
        if facing and (abs(facing.X) > 1e-9 or abs(facing.Y) > 1e-9):
            return (facing.X, facing.Y)
    except Exception:
        pass
    try:
        hand = getattr(elem, "HandOrientation", None)
        if hand and (abs(hand.X) > 1e-9 or abs(hand.Y) > 1e-9):
            return (hand.X, hand.Y)
    except Exception:
        pass
    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        basis = getattr(transform, "BasisX", None)
        if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
            return (basis.X, basis.Y)
        basis = getattr(transform, "BasisY", None)
        if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
            return (basis.X, basis.Y)
    return None


def _get_rotation(elem):
    vec = _orientation_vector(elem)
    if vec is None:
        return 0.0
    try:
        return float(math.degrees(math.atan2(vec[1], vec[0])))
    except Exception:
        return 0.0


def _feet_to_inches(value):
    try:
        return float(value) * 12.0
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
        for bip in _iter_level_bips():
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


def _build_element_linker_payload(
    led_id,
    set_id,
    elem,
    host_point,
    rotation_override=None,
    parent_rotation_deg=None,
    parent_elem_id=None,
):
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
        "Parent Rotation (deg): {}".format(
            "{:.6f}".format(parent_rotation_deg) if parent_rotation_deg is not None else ""
        ),
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
        if not eq_id.upper().startswith("EQ-"):
            continue
        try:
            number = int(eq_id.split("-")[-1])
        except Exception:
            continue
        if number > max_id:
            max_id = number
    return max_id + 1


def _next_set_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            set_id = (linked_set.get("id") or "").strip()
            if not set_id.upper().startswith("SET-"):
                continue
            try:
                number = int(set_id.split("-")[-1])
            except Exception:
                continue
            if number > max_id:
                max_id = number
    return max_id + 1


def _find_equipment_definition_by_name(data, name):
    target = (name or "").strip().lower()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        eq_name = (eq.get("name") or eq.get("id") or "").strip().lower()
        if eq_name == target:
            return eq
    return None


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


def _collect_profile_params(elem):
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

    found = {}
    for key in targets.keys():
        found[key] = ""

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
    for key, value in found.items():
        if key == "dev-Group ID":
            continue
        if value:
            return found
    return {"dev-Group ID": found.get("dev-Group ID", "")}


def _collect_all_parameter_names(elem):
    names = []
    seen = set()
    for param in getattr(elem, "Parameters", []) or []:
        try:
            name = (param.Definition.Name or "").strip()
        except Exception:
            name = ""
        if not name:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        names.append(name)
    names.sort()
    return names


def _read_parameter_value(param):
    if not param:
        return ""
    try:
        storage = param.StorageType.ToString()
    except Exception:
        storage = ""
    try:
        if storage == "String":
            return param.AsString() or ""
        if storage == "Double":
            return param.AsDouble()
        if storage == "Integer":
            return param.AsInteger()
    except Exception:
        pass
    try:
        return param.AsValueString() or ""
    except Exception:
        return ""


def _collect_all_parameter_values(elem):
    values = {}
    for param in getattr(elem, "Parameters", []) or []:
        try:
            name = (param.Definition.Name or "").strip()
        except Exception:
            name = ""
        if not name:
            continue
        if name in ELEMENT_LINKER_PARAM_NAMES:
            continue
        if name in values:
            continue
        values[name] = _read_parameter_value(param)
    return values


def _dedupe_elements(elements):
    unique = []
    seen_ids = set()
    for elem in elements or []:
        if elem is None:
            continue
        try:
            elem_id = _element_id_value(elem.Id)
        except Exception:
            elem_id = None
        if elem_id is None:
            unique.append(elem)
            continue
        if elem_id in seen_ids:
            continue
        seen_ids.add(elem_id)
        unique.append(elem)
    return unique


def _to_short_text(value):
    if value is None:
        return ""
    try:
        text = str(value)
    except Exception:
        text = ""
    text = text.replace("\n", " ").replace("\r", " ").strip()
    if len(text) > 48:
        text = text[:45] + "..."
    return text


def _mapping_parameter_names(params):
    names = []
    for key in sorted((params or {}).keys()):
        if key in ELEMENT_LINKER_PARAM_NAMES:
            continue
        names.append(key)
    return names


def _escape_quotes(text):
    try:
        return str(text or "").replace('"', '\\"')
    except Exception:
        return ""


def _load_mapping_ui_module():
    module_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "EstablishRelationshipMappingWindow.py")
    )
    if not os.path.exists(module_path):
        forms.alert("Mapping UI module not found:\n{}".format(module_path), title=TITLE)
        return None
    try:
        return imp.load_source("ced_establish_relationship_mapping_ui", module_path)
    except Exception as exc:
        forms.alert("Failed to load mapping UI module:\n{}\n\n{}".format(module_path, exc), title=TITLE)
        return None


def _ordered_parent_params(target_param, parent_param_names):
    ordered = []
    seen = set()
    if target_param:
        ordered.append(target_param)
        seen.add(target_param.lower())
    for parent_name in parent_param_names or []:
        key = (parent_name or "").lower()
        if not key or key in seen:
            continue
        seen.add(key)
        ordered.append(parent_name)
    return ordered


def _build_mapping_rows(children_contexts, parent_param_names):
    rows = []
    row_index = 1
    for child_ctx in children_contexts:
        entry = child_ctx.get("entry") or {}
        params = entry.get("parameters") or {}
        all_values = dict(child_ctx.get("all_params") or {})
        tracked_names = _mapping_parameter_names(params)
        available_names = sorted(set(tracked_names) | set(all_values.keys()))
        if not available_names:
            continue
        default_param = tracked_names[0] if tracked_names else available_names[0]
        sibling_lookup = {}
        sibling_options = []
        for sibling_ctx in children_contexts:
            if sibling_ctx is child_ctx:
                continue
            sibling_entry = sibling_ctx.get("entry") or {}
            sibling_led_id = sibling_ctx.get("led_id") or ""
            sibling_label = sibling_ctx.get("label") or "Sibling"
            sibling_all_values = dict(sibling_ctx.get("all_params") or {})
            sibling_param_names = sorted(
                set(_mapping_parameter_names(sibling_entry.get("parameters") or {})) | set(sibling_all_values.keys())
            )
            for sibling_param_name in sibling_param_names:
                display = "{} ({}) :: {}".format(
                    sibling_label,
                    sibling_led_id,
                    sibling_param_name,
                )
                sibling_options.append(display)
                sibling_lookup[display] = (sibling_led_id, sibling_param_name)
        rows.append({
            "row_id": "row-{:05d}".format(row_index),
            "child_label": child_ctx.get("label") or "Child",
            "child_led_id": child_ctx.get("led_id") or "",
            "available_params": available_names,
            "selected_param": default_param,
            "current_values": all_values,
            "parent_param_names": list(parent_param_names or []),
            "sibling_options": sibling_options,
            "sibling_lookup": sibling_lookup,
            "params": params,
        })
        row_index += 1
    return rows


def _apply_mapping_results(mapping_rows, mapping_results):
    row_by_child = {}
    for row in mapping_rows or []:
        child_led_id = row.get("child_led_id")
        if not child_led_id:
            continue
        row_by_child[child_led_id] = row
    decisions_by_child = {}
    for decision in mapping_results or []:
        child_led_id = decision.get("child_led_id")
        if not child_led_id or child_led_id not in row_by_child:
            continue
        decisions_by_child.setdefault(child_led_id, []).append(decision)

    for child_led_id, row in row_by_child.items():
        params = row.get("params") or {}
        current_values = row.get("current_values") or {}

        # Treat selected rows as the full tracked set; preserve special metadata parameters.
        preserved = {}
        for key in list(params.keys()):
            if key in ELEMENT_LINKER_PARAM_NAMES:
                preserved[key] = params.get(key)
        params.clear()
        params.update(preserved)

        for decision in decisions_by_child.get(child_led_id, []):
            mode = decision.get("mode") or FLAG_STATIC
            target = decision.get("target") or ""
            param_name = decision.get("param_name") or ""
            if not param_name:
                continue
            if mode == FLAG_BYPARENT:
                params[param_name] = 'parent_parameter: "{}"'.format(_escape_quotes(target))
                continue
            if mode == FLAG_BYSIBLING:
                sibling_led_id, sibling_param = (row.get("sibling_lookup") or {}).get(target, (None, None))
                if sibling_led_id and sibling_param:
                    params[param_name] = 'sibling_parameter: {}: "{}"'.format(
                        sibling_led_id,
                        _escape_quotes(sibling_param),
                    )
                continue
            params[param_name] = current_values.get(param_name, "")


def _prompt_child_parameter_flags(children_contexts, parent_param_names):
    mapping_rows = _build_mapping_rows(children_contexts, parent_param_names)
    if not mapping_rows:
        return True
    ui_module = _load_mapping_ui_module()
    if ui_module is None:
        return False
    xaml_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "EstablishRelationshipMappingWindow.xaml")
    )
    if not os.path.exists(xaml_path):
        forms.alert("Mapping UI XAML not found:\n{}".format(xaml_path), title=TITLE)
        return False
    try:
        window = ui_module.RelationshipMappingWindow(xaml_path, mapping_rows, TITLE)
    except Exception as exc:
        forms.alert("Failed to create mapping UI:\n\n{}".format(exc), title=TITLE)
        return False
    try:
        applied = bool(window.ShowDialog())
    except Exception as exc:
        forms.alert("Failed to display mapping UI:\n\n{}".format(exc), title=TITLE)
        return False
    if not applied:
        return False
    _apply_mapping_results(mapping_rows, window.results or [])
    return True


def _create_parent_led_entry(
    parent_elem,
    type_set,
    eq_def,
    parent_point,
    parent_rotation,
):
    label, is_group = _build_label_info(parent_elem)
    params = dict(_collect_profile_params(parent_elem) or {})
    led_id = next_led_id(type_set, eq_def)
    set_id = type_set.get("id") or ""
    payload = _build_element_linker_payload(
        led_id,
        set_id,
        parent_elem,
        parent_point,
        rotation_override=parent_rotation,
        parent_rotation_deg=None,
        parent_elem_id=None,
    )
    params[ELEMENT_LINKER_PARAM_NAME] = payload
    entry = {
        "id": led_id,
        "label": label,
        "category": _get_category_name(parent_elem),
        "is_group": bool(is_group),
        "offsets": [{
            "x_inches": 0.0,
            "y_inches": 0.0,
            "z_inches": 0.0,
            "rotation_deg": 0.0,
        }],
        "parameters": params,
        "tags": [],
    }
    return entry, payload


def _create_child_led_entry(
    child_elem,
    type_set,
    eq_def,
    parent_point,
    parent_rotation,
    parent_elem_id,
):
    point = _get_point(child_elem)
    if point is None:
        return None, None
    label, is_group = _build_label_info(child_elem)
    if not label:
        return None, None
    rotation = _get_rotation(child_elem)
    offsets = compute_offsets_from_points(parent_point, parent_rotation, point, rotation)
    offsets["z_inches"] = _level_relative_z_inches(child_elem, point)
    if is_group:
        offsets["rotation_deg"] = _normalize_angle((rotation or 0.0) - (parent_rotation or 0.0))

    params = dict(_collect_profile_params(child_elem) or {})
    led_id = next_led_id(type_set, eq_def)
    set_id = type_set.get("id") or ""
    payload = _build_element_linker_payload(
        led_id,
        set_id,
        child_elem,
        point,
        rotation_override=rotation,
        parent_rotation_deg=parent_rotation,
        parent_elem_id=parent_elem_id,
    )
    params[ELEMENT_LINKER_PARAM_NAME] = payload

    entry = {
        "id": led_id,
        "label": label,
        "category": _get_category_name(child_elem),
        "is_group": bool(is_group),
        "offsets": [offsets],
        "parameters": params,
        "tags": [],
    }
    return entry, payload


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    config_name = forms.ask_for_string(
        prompt="Name this parent/child equipment configuration:",
        default="",
    )
    if not config_name:
        return
    config_name = config_name.strip()
    if not config_name:
        forms.alert("Configuration name cannot be blank.", title=TITLE)
        return

    try:
        data_path, data = _load_active_yaml_data_resilient(doc)
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)

    existing_def = _find_equipment_definition_by_name(data, config_name)
    is_new_definition = existing_def is None

    forms.alert("Select one parent equipment element.", title=TITLE)
    try:
        parent_elem = revit.pick_element(message="Select parent equipment")
    except Exception:
        parent_elem = None
    if parent_elem is None:
        return

    parent_point = _get_point(parent_elem)
    if parent_point is None:
        forms.alert("Could not determine the parent element location.", title=TITLE)
        return

    parent_rotation = _get_rotation(parent_elem)
    parent_param_names = _collect_all_parameter_names(parent_elem)

    forms.alert("Select one or more child equipment elements.", title=TITLE)
    try:
        children_raw = list(revit.pick_elements(message="Select child equipment elements"))
    except Exception:
        children_raw = []
    if not children_raw:
        forms.alert("No child equipment was selected.", title=TITLE)
        return

    children = _dedupe_elements(children_raw)
    filtered_children = []
    try:
        parent_id = _element_id_value(parent_elem.Id)
    except Exception:
        parent_id = None
    for elem in children:
        try:
            elem_id = _element_id_value(elem.Id)
        except Exception:
            elem_id = None
        if parent_id is not None and elem_id == parent_id:
            continue
        filtered_children.append(elem)
    if not filtered_children:
        forms.alert("All selected children matched the parent; nothing to process.", title=TITLE)
        return

    trans_group = TransactionGroup(doc, TITLE)
    trans_group.Start()
    success = False

    try:
        if existing_def is not None:
            overwrite = forms.alert(
                "A profile named '{}' already exists in {}.\n"
                "Select Yes to replace its linked elements with this new relationship setup."
                .format(config_name, yaml_label),
                title=TITLE,
                ok=False,
                yes=True,
                no=True,
            )
            if not overwrite:
                return
            eq_def = existing_def
        else:
            parent_label, parent_is_group = _build_label_info(parent_elem)
            sample_entry = {
                "label": parent_label,
                "category_name": _get_category_name(parent_elem),
                "is_group": parent_is_group,
            }
            next_eq_num = _next_eq_number(data)
            next_set_num = _next_set_number(data)
            eq_def = ensure_equipment_definition(data, config_name, sample_entry)
            eq_def["id"] = "EQ-{:03d}".format(next_eq_num)
            type_set = get_type_set(eq_def)
            type_set["id"] = "SET-{:03d}".format(next_set_num)

        type_set = get_type_set(eq_def)
        if not type_set.get("id"):
            type_set["id"] = "SET-{:03d}".format(_next_set_number(data))
        type_set["name"] = "{} Types".format(config_name)
        type_set["linked_element_definitions"] = []
        eq_def["linked_sets"] = [type_set]

        metadata_updates = []

        parent_led, parent_payload = _create_parent_led_entry(
            parent_elem,
            type_set,
            eq_def,
            parent_point,
            parent_rotation,
        )
        type_set["linked_element_definitions"].append(parent_led)
        metadata_updates.append((parent_elem, parent_payload))

        children_contexts = []
        for child_elem in filtered_children:
            child_led, child_payload = _create_child_led_entry(
                child_elem,
                type_set,
                eq_def,
                parent_point,
                parent_rotation,
                parent_id,
            )
            if child_led is None:
                continue
            type_set["linked_element_definitions"].append(child_led)
            metadata_updates.append((child_elem, child_payload))
            children_contexts.append({
                "entry": child_led,
                "label": child_led.get("label") or "Child",
                "led_id": child_led.get("id") or "",
                "all_params": _collect_all_parameter_values(child_elem),
            })

        if not children_contexts:
            forms.alert("No valid child entries were captured.", title=TITLE)
            return

        mappings_applied = _prompt_child_parameter_flags(children_contexts, parent_param_names)
        if not mappings_applied:
            return

        if metadata_updates:
            txn_name = "Establish Relationship: Store Element Linker metadata ({})".format(len(metadata_updates))
            txn = Transaction(doc, txn_name)
            try:
                txn.Start()
                for element, payload in metadata_updates:
                    _set_element_linker_parameter(element, payload)
                txn.Commit()
            except Exception:
                try:
                    txn.RollBack()
                except Exception:
                    pass

        _save_active_yaml_data_resilient(
            doc,
            data,
            TITLE,
            "{} relationship profile '{}' with {} child element(s)".format(
                "Created" if is_new_definition else "Updated",
                config_name,
                len(children_contexts),
            ),
        )

        forms.alert(
            "Profile '{}':\n"
            "- Parent: 1\n"
            "- Children: {}\n"
            "- Total linked items: {}\n\n"
            "Element_Linker metadata was written to all selected elements."
            .format(
                config_name,
                len(children_contexts),
                len(type_set.get("linked_element_definitions") or []),
            ),
            title=TITLE,
        )
        success = True
    except Exception as exc:
        forms.alert("Failed to establish relationship profile in {}:\n\n{}".format(yaml_label, exc), title=TITLE)
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
