# -*- coding: utf-8 -*-
"""
Add YAML Profiles
-----------------
Helper to append new type entries to CEDLib.lib/profileData.yaml (equipment_definitions
format) used by Place Elements (YAML).

Flow:
1) User selects the target profileData YAML and an existing/new equipment definition.
2) User selects one or more Revit elements (families or model groups).
3) Captures: label (Family : Type), category, is_group flag, zero offsets, and the
   electrical CKT_* parameters (CKT_Rating_CED, CKT_Panel_CEDT, CKT_Schedule Notes_CEDT,
   CKT_Circuit Number_CEDT, CKT_Load Name_CEDT).
"""

import copy
import io
import math
import os

from pyrevit import revit, forms, script

# Add CEDLib.lib to sys.path for shared assets
import sys
LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.profile_schema import (  # noqa: E402
    ensure_equipment_definition,
    get_type_set,
    next_led_id,
)
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

try:
    basestring
except NameError:
    basestring = str

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,

    Group,
    GroupType,
    IndependentTag,
    Transaction,
    TransactionGroup,
    UnitUtils,
    XYZ,
)

ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_SHARED_PARAM = "Element_Linker"
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
SAFE_HASH = u"\uff03"



# --------------------------------------------------------------------------- #
# YAML helpers
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


def _collect_params(elem):
    def _convert_collected_double(target_key, param_obj, raw_value):
        if target_key not in ("Apparent Load Input_CED", "Voltage_CED"):
            return raw_value
        get_unit = getattr(param_obj, "GetUnitTypeId", None)
        if not callable(get_unit):
            return raw_value
        try:
            unit_id = get_unit()
        except Exception:
            unit_id = None
        if not unit_id:
            return raw_value
        try:
            return UnitUtils.ConvertFromInternalUnits(raw_value, unit_id)
        except Exception:
            return raw_value

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
    for p in getattr(elem, "Parameters", []) or []:
        try:
            name = p.Definition.Name
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
            st = p.StorageType.ToString()
        except Exception:
            st = ""
        try:
            if st == "String":
                found[target_key] = p.AsString() or ""
            elif st == "Double":
                found[target_key] = _convert_collected_double(target_key, p, p.AsDouble())
            elif st == "Integer":
                found[target_key] = p.AsInteger()
            else:
                found[target_key] = p.AsValueString() or ""
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


def _get_rotation_degrees(elem):
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            pass
    facing = getattr(elem, "FacingOrientation", None)
    if facing:
        try:
            return math.degrees(math.atan2(facing.Y, facing.X))
        except Exception:
            pass
    return 0.0


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
    rotation_deg = _get_rotation_degrees(elem)
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
    param_names = (ELEMENT_LINKER_SHARED_PARAM, ELEMENT_LINKER_PARAM_NAME)
    for name in param_names:
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


def _annotation_family_type(elem):
    fam_name = None
    type_name = None
    try:
        symbol = getattr(elem, "Symbol", None)
        if symbol:
            fam = getattr(symbol, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else None
            type_name = getattr(symbol, "Name", None)
            if not type_name and hasattr(symbol, "get_Parameter"):
                param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    type_name = param.AsString()
    except Exception:
        fam_name = fam_name
    if fam_name and type_name:
        return fam_name, type_name
    name = getattr(elem, "Name", None)
    return name or "", ""


def _collect_annotation_string_params(annotation_elem):
    results = {}
    for param in getattr(annotation_elem, "Parameters", []) or []:
        try:
            definition = getattr(param, "Definition", None)
            name = getattr(definition, "Name", None)
        except Exception:
            name = None
        if not name:
            continue
        try:
            storage = param.StorageType
            is_string = storage and storage.ToString() == "String"
        except Exception:
            is_string = False
        if not is_string:
            continue
        try:
            if param.IsReadOnly:
                continue
        except Exception:
            pass
        try:
            value = param.AsString()
        except Exception:
            value = None
        if value:
            safe_name = (name or "").replace("#", SAFE_HASH)
            results[safe_name] = value
    return results


def _build_annotation_tag_entry(annotation_elem, host_point):
    try:
        cat = getattr(annotation_elem, "Category", None)
        cat_name = getattr(cat, "Name", "") if cat else ""
    except Exception:
        cat_name = ""
    if not cat_name or "generic annotation" not in cat_name.lower():
        return None
    ann_point = _get_point(annotation_elem)
    if ann_point is None or host_point is None:
        return None
    fam_name, type_name = _annotation_family_type(annotation_elem)
    offsets = {
        "x_inches": _feet_to_inches(ann_point.X - host_point.X),
        "y_inches": _feet_to_inches(ann_point.Y - host_point.Y),
        "z_inches": _feet_to_inches(ann_point.Z - host_point.Z),
        "rotation_deg": _get_rotation_degrees(annotation_elem),
    }
    params = _collect_annotation_string_params(annotation_elem)
    return {
        "family_name": fam_name,
        "type_name": type_name,
        "category_name": cat_name,
        "parameters": params,
        "offsets": offsets,
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
        if not tag:
            continue
        if isinstance(tag, IndependentTag):
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
                "family_name": fam_name or "",
                "type_name": type_name or "",
                "category_name": category_name,
                "parameters": {},
                "offsets": offsets,
            })
            continue
        annotation_entry = _build_annotation_tag_entry(tag, host_point)
        if annotation_entry:
            tags.append(annotation_entry)
    return tags


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

def _build_type_entry(elem, offset_vec, rot_deg, host_point):
    fam_name = None
    type_name = None
    is_group = False
    cat_name = None
    try:
        cat = elem.Category
        if cat:
            cat_name = cat.Name
    except Exception:
        pass

    if isinstance(elem, Group) or isinstance(elem, GroupType):
        is_group = True
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
                if not type_name:
                    try:
                        tparam = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if tparam:
                            type_name = tparam.AsString()
                    except Exception:
                        pass
        except Exception:
            pass

    label = None
    if fam_name and type_name:
        label = u"{} : {}".format(fam_name, type_name)
    elif type_name:
        label = type_name
    elif fam_name:
        label = fam_name
    else:
        label = "Unnamed"

    params = _collect_params(elem)
    tags = _collect_hosted_tags(elem, host_point)

    offsets = {
        "x_inches": _feet_to_inches(offset_vec.X),
        "y_inches": _feet_to_inches(offset_vec.Y),
        "rotation_deg": rot_deg,
        "z_inches": _level_relative_z_inches(elem, host_point),
    }

    type_entry = {
        "label": label,
        "is_group": is_group,
        "instance_config": {
            "offsets": [offsets],
            "parameters": params,
            "tags": tags,
        },
        "category_name": cat_name or "",
    }
    return type_entry


def _next_eq_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        eq_id = (eq.get("id") or "").strip()
        if eq_id.startswith("EQ-"):
            try:
                num = int(eq_id.split("-")[-1])
                if num > max_id:
                    max_id = num
            except Exception:
                continue
    return max_id + 1


def _append_type_entries(equipment_def, template_entries):
    type_set = get_type_set(equipment_def)
    led_list = type_set.setdefault("linked_element_definitions", [])
    for entry in template_entries or []:
        cloned = copy.deepcopy(entry)
        inst_cfg = cloned.get("instance_config") or {}
        params = inst_cfg.get("parameters") or {}
        if isinstance(params, dict):
            params = copy.deepcopy(params)
            params.pop(ELEMENT_LINKER_PARAM_NAME, None)
            params.pop(ELEMENT_LINKER_SHARED_PARAM, None)
        else:
            params = {}
        offsets = inst_cfg.get("offsets") or [{}]
        tags = inst_cfg.get("tags") or []
        offsets = copy.deepcopy(offsets)
        tags = copy.deepcopy(tags)
        led_id = next_led_id(type_set, equipment_def)
        led_list.append({
            "id": led_id,
            "label": cloned.get("label"),
            "category": cloned.get("category_name"),
            "is_group": bool(cloned.get("is_group")),
            "offsets": offsets,
            "parameters": params,
            "tags": tags,
        })


def _truth_group_metadata(equipment_defs):
    groups = {}
    child_to_group = {}
    for entry in equipment_defs or []:
        eq_name = (entry.get("name") or entry.get("id") or "").strip()
        if not eq_name:
            continue
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id:
            source_id = (entry.get("id") or eq_name).strip()
        display = (entry.get(TRUTH_SOURCE_NAME_KEY) or "").strip()
        if not display:
            display = eq_name
        group = groups.setdefault(source_id, {
            "display": display,
            "members": [],
            "primary": None,
        })
        group["members"].append(eq_name)
        eq_id = (entry.get("id") or "").strip()
        if eq_id and eq_id == source_id:
            group["primary"] = eq_name
        child_to_group[eq_name] = source_id
    label_counts = {}
    label_map = {}
    for source_id, data in groups.items():
        members = data.get("members") or []
        primary = data.get("primary") or (members[0] if members else None)
        if not primary:
            continue
        display = data.get("display") or primary
        count = len(members) if members else 1
        base_label = display
        if count > 1:
            base_label = u"{} ({} profiles)".format(display, count)
        label_counts[base_label] = label_counts.get(base_label, 0) + 1
        label_map[source_id] = {
            "base_label": base_label,
            "primary": primary,
        }
    resolved_labels = {}
    for source_id, info in label_map.items():
        base_label = info["base_label"]
        label = base_label
        if label_counts.get(base_label, 0) > 1:
            label = u"{} [{}]".format(base_label, source_id)
        resolved_labels[label] = (info["primary"], source_id)
    return resolved_labels, groups, child_to_group


def _execute_profile_addition(doc, data, yaml_label):
    equipment_defs = data.get("equipment_definitions") or []
    logger = script.get_logger()
    logger.info(
        "[Add YAML] existing defs before prompt: %s",
        [eq.get("name") or eq.get("id") for eq in equipment_defs if isinstance(eq, dict)],
    )
    label_map, truth_groups, child_to_group = _truth_group_metadata(equipment_defs)
    existing_names = sorted({
        (entry.get("name") or entry.get("id") or "").strip()
        for entry in equipment_defs
        if (entry.get("name") or entry.get("id") or "").strip()
    })

    NEW_DEF_OPTION = "<< New equipment definition >>"
    cad_options = [NEW_DEF_OPTION]
    label_to_profile = {}
    label_to_group_id = {}
    if label_map:
        for label in sorted(label_map.keys()):
            cad_options.append(label)
            primary, group_id = label_map[label]
            label_to_profile[label] = primary
            label_to_group_id[label] = group_id
    else:
        cad_options.extend(existing_names)
    cad_choice = forms.SelectFromList.show(
        cad_options,
        title="Select equipment definition (or choose new)",
        multiselect=False,
        button_name="Select",
    )
    if not cad_choice:
        return False
    cad_choice = cad_choice if isinstance(cad_choice, basestring) else cad_choice[0]
    created_new_def = False
    selected_group_id = None
    if cad_choice == NEW_DEF_OPTION:
        cad_name = forms.ask_for_string(
            prompt="Enter a name for the new equipment definition:",
            default="",
        )
        if not cad_name:
            return False
        cad_name = cad_name.strip()
        if not cad_name:
            return False
        existing_match = None
        cad_lower = cad_name.lower()
        for eq in equipment_defs:
            existing_name = (eq.get("name") or eq.get("id") or "").strip()
            if existing_name and existing_name.lower() == cad_lower:
                existing_match = eq
                break
        if existing_match:
            append_to_existing = forms.alert(
                "An equipment definition named '{}' already exists.\n"
                "Selecting Yes will append the selected element type(s) to that definition.\n"
                "Select No to choose a different name.".format(cad_name),
                title="Add YAML Profiles",
                ok=False,
                yes=True,
                no=True,
            )
            if not append_to_existing:
                forms.alert("Add YAML Profiles canceled.", title="Add YAML Profiles")
                return False
            created_new_def = False
            cad_choice = cad_name
            selected_group_id = None
        else:
            created_new_def = True
    elif cad_choice in label_to_profile:
        cad_name = label_to_profile[cad_choice]
        selected_group_id = label_to_group_id.get(cad_choice)
    else:
        cad_name = cad_choice

    try:
        elems = revit.pick_elements(message="Select Revit element(s) to create YAML profile type(s)")
    except Exception:
        e = revit.pick_element(message="Select Revit element to create YAML profile type")
        elems = [e] if e else []
    if not elems:
        return False

    element_locations = []
    for e in elems:
        loc = _get_point(e)
        if loc is not None:
            element_locations.append((e, loc))
    if not element_locations:
        forms.alert("Could not read locations from selected elements.", title="Add YAML Profiles")
        return False

    sum_x = sum(loc.X for _, loc in element_locations)
    sum_y = sum(loc.Y for _, loc in element_locations)
    sum_z = sum(loc.Z for _, loc in element_locations)
    count = float(len(element_locations))
    centroid = XYZ(sum_x / count, sum_y / count, sum_z / count)

    element_records = []
    for elem, loc in element_locations:
        rel_vec = loc - centroid
        type_entry = _build_type_entry(elem, rel_vec, 0.0, loc)
        element_records.append({
            "element": elem,
            "host_point": loc,
            "type_entry": type_entry,
        })

    if not element_records:
        forms.alert("No valid elements were selected.", title="Add YAML Profiles")
        return False
    type_entries = [rec["type_entry"] for rec in element_records]

    equipment_def = ensure_equipment_definition(data, cad_name, type_entries[0])
    if created_new_def:
        next_idx = _next_eq_number(data)
        eq_id = "EQ-{:03d}".format(next_idx)
        set_id = "SET-{:03d}".format(next_idx)
        equipment_def["id"] = eq_id
        for linked_set in equipment_def.get("linked_sets") or []:
            linked_set["id"] = set_id
            linked_set["name"] = "{} Types".format(cad_name)

    type_set = get_type_set(equipment_def)
    eq_id = equipment_def.get("id")
    set_id = type_set.get("id")
    if created_new_def and not selected_group_id:
        selected_group_id = eq_id or cad_name

    metadata_updates = []
    led_list = type_set.setdefault("linked_element_definitions", [])
    for rec in element_records:
        entry = rec["type_entry"]
        lbl = (entry.get("label") or "").strip()
        if not lbl:
            continue
        inst_cfg = entry.setdefault("instance_config", {})
        entry_offsets = inst_cfg.get("offsets") or [{}]
        entry_params = inst_cfg.setdefault("parameters", {})
        entry_tags = inst_cfg.get("tags") or []

        led_id = next_led_id(type_set, equipment_def)
        payload = _build_element_linker_payload(led_id, set_id, rec["element"], rec["host_point"])
        entry_params[ELEMENT_LINKER_PARAM_NAME] = payload
        metadata_updates.append((rec["element"], payload))

        led_list.append({
            "id": led_id,
            "label": lbl,
            "category": entry.get("category_name"),
            "is_group": bool(entry.get("is_group")),
            "offsets": entry_offsets,
            "parameters": entry_params,
            "tags": entry_tags,
        })

    if metadata_updates and doc:
        txn_name = "Add YAML Profiles: Store Element Linker metadata ({})".format(len(metadata_updates))
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

    target_group_id = selected_group_id or child_to_group.get(cad_name) or (equipment_def.get("id") or cad_name)
    if target_group_id:
        group_info = truth_groups.get(target_group_id) or {}
        display_name = group_info.get("display") or cad_name
        equipment_def[TRUTH_SOURCE_ID_KEY] = target_group_id
        equipment_def[TRUTH_SOURCE_NAME_KEY] = display_name
        members = group_info.get("members") or []
        for member in members:
            if member == cad_name:
                continue
            target_def = ensure_equipment_definition(data, member, type_entries[0])
            target_def[TRUTH_SOURCE_ID_KEY] = target_group_id
            target_def[TRUTH_SOURCE_NAME_KEY] = display_name
            _append_type_entries(target_def, [rec["type_entry"] for rec in element_records])

    logger.info(
        "[Add YAML] equipment definitions now: %s",
        [eq.get("name") or eq.get("id") for eq in data.get("equipment_definitions") or []],
    )
    try:
        save_active_yaml_data(
            None,
            data,
            "Add YAML Profiles",
            "Added {} type(s) to '{}'".format(len(type_entries), cad_name),
        )
        forms.alert(
            "Added {} type(s) under equipment definition '{}' in {}.\nReload Place Elements (YAML) to use them.".format(
                len(type_entries),
                cad_name,
                yaml_label,
            ),
            title="Add YAML Profiles",
        )
        return True
    except Exception as ex:
        forms.alert("Failed to update {}:\n\n{}".format(yaml_label, ex), title="Add YAML Profiles")
        return False


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #


def main():
    doc = getattr(revit, "doc", None)
    trans_group = TransactionGroup(doc, "Add YAML Profiles") if doc else None
    if trans_group:
        trans_group.Start()
    success = False
    try:
        try:
            yaml_path, data = load_active_yaml_data()
        except RuntimeError as exc:
            forms.alert(str(exc), title="Add YAML Profiles")
            return
        yaml_label = get_yaml_display_name(yaml_path)
        success = _execute_profile_addition(doc, data, yaml_label)
    finally:
        if trans_group:
            try:
                if success:
                    trans_group.Assimilate()
                else:
                    trans_group.RollBack()
            except Exception:
                pass


if __name__ == "__main__":
    main()
