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

import io
import json
import os
import datetime

from pyrevit import revit, forms

# Add CEDLib.lib to sys.path for shared assets
import sys
LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from profile_schema import (  # noqa: E402
    ensure_equipment_definition,
    get_type_set,
    load_data as load_profile_data,
    next_led_id,
    save_data as save_profile_data,
)
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")

try:
    basestring
except NameError:
    basestring = str

from Autodesk.Revit.DB import Group, GroupType, XYZ, BuiltInParameter, IndependentTag  # noqa: E402

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


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #


def _log_entry(entry, log_path):
    try:
        with io.open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=True) + "\n")
    except Exception:
        pass


def _feet_to_inches(value):
    try:
        return float(value) * 12.0
    except Exception:
        return 0.0


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
                found[target_key] = p.AsDouble()
            elif st == "Integer":
                found[target_key] = p.AsInteger()
            else:
                found[target_key] = p.AsValueString() or ""
        except Exception:
            continue
    if "dev-Group ID" not in found:
        found["dev-Group ID"] = ""
    if is_electrical:
        return found
    if any(value for key, value in found.items() if key != "dev-Group ID" and value):
        return found
    return {"dev-Group ID": found.get("dev-Group ID", "")}


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
                    try:
                        sparam = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if sparam:
                            type_name = sparam.AsString()
                    except Exception:
                        pass
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
            offsets["z_inches"] = _feet_to_inches(delta.Z)
        tags.append({
            "family_name": fam_name,
            "type_name": type_name,
            "category_name": category_name,
            "parameters": {},
            "offsets": offsets,
        })
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

    type_entry = {
        "label": label,
        "is_group": is_group,
        "instance_config": {
            "offsets": [{
                "x_inches": _feet_to_inches(offset_vec.X),
                "y_inches": _feet_to_inches(offset_vec.Y),
                "z_inches": _feet_to_inches(offset_vec.Z),
                "rotation_deg": rot_deg,
            }],
            "parameters": params,
            "tags": tags,
        },
        "category_name": cat_name or "",
    }
    return type_entry


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #


def main():
    data_path = _pick_profile_data_path()
    if not data_path:
        return
    log_path = os.path.join(os.path.dirname(data_path), "profileData.log")

    data = _load_profile_store(data_path)
    existing_names = sorted({
        (entry.get("name") or entry.get("id") or "").strip()
        for entry in data.get("equipment_definitions") or []
        if (entry.get("name") or entry.get("id") or "").strip()
    })

    NEW_DEF_OPTION = "<< New equipment definition >>"
    cad_options = [NEW_DEF_OPTION] + existing_names
    cad_choice = forms.SelectFromList.show(
        cad_options,
        title="Select equipment definition (or choose new)",
        multiselect=False,
        button_name="Select"
    )
    if not cad_choice:
        return
    cad_choice = cad_choice if isinstance(cad_choice, basestring) else cad_choice[0]
    if cad_choice == NEW_DEF_OPTION:
        cad_name = forms.ask_for_string(
            prompt="Enter a name for the new equipment definition:",
            default=""
        )
        if not cad_name:
            return
        cad_name = cad_name.strip()
        if not cad_name:
            return
    else:
        cad_name = cad_choice

    # 2) Pick elements
    try:
        elems = revit.pick_elements(message="Select Revit element(s) to create YAML profile type(s)")
    except Exception:
        # fallback single
        e = revit.pick_element(message="Select Revit element to create YAML profile type")
        elems = [e] if e else []
    if not elems:
        return

    # Capture zero-offset entries for each selected element
    element_locations = []
    for e in elems:
        loc = _get_point(e)
        if loc is not None:
            element_locations.append((e, loc))
    if not element_locations:
        forms.alert("Could not read locations from selected elements.", title="Add YAML Profiles")
        return

    # Compute centroid to preserve spacing between elements
    sum_x = sum(loc.X for _, loc in element_locations)
    sum_y = sum(loc.Y for _, loc in element_locations)
    sum_z = sum(loc.Z for _, loc in element_locations)
    count = float(len(element_locations))
    centroid = XYZ(sum_x / count, sum_y / count, sum_z / count)

    type_entries = []
    for elem, loc in element_locations:
        rel_vec = loc - centroid
        type_entries.append(_build_type_entry(elem, rel_vec, 0.0, loc))

    if not type_entries:
        forms.alert("No valid elements were selected.", title="Add YAML Profiles")
        return

    data = _load_profile_store(data_path)
    equipment_def = ensure_equipment_definition(data, cad_name, type_entries[0])
    type_set = get_type_set(equipment_def)
    led_list = type_set.setdefault("linked_element_definitions", [])

    for entry in type_entries:
        lbl = (entry.get("label") or "").strip()
        if not lbl:
            continue
        entry_inst = entry.get("instance_config") or {}
        entry_offsets = entry_inst.get("offsets") or [{}]
        entry_params = entry_inst.get("parameters") or {}
        entry_tags = entry_inst.get("tags") or []

        led_list.append({
            "id": next_led_id(type_set, equipment_def),
            "label": lbl,
            "category": entry.get("category_name"),
            "is_group": bool(entry.get("is_group")),
            "offsets": entry_offsets,
            "parameters": entry_params,
            "tags": entry_tags,
        })

    try:
        save_profile_data(data_path, data)
        _log_entry({
            "timestamp": datetime.datetime.utcnow().isoformat() + "Z",
            "action": "add",
            "cad_name": cad_name,
            "type_labels": [t.get("label") for t in type_entries],
            "user": os.getenv("USERNAME") or os.getenv("USER") or "unknown",
        }, log_path)
        forms.alert("Added {} type(s) under CAD profile '{}'.\nReload Place Elements (YAML) to use them.".format(len(type_entries), cad_name), title="Add YAML Profiles")
    except Exception as ex:
        forms.alert("Failed to save profileData.yaml:\n\n{}".format(ex), title="Add YAML Profiles")


if __name__ == "__main__":
    main()
