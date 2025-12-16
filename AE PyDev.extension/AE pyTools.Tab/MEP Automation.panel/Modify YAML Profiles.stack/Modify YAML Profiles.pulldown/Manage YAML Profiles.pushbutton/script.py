# -*- coding: utf-8 -*-



"""



Manage YAML Profiles



------------------



Edits the active YAML payload stored in Extensible Storage so Place Elements



and other tools always reference the in-model definition.



Allows editing offsets, parameters, tags, category, and is_group for each equipment definition/type.



"""



import copy



import imp



import json



import math



import os



from pyrevit import script, forms, revit



from Autodesk.Revit.DB import (



    BuiltInParameter,



    Group,



    GroupType,



    IndependentTag,



    Transaction,



    TransactionGroup,



    XYZ,



)



# Add CEDLib.lib to sys.path for shared UI/logic classes



import sys



def _find_cedlib_root():



    current = os.path.abspath(os.path.dirname(__file__))



    while True:



        candidate = os.path.join(current, "CEDLib.lib")



        if os.path.isdir(candidate):



            return candidate



        parent = os.path.dirname(current)



        if parent == current:



            break



        current = parent



    raise RuntimeError("Unable to locate CEDLib.lib relative to {}".format(__file__))



try:



    LIB_ROOT = _find_cedlib_root()



except RuntimeError as exc:



    forms.alert(str(exc), title="Manage YAML Profiles")



    raise



if LIB_ROOT not in sys.path:



    sys.path.append(LIB_ROOT)



from UIClasses.ProfileEditorWindow import ProfileEditorWindow  # noqa: E402



from LogicClasses.profile_schema import (  # noqa: E402



    ensure_equipment_definition,



    equipment_defs_to_legacy,



    get_type_set,



    legacy_to_equipment_defs,



    next_led_id,



)



from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402



from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402



try:
    basestring
except NameError:
    basestring = str



TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"



TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"



ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"



ELEMENT_LINKER_SHARED_PARAM = "Element_Linker"



TITLE = "Manage YAML Profiles"











try:



    basestring



except NameError:



    basestring = str



# --------------------------------------------------------------------------- #



# Pending orphan persistence helpers



# --------------------------------------------------------------------------- #



def _pending_store_path():



    try:



        return script.get_appdata_file("manage_yaml_pending_orphans.json")



    except Exception:



        return os.path.join(os.path.expanduser("~"), "manage_yaml_pending_orphans.json")



def _doc_store_key(doc):



    path = getattr(doc, "PathName", "") or ""



    if path:



        return path.lower()



    title = getattr(doc, "Title", "") or ""



    return title.lower()



def _load_pending_orphans(doc):



    path = _pending_store_path()



    if not os.path.exists(path):



        return []



    try:



        with open(path, "r") as handle:



            data = json.load(handle) or {}



    except Exception:



        return []



    key = _doc_store_key(doc)



    entries = data.get(key) or []



    return [entry for entry in entries if isinstance(entry, basestring) and entry.strip()]



def _save_pending_orphans(doc, pending):



    path = _pending_store_path()



    data = {}



    if os.path.exists(path):



        try:



            with open(path, "r") as handle:



                data = json.load(handle) or {}



        except Exception:



            data = {}



    key = _doc_store_key(doc)



    filtered = [entry for entry in (pending or []) if isinstance(entry, basestring) and entry.strip()]



    if filtered:



        data[key] = filtered



    elif key in data:



        data.pop(key, None)



    directory = os.path.dirname(path)



    if directory and not os.path.exists(directory):



        try:



            os.makedirs(directory)



        except Exception:



            pass



    try:
        with open(path, "w") as handle:
            json.dump(data, handle)



    except Exception:



        pass











def _next_eq_number(data):



    max_id = 0



    for entry in data.get("equipment_definitions") or []:



        eq_id = (entry.get("id") or "").strip()



        if eq_id.startswith("EQ-"):



            try:



                num = int(eq_id.split("-")[-1])



                if num > max_id:



                    max_id = num



            except Exception:



                continue



    return max_id + 1











def _find_definition_by_name(data, cad_name):



    target = (cad_name or "").strip().lower()



    if not target:



        return None



    for entry in data.get("equipment_definitions") or []:



        name = (entry.get("name") or entry.get("id") or "").strip().lower()



        if name == target:



            return entry



    return None











def _unique_truth_id(data, equipment_def, cad_label):



    base = (equipment_def.get("id") or cad_label or "").strip()



    if not base:



        base = cad_label or "Profile"



    existing = set()



    for entry in data.get("equipment_definitions") or []:



        if entry is equipment_def:



            continue



        val = (entry.get(TRUTH_SOURCE_ID_KEY) or entry.get("id") or entry.get("name") or "").strip()



        if val:



            existing.add(val.lower())



    candidate = base



    counter = 1



    while candidate and candidate.lower() in existing:



        counter += 1



        candidate = "{}#{}".format(base, counter)



    if not candidate:



        candidate = "{}#{}".format((cad_label or "Profile").strip() or "Profile", counter)



    return candidate



# --------------------------------------------------------------------------- #



# YAML helpers



# --------------------------------------------------------------------------- #



# --------------------------------------------------------------------------- #



# Simple shims so the existing ProfileEditorWindow can work on stored YAML



# --------------------------------------------------------------------------- #



class OffsetShim(object):



    def __init__(self, dct=None):



        dct = dct or {}



        self.x_inches = float(dct.get("x_inches", 0.0) or 0.0)



        self.y_inches = float(dct.get("y_inches", 0.0) or 0.0)



        self.z_inches = float(dct.get("z_inches", 0.0) or 0.0)



        self.rotation_deg = float(dct.get("rotation_deg", 0.0) or 0.0)



class InstanceConfigShim(object):



    def __init__(self, dct=None):



        dct = dct or {}



        offs = dct.get("offsets") or [{}]



        self.offsets = [OffsetShim(o) for o in offs]



        self.parameters = dict(dct.get("parameters") or {})



        raw_tags = dct.get("tags") or []



        shim_tags = []



        for tg in raw_tags:



            if isinstance(tg, dict):



                shim_tags.append({



                    "category_name": tg.get("category_name"),



                    "family_name": tg.get("family_name"),



                    "type_name": tg.get("type_name"),



                    "parameters": tg.get("parameters") or {},



                    "offsets": tg.get("offsets") or {},



                })



            else:



                shim_tags.append(tg)



        self.tags = shim_tags



    def get_offset(self, idx):



        if not self.offsets:



            self.offsets = [OffsetShim()]



        try:



            return self.offsets[idx]



        except Exception:



            return self.offsets[0]



class TypeConfigShim(object):



    def __init__(self, dct=None):



        dct = dct or {}



        self.label = dct.get("label")



        self.led_id = dct.get("led_id")



        self.element_def_id = dct.get("id") or dct.get("led_id")



        self.category_name = dct.get("category_name")



        self.is_group = bool(dct.get("is_group", False))



        self.instance_config = InstanceConfigShim(dct.get("instance_config") or {})



class ProfileShim(object):



    def __init__(self, dct=None):



        dct = dct or {}



        self.cad_name = dct.get("cad_name")



        self._types = [TypeConfigShim(t) for t in (dct.get("types") or [])]



    def get_types(self):



        return list(self._types)



    def find_type_by_label(self, label):



        for t in self._types:



            if getattr(t, "label", None) == label:



                return t



        return None



def _dict_from_shims(profiles):



    out = {"profiles": []}



    for p in profiles.values():



        types = []



        for t in p.get_types():



            inst = t.instance_config



            offsets = []



            for off in getattr(inst, "offsets", []) or []:



                offsets.append({



                    "x_inches": off.x_inches,



                    "y_inches": off.y_inches,



                    "z_inches": off.z_inches,



                    "rotation_deg": off.rotation_deg,



                })



            params = getattr(inst, "parameters", {}) or {}



            types.append({



                "label": t.label,



                "id": getattr(t, "element_def_id", None) or getattr(t, "led_id", None),



                "led_id": getattr(t, "led_id", None),



                "category_name": t.category_name,



                "is_group": t.is_group,



                "instance_config": {



                    "offsets": offsets or [{}],



                    "parameters": params,



                    "tags": _serialize_tags(getattr(inst, "tags", []) or []),



                },



            })



        out["profiles"].append({



            "cad_name": p.cad_name,



            "types": types,



        })



    return out



# --------------------------------------------------------------------------- #



# Orphan capture helpers



# --------------------------------------------------------------------------- #



def _feet_to_inches(value):



    try:



        return float(value) * 12.0



    except Exception:



        return 0.0



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



def _collect_params(elem):



    try:



        cat = getattr(elem, "Category", None)



        cat_name = getattr(cat, "Name", "") if cat else ""



    except Exception:



        cat_name = ""



    cat_lower = (cat_name or "").lower()



    is_electrical = ("electrical" in cat_lower) or ("lighting" in cat_lower) or ("data" in cat_lower)



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



    found = {key: "" for key in targets.keys()}



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



def _collect_hosted_tags(elem, host_point):



    doc = getattr(elem, "Document", None)



    if doc is None or host_point is None:



        return []



    try:



        dep_ids = list(elem.GetDependentElements(None))



    except Exception:



        dep_ids = []



    tags = []



    for dep_id in dep_ids:



        try:



            tag = doc.GetElement(dep_id)



        except Exception:



            tag = None



        if not isinstance(tag, IndependentTag):



            continue



        try:



            tag_point = tag.TagHeadPosition



        except Exception:



            tag_point = None



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



                fam = getattr(tag_symbol, "Family", None)



                fam_name = getattr(fam, "Name", None) if fam else getattr(tag_symbol, "FamilyName", None)



            except Exception:



                fam_name = None



            try:



                type_name = getattr(tag_symbol, "Name", None)



                if not type_name and hasattr(tag_symbol, "get_Parameter"):



                    param = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)



                    if param:



                        type_name = param.AsString()



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



        if not fam_name and not type_name:



            continue



        offsets = {"x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0}



        if tag_point:



            delta = tag_point - host_point



            offsets["x_inches"] = _feet_to_inches(delta.X)



            offsets["y_inches"] = _feet_to_inches(delta.Y)



            offsets["z_inches"] = _feet_to_inches(delta.Z)



        tags.append({



            "family_name": fam_name or "",



            "type_name": type_name or "",



            "category_name": category_name,



            "parameters": {},



            "offsets": offsets,



        })



    return tags



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



        fallbacks = (



            BuiltInParameter.SCHEDULE_LEVEL_PARAM,



            BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,



            BuiltInParameter.FAMILY_LEVEL_PARAM,



            BuiltInParameter.INSTANCE_LEVEL_PARAM,



        )



        for bip in fallbacks:



            try:



                param = elem.get_Parameter(bip)



            except Exception:



                param = None



            if not param:



                continue



            try:



                eid = param.AsElementId()



                if eid and doc:



                    level_elem = doc.GetElement(eid)



            except Exception:



                level_elem = None



            if level_elem:



                break



    level_z = 0.0



    if level_elem:



        try:



            level_z = getattr(level_elem, "Elevation", 0.0) or 0.0



        except Exception:



            level_z = 0.0



    world_z = world_point.Z if world_point else 0.0



    return _feet_to_inches(world_z - level_z)



def _get_level_element_id(elem):



    try:



        level_id = getattr(elem, "LevelId", None)



        if level_id and getattr(level_id, "IntegerValue", 0) > 0:



            return level_id.IntegerValue



    except Exception:



        pass



    fallback_params = (



        BuiltInParameter.SCHEDULE_LEVEL_PARAM,



        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,



        BuiltInParameter.FAMILY_LEVEL_PARAM,



        BuiltInParameter.INSTANCE_LEVEL_PARAM,



    )



    for bip in fallback_params:



        try:



            param = elem.get_Parameter(bip)



        except Exception:



            param = None



        if not param:



            continue



        try:



            eid = param.AsElementId()



            if eid and getattr(eid, "IntegerValue", 0) > 0:



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



    for name in (ELEMENT_LINKER_SHARED_PARAM, ELEMENT_LINKER_PARAM_NAME):



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



def _build_type_entry(elem, offset_vec, rot_deg, host_point):



    fam_name = None



    type_name = None



    is_group = False



    category_name = None



    try:



        cat = elem.Category



        if cat:



            category_name = cat.Name



    except Exception:



        category_name = None



    if isinstance(elem, (Group, GroupType)):



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



                if not type_name and hasattr(sym, "get_Parameter"):



                    name_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)



                    if name_param:



                        type_name = name_param.AsString()



        except Exception:



            pass



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



    return {



        "label": label,



        "is_group": is_group,



        "instance_config": {



            "offsets": [offsets],



            "parameters": params,



            "tags": tags,



        },



        "category_name": category_name or "",



    }



def _capture_orphan_profile(doc, cad_name, state, refresh_callback, yaml_label):



    if doc is None:



        forms.alert("No active document detected.", title=TITLE)



        return False



    data = state.get("raw_data") or {}



    if not cad_name:



        forms.alert("Profile name cannot be empty.", title=TITLE)



        return False



    cad_label = cad_name.strip()



    existing_names = {



        (entry.get("name") or entry.get("id") or "").strip().lower()



        for entry in data.get("equipment_definitions") or []



    }



    if cad_label.lower() in existing_names:



        forms.alert("An equipment definition named '{}' already exists.".format(cad_label), title=TITLE)



        return False



    forms.alert(



        "Select the Revit elements you want to store under '{}'.\nClick Finish when you've picked them."



        .format(cad_label),



        title=TITLE,



    )



    try:



        elems = revit.pick_elements(message="Select element(s) for '{}'".format(cad_label))



    except Exception:



        try:



            elem = revit.pick_element(message="Select element for '{}'".format(cad_label))



        except Exception:



            elem = None



        elems = [elem] if elem else []



    if not elems:



        forms.alert("No elements were selected for '{}'.".format(cad_label), title=TITLE)



        return False



    element_locations = []



    for elem in elems:



        loc = _get_point(elem)



        if loc is not None:



            element_locations.append((elem, loc))



    if not element_locations:



        forms.alert("Could not read locations from the selected elements.", title=TITLE)



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



        forms.alert("No valid elements were found.", title=TITLE)



        return False



    existing_def = _find_definition_by_name(data, cad_label)



    equipment_def = ensure_equipment_definition(data, cad_label, element_records[0]["type_entry"])



    type_set = get_type_set(equipment_def)



    if existing_def is None:



        next_idx = _next_eq_number(data)



        eq_id = "EQ-{:03d}".format(next_idx)



        set_id = "SET-{:03d}".format(next_idx)



        equipment_def["id"] = eq_id



        type_set["id"] = set_id



        type_set["name"] = "{} Types".format(cad_label)



    else:



        set_id = type_set.get("id")



        eq_id = equipment_def.get("id") or cad_label



    led_list = type_set.setdefault("linked_element_definitions", [])



    metadata_updates = []



    for record in element_records:



        entry = record["type_entry"]



        inst_cfg = entry.setdefault("instance_config", {})



        entry_offsets = inst_cfg.get("offsets") or [{}]



        entry_params = inst_cfg.setdefault("parameters", {})



        entry_tags = inst_cfg.get("tags") or []



        led_id = next_led_id(type_set, equipment_def)



        payload = _build_element_linker_payload(led_id, set_id, record["element"], record["host_point"])



        entry_params[ELEMENT_LINKER_PARAM_NAME] = payload



        metadata_updates.append((record["element"], payload))



        led_list.append({



            "id": led_id,



            "label": entry.get("label"),



            "category": entry.get("category_name"),



            "is_group": bool(entry.get("is_group")),



            "offsets": entry_offsets,



            "parameters": entry_params,



            "tags": entry_tags,



        })



    truth_id = _unique_truth_id(data, equipment_def, cad_label)



    equipment_def[TRUTH_SOURCE_ID_KEY] = truth_id



    equipment_def[TRUTH_SOURCE_NAME_KEY] = cad_label



    if metadata_updates and doc:



        txn = Transaction(doc, "Create Orphan Profile ({})".format(cad_label))



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



    try:



        save_active_yaml_data(



            None,



            state["raw_data"],



            TITLE,



            "Captured {} orphan element(s) for '{}'".format(len(element_records), cad_label),



        )



    except Exception as exc:



        forms.alert("Failed to save {}:\n\n{}".format(yaml_label, exc), title=TITLE)



        return False



    refresh_callback()



    forms.alert(



        "Saved {} orphan element(s) to '{}'.\nReload placement tools to use the updates.".format(



            len(element_records),



            cad_label,



        ),



        title=TITLE,



    )



    return True



def _has_negative_z(value):



    if value is None:



        return False



    try:



        return float(value) < 0.0



    except Exception:



        return False



def _find_negative_z_offsets(profile_dict):



    negatives = []



    for profile in profile_dict.get("profiles") or []:



        cad_name = profile.get("cad_name") or "<Unnamed CAD>"



        for type_entry in profile.get("types") or []:



            label = type_entry.get("label") or type_entry.get("id") or "<Unnamed Type>"



            inst = type_entry.get("instance_config") or {}



            for idx, offset in enumerate(inst.get("offsets") or []):



                if _has_negative_z(offset.get("z_inches")):



                    negatives.append({



                        "cad": cad_name,



                        "label": label,



                        "index": idx + 1,



                        "value": float(offset.get("z_inches") or 0.0),



                        "source": "offset",



                    })



            for tag in inst.get("tags") or []:



                offsets = tag.get("offsets") or {}



                if _has_negative_z(offsets.get("z_inches")):



                    negatives.append({



                        "cad": cad_name,



                        "label": label,



                        "index": None,



                        "value": float(offsets.get("z_inches") or 0.0),



                        "source": "tag",



                    })



    return negatives



def _shims_from_dict(data):



    profiles = {}



    for p in data.get("profiles") or []:



        cad = p.get("cad_name")



        if not cad:



            continue



        profiles[cad] = ProfileShim(p)



    return profiles



def _build_relations_index(equipment_defs):



    id_to_name = {}



    for entry in equipment_defs or []:



        eq_id = (entry.get("id") or "").strip()



        eq_name = (entry.get("name") or eq_id or "").strip()



        if eq_id:



            id_to_name[eq_id] = eq_name



    relations = {}



    for entry in equipment_defs or []:



        profile_name = (entry.get("name") or entry.get("id") or "").strip()



        entry_id = (entry.get("id") or "").strip()



        rel = entry.get("linked_relations") or {}



        parent_block = rel.get("parent") or {}



        parent_id = (parent_block.get("equipment_id") or "").strip()



        parent_led = (parent_block.get("parent_led_id") or "").strip()



        child_entries = []



        for child in rel.get("children") or []:



            cid = (child.get("equipment_id") or "").strip()



            if not cid:



                continue



            anchor_led = (child.get("anchor_led_id") or "").strip()



            child_entries.append({



                "id": cid,



                "name": id_to_name.get(cid, ""),



                "anchor_led_id": anchor_led,



            })



        data = {



            "parent_id": parent_id,



            "parent_name": id_to_name.get(parent_id, ""),



            "parent_led_id": parent_led,



            "children": child_entries,



        }



        if profile_name:



            relations[profile_name] = data



        if entry_id and entry_id not in relations:



            relations[entry_id] = data



    return relations



def _build_truth_groups(equipment_defs):



    """



    Build mapping of source-of-truth groups.



    Returns:



        groups: {source_key: {"display_name": str, "source_profile_name": str, "source_id": str, "members": [names]}}



        child_to_root: {cad_name: source_key}



    """



    groups = {}



    child_to_root = {}



    id_to_name = {}



    for entry in equipment_defs or []:



        eq_id = (entry.get("id") or "").strip()



        eq_name = (entry.get("name") or entry.get("id") or "").strip()



        if eq_id:



            id_to_name[eq_id] = eq_name



    for entry in equipment_defs or []:



        eq_id = (entry.get("id") or "").strip()



        eq_name = (entry.get("name") or entry.get("id") or "").strip()



        if not eq_name:



            continue



        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()



        source_key = source_id or eq_id or eq_name



        source_profile_name = id_to_name.get(source_id) or eq_name



        display_name = (entry.get(TRUTH_SOURCE_NAME_KEY) or source_profile_name or eq_name).strip()



        if not display_name:



            display_name = source_profile_name



        data = groups.setdefault(source_key, {



            "display_name": display_name,



            "source_profile_name": source_profile_name,



            "source_id": source_id or eq_id,



            "members": [],



        })



        if source_id and eq_id == source_id:



            stored_display = (entry.get(TRUTH_SOURCE_NAME_KEY) or "").strip()



            if stored_display:



                data["display_name"] = stored_display



            data["source_profile_name"] = eq_name



        members = data.setdefault("members", [])



        if eq_name not in members:



            members.append(eq_name)



        child_to_root[eq_name] = source_key



    return groups, child_to_root



def _apply_truth_links(profile_dict, truth_groups):



    if not truth_groups:



        return



    profiles = profile_dict.get("profiles") or []



    by_name = {}



    for entry in profiles:



        cad = entry.get("cad_name")



        if cad:



            by_name[cad] = entry



    for source_key, data in (truth_groups or {}).items():



        source_name = data.get("source_profile_name") or source_key



        members = data.get("members") or []



        root_entry = by_name.get(source_name)



        if not root_entry:



            continue



        for cad_name in members:



            if cad_name == source_name:



                continue



            target = by_name.get(cad_name)



            if not target:



                continue



            target["types"] = copy.deepcopy(root_entry.get("types") or [])



def _apply_truth_metadata(equipment_defs, truth_groups):



    if not truth_groups:



        return



    membership = {}



    for source_key, data in truth_groups.items():



        display_name = (data.get("display_name") or data.get("source_profile_name") or source_key).strip()



        source_id = (data.get("source_id") or source_key or "").strip()



        for member in data.get("members") or []:



            membership[member] = (display_name, source_id)



    for entry in equipment_defs or []:



        eq_name = (entry.get("name") or entry.get("id") or "").strip()



        eq_id = (entry.get("id") or "").strip()



        display, source_id = membership.get(eq_name, (None, None))



        if source_id:



            entry[TRUTH_SOURCE_ID_KEY] = source_id



        elif eq_id:



            entry[TRUTH_SOURCE_ID_KEY] = eq_id



        if display:



            entry[TRUTH_SOURCE_NAME_KEY] = display



        elif eq_name:



            entry[TRUTH_SOURCE_NAME_KEY] = eq_name



def main():



    doc = getattr(revit, "doc", None)



    trans_group = TransactionGroup(doc, TITLE) if doc else None



    if trans_group:



        trans_group.Start()



    success = False



    try:



        try:



            data_path, raw_data = load_active_yaml_data()



        except RuntimeError as exc:



            forms.alert(str(exc), title=TITLE)



            return



        yaml_label = get_yaml_display_name(data_path)



        # XAML lives alongside the UI class in CEDLib.lib/UIClasses



        xaml_path = os.path.join(LIB_ROOT, "UIClasses", "ProfileEditorWindow.xaml")



        if not os.path.exists(xaml_path):



            forms.alert("ProfileEditorWindow.xaml not found under CEDLib.lib/UIClasses.", title=TITLE)



            return



        state = {



            "raw_data": raw_data,



            "yaml_label": yaml_label,



        }



        def _refresh_state_from_defs():



            equipment_defs = state["raw_data"].get("equipment_definitions") or []



            state["relations"] = _build_relations_index(equipment_defs)



            truth_groups_local, child_to_root_local = _build_truth_groups(equipment_defs)



            state["truth_groups"] = truth_groups_local



            state["child_to_root"] = child_to_root_local



            legacy_dict_local = {"profiles": equipment_defs_to_legacy(equipment_defs)}



            state["shim_profiles"] = _shims_from_dict(legacy_dict_local)



        _refresh_state_from_defs()



        state["pending_orphans"] = _load_pending_orphans(doc)



        def _run_delete_flow(selection):



            delete_path = os.path.join(os.path.dirname(__file__), "..", "Delete YAML Profiles.pushbutton", "script.py")



            delete_path = os.path.abspath(delete_path)



            if not os.path.exists(delete_path):



                forms.alert("Delete YAML Profiles script not found.", title=TITLE)



                return None



            try:



                delete_mod = sys.modules.get("ced_delete_yaml_profiles")



                if not delete_mod:



                    delete_mod = imp.load_source("ced_delete_yaml_profiles", delete_path)



            except Exception as exc:



                forms.alert("Failed to load delete script:\n\n{}".format(exc), title=TITLE)



                return None



            try:



                _, raw_data_for_delete = load_active_yaml_data()



            except RuntimeError as exc:



                forms.alert(str(exc), title=TITLE)



                return None



            equipment_defs = raw_data_for_delete.get("equipment_definitions") or []



            truth_groups_local, child_to_root_local = _build_truth_groups(equipment_defs)



            profile_name = (selection.get("profile_name") or "").strip() if selection else ""



            type_id = (selection.get("type_id") or "").strip() if selection else ""



            root_key = (selection.get("root_key") or "").strip() if selection else ""



            delete_profile_only = bool(selection.get("delete_profile")) if selection else False



            # Determine target profiles (root + mirrors)



            target_profiles = []



            if root_key and truth_groups_local.get(root_key):



                target_profiles = truth_groups_local[root_key].get("members") or []



            elif profile_name:



                target_profiles = [profile_name]



            if delete_profile_only and profile_name:



                target_profiles = [profile_name]



            removed_entries = []



            changed = False



            def _find_definition(name):



                for entry in equipment_defs:



                    entry_name = (entry.get("name") or entry.get("id") or "").strip()



                    if entry_name == name:



                        return entry



                return None



            # Delete types or whole profiles across group



            for target in target_profiles:



                definition = _find_definition(target)



                if not definition:



                    continue



                if type_id:



                    ids_to_remove = set()



                    ids_to_remove.add(type_id)



                    try:



                        delta_changed, removed_defs = delete_mod._erase_entries(equipment_defs, target, ids_to_remove)



                    except Exception:



                        # Fall back silently if helper unavailable



                        delta_changed, removed_defs = False, []



                    changed = changed or delta_changed



                    removed_entries.extend(removed_defs or [])



                else:



                    # Remove entire profile/definition



                    try:



                        equipment_defs.remove(definition)



                        removed_entries.append(definition)



                        changed = True



                    except ValueError:



                        pass



            if removed_entries:



                try:



                    cascade_entries = delete_mod._cascade_remove_children(equipment_defs, removed_entries)



                    removed_entries.extend(cascade_entries or [])



                except Exception:



                    pass



                removed_ids = [(entry.get("id") or "").strip() for entry in removed_entries if isinstance(entry, dict)]



                try:



                    delete_mod._cleanup_relations(equipment_defs, removed_ids)



                except Exception:



                    pass



            try:



                delete_mod._prune_anchor_only_definitions(raw_data_for_delete)



            except Exception:



                pass



            if not changed and not removed_entries:



                return None



            raw_data_for_delete["equipment_definitions"] = equipment_defs



            save_active_yaml_data(



                None,



                raw_data_for_delete,



                TITLE,



                "Deleted type(s) via Manage YAML Profiles",



            )



            try:



                _, refreshed_raw = load_active_yaml_data()



            except RuntimeError as exc:



                forms.alert(str(exc), title=TITLE)



                return None



            state["raw_data"] = refreshed_raw



            _refresh_state_from_defs()



            return {



                "profiles": state["shim_profiles"],



                "relations": state["relations"],



                "truth_groups": state["truth_groups"],



                "child_to_root": state["child_to_root"],



            }



        while True:



            pending_list = state.get("pending_orphans") or []



            if pending_list:



                cad_name = pending_list[0]



                doc = getattr(revit, "doc", None)



                ready = forms.alert(



                    "Capture pending orphan profile '{}' now?\nSelect Yes to pick its elements, or No to keep waiting."



                    .format(cad_name),



                    title=TITLE,



                    ok=False,



                    yes=True,



                    no=True,



                )



                if ready:



                    if _capture_orphan_profile(doc, cad_name, state, _refresh_state_from_defs, yaml_label):



                        success = True



                        pending_list.pop(0)



                        state["pending_orphans"] = pending_list



                        _save_pending_orphans(doc, pending_list)



                        _refresh_state_from_defs()



                        continue



                    return



                forms.alert(



                    "No problem. Use Revit tools to place or adjust '{}', then run Manage YAML Profiles again to finish capturing it."



                    .format(cad_name),



                    title=TITLE,



                )



                return



            window = ProfileEditorWindow(



                xaml_path,



                state["shim_profiles"],



                state["relations"],



                truth_groups=state["truth_groups"],



                child_to_root=state["child_to_root"],



                delete_callback=_run_delete_flow,



            )



            result = window.show_dialog()



            state["shim_profiles"] = getattr(window, "_profiles", state["shim_profiles"])



            state["truth_groups"] = getattr(window, "_truth_groups", state["truth_groups"])



            state["child_to_root"] = getattr(window, "_child_to_root", state["child_to_root"])



            orphan_requests = getattr(window, "orphan_requests", []) or []



            if orphan_requests:



                doc = getattr(revit, "doc", None)



                pending_list = state.get("pending_orphans") or []



                for cad_name in orphan_requests:



                    ready = forms.alert(



                        "Capture orphan profile '{}' now?\nSelect Yes to pick elements, or No to place equipment first and resume later."



                        .format(cad_name),



                        title=TITLE,



                        ok=False,



                        yes=True,



                        no=True,



                    )



                    if not ready:



                        if cad_name not in pending_list:



                            pending_list.append(cad_name)



                            state["pending_orphans"] = pending_list



                            _save_pending_orphans(doc, pending_list)



                        forms.alert(



                            "Saved '{}' for later capture. Use Revit tools as needed, then run Manage YAML Profiles again to finish."



                            .format(cad_name),



                            title=TITLE,



                        )



                        return



                    if _capture_orphan_profile(doc, cad_name, state, _refresh_state_from_defs, yaml_label):



                        success = True



                        _refresh_state_from_defs()



                    else:



                        return



                _save_pending_orphans(doc, pending_list)



                continue



            if not result:



                return



            try:



                updated_dict = _dict_from_shims(state["shim_profiles"])



                _apply_truth_links(updated_dict, state["truth_groups"])



                negatives = _find_negative_z_offsets(updated_dict)



                if negatives:



                    lines = ["Negative Z-offsets detected:"]



                    for entry in negatives[:5]:



                        if entry["source"] == "tag":



                            lines.append(" - {} / {} tag offsets = {:.2f}\"".format(entry["cad"], entry["label"], entry["value"]))



                        else:



                            lines.append(" - {} / {} offset #{} = {:.2f}\"".format(entry["cad"], entry["label"], entry["index"], entry["value"]))



                    if len(negatives) > len(lines) - 1:



                        lines.append(" - (+{} more)".format(len(negatives) - (len(lines) - 1)))



                    lines.append("")



                    lines.append("Continue saving anyway?")



                    proceed = forms.alert(



                        "\n".join(lines),



                        title=TITLE,



                        ok=False,



                        yes=True,



                        no=True,



                    )



                    if not proceed:



                        forms.alert("Save canceled. No changes were written.", title=TITLE)



                        return



                updated_defs = legacy_to_equipment_defs(



                    updated_dict.get("profiles") or [],



                    state["raw_data"].get("equipment_definitions") or [],



                )



                _apply_truth_metadata(updated_defs, state["truth_groups"])



                state["raw_data"]["equipment_definitions"] = updated_defs



                save_active_yaml_data(



                    None,



                    state["raw_data"],



                    TITLE,



                    "Updated YAML profiles via editor window",



                )



                forms.alert(



                    "Saved profile changes to {}.\nReload Place Elements (YAML) to use the updates.".format(yaml_label),



                    title=TITLE,



                )



                success = True



            except Exception as ex:



                forms.alert("Failed to save {}:\n\n{}".format(yaml_label, ex), title=TITLE)



            return



    finally:



        if trans_group:



            try:



                if success:



                    trans_group.Assimilate()



                else:



                    trans_group.RollBack()



            except Exception:



                pass



def _serialize_tags(tags):



    serialized = []



    for tg in tags:



        if isinstance(tg, dict):



            serialized.append(tg)



            continue



        family = getattr(tg, "family_name", None) or getattr(tg, "family", None)



        type_name = getattr(tg, "type_name", None) or getattr(tg, "type", None)



        cat = getattr(tg, "category_name", None) or getattr(tg, "category", None)



        offsets = getattr(tg, "offsets", None)



        offsets_dict = {}



        if isinstance(offsets, dict):



            offsets_dict = {



                "x_inches": float(offsets.get("x_inches", 0.0) or 0.0),



                "y_inches": float(offsets.get("y_inches", 0.0) or 0.0),



                "z_inches": float(offsets.get("z_inches", 0.0) or 0.0),



                "rotation_deg": float(offsets.get("rotation_deg", 0.0) or 0.0),



            }



        elif hasattr(offsets, "x_inches"):



            offsets_dict = {



                "x_inches": float(getattr(offsets, "x_inches", 0.0) or 0.0),



                "y_inches": float(getattr(offsets, "y_inches", 0.0) or 0.0),



                "z_inches": float(getattr(offsets, "z_inches", 0.0) or 0.0),



                "rotation_deg": float(getattr(offsets, "rotation_deg", 0.0) or 0.0),



            }



        serialized.append({



            "category_name": cat,



            "family_name": family,



            "type_name": type_name,



            "parameters": getattr(tg, "parameters", {}) or {},



            "offsets": offsets_dict,



        })



    return serialized



if __name__ == "__main__":



    main()



