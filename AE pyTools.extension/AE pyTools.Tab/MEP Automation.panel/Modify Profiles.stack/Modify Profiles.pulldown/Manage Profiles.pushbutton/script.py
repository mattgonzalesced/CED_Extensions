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
    FamilySymbol,
    FilteredElementCollector,
    Group,
    GroupType,
    IndependentTag,
    TextNote,
    TextNoteLeaderTypes,
    Transaction,
    TransactionGroup,
    UnitUtils,
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



from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data, load_active_yaml_text  # noqa: E402
from ExtensibleStorage import ExtensibleStorage  # noqa: E402



try:
    basestring
except NameError:
    basestring = str



TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"



TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"



ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"



ELEMENT_LINKER_SHARED_PARAM = "Element_Linker"



SAFE_HASH = u"\uff03"


def _pick_loaded_family_type(current_label=None):
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active Revit document found.", title=TITLE)
        return None
    options = []
    option_map = {}
    try:
        symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements())
    except Exception:
        symbols = []
    for sym in symbols:
        try:
            fam = getattr(sym, "Family", None)
            fam_name = getattr(fam, "Name", None) or getattr(sym, "FamilyName", None)
            type_param = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            type_name = type_param.AsString() if type_param else None
            if not type_name:
                type_name = getattr(sym, "Name", None)
            if not fam_name or not type_name:
                continue
            label = u"{} : {}".format(fam_name, type_name)
            if label in option_map:
                continue
            try:
                category_name = sym.Category.Name if sym.Category else None
            except Exception:
                category_name = None
            option_map[label] = {
                "family": fam_name,
                "type": type_name,
                "category": category_name,
            }
            options.append(label)
        except Exception:
            continue
    if not options:
        forms.alert("No loaded family types found in the model.", title=TITLE)
        return None
    options.sort(key=lambda value: value.lower())
    selection = forms.SelectFromList.show(
        options,
        title="Select Family Type",
        button_name="Select Type",
        multiselect=False,
    )
    if not selection:
        return None
    chosen = selection[0] if isinstance(selection, list) else selection
    return option_map.get(chosen)


def _pick_loaded_group_type(current_label=None):
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active Revit document found.", title=TITLE)
        return None
    options = []
    option_map = {}
    try:
        group_types = list(FilteredElementCollector(doc).OfClass(GroupType).ToElements())
    except Exception:
        group_types = []
    for gtype in group_types:
        try:
            cat = getattr(gtype, "Category", None)
            cat_name = cat.Name if cat else None
        except Exception:
            cat_name = None
        if cat_name and cat_name.lower() != "model groups":
            continue
        name = getattr(gtype, "Name", None)
        if not name:
            continue
        if name in option_map:
            continue
        option_map[name] = {"label": name, "category": cat_name or "Model Groups"}
        options.append(name)
    if not options:
        forms.alert("No model group types found in the model.", title=TITLE)
        return None
    options.sort(key=lambda value: value.lower())
    selection = forms.SelectFromList.show(
        options,
        title="Select Model Group Type",
        button_name="Select Group",
        multiselect=False,
    )
    if not selection:
        return None
    chosen = selection[0] if isinstance(selection, list) else selection
    return option_map.get(chosen)












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


def _normalize_yaml_path(path):
    if not path:
        return ""
    try:
        normalized = os.path.abspath(path)
    except Exception:
        normalized = path
    return normalized.replace("\\", "/").lower()


def _ensure_active_yaml_access(state):
    expected = state.get("normalized_yaml_path") or _normalize_yaml_path(state.get("yaml_path"))
    if not expected:
        return True
    try:
        current_path, _ = load_active_yaml_text()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return False
    current_norm = _normalize_yaml_path(current_path)
    if current_norm != expected:
        current_label = get_yaml_display_name(current_path)
        expected_label = state.get("yaml_label") or "<unknown>"
        forms.alert(
            "Active YAML changed from '{}' to '{}'. Another user likely took control.\n"
            "Close Manage YAML and reopen after selecting the desired YAML.".format(expected_label, current_label),
            title=TITLE,
        )
        return False
    return True


def _current_editor_user(doc):
    try:
        user = doc.Application.Username
        if user:
            return user
    except Exception:
        pass
    return os.getenv("USERNAME") or os.getenv("USER") or "unknown"



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



def _normalize_keynote_family(value):
    if not value:
        return ""
    text = str(value)
    if ":" in text:
        text = text.split(":", 1)[0]
    return "".join([ch for ch in text.lower() if ch.isalnum()])


def _is_ga_keynote_symbol(family_name):
    return _normalize_keynote_family(family_name) == "gakeynotesymbolced"


def _normalize_keynote_params(params):
    if not isinstance(params, dict):
        return {}
    if "Keynote Value" in params:
        if "Key Value" in params:
            params.pop("Key Value", None)
        return params
    if "Key Value" in params:
        params["Keynote Value"] = params.pop("Key Value")
    return params


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


def _is_builtin_keynote_tag(tag_entry):
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
        category = tag_entry.get("category_name") or tag_entry.get("category") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
        category = getattr(tag_entry, "category_name", None) or getattr(tag_entry, "category", None) or ""
    if _is_ga_keynote_symbol(family):
        return False
    fam_text = (family or "").lower()
    cat_text = (category or "").lower()
    if "keynote tags" in cat_text:
        return True
    if "keynote tag" in fam_text:
        return True
    return False


def _is_keynote_entry(tag_entry):
    if not tag_entry:
        return False
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
    return _is_ga_keynote_symbol(family)


def _split_keynote_entries(entries):
    normal = []
    keynotes = []
    for entry in entries or []:
        if _is_builtin_keynote_tag(entry):
            continue
        if _is_keynote_entry(entry):
            keynotes.append(entry)
        else:
            normal.append(entry)
    return normal, keynotes


def _normalize_tag_entry(tag_entry):
    if isinstance(tag_entry, dict):
        return {
            "category_name": tag_entry.get("category_name"),
            "family_name": tag_entry.get("family_name"),
            "type_name": tag_entry.get("type_name"),
            "parameters": tag_entry.get("parameters") or {},
            "offsets": tag_entry.get("offsets") or {},
        }
    return tag_entry


def _is_tag_like(elem):
    if isinstance(elem, IndependentTag):
        return True
    try:
        _ = elem.TagHeadPosition
    except Exception:
        return False
    return True


def _collect_tag_parameters(tag_elem, include_read_only=True):
    results = {}
    if tag_elem is None:
        return results
    for param in getattr(tag_elem, "Parameters", []) or []:
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
                value = elem_id.IntegerValue if elem_id else None
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
        safe_name = (name or "").replace("#", SAFE_HASH)
        results[safe_name] = value
    return results


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
        raw_keynotes = dct.get("keynotes") or []



        shim_tags = [_normalize_tag_entry(tg) for tg in raw_tags]
        shim_keynotes = [_normalize_tag_entry(tg) for tg in raw_keynotes]
        if shim_keynotes:
            shim_tags = [tg for tg in shim_tags if not _is_keynote_entry(tg)]
            self.tags = shim_tags
            self.keynotes = shim_keynotes
        else:
            normal_tags, keynote_tags = _split_keynote_entries(shim_tags)
            self.tags = normal_tags
            self.keynotes = keynote_tags

        raw_text_notes = dct.get("text_notes") or []

        self.text_notes = []

        for note in raw_text_notes:

            if isinstance(note, dict):

                self.text_notes.append(dict(note))



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
                    "keynotes": _serialize_tags(getattr(inst, "keynotes", []) or []),



                    "text_notes": _serialize_text_notes(getattr(inst, "text_notes", []) or []),



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
    cat_lower = (cat_name or "").strip().lower()
    power_categories = {"electrical fixtures", "electrical equipment", "electrical devices"}
    circuit_categories = power_categories | {"lighting fixtures", "lighting devices", "data devices"}
    capture_power = cat_lower in power_categories
    capture_circuits = cat_lower in circuit_categories

    base_targets = {"dev-Group ID": ["dev-Group ID", "dev_Group ID"]}
    power_targets = {
        "Number of Poles_CED": ["Number of Poles_CED", "Number of Poles_CEDT"],
        "Apparent Load Input_CED": ["Apparent Load Input_CED", "Apparent Load Input_CEDT"],
        "Voltage_CED": ["Voltage_CED", "Voltage_CEDT"],
        "Load Classification_CED": ["Load Classification_CED", "Load Classification_CEDT"],
        "FLA Input_CED": ["FLA Input_CED", "FLA Input_CEDT"],
        "Wattage Input_CED": ["Wattage Input_CED", "Wattage Input_CEDT"],
        "Power Factor_CED": ["Power Factor_CED", "Power Factor_CEDT"],
        "Product Datasheet URL": [
            "Product Datasheet URL",
            "Product Datasheet URL_CED",
            "Product Datasheet URL_CEDT",
        ],
        "Product Specification": [
            "Product Specification",
            "Product Specification_CED",
            "Product Specification_CEDT",
        ],
        "SLD_Component ID_CED": ["SLD_Component ID_CED"],
        "SLD_Symbol ID_CED": ["SLD_Symbol ID_CED"],
    }
    circuit_targets = {
        "CKT_Rating_CED": ["CKT_Rating_CED"],
        "CKT_Panel_CEDT": ["CKT_Panel_CED", "CKT_Panel_CEDT"],
        "CKT_Schedule Notes_CEDT": ["CKT_Schedule Notes_CED", "CKT_Schedule Notes_CEDT"],
        "CKT_Circuit Number_CEDT": ["CKT_Circuit Number_CED", "CKT_Circuit Number_CEDT"],
        "CKT_Load Name_CEDT": ["CKT_Load Name_CED", "CKT_Load Name_CEDT"],
    }

    targets = dict(base_targets)
    if capture_power:
        targets.update(power_targets)
    if capture_circuits:
        targets.update(circuit_targets)

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
                found[target_key] = _convert_collected_double(target_key, param, param.AsDouble())
            elif storage == "Integer":
                found[target_key] = param.AsInteger()
            elif storage == "ElementId" and target_key == "Load Classification_CED":
                classification_name = None
                try:
                    elem_id = param.AsElementId()
                except Exception:
                    elem_id = None
                if elem_id:
                    doc = getattr(elem, "Document", None)
                    if doc:
                        try:
                            class_elem = doc.GetElement(elem_id)
                        except Exception:
                            class_elem = None
                        if class_elem is not None:
                            classification_name = getattr(class_elem, "Name", None)
                found[target_key] = classification_name or (param.AsValueString() or "")
            else:
                found[target_key] = param.AsValueString() or ""
        except Exception:
            continue
    if "dev-Group ID" not in found:
        found["dev-Group ID"] = ""
    if capture_power and "Voltage_CED" in targets and not found.get("Voltage_CED"):
        found["Voltage_CED"] = 120
    if capture_power or capture_circuits:
        return found
    if any(value for key, value in found.items() if key != "dev-Group ID" and value):
        return found
    return {"dev-Group ID": found.get("dev-Group ID", "")}


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
    label = getattr(elem, "Name", None)
    return label or "", ""


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


def _collect_keynote_parameters(tag_elem):
    if tag_elem is None:
        return {}
    params = _collect_tag_parameters(tag_elem, include_read_only=True)
    type_params = {}
    type_elem = None
    try:
        sym = getattr(tag_elem, "Symbol", None)
        if sym:
            type_elem = sym
    except Exception:
        type_elem = None
    if type_elem is None:
        try:
            doc = getattr(tag_elem, "Document", None)
            type_id = tag_elem.GetTypeId()
            if doc and type_id:
                type_elem = doc.GetElement(type_id)
        except Exception:
            type_elem = None
    if type_elem is not None:
        type_params = _collect_tag_parameters(type_elem, include_read_only=True)
    merged = {}
    merged.update(type_params)
    merged.update(params)
    return merged


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
    entry = {
        "family_name": fam_name,
        "type_name": type_name,
        "category_name": cat_name,
        "parameters": params,
        "offsets": offsets,
    }
    if _is_keynote_entry(entry):
        entry["parameters"] = _normalize_keynote_params(
            _collect_keynote_parameters(annotation_elem)
        )
    return entry


def _build_text_note_entry(note_elem, host_point):
    if TextNote is None or host_point is None:
        return None
    try:
        if note_elem is None or not isinstance(note_elem, TextNote):
            return None
    except Exception:
        return None
    try:
        text_value = (note_elem.Text or "").strip()
    except Exception:
        text_value = ""
    if not text_value:
        return None
    note_point = getattr(note_elem, "Coord", None)
    if note_point is None:
        note_point = _get_point(note_elem)
    if note_point is None:
        return None
    offsets = {
        "x_inches": _feet_to_inches(note_point.X - host_point.X),
        "y_inches": _feet_to_inches(note_point.Y - host_point.Y),
        "z_inches": _feet_to_inches(note_point.Z - host_point.Z),
        "rotation_deg": 0.0,
    }
    try:
        rotation_rad = getattr(note_elem, "Rotation", 0.0) or 0.0
        offsets["rotation_deg"] = math.degrees(rotation_rad)
    except Exception:
        offsets["rotation_deg"] = 0.0
    width_inches = 0.0
    try:
        width_val = getattr(note_elem, "Width", None)
        if width_val not in (None, False):
            width_inches = float(width_val) * 12.0
    except Exception:
        width_inches = 0.0
    note_type_name = ""
    note_family_name = ""
    try:
        doc = getattr(note_elem, "Document", None)
        type_id = note_elem.GetTypeId()
    except Exception:
        doc = None
        type_id = None
    if doc is not None and type_id:
        try:
            note_type = doc.GetElement(type_id)
        except Exception:
            note_type = None
        if note_type is not None:
            note_type_name = (getattr(note_type, "Name", None) or "").strip()
            note_family_name = _get_text_note_family_label(note_type)
            if not note_type_name and hasattr(note_type, "get_Parameter"):
                fallback_params = (
                    BuiltInParameter.ALL_MODEL_TYPE_NAME,
                    BuiltInParameter.SYMBOL_NAME_PARAM,
                )
                for bip in fallback_params:
                    if not bip:
                        continue
                    try:
                        param = note_type.get_Parameter(bip)
                    except Exception:
                        param = None
                    if param:
                        try:
                            note_type_name = (param.AsString() or "").strip()
                        except Exception:
                            note_type_name = ""
                        if note_type_name:
                            break
    display_type = note_type_name
    if note_family_name and note_type_name and note_family_name not in note_type_name:
        display_type = u"{} : {}".format(note_family_name, note_type_name).strip()
    elif note_family_name and not note_type_name:
        display_type = note_family_name.strip()
    elif note_type_name:
        display_type = note_type_name.strip()
    else:
        display_type = ""
    leaders = _capture_text_note_leaders(note_elem, host_point)
    return {
        "text": text_value,
        "type_name": display_type,
        "width_inches": width_inches,
        "offsets": offsets,
        "leaders": leaders,
    }


def _text_note_leader_type_label(leader):
    """
    Attempt to coerce the leader type into a string so we can persist it in YAML.
    Some Revit builds expose LeaderType, others require curve inspection, so try multiple fallbacks.
    """
    if leader is None:
        return None
    # Direct property / method exposure
    for attr in ("LeaderType", "GetLeaderType", "LeaderStyle", "GetLeaderStyle", "LeaderShape"):
        try:
            raw = getattr(leader, attr, None)
            if callable(raw):
                raw = raw()
            if raw is not None:
                to_string = getattr(raw, "ToString", None)
                label = to_string() if callable(to_string) else str(raw)
                if label:
                    return label
        except Exception:
            continue
    # Curve fallback: arc curves indicate ArcLeader, straight lines remain straight.
    curve = getattr(leader, "Curve", None)
    if curve is None:
        try:
            curve_getter = getattr(leader, "GetCurve", None)
            curve = curve_getter() if callable(curve_getter) else None
        except Exception:
            curve = None
    if curve is not None:
        try:
            type_info = curve.GetType()
            type_name = getattr(type_info, "Name", None) or ""
        except Exception:
            try:
                type_name = type(curve).__name__
            except Exception:
                type_name = ""
        lowered = (type_name or "").lower()
        if "arc" in lowered:
            return "ArcLeader"
        if "line" in lowered or "straight" in lowered:
            return "StraightLeader"
        if "free" in lowered:
            return "FreeLeader"
    try:
        has_elbow = getattr(leader, "HasElbow", None)
        if callable(has_elbow):
            has_elbow = has_elbow()
    except Exception:
        has_elbow = None
    # No reliable signal detected
    return None


def _capture_text_note_leaders(note_elem, host_point):
    leaders = []
    if note_elem is None or host_point is None:
        return leaders
    try:
        leader_list = list(getattr(note_elem, "GetLeaders", lambda: [])() or [])
    except Exception:
        leader_list = []
    logger = script.get_logger()
    for leader in leader_list:
        data = {}
        leader_type_label = _text_note_leader_type_label(leader)
        if leader_type_label:
            data["type"] = leader_type_label
        try:
            end_pos = getattr(leader, "EndPosition", None)
        except Exception:
            end_pos = None
        if end_pos:
            data["end"] = _point_offsets_dict(end_pos, host_point)
        try:
            elbow_pos = getattr(leader, "ElbowPosition", None)
        except Exception:
            elbow_pos = None
        if elbow_pos:
            data["elbow"] = _point_offsets_dict(elbow_pos, host_point)
        if data:
            leaders.append(data)
    if logger:
        try:
            logger.info(
                "[Manage YAML] Captured %s leader(s) for TextNote %s: %s",
                len(leaders),
                getattr(getattr(note_elem, "Id", None), "IntegerValue", getattr(note_elem, "Id", None)),
                [entry.get("type") or "<unspecified>" for entry in leaders],
            )
        except Exception:
            pass
    return leaders


def _point_offsets_dict(point, origin):
    if point is None or origin is None:
        return None
    return {
        "x_inches": _feet_to_inches(point.X - origin.X),
        "y_inches": _feet_to_inches(point.Y - origin.Y),
        "z_inches": _feet_to_inches(point.Z - origin.Z),
    }


def _get_text_note_family_label(note_type):
    if note_type is None:
        return ""
    try:
        fam = getattr(note_type, "Family", None)
        fam_name = getattr(fam, "Name", None) if fam else None
        if fam_name:
            return fam_name
    except Exception:
        pass
    try:
        family_name = getattr(note_type, "FamilyName", None)
        if family_name:
            return family_name
    except Exception:
        pass
    if hasattr(note_type, "get_Parameter"):
        for bip in (BuiltInParameter.ALL_MODEL_FAMILY_NAME, BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM):
            if not bip:
                continue
            try:
                param = note_type.get_Parameter(bip)
            except Exception:
                param = None
            if param:
                try:
                    value = (param.AsString() or "").strip()
                except Exception:
                    value = ""
                if value:
                    return value
    return ""


def _collect_hosted_tags(elem, host_point):



    doc = getattr(elem, "Document", None)



    if doc is None or host_point is None:



        return [], []



    try:



        dep_ids = list(elem.GetDependentElements(None))



    except Exception:



        dep_ids = []



    tags = []
    keynotes = []
    text_notes = []



    for dep_id in dep_ids:



        try:



            dep_elem = doc.GetElement(dep_id)



        except Exception:



            dep_elem = None



        if _is_tag_like(dep_elem):



            try:



                tag_point = dep_elem.TagHeadPosition



            except Exception:



                tag_point = None



            tag_symbol = None



            try:



                tag_symbol = doc.GetElement(dep_elem.GetTypeId())



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



                    sym = getattr(dep_elem, "Symbol", None)



                    fam = getattr(sym, "Family", None) if sym else None



                    fam_name = getattr(fam, "Name", None) if fam else fam_name



                except Exception:



                    pass



            if not type_name:



                try:



                    tag_type = getattr(dep_elem, "TagType", None)



                    type_name = getattr(tag_type, "Name", None)



                except Exception:



                    type_name = None



            if not category_name:



                try:



                    cat = getattr(dep_elem, "Category", None)



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



            entry = {
                "family_name": fam_name or "",
                "type_name": type_name or "",
                "category_name": category_name,
                "parameters": {},
                "offsets": offsets,
            }
            if _is_builtin_keynote_tag(entry):
                continue
            if _is_keynote_entry(entry):
                entry["parameters"] = _normalize_keynote_params(
                    _collect_keynote_parameters(dep_elem)
                )
                keynotes.append(entry)
            else:
                tags.append(entry)



            continue



        text_note_entry = _build_text_note_entry(dep_elem, host_point)



        if text_note_entry:
            text_note_entry['_ced_target_locked'] = False



            text_notes.append(text_note_entry)



    return tags, keynotes, text_notes


def _find_closest_record_index(records, note_elem):
    if not records or note_elem is None:
        return None
    note_point = getattr(note_elem, "Coord", None)
    if note_point is None:
        note_point = _get_point(note_elem)
    if note_point is None:
        return None
    closest_idx = None
    closest_dist = None
    for idx, record in enumerate(records):
        host_point = record.get("host_point")
        if host_point is None:
            continue
        try:
            dist = host_point.DistanceTo(note_point)
        except Exception:
            try:
                dx = host_point.X - note_point.X
                dy = host_point.Y - note_point.Y
                dz = host_point.Z - note_point.Z
                dist = math.sqrt(dx * dx + dy * dy + dz * dz)
            except Exception:
                continue
        if closest_idx is None or dist < closest_dist:
            closest_idx = idx
            closest_dist = dist
    return closest_idx


def _text_note_preview(note_elem, limit=60):
    text_value = ""
    try:
        text_value = getattr(note_elem, "Text", "") or ""
    except Exception:
        text_value = ""
    text_value = text_value.replace("\r", " ").replace("\n", " ").strip()
    if limit and len(text_value) > limit:
        text_value = text_value[: limit - 3] + "..."
    return text_value or "<text note>"


def _record_display_label(record, index):
    type_entry = record.get("type_entry") or {}
    label = type_entry.get("label") or type_entry.get("id")
    if not label:
        element = record.get("element")
        try:
            symbol = getattr(element, "Symbol", None)
            fam = getattr(symbol, "Family", None)
            fam_name = getattr(fam, "Name", None)
            type_name = getattr(symbol, "Name", None)
            if fam_name or type_name:
                label = u"{} : {}".format(fam_name or "", type_name or "").strip(": ")
        except Exception:
            label = None
    if not label:
        label = u"Element #{}".format(index + 1)
    category = type_entry.get("category_name") or ""
    if category:
        return u"{} ({})".format(label, category)
    return label


def _select_text_note_record(element_records, note_elem, preview_text=None):
    if not element_records:
        return None
    if len(element_records) == 1:
        return 0
    option_map = {}
    options = []
    for idx, record in enumerate(element_records):
        label = _record_display_label(record, idx)
        option = u"{:02d}. {}".format(idx + 1, label)
        option_map[option] = idx
        options.append(option)
    preview = preview_text or _text_note_preview(note_elem)
    title = u"Select equipment for text note"
    if preview:
        title = u"Select equipment for note: {}".format(preview)
    try:
        selection = forms.SelectFromList.show(
            options,
            title=title,
            button_name="Point Leader",
            multiselect=False,
        )
    except Exception:
        selection = None
    if selection:
        chosen = selection[0]
        return option_map.get(chosen)
    return None


def _assign_selected_text_notes(element_records, text_note_elems):
    if not element_records or not text_note_elems:
        return
    for note in text_note_elems:
        target_idx = _select_text_note_record(element_records, note)
        if target_idx is None:
            target_idx = _find_closest_record_index(element_records, note)
        if target_idx is None:
            continue
        host_point = element_records[target_idx].get("host_point")
        if host_point is None:
            continue
        note_entry = _build_text_note_entry(note, host_point)
        if not note_entry:
            continue
        note_entry["_ced_target_locked"] = True
        type_entry = element_records[target_idx].get("type_entry") or {}
        inst_cfg = type_entry.setdefault("instance_config", {})
        entries = inst_cfg.setdefault("text_notes", [])
        entries.append(note_entry)
    _rebalance_text_notes(element_records)


def _assign_ga_keynotes(element_records, keynote_elems):
    if not element_records or not keynote_elems:
        return
    for note_elem in keynote_elems:
        target_idx = _find_closest_record_index(element_records, note_elem)
        if target_idx is None:
            continue
        host_point = element_records[target_idx].get("host_point")
        if host_point is None:
            continue
        note_entry = _build_annotation_tag_entry(note_elem, host_point)
        if not note_entry or not _is_keynote_entry(note_entry):
            continue
        type_entry = element_records[target_idx].get("type_entry") or {}
        inst_cfg = type_entry.setdefault("instance_config", {})
        entries = inst_cfg.setdefault("keynotes", [])
        entries.append(note_entry)


def _offset_dict_to_world(offsets, origin):
    if origin is None or offsets is None:
        return None
    return XYZ(
        origin.X + _inch_to_ft(offsets.get("x_inches", 0.0) or 0.0),
        origin.Y + _inch_to_ft(offsets.get("y_inches", 0.0) or 0.0),
        origin.Z + _inch_to_ft(offsets.get("z_inches", 0.0) or 0.0),
    )


def _world_to_offset_dict(point, origin):
    if origin is None or point is None:
        return None
    return {
        "x_inches": _feet_to_inches(point.X - origin.X),
        "y_inches": _feet_to_inches(point.Y - origin.Y),
        "z_inches": _feet_to_inches(point.Z - origin.Z),
    }


def _move_note_to_record(note_entry, source_record, dest_record):
    if not note_entry or not source_record or not dest_record:
        return False
    source_host = source_record.get("host_point")
    dest_host = dest_record.get("host_point")
    if source_host is None or dest_host is None:
        return False
    offsets = note_entry.get("offsets") or {}
    world_note = _offset_dict_to_world(offsets, source_host)
    new_offsets = _world_to_offset_dict(world_note, dest_host)
    if new_offsets:
        note_entry["offsets"] = new_offsets
    for leader in note_entry.get("leaders") or []:
        if not isinstance(leader, dict):
            continue
        for key in ("end", "elbow"):
            loc = leader.get(key)
            if not loc:
                continue
            world_loc = _offset_dict_to_world(loc, source_host)
            rebased = _world_to_offset_dict(world_loc, dest_host)
            if rebased:
                leader[key] = rebased
    dest_type = dest_record.get("type_entry") or {}
    dest_inst = dest_type.setdefault("instance_config", {})
    dest_notes = dest_inst.setdefault("text_notes", [])
    note_entry["_ced_target_locked"] = True
    dest_notes.append(note_entry)
    return True


def _rebalance_text_notes(element_records):
    if not element_records or len(element_records) < 2:
        return
    for idx, record in enumerate(element_records):
        type_entry = record.get("type_entry") or {}
        inst_cfg = type_entry.setdefault("instance_config", {})
        notes = list(inst_cfg.get("text_notes") or [])
        remaining = []
        for note in notes:
            if note.get("_ced_target_locked"):
                remaining.append(note)
                continue
            preview = (note.get("text") or "").strip()
            target_idx = _select_text_note_record(element_records, None, preview_text=preview)
            if target_idx is None:
                remaining.append(note)
                continue
            if target_idx == idx:
                note["_ced_target_locked"] = True
                remaining.append(note)
                continue
            moved = _move_note_to_record(note, record, element_records[target_idx])
            if not moved:
                remaining.append(note)
        inst_cfg["text_notes"] = remaining



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
        fallback_names = (
            "SCHEDULE_LEVEL_PARAM",
            "INSTANCE_REFERENCE_LEVEL_PARAM",
            "FAMILY_LEVEL_PARAM",
            "INSTANCE_LEVEL_PARAM",
        )

        for name in fallback_names:
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



def _get_level_element_id(elem):



    try:



        level_id = getattr(elem, "LevelId", None)



        if level_id and getattr(level_id, "IntegerValue", 0) > 0:



            return level_id.IntegerValue



    except Exception:



        pass



    fallback_params = (
        getattr(BuiltInParameter, "SCHEDULE_LEVEL_PARAM", None),
        getattr(BuiltInParameter, "INSTANCE_REFERENCE_LEVEL_PARAM", None),
        getattr(BuiltInParameter, "FAMILY_LEVEL_PARAM", None),
        getattr(BuiltInParameter, "INSTANCE_LEVEL_PARAM", None),
    )

    for bip in fallback_params:
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



    tags, keynotes, text_notes = _collect_hosted_tags(elem, host_point)



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
            "keynotes": keynotes,



            "text_notes": text_notes,



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



    explicit_text_notes = []

    selected_elements = []
    keynote_elements = []

    for elem in elems:



        if isinstance(elem, TextNote):



            explicit_text_notes.append(elem)



            continue

        if _is_ga_keynote_symbol_element(elem):
            keynote_elements.append(elem)
            continue



        selected_elements.append(elem)



    if not selected_elements:



        if keynote_elements:
            forms.alert(
                "Select at least one host element for '{}'. Keynote symbols are stored with the nearest host."
                .format(cad_label),
                title=TITLE,
            )
        else:
            forms.alert(
                "Select at least one host element for '{}' (text notes can be selected in addition).".format(cad_label),
                title=TITLE,
            )



        return False



    element_locations = []



    for elem in selected_elements:



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



        rot_deg = _get_rotation_degrees(elem)
        type_entry = _build_type_entry(elem, rel_vec, rot_deg, loc)



        element_records.append({



            "element": elem,



            "host_point": loc,



            "type_entry": type_entry,



        })



    _assign_selected_text_notes(element_records, explicit_text_notes)
    _assign_ga_keynotes(element_records, keynote_elements)
    _rebalance_text_notes(element_records)



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
        entry_keynotes = inst_cfg.get("keynotes") or []



        entry_text_notes = inst_cfg.get("text_notes") or []



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
            "keynotes": entry_keynotes,



            "text_notes": entry_text_notes,



        })



    truth_id = _unique_truth_id(data, equipment_def, cad_label)



    equipment_def[TRUTH_SOURCE_ID_KEY] = truth_id



    equipment_def[TRUTH_SOURCE_NAME_KEY] = cad_label



    if metadata_updates and doc:



        txn = Transaction(doc, "Create Independent Profile ({})".format(cad_label))



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



    if not _ensure_active_yaml_access(state):



        return False



    try:



        save_active_yaml_data(



            None,



            state["raw_data"],



            TITLE,



            "Captured {} independent element(s) for '{}'".format(len(element_records), cad_label),



        )



    except Exception as exc:



        forms.alert("Failed to save {}:\n\n{}".format(yaml_label, exc), title=TITLE)



        return False



    refresh_callback()



    forms.alert(



        "Saved {} independent element(s) to '{}'.\nReload placement tools to use the updates.".format(



            len(element_records),



            cad_label,



        ),



        title=TITLE,



    )



    return True



def _capture_profile_additions(doc, cad_name, state, refresh_callback, yaml_label):

    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return False

    data = state.get("raw_data") or {}
    cad_label = (cad_name or "").strip()
    if not cad_label:
        forms.alert("Profile name cannot be empty.", title=TITLE)
        return False

    equipment_def = _find_definition_by_name(data, cad_label)
    if not equipment_def:
        forms.alert("Could not locate profile '{}' in {}.".format(cad_label, yaml_label), title=TITLE)
        return False

    forms.alert(
        "Select the elements to add to '{}'.\nClick Finish when you've picked them."
        .format(cad_label),
        title=TITLE,
    )

    try:
        elems = revit.pick_elements(message="Select element(s) to add to '{}'".format(cad_label))
    except Exception:
        try:
            elem = revit.pick_element(message="Select element to add to '{}'".format(cad_label))
        except Exception:
            elem = None
        elems = [elem] if elem else []

    if not elems:
        forms.alert("No elements were selected for '{}'.".format(cad_label), title=TITLE)
        return False

    explicit_text_notes = []
    selected_elements = []
    keynote_elements = []
    for elem in elems:
        if isinstance(elem, TextNote):
            explicit_text_notes.append(elem)
            continue
        if _is_ga_keynote_symbol_element(elem):
            keynote_elements.append(elem)
            continue
        selected_elements.append(elem)

    if not selected_elements:
        if keynote_elements:
            forms.alert(
                "Select at least one host element for '{}'. Keynote symbols are stored with the nearest host."
                .format(cad_label),
                title=TITLE,
            )
        else:
            forms.alert(
                "Select at least one host element for '{}' (text notes can be selected in addition)."
                .format(cad_label),
                title=TITLE,
            )
        return False

    element_locations = []
    for elem in selected_elements:
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

    _assign_selected_text_notes(element_records, explicit_text_notes)
    _assign_ga_keynotes(element_records, keynote_elements)
    _rebalance_text_notes(element_records)

    if not element_records:
        forms.alert("No valid elements were found.", title=TITLE)
        return False

    type_set = get_type_set(equipment_def)
    set_id = type_set.get("id") or (equipment_def.get("id") or cad_label)
    led_list = type_set.setdefault("linked_element_definitions", [])

    metadata_updates = []
    for record in element_records:
        entry = record["type_entry"]
        inst_cfg = entry.setdefault("instance_config", {})
        entry_offsets = inst_cfg.get("offsets") or [{}]
        entry_params = inst_cfg.setdefault("parameters", {})
        entry_tags = inst_cfg.get("tags") or []
        entry_keynotes = inst_cfg.get("keynotes") or []
        entry_text_notes = inst_cfg.get("text_notes") or []

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
            "keynotes": entry_keynotes,
            "text_notes": entry_text_notes,
        })

    if metadata_updates and doc:
        txn = Transaction(doc, "Add Equipment to Profile ({})".format(cad_label))
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

    if not _ensure_active_yaml_access(state):
        return False

    try:
        save_active_yaml_data(
            None,
            state["raw_data"],
            TITLE,
            "Added {} element(s) to '{}'".format(len(element_records), cad_label),
        )
    except Exception as exc:
        forms.alert("Failed to save {}\n\n{}".format(yaml_label, exc), title=TITLE)
        return False

    refresh_callback()
    forms.alert(
        "Added {} element(s) to '{}'.\nReload placement tools to use the updates."
        .format(len(element_records), cad_label),
        title=TITLE,
    )
    return True


def _has_negative_z(value):



    if value is None:



        return False



    try:



        numeric = float(value)



    except Exception:



        return False



    if abs(numeric) < 1e-6:



        return False



    return numeric < 0.0



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



            tag_entries = list(inst.get("tags") or [])
            tag_entries.extend(inst.get("keynotes") or [])
            for tag in tag_entries:



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


def _log_keynote_counts(stage_label, entries):
    logger = script.get_logger()
    for cad_name, type_label, tags, keynotes in entries:
        logger.info(
            "[Manage Profiles] %s cad=%s type=%s tags=%s keynotes=%s",
            stage_label,
            cad_name or "",
            type_label or "",
            len(tags or []),
            len(keynotes or []),
        )


def _iter_equipment_def_entries(equipment_defs):
    for eq in equipment_defs or []:
        cad_name = (eq.get("name") or eq.get("id") or "").strip()
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                type_label = led.get("label") or led.get("id") or ""
                tags = led.get("tags") or []
                keynotes = led.get("keynotes") or []
                yield cad_name, type_label, tags, keynotes


def _iter_legacy_type_entries(legacy_profiles):
    for prof in legacy_profiles or []:
        cad_name = (prof.get("cad_name") or "").strip()
        for type_entry in prof.get("types") or []:
            if not isinstance(type_entry, dict):
                continue
            inst_cfg = type_entry.get("instance_config") or {}
            type_label = type_entry.get("label") or ""
            tags = inst_cfg.get("tags") or []
            keynotes = inst_cfg.get("keynotes") or []
            yield cad_name, type_label, tags, keynotes



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



    membership_by_id = {}



    name_to_ids = {}



    for entry in equipment_defs or []:



        eq_name = (entry.get("name") or entry.get("id") or "").strip()



        eq_id = (entry.get("id") or "").strip()



        if eq_name and eq_id:



            name_to_ids.setdefault(eq_name, []).append(eq_id)



    for source_key, data in truth_groups.items():



        display_name = (data.get("display_name") or data.get("source_profile_name") or source_key).strip()



        source_id = (data.get("source_id") or source_key or "").strip()



        for member in data.get("members") or []:



            member_name = (member or "").strip()



            if not member_name:



                continue



            ids = name_to_ids.get(member_name) or []



            if len(ids) == 1:



                membership_by_id[ids[0]] = (display_name, source_id)



    for entry in equipment_defs or []:



        eq_name = (entry.get("name") or entry.get("id") or "").strip()



        eq_id = (entry.get("id") or "").strip()



        display, source_id = membership_by_id.get(eq_id, (None, None))



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



    lock_user = None



    try:



        try:



            data_path, raw_data = load_active_yaml_data()



        except RuntimeError as exc:



            forms.alert(str(exc), title=TITLE)



            return



        if doc:



            candidate_user = _current_editor_user(doc)



            conflict = ExtensibleStorage.acquire_editor_lock(doc, candidate_user)



            if conflict:



                holder = conflict.get("user") or "another user"



                forms.alert(



                    "YAML storage is currently being edited by '{}'. Please wait for that user to finish, sync, "



                    "then resync your model before editing.".format(holder),



                    title=TITLE,



                )



                return



            lock_user = candidate_user



        yaml_label = get_yaml_display_name(data_path)



        # XAML lives alongside the UI class in CEDLib.lib/UIClasses



        xaml_path = os.path.join(LIB_ROOT, "UIClasses", "ProfileEditorWindow.xaml")



        if not os.path.exists(xaml_path):



            forms.alert("ProfileEditorWindow.xaml not found under CEDLib.lib/UIClasses.", title=TITLE)



            return



        state = {
            "raw_data": raw_data,
            "yaml_label": yaml_label,
            "yaml_path": data_path,
            "normalized_yaml_path": _normalize_yaml_path(data_path),
            "enable_truth_links": False,
        }



        def _refresh_state_from_defs():



            equipment_defs = state["raw_data"].get("equipment_definitions") or []



            state["relations"] = _build_relations_index(equipment_defs)



            truth_groups_local, child_to_root_local = _build_truth_groups(equipment_defs)



            state["truth_groups"] = truth_groups_local



            state["child_to_root"] = child_to_root_local



            _log_keynote_counts("equipment_defs", _iter_equipment_def_entries(equipment_defs))
            legacy_profiles_local = equipment_defs_to_legacy(equipment_defs)
            _log_keynote_counts("legacy_profiles", _iter_legacy_type_entries(legacy_profiles_local))
            legacy_dict_local = {"profiles": legacy_profiles_local}



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



            if not _ensure_active_yaml_access(state):



                return None



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



                    "Capture pending independent profile '{}' now?\nSelect Yes to pick its elements, or No to keep waiting."



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



                change_type_callback=_pick_loaded_family_type,



                change_group_callback=_pick_loaded_group_type,



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



                        "Capture independent profile '{}' now?\nSelect Yes to pick elements, or No to place equipment first and resume later."



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
            add_equipment_request = getattr(window, "add_equipment_request", None)

            if add_equipment_request:

                doc = getattr(revit, "doc", None)

                if _capture_profile_additions(doc, add_equipment_request, state, _refresh_state_from_defs, yaml_label):

                    success = True

                    _refresh_state_from_defs()

                    continue

                return





            if not result:



                return



            try:



                updated_dict = _dict_from_shims(state["shim_profiles"])



                if state.get("enable_truth_links"):
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



                if not _ensure_active_yaml_access(state):



                    return



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



        if doc and lock_user:



            try:



                ExtensibleStorage.release_editor_lock(doc, lock_user)



            except Exception:



                pass



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



def _serialize_text_notes(text_notes):



    serialized = []



    for note in text_notes:



        if isinstance(note, dict):



            clean = dict(note)



            clean.pop("_ced_target_locked", None)



            offsets = clean.get("offsets") or {}



            serialized.append({



                "text": clean.get("text") or "",



                "type_name": clean.get("type_name"),



                "width_inches": float(clean.get("width_inches", 0.0) or 0.0),



                "offsets": {



                    "x_inches": float(offsets.get("x_inches", 0.0) or 0.0),



                    "y_inches": float(offsets.get("y_inches", 0.0) or 0.0),



                    "z_inches": float(offsets.get("z_inches", 0.0) or 0.0),



                    "rotation_deg": float(offsets.get("rotation_deg", 0.0) or 0.0),



                },



                "leaders": clean.get("leaders") or [],



            })



            continue



        offsets = getattr(note, "offsets", None)



        serialized.append({



            "text": getattr(note, "text", "") or "",



            "type_name": getattr(note, "type_name", None),



            "width_inches": float(getattr(note, "width_inches", 0.0) or 0.0),



            "offsets": {



                "x_inches": float(getattr(offsets, "x_inches", 0.0) or 0.0) if offsets else 0.0,



                "y_inches": float(getattr(offsets, "y_inches", 0.0) or 0.0) if offsets else 0.0,



                "z_inches": float(getattr(offsets, "z_inches", 0.0) or 0.0) if offsets else 0.0,



                "rotation_deg": float(getattr(offsets, "rotation_deg", 0.0) or 0.0) if offsets else 0.0,



            },



            "leaders": list(getattr(note, "leaders", []) or []),



        })



    return serialized



if __name__ == "__main__":



    main()







