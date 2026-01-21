# -*- coding: utf-8 -*-
"""
Edit/Create YAML Profiles
-------------------------
Select a linked parent element, optionally load any existing equipment profile,
then capture offsets/tags for the selected equipment to update the
active YAML definition.
"""

import copy
import json
import math
import os
import sys

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Group,
    IndependentTag,
    TextNote,
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

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from LogicClasses.linked_equipment import compute_offsets_from_points  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402
from LogicClasses.profile_schema import ensure_equipment_definition, get_type_set, next_led_id, equipment_defs_to_legacy  # noqa: E402

TITLE = "Edit/Create YAML Profiles"
LOG = script.get_logger()
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", ELEMENT_LINKER_PARAM_NAME)
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
SESSION_KEY = "edit_create_yaml_session"
SAFE_HASH = u"\uff03"

try:
    basestring
except NameError:
    basestring = str


# --------------------------------------------------------------------------- #
# Session helpers
# --------------------------------------------------------------------------- #


def _session_file():
    try:
        return script.get_appdata_file("edit_create_yaml_session.json")
    except Exception:
        return os.path.join(os.path.expanduser("~"), "edit_create_yaml_session.json")


def _session_storage():
    path = _session_file()
    if os.path.exists(path):
        try:
            with open(path, "r") as handle:
                data = json.load(handle)
                if isinstance(data, dict):
                    return data
        except Exception:
            return {}
    return {}


def _doc_session_key(doc):
    path = getattr(doc, "PathName", "") or ""
    if path:
        return path.lower()
    title = getattr(doc, "Title", "") or ""
    return title.lower()


def _load_session(doc):
    key = _doc_session_key(doc)
    store = _session_storage()
    return store.get(key)


def _save_session(doc, session):
    key = _doc_session_key(doc)
    store = _session_storage()
    if session:
        store[key] = session
    elif key in store:
        store.pop(key, None)
    try:
        directory = os.path.dirname(_session_file())
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        with open(_session_file(), "w") as handle:
            json.dump(store, handle)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Identifier helpers
# --------------------------------------------------------------------------- #


def _next_eq_number(data, exclude_defs=None):
    excluded = {id(entry) for entry in (exclude_defs or []) if entry is not None}
    max_id = 0
    for entry in data.get("equipment_definitions") or []:
        if excluded and id(entry) in excluded:
            continue
        eq_id = (entry.get("id") or "").strip()
        if not eq_id:
            continue
        suffix = eq_id.split("-")[-1]
        try:
            num = int(suffix)
        except Exception:
            continue
        if num > max_id:
            max_id = num
    return max_id + 1


# --------------------------------------------------------------------------- #
# Selection helpers
# --------------------------------------------------------------------------- #


def _linked_element_from_reference(doc, reference):
    linked_id = getattr(reference, "LinkedElementId", None)
    if isinstance(linked_id, ElementId) and linked_id != ElementId.InvalidElementId:
        try:
            host_elem = doc.GetElement(reference.ElementId)
        except Exception:
            host_elem = None
        if isinstance(host_elem, RevitLinkInstance):
            link_doc = host_elem.GetLinkDocument()
            if link_doc:
                transform = host_elem.GetTransform()
                if not isinstance(transform, Transform):
                    transform = None
                try:
                    linked_elem = link_doc.GetElement(linked_id)
                except Exception:
                    linked_elem = None
                return linked_elem, transform
    try:
        elem = doc.GetElement(reference.ElementId)
    except Exception:
        elem = None
    return elem, None


def _pick_parent_element(prompt):
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return None, None
    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, prompt)
    except Exception:
        return None, None
    elem, transform = _linked_element_from_reference(doc, reference)
    if isinstance(elem, RevitLinkInstance):
        forms.alert("Select the specific element inside the link after this message.", title=TITLE)
        try:
            link_ref = uidoc.Selection.PickObject(ObjectType.LinkedElement, prompt)
        except Exception:
            return None, None
        return _linked_element_from_reference(doc, link_ref)
    return elem, transform


# --------------------------------------------------------------------------- #
# Geometry helpers
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
            pass
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


def _get_rotation_degrees(elem):
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            pass
    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        basis = getattr(transform, "BasisX", None)
        if basis:
            try:
                return math.degrees(math.atan2(basis.Y, basis.X))
            except Exception:
                pass
    return 0.0


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
        ang = math.radians(rotation_deg or 0.0)
    except Exception:
        ang = 0.0
    vec = XYZ(math.cos(ang), math.sin(ang), 0.0)
    try:
        vec_world = transform.OfVector(vec)
        return math.degrees(math.atan2(vec_world.Y, vec_world.X))
    except Exception:
        return rotation_deg


def _level_relative_z_inches(elem, world_point):
    if elem is None:
        return 0.0

    direct = _instance_elevation_inches(elem)
    if direct is not None:
        return direct

    doc = getattr(elem, "Document", None)
    level = None
    level_id = getattr(elem, "LevelId", None)
    if doc and level_id:
        try:
            level = doc.GetElement(level_id)
        except Exception:
            level = None
    if not level:
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
                    level = doc.GetElement(eid)
            except Exception:
                level = None
            if level:
                break
    level_z = 0.0
    if level:
        try:
            level_z = float(level.Elevation or 0.0)
        except Exception:
            level_z = 0.0
    world_z = getattr(world_point, "Z", None) or 0.0
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


# --------------------------------------------------------------------------- #
# YAML helpers
# --------------------------------------------------------------------------- #


def _collect_params(elem):
    def _convert_double_value(target_key, param_obj, raw_value):
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

    found = {key: "" for key in targets}
    for param in getattr(elem, "Parameters", []) or []:
        try:
            name = param.Definition.Name
        except Exception:
            continue
        target = None
        for out_key, aliases in targets.items():
            if name in aliases:
                target = out_key
                break
        if not target:
            continue
        try:
            storage = param.StorageType
        except Exception:
            storage = None
        try:
            storage_str = storage.ToString() if storage else ""
        except Exception:
            storage_str = ""
        try:
            if storage_str == "String":
                found[target] = param.AsString() or ""
            elif storage_str == "Double":
                found[target] = _convert_double_value(target, param, param.AsDouble())
            elif storage_str == "Integer":
                found[target] = param.AsInteger()
            else:
                found[target] = param.AsValueString() or ""
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


def _tag_signature(name):
    return " ".join((name or "").strip().split()).lower()


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


def _is_tag_like(elem):
    if isinstance(elem, IndependentTag):
        return True
    try:
        _ = elem.TagHeadPosition
    except Exception:
        return False
    return True


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


def _is_keynote_entry(tag_entry):
    if not tag_entry:
        return False
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
        type_name = tag_entry.get("type_name") or tag_entry.get("type") or ""
        category = tag_entry.get("category_name") or tag_entry.get("category") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
        type_name = getattr(tag_entry, "type_name", None) or getattr(tag_entry, "type", None) or ""
        category = getattr(tag_entry, "category_name", None) or getattr(tag_entry, "category", None) or ""
    text = "{} {} {}".format(family, type_name, category).lower()
    return "keynote" in text


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
    label = _build_label(elem)
    if ":" in label:
        fam_part, type_part = label.split(":", 1)
        return fam_part.strip(), type_part.strip()
    return label, ""


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
    cat_lower = (cat_name or "").lower()
    if not cat_name or ("generic annotation" not in cat_lower and "keynote" not in cat_lower):
        return None
    ann_point = _get_point(annotation_elem)
    if ann_point is None or host_point is None:
        return None
    fam_name, type_name = _annotation_family_type(annotation_elem)
    offsets = {
        "x_inches": _feet_to_inches(ann_point.X - host_point.X),
        "y_inches": _feet_to_inches(ann_point.Y - host_point.Y),
        "z_inches": _feet_to_inches(ann_point.Z - host_point.Z),
        "rotation_deg": _normalize_angle(_get_rotation_degrees(annotation_elem)),
    }
    params = _collect_annotation_string_params(annotation_elem)
    return {
        "family_name": fam_name,
        "type_name": type_name,
        "category_name": cat_name,
        "parameters": params,
        "offsets": offsets,
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
        except Exception:
            type_name = None
        try:
            cat = getattr(tag_symbol, "Category", None)
            category_name = getattr(cat, "Name", None) if cat else None
        except Exception:
            category_name = None
    if not (fam_name and type_name):
        try:
            symbols = getattr(tag, "Symbol", None)
        except Exception:
            symbols = None
        if symbols:
            try:
                fam = getattr(symbols, "Family", None)
                fam_name = getattr(fam, "Name", None) if fam else fam_name
            except Exception:
                pass
            try:
                type_name = getattr(symbols, "Name", None) or type_name
                if not type_name and hasattr(symbols, "get_Parameter"):
                    param = symbols.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if param:
                        type_name = param.AsString()
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
    family_field = fam_name or ""
    type_field = type_name or ""
    if not type_field and tag_symbol:
        try:
            param = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if param:
                type_field = param.AsString() or ""
        except Exception:
            pass
    if not type_field:
        try:
            tag_type = getattr(tag, "TagType", None)
            if tag_type:
                type_field = getattr(tag_type, "Name", None) or type_field
        except Exception:
            pass
    if not type_field and ":" in (family_field or ""):
        fam_part, type_part = family_field.split(":", 1)
        family_field = fam_part.strip()
        type_field = type_part.strip()
    return {
        "family_name": family_field,
        "type_name": type_field,
        "category_name": category_name,
        "parameters": {},
        "offsets": offsets,
        "leader_elbow": leader_elbow,
        "leader_end": leader_end,
    }


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


def _text_note_leader_type_label(leader):
    if leader is None:
        return None
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
    return None


def _capture_text_note_leaders(note_elem, host_point):
    leaders = []
    if note_elem is None or host_point is None:
        return leaders
    try:
        leader_list = list(getattr(note_elem, "GetLeaders", lambda: [])() or [])
    except Exception:
        leader_list = []
    for leader in leader_list:
        data = {}
        label = _text_note_leader_type_label(leader)
        if label:
            data["type"] = label
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
    return leaders


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
        rotation_val = getattr(note_elem, "Rotation", None)
        if rotation_val not in (None, False):
            offsets["rotation_deg"] = math.degrees(rotation_val)
    except Exception:
        pass
    width_inches = 0.0
    try:
        width_val = getattr(note_elem, "Width", None)
        if width_val not in (None, False):
            width_inches = float(width_val) * 12.0
    except Exception:
        width_inches = 0.0
    note_type_name = ""
    note_family_name = ""
    doc = getattr(note_elem, "Document", None)
    type_id = None
    try:
        type_id = note_elem.GetTypeId()
    except Exception:
        type_id = None
    note_type = None
    if doc is not None and type_id:
        try:
            note_type = doc.GetElement(type_id)
        except Exception:
            note_type = None
    if note_type is not None:
        note_type_name = (getattr(note_type, "Name", None) or "").strip()
        if not note_type_name and hasattr(note_type, "get_Parameter"):
            for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM):
                if not bip:
                    continue
                try:
                    param = note_type.get_Parameter(bip)
                except Exception:
                    param = None
                if param:
                    try:
                        candidate = (param.AsString() or "").strip()
                    except Exception:
                        candidate = ""
                    if candidate:
                        note_type_name = candidate
                        break
        note_family_name = _get_text_note_family_label(note_type)
    display_type = ""
    if note_family_name and note_type_name:
        display_type = u"{} : {}".format(note_family_name.strip(), note_type_name.strip())
    elif note_family_name:
        display_type = note_family_name.strip()
    else:
        display_type = note_type_name.strip()
    leaders = _capture_text_note_leaders(note_elem, host_point)
    return {
        "text": text_value,
        "type_name": display_type,
        "width_inches": width_inches,
        "offsets": offsets,
        "leaders": leaders,
    }


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


def _find_closest_child_entry(entries, note_elem):
    if not entries or note_elem is None:
        return None
    note_point = getattr(note_elem, "Coord", None)
    if note_point is None:
        note_point = _get_point(note_elem)
    if note_point is None:
        return None
    return _find_closest_entry_by_point(entries, note_point)


def _assign_selected_text_notes(child_entries, note_elems):
    if not child_entries or not note_elems:
        return
    for note in note_elems:
        target_idx = _find_closest_child_entry(child_entries, note)
        if target_idx is None:
            continue
        host_point = child_entries[target_idx].get("point")
        if host_point is None:
            continue
        entry = _build_text_note_entry(note, host_point)
        if not entry:
            continue
        notes = child_entries[target_idx].setdefault("text_notes", [])
        notes.append(entry)


def _assign_selected_tags(child_entries, tag_elems):
    if not child_entries or not tag_elems:
        return
    host_index = {}
    for idx, entry in enumerate(child_entries):
        elem = entry.get("element")
        if elem is None:
            continue
        try:
            elem_id = elem.Id.IntegerValue
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
                host_id_val = host_id.IntegerValue
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
        target_list_name = "keynotes" if _is_keynote_entry(entry) else "tags"
        tags = child_entries[target_idx].setdefault(target_list_name, [])
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


def _collect_hosted_tags(elem, host_point, active_view_id=None):
    doc = getattr(elem, "Document", None)
    if not doc or host_point is None:
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
        if dep_elem is None:
            continue
        if active_view_id is not None:
            try:
                owner_view_id = getattr(dep_elem, "OwnerViewId", None)
            except Exception:
                owner_view_id = None
            if not owner_view_id:
                continue
            try:
                if owner_view_id.IntegerValue != active_view_id:
                    continue
            except Exception:
                pass
        if _is_tag_like(dep_elem):
            entry = _build_independent_tag_entry(dep_elem, host_point)
            if entry:
                if _is_keynote_entry(entry):
                    keynotes.append(entry)
                else:
                    tags.append(entry)
            continue
        annotation_entry = _build_annotation_tag_entry(dep_elem, host_point)
        if annotation_entry:
            if _is_keynote_entry(annotation_entry):
                keynotes.append(annotation_entry)
            else:
                tags.append(annotation_entry)
            continue
        text_entry = _build_text_note_entry(dep_elem, host_point)
        if text_entry:
            text_notes.append(text_entry)
    return tags, keynotes, text_notes


def _normalize_angle(value):
    if value is None:
        return 0.0
    ang = float(value)
    while ang <= -180.0:
        ang += 360.0
    while ang > 180.0:
        ang -= 360.0
    return ang


def _build_label(elem):
    fam_name = None
    type_name = None
    if isinstance(elem, Group):
        try:
            fam_name = elem.Name
            type_name = elem.Name
        except Exception:
            pass
    else:
        try:
            sym = getattr(elem, "Symbol", None)
            if sym:
                fam = getattr(sym, "Family", None)
                fam_name = getattr(fam, "Name", None) if fam else None
                type_name = getattr(sym, "Name", None)
                if not fam_name:
                    fam_name = getattr(sym, "FamilyName", None)
                if not type_name and hasattr(sym, "get_Parameter"):
                    try:
                        name_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if name_param:
                            type_name = name_param.AsString()
                    except Exception:
                        pass
        except Exception:
            pass
    if not fam_name and hasattr(elem, "Name"):
        try:
            fam_name = elem.Name
        except Exception:
            fam_name = None
    if not type_name and not fam_name and hasattr(elem, "Name"):
        try:
            type_name = elem.Name
        except Exception:
            type_name = None
    if fam_name and type_name:
        return "{} : {}".format(fam_name, type_name)
    return type_name or fam_name or "Unnamed"


def _get_category(elem):
    try:
        cat = elem.Category
        if cat:
            return cat.Name or ""
    except Exception:
        pass
    return ""


def _candidate_equipment_names(elem):
    names = []
    try:
        if hasattr(elem, "Name") and elem.Name:
            names.append(elem.Name.strip())
    except Exception:
        pass
    label = _build_label(elem)
    if label:
        names.insert(0, label)
    uniq = []
    seen = set()
    for name in names:
        norm = (name or "").strip()
        if not norm:
            continue
        key = norm.lower()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(norm)
    return uniq


def _normalize_key(value):
    value = (value or "").strip().lower()
    if not value:
        return ""
    return "".join(ch for ch in value if ch.isalnum())


def _find_equipment_definition_by_name(data, name):
    raw = (name or "").strip()
    target = raw.lower()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        current_raw = (eq.get("name") or eq.get("id") or "").strip()
        if not current_raw:
            continue
        current = current_raw.lower()
        if current == target:
            return eq
    return None


def _sanitize_equipment_definitions(equipment_defs):
    cleaned_defs = []
    for eq in equipment_defs or []:
        if not isinstance(eq, dict):
            continue
        sanitized = dict(eq)
        linked_sets = []
        for linked_set in sanitized.get("linked_sets") or []:
            if not isinstance(linked_set, dict):
                continue
            ls_copy = dict(linked_set)
            led_defs = []
            for led in ls_copy.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_defs.append(dict(led))
            ls_copy["linked_element_definitions"] = led_defs
            linked_sets.append(ls_copy)
        sanitized["linked_sets"] = linked_sets
        cleaned_defs.append(sanitized)
    return cleaned_defs


def _sanitize_profiles(profiles):
    cleaned = []
    for prof in profiles or []:
        if not isinstance(prof, dict):
            continue
        prof_copy = dict(prof)
        type_list = []
        for entry in prof_copy.get("types") or []:
            if not isinstance(entry, dict):
                continue
            type_entry = dict(entry)
            inst_cfg = type_entry.get("instance_config") or {}
            offsets = inst_cfg.get("offsets")
            if not isinstance(offsets, list) or not offsets:
                offsets = [{}]
            inst_cfg["offsets"] = [off if isinstance(off, dict) else {} for off in offsets]
            tags = inst_cfg.get("tags") or []
            inst_cfg["tags"] = [tag if isinstance(tag, dict) else {} for tag in tags]
            keynotes = inst_cfg.get("keynotes") or []
            inst_cfg["keynotes"] = [tag if isinstance(tag, dict) else {} for tag in keynotes]
            params = inst_cfg.get("parameters")
            inst_cfg["parameters"] = params if isinstance(params, dict) else {}
            type_entry["instance_config"] = inst_cfg
            type_list.append(type_entry)
        prof_copy["types"] = type_list
        cleaned.append(prof_copy)
    return cleaned


def _build_repository(data):
    cleaned_defs = _sanitize_equipment_definitions(data.get("equipment_definitions") or [])
    legacy_profiles = equipment_defs_to_legacy(cleaned_defs)
    cleaned_profiles = _sanitize_profiles(legacy_profiles)
    eq_defs = ProfileRepository._parse_profiles(cleaned_profiles)
    return ProfileRepository(eq_defs)


def _extract_set_id(payload_text):
    if not payload_text:
        return ""
    text = str(payload_text)
    if "\n" in text:
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, remainder = line.partition(":")
            if key.strip().lower() == "set definition id":
                return remainder.strip()
    marker = "Set Definition ID:"
    idx = text.find(marker)
    if idx < 0:
        return ""
    remainder = text[idx + len(marker):]
    return remainder.split(",", 1)[0].strip()


def _extract_led_id(payload_text):
    if not payload_text:
        return ""
    text = str(payload_text)
    if "\n" in text:
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, remainder = line.partition(":")
            if key.strip().lower() == "linked element definition id":
                return remainder.strip()
    marker = "Linked Element Definition ID:"
    idx = text.find(marker)
    if idx < 0:
        return ""
    remainder = text[idx + len(marker):]
    return remainder.split(",", 1)[0].strip()


def _get_element_linker_text(elem):
    if elem is None:
        return ""
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        try:
            text = param.AsString()
        except Exception:
            text = None
        if not text:
            try:
                text = param.AsValueString()
            except Exception:
                text = None
        if text and str(text).strip():
            return str(text)
    return ""


def _tag_host_element_ids(tag):
    if tag is None:
        return []
    try:
        getter = getattr(tag, "GetTaggedLocalElementIds", None)
        if callable(getter):
            return list(getter() or [])
    except Exception:
        pass
    ids = []
    for attr in ("TaggedLocalElementId", "TaggedElementId"):
        try:
            value = getattr(tag, attr, None)
        except Exception:
            value = None
        if value:
            ids.append(value)
    return ids


def _tag_display_label(tag):
    if tag is None:
        return "<Tag?>"
    doc = getattr(tag, "Document", None)
    tag_type = None
    if doc is not None:
        try:
            tag_type = doc.GetElement(tag.GetTypeId())
        except Exception:
            tag_type = None
    fam_name = None
    type_name = None
    if tag_type:
        try:
            fam = getattr(tag_type, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else getattr(tag_type, "FamilyName", None)
        except Exception:
            fam_name = None
        try:
            type_name = getattr(tag_type, "Name", None)
        except Exception:
            type_name = None
    if fam_name and type_name:
        return u"{} : {}".format(fam_name, type_name)
    return type_name or fam_name or "<Tag?>"


def _fmt_point(point):
    if point is None:
        return "<none>"
    try:
        return "{:.3f},{:.3f},{:.3f}".format(point.X, point.Y, point.Z)
    except Exception:
        return "<none>"


def _ensure_element_id(value):
    if isinstance(value, ElementId):
        return value
    if value in (None, ""):
        return None
    try:
        return ElementId(int(value))
    except Exception:
        return None


def _tag_label_key(label):
    if not label:
        return ""
    return " ".join(str(label).strip().lower().split())


def _tag_key_from_def(tag_def):
    if not isinstance(tag_def, dict):
        return ""
    family = tag_def.get("family") or tag_def.get("family_name") or ""
    type_name = tag_def.get("type") or tag_def.get("type_name") or ""
    if family and type_name:
        return _tag_label_key(u"{} : {}".format(family, type_name))
    return _tag_label_key(type_name or family)


def _cleanup_far_tags_by_type(doc, active_view_id, tag_type_keys, max_distance_ft, expected_points=None, tag_ids=None):
    if not doc or not active_view_id or not max_distance_ft:
        return 0
    try:
        limit = float(max_distance_ft)
    except Exception:
        return 0
    view_id = _ensure_element_id(active_view_id)
    if not view_id:
        return 0
    allowed_ids = None
    if tag_ids:
        try:
            allowed_ids = {int(val) for val in tag_ids}
        except Exception:
            allowed_ids = None
    try:
        tags = list(FilteredElementCollector(doc, view_id).OfClass(IndependentTag))
    except Exception:
        tags = []
    if not tags:
        return 0
    if not expected_points:
        LOG.info("[Edit-Create Profiles] Autoload tag cleanup skipped: no expected tag points.")
        return 0
    def _distance_between_points(a, b):
        if a is None or b is None:
            return None
        try:
            return a.DistanceTo(b)
        except Exception:
            try:
                dx = a.X - b.X
                dy = a.Y - b.Y
                dz = a.Z - b.Z
                return math.sqrt(dx * dx + dy * dy + dz * dz)
            except Exception:
                return None

    def _distance_to_nearest_expected(point):
        if point is None or not expected_points:
            return None
        nearest = None
        for expected_point in expected_points:
            dist = _distance_between_points(expected_point, point)
            if dist is None:
                continue
            if nearest is None or dist < nearest:
                nearest = dist
        return nearest

    removed = 0
    seen = 0
    debug = []
    txn = Transaction(doc, "Clean Autoload Tag Types")
    try:
        txn.Start()
        for tag in tags:
            if allowed_ids is not None:
                try:
                    tag_id_val = tag.Id.IntegerValue
                except Exception:
                    tag_id_val = None
                if tag_id_val not in allowed_ids:
                    continue
            if tag_type_keys:
                label = _tag_label_key(_tag_display_label(tag))
                if label not in tag_type_keys:
                    continue
            seen += 1
            if len(debug) < 5:
                try:
                    debug.append(_tag_display_label(tag))
                except Exception:
                    debug.append("<tag>")
            try:
                tag_point = tag.TagHeadPosition
            except Exception:
                tag_point = None
            if tag_point is None:
                LOG.info(
                    "[Edit-Create Profiles] Autoload tag cleanup: '%s' missing tag head position.",
                    _tag_display_label(tag),
                )
                continue
            dist = _distance_to_nearest_expected(tag_point)
            if allowed_ids is not None:
                label = _tag_display_label(tag)
                is_keynote = "keynote" in (label or "").lower()
                LOG.info(
                    "[Edit-Create Profiles] Autoload tag check: '%s' head=(%s) nearest=%.2f ft keynote=%s",
                    label,
                    _fmt_point(tag_point),
                    dist if dist is not None else -1.0,
                    is_keynote,
                )
            source_label = "expected tag"
            if dist is None or dist <= limit:
                continue
            try:
                doc.Delete(tag.Id)
                removed += 1
                LOG.info(
                    "[Edit-Create Profiles] Removed tag '%s' %.2f ft from %s (limit %.2f ft).",
                    _tag_display_label(tag),
                    dist,
                    source_label,
                    limit,
                )
            except Exception:
                continue
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass
    LOG.info(
        "[Edit-Create Profiles] Autoload tag cleanup scanned %s tag(s); removed %s.",
        seen,
        removed,
    )
    if seen == 0 and allowed_ids is not None:
        LOG.info(
            "[Edit-Create Profiles] Autoload tag cleanup debug: allowed_ids=%s sample=%s",
            len(allowed_ids),
            ", ".join(debug) if debug else "<none>",
        )
    return removed


def _cleanup_existing_profile_tags(doc, active_view_id, tag_type_keys, set_ids, led_ids, expected_points, expected_hosts, max_distance_ft):
    if not doc or not active_view_id or not tag_type_keys or not max_distance_ft:
        return 0
    if not expected_points:
        LOG.info("[Edit-Create Profiles] Existing tag cleanup skipped: no expected tag points.")
        return 0
    if not set_ids and not led_ids:
        LOG.info("[Edit-Create Profiles] Existing tag cleanup skipped: no set/led ids available.")
        return 0
    try:
        limit = float(max_distance_ft)
    except Exception:
        return 0
    view_id = _ensure_element_id(active_view_id)
    if not view_id:
        LOG.info("[Edit-Create Profiles] Existing tag cleanup skipped: invalid view id.")
        return 0
    try:
        tags = list(FilteredElementCollector(doc, view_id).OfClass(IndependentTag))
    except Exception:
        tags = []
    if not tags:
        LOG.info("[Edit-Create Profiles] Existing tag cleanup found 0 tags in active view.")
        return 0

    def _distance_between_points(a, b):
        if a is None or b is None:
            return None
        try:
            return a.DistanceTo(b)
        except Exception:
            try:
                dx = a.X - b.X
                dy = a.Y - b.Y
                dz = a.Z - b.Z
                return math.sqrt(dx * dx + dy * dy + dz * dz)
            except Exception:
                return None

    def _distance_to_nearest_expected(point):
        if point is None:
            return None
        nearest = None
        for expected_point in expected_points:
            dist = _distance_between_points(expected_point, point)
            if dist is None:
                continue
            if nearest is None or dist < nearest:
                nearest = dist
        return nearest

    def _host_matches_expected(host_point):
        if host_point is None or not expected_hosts:
            return False
        nearest = None
        for expected_host in expected_hosts:
            dist = _distance_between_points(expected_host, host_point)
            if dist is None:
                continue
            if nearest is None or dist < nearest:
                nearest = dist
        return nearest is not None and nearest <= limit

    removed = 0
    scanned = 0
    matched_type = 0
    matched_host = 0
    txn = Transaction(doc, "Clean Existing Autoload Tags")
    try:
        txn.Start()
        for tag in tags:
            label = _tag_label_key(_tag_display_label(tag))
            if label not in tag_type_keys:
                continue
            matched_type += 1
            host_elem = None
            for host_id in _tag_host_element_ids(tag):
                try:
                    host_elem = doc.GetElement(host_id)
                except Exception:
                    host_elem = None
                if host_elem is not None:
                    break
            if host_elem is None:
                continue
            payload = _get_element_linker_text(host_elem)
            set_id = _extract_set_id(payload)
            led_id = _extract_led_id(payload)
            if set_id and set_id in set_ids:
                pass
            elif led_id and led_id in led_ids:
                pass
            else:
                host_point = _get_point(host_elem)
                if not _host_matches_expected(host_point):
                    continue
            matched_host += 1
            scanned += 1
            try:
                tag_point = tag.TagHeadPosition
            except Exception:
                tag_point = None
            if tag_point is None:
                continue
            dist = _distance_to_nearest_expected(tag_point)
            if dist is None or dist <= limit:
                continue
            try:
                doc.Delete(tag.Id)
                removed += 1
                LOG.info(
                    "[Edit-Create Profiles] Removed existing tag '%s' %.2f ft from expected tag (limit %.2f ft).",
                    _tag_display_label(tag),
                    dist,
                    limit,
                )
            except Exception:
                continue
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass
    LOG.info(
        "[Edit-Create Profiles] Existing tag cleanup matched type=%s host=%s scanned=%s removed=%s.",
        matched_type,
        matched_host,
        scanned,
        removed,
    )
    return removed


def _cleanup_far_autoload_tags(doc, active_view_id, set_ids, max_distance_ft):
    if not doc or not active_view_id or not set_ids or not max_distance_ft:
        return 0
    try:
        limit = float(max_distance_ft)
    except Exception:
        return 0
    view_id = _ensure_element_id(active_view_id)
    if not view_id:
        return 0
    host_lookup = {}
    try:
        elements = FilteredElementCollector(doc, view_id).WhereElementIsNotElementType()
    except Exception:
        elements = []
    for elem in elements:
        payload = _get_element_linker_text(elem)
        if not payload:
            continue
        set_id = _extract_set_id(payload)
        if not set_id or set_id not in set_ids:
            continue
        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            continue
        host_lookup[elem_id] = elem
    if not host_lookup:
        return 0
    try:
        tags = list(FilteredElementCollector(doc, view_id).OfClass(IndependentTag))
    except Exception:
        tags = []
    if not tags:
        return 0
    removed = 0
    txn = Transaction(doc, "Clean Autoload Tags")
    try:
        txn.Start()
        for tag in tags:
            host_elem = None
            for host_id in _tag_host_element_ids(tag):
                try:
                    host_id_val = host_id.IntegerValue
                except Exception:
                    try:
                        host_id_val = int(host_id)
                    except Exception:
                        host_id_val = None
                if host_id_val is None:
                    continue
                host_elem = host_lookup.get(host_id_val)
                if host_elem is not None:
                    break
            if host_elem is None:
                continue
            host_point = _get_point(host_elem)
            if host_point is None:
                continue
            try:
                tag_point = tag.TagHeadPosition
            except Exception:
                tag_point = None
            if tag_point is None:
                continue
            try:
                dist = host_point.DistanceTo(tag_point)
            except Exception:
                try:
                    dx = host_point.X - tag_point.X
                    dy = host_point.Y - tag_point.Y
                    dz = host_point.Z - tag_point.Z
                    dist = math.sqrt(dx * dx + dy * dy + dz * dz)
                except Exception:
                    dist = None
            if dist is None or dist <= limit:
                continue
            try:
                doc.Delete(tag.Id)
                removed += 1
                LOG.info(
                    "[Edit-Create Profiles] Removed autoload tag '%s' %.2f ft from host (limit %.2f ft).",
                    _tag_display_label(tag),
                    dist,
                    limit,
                )
            except Exception:
                continue
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass
    return removed


def _filter_tags_for_autoload(data, cad_name, max_distance_ft):
    if not data or not cad_name or not max_distance_ft:
        return data
    try:
        filtered = copy.deepcopy(data)
    except Exception:
        return data
    target = (cad_name or "").strip().lower()
    for eq in filtered.get("equipment_definitions") or []:
        name = (eq.get("name") or eq.get("id") or "").strip().lower()
        if name != target:
            continue
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                def _filter_tag_list(items):
                    kept = []
                    for tag in items:
                        if not isinstance(tag, dict):
                            kept.append(tag)
                            continue
                        offsets = tag.get("offsets") or {}
                        if not isinstance(offsets, dict):
                            kept.append(tag)
                            continue
                        try:
                            x_ft = float(offsets.get("x_inches", 0.0) or 0.0) / 12.0
                            y_ft = float(offsets.get("y_inches", 0.0) or 0.0) / 12.0
                            z_ft = float(offsets.get("z_inches", 0.0) or 0.0) / 12.0
                        except Exception:
                            kept.append(tag)
                            continue
                        dist = math.sqrt((x_ft * x_ft) + (y_ft * y_ft) + (z_ft * z_ft))
                        if dist <= max_distance_ft:
                            kept.append(tag)
                    return kept
                tags = led.get("tags")
                if isinstance(tags, list):
                    led["tags"] = _filter_tag_list(tags)
                keynotes = led.get("keynotes")
                if isinstance(keynotes, list):
                    led["keynotes"] = _filter_tag_list(keynotes)
    return filtered


def _expected_tag_points(repo, cad_name, parent_point, parent_rotation):
    if not repo or not cad_name or parent_point is None:
        return []
    points = []
    labels = repo.labels_for_cad(cad_name)
    if not labels:
        return points
    for label in labels:
        linked_def = repo.definition_for_label(cad_name, label)
        if not linked_def:
            continue
        placement = linked_def.get_placement()
        if not placement:
            continue
        offsets = placement.get_offset_xyz() or (0.0, 0.0, 0.0)
        ox, oy, oz = offsets
        if parent_rotation:
            try:
                ang = math.radians(float(parent_rotation))
            except Exception:
                ang = 0.0
            cos_a = math.cos(ang)
            sin_a = math.sin(ang)
            rot_x = ox * cos_a - oy * sin_a
            rot_y = ox * sin_a + oy * cos_a
            host_point = XYZ(parent_point.X + rot_x, parent_point.Y + rot_y, parent_point.Z + oz)
        else:
            host_point = XYZ(parent_point.X + ox, parent_point.Y + oy, parent_point.Z + oz)
        if host_point.Z < 0.0:
            host_point = XYZ(host_point.X, host_point.Y, 1.0)
        for tag_def in placement.get_tags() or []:
            offset = tag_def.get("offset") or (0.0, 0.0, 0.0)
            try:
                tag_point = XYZ(
                    host_point.X + (offset[0] or 0.0),
                    host_point.Y + (offset[1] or 0.0),
                    host_point.Z + (offset[2] or 0.0),
                )
            except Exception:
                tag_point = None
            if tag_point is not None:
                points.append(tag_point)
    return points


def _expected_host_points(repo, cad_name, parent_point, parent_rotation):
    if not repo or not cad_name or parent_point is None:
        return []
    points = []
    labels = repo.labels_for_cad(cad_name)
    if not labels:
        return points
    for label in labels:
        linked_def = repo.definition_for_label(cad_name, label)
        if not linked_def:
            continue
        placement = linked_def.get_placement()
        if not placement:
            continue
        offsets = placement.get_offset_xyz() or (0.0, 0.0, 0.0)
        ox, oy, oz = offsets
        if parent_rotation:
            try:
                ang = math.radians(float(parent_rotation))
            except Exception:
                ang = 0.0
            cos_a = math.cos(ang)
            sin_a = math.sin(ang)
            rot_x = ox * cos_a - oy * sin_a
            rot_y = ox * sin_a + oy * cos_a
            host_point = XYZ(parent_point.X + rot_x, parent_point.Y + rot_y, parent_point.Z + oz)
        else:
            host_point = XYZ(parent_point.X + ox, parent_point.Y + oy, parent_point.Z + oz)
        if host_point.Z < 0.0:
            host_point = XYZ(host_point.X, host_point.Y, 1.0)
        points.append(host_point)
    return points


def _collect_tag_ids_in_view(doc, active_view_id):
    if not doc or not active_view_id:
        return set()
    view_id = _ensure_element_id(active_view_id)
    if not view_id:
        return set()
    try:
        tags = list(FilteredElementCollector(doc, view_id).OfClass(IndependentTag))
    except Exception:
        tags = []
    ids = set()
    for tag in tags:
        try:
            ids.add(tag.Id.IntegerValue)
        except Exception:
            continue
    return ids


def _collect_tag_ids_near_points(doc, active_view_id, points, radius_ft):
    if not doc or not active_view_id or not points or not radius_ft:
        return set()
    view_id = _ensure_element_id(active_view_id)
    if not view_id:
        return set()
    try:
        limit = float(radius_ft)
    except Exception:
        return set()
    try:
        tags = list(FilteredElementCollector(doc, view_id).OfClass(IndependentTag))
    except Exception:
        tags = []
    ids = set()
    for tag in tags:
        try:
            tag_point = tag.TagHeadPosition
        except Exception:
            tag_point = None
        if tag_point is None:
            continue
        for point in points:
            try:
                dist = point.DistanceTo(tag_point)
            except Exception:
                try:
                    dx = point.X - tag_point.X
                    dy = point.Y - tag_point.Y
                    dz = point.Z - tag_point.Z
                    dist = math.sqrt(dx * dx + dy * dy + dz * dz)
                except Exception:
                    dist = None
            if dist is not None and dist <= limit:
                try:
                    ids.add(tag.Id.IntegerValue)
                except Exception:
                    pass
                break
    return ids


def _place_existing_configuration(doc, data, cad_name, parent_point, parent_rotation, active_view_id=None):
    if not cad_name or parent_point is None:
        return
    filtered_data = _filter_tags_for_autoload(data, cad_name, 5.0)
    repo = _build_repository(filtered_data)
    eq_def = _find_equipment_definition_by_name(data, cad_name)
    set_ids = set()
    if eq_def:
        for linked_set in eq_def.get("linked_sets") or []:
            set_id = (linked_set.get("id") or "").strip()
            if set_id:
                set_ids.add(set_id)
    labels = repo.labels_for_cad(cad_name)
    if not labels:
        return
    tag_type_keys = set()
    total_tag_defs = 0
    total_keynote_defs = 0
    for label in labels:
        linked_def = repo.definition_for_label(cad_name, label)
        if not linked_def:
            continue
        placement = linked_def.get_placement()
        if not placement:
            continue
        for tag_def in placement.get_tags() or []:
            total_tag_defs += 1
            if _is_keynote_entry(tag_def):
                total_keynote_defs += 1
            key = _tag_key_from_def(tag_def)
            if key:
                tag_type_keys.add(key)
    LOG.info(
        "[Edit-Create Profiles] Autoload tag defs: total=%s keynotes=%s labels=%s",
        total_tag_defs,
        total_keynote_defs,
        len(labels),
    )
    selection_map = {cad_name: labels}
    rows = [{
        "Name": cad_name,
        "Count": "1",
        "Position X": str(parent_point.X * 12.0),
        "Position Y": str(parent_point.Y * 12.0),
        "Position Z": str(parent_point.Z * 12.0),
        "Rotation": str(parent_rotation or 0.0),
    }]
    pre_tag_ids = _collect_tag_ids_in_view(doc, active_view_id)
    expected_points = _expected_tag_points(repo, cad_name, parent_point, parent_rotation)
    expected_hosts = _expected_host_points(repo, cad_name, parent_point, parent_rotation)
    led_ids = set()
    if eq_def:
        for linked_set in eq_def.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if isinstance(led, dict):
                    led_id = (led.get("id") or "").strip()
                    if led_id:
                        led_ids.add(led_id)
    if active_view_id and tag_type_keys and (set_ids or led_ids):
        LOG.info(
            "[Edit-Create Profiles] Existing tag cleanup scope: set_ids=%s led_ids=%s expected_points=%s expected_hosts=%s",
            len(set_ids),
            len(led_ids),
            len(expected_points),
            len(expected_hosts),
        )
        _cleanup_existing_profile_tags(
            doc,
            active_view_id,
            tag_type_keys,
            set_ids,
            led_ids,
            expected_points,
            expected_hosts,
            5.0,
        )
    engine = PlaceElementsEngine(
        doc,
        repo,
        allow_tags=True,
        allow_text_notes=True,
        max_tag_distance_feet=5.0,
        transaction_name="Load Profile for Edit/Create",
    )
    try:
        engine.place_from_csv(rows, selection_map)
        try:
            doc.Regenerate()
        except Exception:
            pass
        if active_view_id and tag_type_keys:
            post_tag_ids = _collect_tag_ids_in_view(doc, active_view_id)
            new_tag_ids = post_tag_ids - pre_tag_ids
            if new_tag_ids:
                near_expected = _collect_tag_ids_near_points(doc, active_view_id, expected_points, 10.0)
                filtered_new = set(new_tag_ids) - set(near_expected)
                if filtered_new != new_tag_ids:
                    LOG.info(
                        "[Edit-Create Profiles] Autoload tag cleanup narrowed new_tags=%s -> %s near_expected=%s",
                        len(new_tag_ids),
                        len(filtered_new),
                        len(near_expected),
                    )
                new_tag_ids = filtered_new
            LOG.info(
                "[Edit-Create Profiles] Autoload tag cleanup setup: new_tags=%s expected_points=%s tag_types=%s",
                len(new_tag_ids),
                len(expected_points),
                len(tag_type_keys),
            )
            _cleanup_far_tags_by_type(
                doc,
                active_view_id,
                None,
                10.0,
                expected_points=expected_points,
                tag_ids=new_tag_ids,
            )
    except Exception as exc:
        LOG.warning("Failed to place existing profile: %s", exc)


def _build_element_linker_payload(led_id, set_id, elem, host_point, rotation_override=None, parent_rotation=None, parent_elem_id=None):
    point = host_point or _get_point(elem)
    rot = rotation_override if rotation_override is not None else _get_rotation_degrees(elem)
    level_id = getattr(getattr(elem, "LevelId", None), "IntegerValue", None)
    elem_id = getattr(getattr(elem, "Id", None), "IntegerValue", None)
    facing = getattr(elem, "FacingOrientation", None)
    lines = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Location XYZ (ft): {:.6f},{:.6f},{:.6f}".format(point.X, point.Y, point.Z) if point else "Location XYZ (ft):",
        "Rotation (deg): {:.6f}".format(rot or 0.0),
        "Parent Rotation (deg): {}".format("{:.6f}".format(parent_rotation) if parent_rotation is not None else ""),
        "Parent ElementId: {}".format(parent_elem_id if parent_elem_id is not None else ""),
        "LevelId: {}".format(level_id if level_id is not None else ""),
        "ElementId: {}".format(elem_id if elem_id is not None else ""),
        "FacingOrientation: {}".format("{:.6f},{:.6f},{:.6f}".format(facing.X, facing.Y, facing.Z) if facing else ""),
    ]
    return "\n".join(lines)


def _set_element_linker_parameter(elem, value):
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
        })
        group["members"].append(eq_name)
        child_to_group[eq_name] = source_id
    return groups, child_to_group


def _gather_child_entries(elements, parent_point, parent_rotation, parent_elem, active_view_id=None):
    entries = []
    for elem in elements:
        if elem is None or isinstance(elem, IndependentTag) or isinstance(elem, TextNote):
            continue
        if parent_elem is not None and getattr(elem, "Id", None) == getattr(parent_elem, "Id", None):
            continue
        point = _get_point(elem)
        if point is None:
            continue
        rotation = _get_rotation_degrees(elem)
        offsets = compute_offsets_from_points(parent_point, parent_rotation, point, rotation)
        offsets["z_inches"] = _level_relative_z_inches(elem, point)
        if isinstance(elem, Group):
            offsets["rotation_deg"] = _normalize_angle(rotation - parent_rotation)
        params = _collect_params(elem)
        led_entry = {
            "element": elem,
            "point": point,
            "rotation_deg": rotation,
            "label": _build_label(elem),
            "category": _get_category(elem),
            "is_group": isinstance(elem, Group),
            "offsets": offsets,
            "parameters": params,
            "tags": [],
            "keynotes": [],
            "text_notes": [],
        }
        entries.append(led_entry)
    return entries


def _context_from_parent(parent_elem, cad_name, parent_point, parent_rotation):
    context = {
        "cad_name": cad_name,
        "parent_point": (parent_point.X, parent_point.Y, parent_point.Z),
        "parent_rotation": parent_rotation,
    }
    try:
        context["parent_element_id"] = parent_elem.Id.IntegerValue
    except Exception:
        context["parent_element_id"] = None
    return context


def _xyz_from_tuple(values):
    if not isinstance(values, (list, tuple)) or len(values) != 3:
        return XYZ(0, 0, 0)
    return XYZ(float(values[0]), float(values[1]), float(values[2]))


def _match_truth_group(name, truth_groups, data):
    if not truth_groups:
        return None
    raw = (name or "").strip()
    norm = _normalize_key(raw)
    if not norm:
        return None
    for group in truth_groups.values():
        display = (group.get("display") or "").strip()
        display_norm = _normalize_key(display)
        if display_norm != norm:
            continue
        for member in group.get("members") or []:
            eq_def = _find_equipment_definition_by_name(data, member)
            if eq_def:
                resolved_name = display or member or eq_def.get("name") or eq_def.get("id")
                return eq_def, resolved_name, False
    return None


def _resolve_equipment_definition(parent_elem, data, truth_groups=None):
    candidates = _candidate_equipment_names(parent_elem) or []
    for name in candidates:
        eq_def = _find_equipment_definition_by_name(data, name)
        if eq_def:
            resolved_name = (eq_def.get("name") or eq_def.get("id") or name).strip() or name
            return eq_def, resolved_name, False
        group_match = _match_truth_group(name, truth_groups, data)
        if group_match:
            eq_def, member_name, _ = group_match
            resolved_name = (eq_def.get("name") or eq_def.get("id") or member_name).strip() or member_name
            return eq_def, resolved_name, False
    default_name = candidates[0] if candidates else "New Profile"
    cad_name = default_name or "New Profile"
    sample_entry = {
        "label": default_name or cad_name,
        "category_name": _get_category(parent_elem),
    }
    eq_def = ensure_equipment_definition(data, cad_name, sample_entry)
    linked_set = get_type_set(eq_def)
    next_idx = _next_eq_number(data, exclude_defs=[eq_def])
    eq_id = "EQ-{:03d}".format(next_idx)
    set_id = "SET-{:03d}".format(next_idx)
    eq_def["id"] = eq_id
    linked_set["id"] = set_id
    linked_set["name"] = "{} Types".format(cad_name)
    source_id = eq_def.get("id") or cad_name
    if source_id:
        if not eq_def.get(TRUTH_SOURCE_ID_KEY):
            eq_def[TRUTH_SOURCE_ID_KEY] = source_id
        if not eq_def.get(TRUTH_SOURCE_NAME_KEY):
            eq_def[TRUTH_SOURCE_NAME_KEY] = cad_name
    return eq_def, cad_name, True


def _write_metadata_updates(updates):
    if not updates:
        return
    doc = revit.doc
    if doc is None:
        return
    txn = Transaction(doc, "Set Element_Linker metadata")
    try:
        txn.Start()
        for elem, payload in updates:
            _set_element_linker_parameter(elem, payload)
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass


def _run_selection_flow(doc, data, context, truth_groups, child_to_group, parent_elem=None):
    cad_name = context.get("cad_name")
    if not cad_name:
        forms.alert("Missing profile name for this session.", title=TITLE)
        return
    eq_def = _find_equipment_definition_by_name(data, cad_name)
    if not eq_def:
        forms.alert("Could not locate equipment definition '{}'.".format(cad_name), title=TITLE)
        _save_session(doc, None)
        return
    parent_point = _xyz_from_tuple(context.get("parent_point") or (0.0, 0.0, 0.0))
    parent_rotation = context.get("parent_rotation") or 0.0
    if parent_elem is None:
        elem_id = context.get("parent_element_id")
        if elem_id is not None:
            try:
                parent_elem = doc.GetElement(ElementId(int(elem_id)))
            except Exception:
                parent_elem = None
    parent_elem_id = None
    if parent_elem is not None:
        try:
            parent_elem_id = parent_elem.Id.IntegerValue
        except Exception:
            parent_elem_id = None
    if parent_elem_id is None:
        parent_elem_id = context.get("parent_element_id")

    trans_group = TransactionGroup(doc, TITLE)
    trans_group.Start()
    success = False
    try:
        try:
            picked = list(revit.pick_elements(message="Select equipment/tag elements for '{}'".format(cad_name)))
        except Exception:
            picked = []
        if not picked:
            forms.alert("No elements were selected.", title=TITLE)
            return

        text_note_elems = []
        tag_elems = []
        host_candidates = []
        for elem in picked:
            if isinstance(elem, TextNote):
                text_note_elems.append(elem)
                continue
            if isinstance(elem, IndependentTag):
                tag_elems.append(elem)
                continue
            host_candidates.append(elem)

        active_view = getattr(doc, "ActiveView", None)
        active_view_id = getattr(getattr(active_view, "Id", None), "IntegerValue", None)
        child_entries = _gather_child_entries(
            host_candidates,
            parent_point,
            parent_rotation,
            parent_elem,
            active_view_id=active_view_id,
        )
        if not child_entries:
            forms.alert("None of the selected elements produced valid entries.", title=TITLE)
            return
        _assign_selected_text_notes(child_entries, text_note_elems)
        _assign_selected_tags(child_entries, tag_elems)

        linked_set = get_type_set(eq_def)
        linked_set["linked_element_definitions"] = []
        metadata_updates = []
        for entry in child_entries:
            led_id = next_led_id(linked_set, eq_def)
            offsets = dict(entry["offsets"])
            entry_params = dict(entry.get("parameters") or {})
            entry_tags = list(entry.get("tags") or [])
            entry_keynotes = list(entry.get("keynotes") or [])
            entry_text_notes = list(entry.get("text_notes") or [])
            led_entry = {
                "id": led_id,
                "label": entry["label"],
                "category": entry["category"],
                "is_group": entry["is_group"],
                "offsets": [offsets],
                "parameters": entry_params,
                "tags": entry_tags,
                "keynotes": entry_keynotes,
                "text_notes": entry_text_notes,
            }
            payload = _build_element_linker_payload(
                led_id,
                linked_set.get("id") or "",
                entry["element"],
                entry["point"],
                entry["rotation_deg"],
                parent_rotation,
                parent_elem_id,
            )
            led_entry["parameters"][ELEMENT_LINKER_PARAM_NAME] = payload
            metadata_updates.append((entry["element"], payload))
            linked_set["linked_element_definitions"].append(led_entry)

        group_id = eq_def.get(TRUTH_SOURCE_ID_KEY) or child_to_group.get(cad_name) or (eq_def.get("id") or cad_name)
        display_name = truth_groups.get(group_id, {}).get("display") if truth_groups else None
        eq_def[TRUTH_SOURCE_ID_KEY] = group_id
        eq_def[TRUTH_SOURCE_NAME_KEY] = display_name or cad_name

        _write_metadata_updates(metadata_updates)

        save_active_yaml_data(
            None,
            data,
            TITLE,
            "Updated profile '{}' with {} item(s)".format(cad_name, len(child_entries)),
        )
        forms.alert(
            "Saved {} element(s) to '{}'.\nReload any tools that cache YAML."
            .format(len(child_entries), cad_name),
            title=TITLE,
        )
        _save_session(doc, None)
        success = True
    finally:
        try:
            if success:
                trans_group.Assimilate()
            else:
                trans_group.RollBack()
        except Exception:
            pass


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    truth_groups, child_to_group = _truth_group_metadata(data.get("equipment_definitions") or [])

    session = _load_session(doc)
    if session:
        cad_name = session.get("cad_name") or "<Unknown>"
        resume = forms.alert(
            "Resume editing profile '{}'?".format(cad_name),
            title=TITLE,
            yes=True,
            no=True,
        )
        if resume:
            _run_selection_flow(doc, data, session, truth_groups, child_to_group)
            return
        _save_session(doc, None)
        forms.alert("Edit/Create session cleared. Continue selecting a new parent.", title=TITLE)

    forms.alert("Select the linked parent element you want to edit or create a profile for.", title=TITLE)
    parent_elem, transform = _pick_parent_element("Select parent element")
    if not parent_elem:
        return
    parent_point = _transform_point(_get_point(parent_elem), transform)
    if parent_point is None:
        forms.alert("Unable to determine the parent's location.", title=TITLE)
        return
    parent_rotation = _transform_rotation(_get_rotation_degrees(parent_elem), transform)

    eq_def, cad_name, created_new = _resolve_equipment_definition(parent_elem, data, truth_groups)
    if not eq_def or not cad_name:
        return
    if not created_new:
        active_view = getattr(doc, "ActiveView", None)
        active_view_id = getattr(getattr(active_view, "Id", None), "IntegerValue", None)
        try:
            _place_existing_configuration(
                doc,
                data,
                cad_name,
                parent_point,
                parent_rotation,
                active_view_id=active_view_id,
            )
        except Exception as exc:
            LOG.warning("Autoload failed: %s", exc)

    context = _context_from_parent(parent_elem, cad_name, parent_point, parent_rotation)
    proceed_now = forms.alert(
        "Place or adjust the equipment using Revit tools.\n"
        "Select 'Yes' to capture the profile now, or 'No' to pause and continue later.",
        title=TITLE,
        yes=True,
        no=True,
    )
    if not proceed_now:
        _save_session(doc, context)
        forms.alert(
            "Session saved for '{}'. Use other Revit commands as needed, then run Edit/Create again to finish."
            .format(cad_name),
            title=TITLE,
        )
        return

    _run_selection_flow(doc, data, context, truth_groups, child_to_group, parent_elem)


if __name__ == "__main__":
    main()
