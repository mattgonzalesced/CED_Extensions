# -*- coding: utf-8 -*-
"""
Edit/Create YAML Profiles
-------------------------
Select a linked parent element, optionally load any existing equipment profile,
then capture offsets/tags for the selected equipment to update the
active YAML definition.
"""

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

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from LogicClasses.linked_equipment import compute_offsets_from_points  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402
from profile_schema import ensure_equipment_definition, get_type_set, next_led_id, equipment_defs_to_legacy  # noqa: E402

TITLE = "Edit/Create YAML Profiles"
LOG = script.get_logger()
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", ELEMENT_LINKER_PARAM_NAME)
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
SESSION_KEY = "edit_create_yaml_session"


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
    doc = getattr(elem, "Document", None)
    level = None
    level_id = getattr(elem, "LevelId", None)
    if doc and level_id:
        try:
            level = doc.GetElement(level_id)
        except Exception:
            level = None
    if not level:
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


# --------------------------------------------------------------------------- #
# YAML helpers
# --------------------------------------------------------------------------- #


def _collect_params(elem):
    targets = {
        "dev-Group ID": ["dev-Group ID", "dev_Group ID"],
        "Number of Poles_CED": ["Number of Poles_CED", "Number of Poles_CEDT"],
        "Apparent Load Input_CED": ["Apparent Load Input_CED", "Apparent Load Input_CEDT"],
        "Voltage_CED": ["Voltage_CED", "Voltage_CEDT"],
        "CKT_Rating_CED": ["CKT_Rating_CED"],
        "CKT_Panel_CEDT": ["CKT_Panel_CED", "CKT_Panel_CEDT"],
        "CKT_Schedule Notes_CEDT": ["CKT_Schedule Notes_CED", "CKT_Schedule Notes_CEDT"],
        "CKT_Circuit Number_CEDT": ["CKT_Circuit Number_CED", "CKT_Circuit Number_CEDT"],
        "CKT_Load Name_CEDT": ["CKT_Load Name_CED", "CKT_Load Name_CEDT"],
    }
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
            if storage and storage.ToString() == "String":
                found[target] = param.AsString() or ""
            elif storage and storage.ToString() == "Double":
                found[target] = param.AsDouble()
            elif storage and storage.ToString() == "Integer":
                found[target] = param.AsInteger()
            else:
                found[target] = param.AsValueString() or ""
        except Exception:
            continue
    return found


def _tag_signature(name):
    return " ".join((name or "").strip().split()).lower()


def _collect_hosted_tags(elem, host_point):
    doc = getattr(elem, "Document", None)
    if not doc or host_point is None:
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
        if tag_point is None:
            continue
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
        offsets = {
            "x_inches": _feet_to_inches((tag_point.X - host_point.X)),
            "y_inches": _feet_to_inches((tag_point.Y - host_point.Y)),
            "z_inches": _feet_to_inches((tag_point.Z - host_point.Z)),
            "rotation_deg": 0.0,
        }
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
        tags.append({
            "family_name": family_field,
            "type_name": type_field,
            "category_name": category_name,
            "parameters": {},
            "offsets": offsets,
        })
    return tags


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


def _place_existing_configuration(doc, data, cad_name, parent_point, parent_rotation):
    if not cad_name or parent_point is None:
        return
    repo = _build_repository(data)
    labels = repo.labels_for_cad(cad_name)
    if not labels:
        return
    selection_map = {cad_name: labels}
    rows = [{
        "Name": cad_name,
        "Count": "1",
        "Position X": str(parent_point.X * 12.0),
        "Position Y": str(parent_point.Y * 12.0),
        "Position Z": str(parent_point.Z * 12.0),
        "Rotation": str(parent_rotation or 0.0),
    }]
    engine = PlaceElementsEngine(
        doc,
        repo,
        allow_tags=True,
        transaction_name="Load Profile for Edit/Create",
    )
    try:
        engine.place_from_csv(rows, selection_map)
    except Exception as exc:
        LOG.warning("Failed to place existing profile: %s", exc)


def _build_element_linker_payload(led_id, set_id, elem, host_point, rotation_override=None, parent_rotation=None):
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


def _gather_child_entries(elements, parent_point, parent_rotation, parent_elem):
    entries = []
    for elem in elements:
        if elem is None or isinstance(elem, IndependentTag):
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
        led_entry = {
            "element": elem,
            "point": point,
            "rotation_deg": rotation,
            "label": _build_label(elem),
            "category": _get_category(elem),
            "is_group": isinstance(elem, Group),
            "offsets": offsets,
            "parameters": _collect_params(elem),
            "tags": _collect_hosted_tags(elem, point),
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
    cad_name = forms.ask_for_string(
        prompt="Enter a name for this equipment profile",
        title=TITLE,
        default=default_name,
    )
    if not cad_name:
        return None, None, False
    sample_entry = {
        "label": default_name or cad_name,
        "category_name": _get_category(parent_elem),
    }
    eq_def = ensure_equipment_definition(data, cad_name, sample_entry)
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

        child_entries = _gather_child_entries(picked, parent_point, parent_rotation, parent_elem)
        if not child_entries:
            forms.alert("None of the selected elements produced valid entries.", title=TITLE)
            return

        linked_set = get_type_set(eq_def)
        linked_set["linked_element_definitions"] = []
        metadata_updates = []
        for entry in child_entries:
            led_id = next_led_id(linked_set, eq_def)
            offsets = dict(entry["offsets"])
            led_entry = {
                "id": led_id,
                "label": entry["label"],
                "category": entry["category"],
                "is_group": entry["is_group"],
                "offsets": [offsets],
                "parameters": dict(entry["parameters"]),
                "tags": entry["tags"],
            }
            payload = _build_element_linker_payload(
                led_id,
                linked_set.get("id") or "",
                entry["element"],
                entry["point"],
                entry["rotation_deg"],
                parent_rotation,
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
        try:
            _place_existing_configuration(doc, data, cad_name, parent_point, parent_rotation)
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
