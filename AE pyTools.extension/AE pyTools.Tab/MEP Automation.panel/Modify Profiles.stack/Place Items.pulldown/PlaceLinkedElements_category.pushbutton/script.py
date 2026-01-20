# -*- coding: utf-8 -*-
"""Place every equipment definition on matching linked elements automatically using the active YAML store."""

import math
import os
import re
import sys

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import BuiltInParameter, Category, FamilyInstance, FamilySymbol, FilteredElementCollector, Group, GroupType, Level, RevitLinkInstance, XYZ

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Place Linked Elements (Category Filter)"
LOG = script.get_logger()

# Hardcoded MEP and Architecture categories
MEP_CATEGORIES = [
    "Air Terminals",
    "Cable Tray Fittings",
    "Cable Tray Runs",
    "Cable Trays",
    "Casework",
    "Communication Devices",
    "Conduit Fittings",
    "Conduit Runs",
    "Conduits",
    "Data Devices",
    "Duct Accessories",
    "Duct Fittings",
    "Duct Insulations",
    "Duct Linings",
    "Duct Placeholders",
    "Duct Systems",
    "Ducts",
    "Electrical Circuits",
    "Electrical Equipment",
    "Electrical Fixtures",
    "Fabrication Containment",
    "Fabrication Ductwork",
    "Fabrication Hangers",
    "Fabrication Pipework",
    "Fire Alarm Devices",
    "Fire Protection",
    "Flex Ducts",
    "Flex Pipes",
    "Furniture",
    "Furniture Systems",
    "Generic Models",
    "HVAC Zones",
    "Lighting Devices",
    "Lighting Fixtures",
    "Mechanical Control Devices",
    "Mechanical Equipment",
    "Nurse Call Devices",
    "Pipe Accessories",
    "Pipe Fittings",
    "Pipe Insulations",
    "Pipe Placeholders",
    "Pipes",
    "Piping Systems",
    "Planting",
    "Plumbing Equipment",
    "Plumbing Fixtures",
    "Security Devices",
    "Site",
    "Specialty Equipment",
    "Sprinklers",
    "Telephone Devices",
    "Temporary",
]

try:
    basestring
except NameError:
    basestring = str

TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
LEVEL_NUMBER_RE = re.compile(r"(?:^|\b)(?:level|lvl|l)\s*0*([0-9]+)\b", re.IGNORECASE)
PARENT_ID_RE = re.compile(r"parent element(?:id| id)\s*:\s*([0-9]+)", re.IGNORECASE)
PARENT_ID_KEYS = ("Parent ElementId", "Parent Element ID")


def _build_repository(data):
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def _place_requests(doc, repo, selection_map, rows, default_level=None):
    if not selection_map or not rows:
        return {"placed": 0}
    engine = PlaceElementsEngine(
        doc,
        repo,
        default_level=default_level,
        allow_tags=False,
        transaction_name="Place Linked Elements",
        apply_recorded_level=False,
    )
    return engine.place_from_csv(rows, selection_map)


def _get_category_name_from_label(doc, label):
    """Get the Revit category name from a 'Family : Type' or group label."""
    if not label or not doc:
        return None

    # Try as FamilySymbol first
    symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol))
    for sym in symbols:
        try:
            family = getattr(sym, "Family", None)
            fam_name = getattr(family, "Name", None) if family else None
            if not fam_name:
                continue
            type_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = type_param.AsString() if type_param else None
            if not type_name and hasattr(sym, "Name"):
                type_name = sym.Name
            if not type_name:
                continue
            sym_label = u"{} : {}".format(fam_name, type_name)
            if sym_label == label:
                category = sym.Category
                if category:
                    return category.Name
        except Exception:
            continue

    # Try as Group
    group_types = list(FilteredElementCollector(doc).OfClass(GroupType))
    for gtype in group_types:
        try:
            group_name = getattr(gtype, "Name", None)
            if group_name == label:
                category = gtype.Category
                if category:
                    return category.Name
        except Exception:
            continue

    return None


def _build_row(name, point, rotation_deg, parent_element_id=None, level_id=None):
    row = {
        "Name": name,
        "Count": "1",
        "Position X": str(point.X * 12.0),
        "Position Y": str(point.Y * 12.0),
        "Position Z": str(point.Z * 12.0),
        "Rotation": str(rotation_deg or 0.0)
    }
    if parent_element_id is not None:
        row["Parent ElementId"] = str(parent_element_id)
    if level_id is not None:
        row["LevelId"] = str(level_id)
    return row


def _placement_key(point, rotation_deg, level_id, parent_element_id):
    if point is None:
        return None
    try:
        x = round(float(point.X), 6)
        y = round(float(point.Y), 6)
        z = round(float(point.Z), 6)
    except Exception:
        return None
    try:
        rot = round(float(rotation_deg or 0.0), 3)
    except Exception:
        rot = 0.0
    return (level_id, parent_element_id, x, y, z, rot)


def _normalize_name(value):
    if not value:
        return ""
    normalized = " ".join(str(value).strip().lower().split())
    return normalized


def _normalize_level_name(value):
    if not value:
        return ""
    normalized = " ".join(str(value).strip().lower().split())
    return normalized


def _compact_level_text(value):
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def _extract_level_number(value):
    if not value:
        return None
    match = LEVEL_NUMBER_RE.search(str(value))
    if match:
        try:
            return int(match.group(1))
        except Exception:
            return None
    return None


def _score_level_match(link_norm, host_norm):
    if not link_norm or not host_norm:
        return 0
    if host_norm == link_norm:
        return 100
    score = 0
    if host_norm in link_norm:
        score = max(score, 80 + len(host_norm))
    if link_norm in host_norm:
        score = max(score, 70 + len(link_norm))
    link_compact = _compact_level_text(link_norm)
    host_compact = _compact_level_text(host_norm)
    if link_compact and host_compact:
        if link_compact == host_compact:
            score = max(score, 90)
        if host_compact in link_compact or link_compact in host_compact:
            score = max(score, 60 + min(len(link_compact), len(host_compact)))
    link_num = _extract_level_number(link_norm)
    host_num = _extract_level_number(host_norm)
    if link_num and host_num and link_num == host_num:
        score = max(score, 50)
    return score


def _find_level_one(doc):
    if doc is None:
        return None
    targets = {"level1", "level01", "lvl1", "l1"}
    try:
        levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        levels = []
    for level in levels:
        name = getattr(level, "Name", None)
        key = _compact_level_text(name)
        if key in targets:
            return level
    return None


def _find_closest_level(levels, z_value):
    if not levels or z_value is None:
        return None
    closest = None
    closest_dist = None
    for level in levels:
        try:
            elev = float(level.Elevation)
        except Exception:
            continue
        dist = abs(elev - z_value)
        if closest_dist is None or dist < closest_dist:
            closest = level
            closest_dist = dist
    if closest is None:
        return None
    name = getattr(closest, "Name", None)
    if name and "xx - legend" in str(name).lower():
        above = []
        try:
            closest_elev = float(closest.Elevation)
        except Exception:
            return closest
        for level in levels:
            try:
                elev = float(level.Elevation)
            except Exception:
                continue
            if elev > closest_elev:
                above.append((elev, level))
        if above:
            above.sort(key=lambda item: item[0])
            return above[0][1]
    return closest


def _get_element_level(elem):
    if elem is None:
        return None
    doc = getattr(elem, "Document", None)
    level_id = getattr(elem, "LevelId", None)
    if level_id and doc:
        try:
            if getattr(level_id, "IntegerValue", None) not in (None, -1):
                level = doc.GetElement(level_id)
                if level is not None:
                    return level
        except Exception:
            pass
    if not hasattr(elem, "get_Parameter"):
        return None
    level_param_names = (
        "FAMILY_LEVEL_PARAM",
        "INSTANCE_LEVEL_PARAM",
        "INSTANCE_REFERENCE_LEVEL_PARAM",
        "SCHEDULE_LEVEL_PARAM",
        "LEVEL_PARAM",
        "WALL_BASE_CONSTRAINT",
    )
    for param_name in level_param_names:
        bip = getattr(BuiltInParameter, param_name, None)
        if bip is None:
            continue
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if not param:
            continue
        try:
            level_id = param.AsElementId()
        except Exception:
            level_id = None
        if level_id and doc:
            try:
                level = doc.GetElement(level_id)
            except Exception:
                level = None
            if level is not None:
                return level
    return None


def _get_parent_element(elem):
    if elem is None:
        return None
    parent = None
    if isinstance(elem, FamilyInstance):
        try:
            parent = getattr(elem, "SuperComponent", None)
        except Exception:
            parent = None
        if parent is None:
            try:
                parent = getattr(elem, "Host", None)
            except Exception:
                parent = None
    return parent


def _get_parent_level(elem):
    parent = _get_parent_element(elem)
    if parent is None:
        return None
    return _get_element_level(parent)


def _get_level_name_from_param(elem):
    if elem is None or not hasattr(elem, "get_Parameter"):
        return None
    level_param_names = (
        "FAMILY_LEVEL_PARAM",
        "INSTANCE_LEVEL_PARAM",
        "INSTANCE_REFERENCE_LEVEL_PARAM",
        "SCHEDULE_LEVEL_PARAM",
        "LEVEL_PARAM",
        "WALL_BASE_CONSTRAINT",
    )
    for param_name in level_param_names:
        bip = getattr(BuiltInParameter, param_name, None)
        if bip is None:
            continue
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if not param or not param.HasValue:
            continue
        value = None
        try:
            value = param.AsValueString()
        except Exception:
            value = None
        if not value:
            try:
                value = param.AsString()
            except Exception:
                value = None
        if value:
            text = str(value).strip()
            if text:
                return text
    for name in ("Level", "Reference Level", "Schedule Level", "Base Constraint"):
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if param and param.HasValue:
            value = None
            try:
                value = param.AsString()
            except Exception:
                value = None
            if not value:
                try:
                    value = param.AsValueString()
                except Exception:
                    value = None
            if value:
                text = str(value).strip()
                if text:
                    return text
    return None


def _prefer_parent_level(level, parent_level):
    if parent_level is None:
        return level
    if level is None:
        return parent_level
    parent_name = getattr(parent_level, "Name", None)
    level_name = getattr(level, "Name", None)
    parent_num = _extract_level_number(parent_name)
    level_num = _extract_level_number(level_name)
    if parent_num and (level_num is None or parent_num != level_num):
        return parent_level
    return level


def _collect_host_levels(doc):
    by_name = {}
    try:
        levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        levels = []
    for level in levels:
        name = getattr(level, "Name", None)
        if not name:
            continue
        norm = _normalize_level_name(name)
        if norm:
            by_name.setdefault(norm, []).append(level)
    return by_name


def _resolve_host_level_name(level_name, host_level_by_name):
    if not level_name or not host_level_by_name:
        return None
    norm = _normalize_level_name(level_name)
    candidates = host_level_by_name.get(norm)
    if candidates:
        return candidates[0]
    link_num = _extract_level_number(norm)
    candidate_map = host_level_by_name
    if link_num is not None:
        filtered = {}
        for key, levels in host_level_by_name.items():
            if not key or not levels:
                continue
            if _extract_level_number(key) == link_num:
                filtered[key] = levels
        if filtered:
            candidate_map = filtered
    best = None
    best_score = 0
    for key, levels in candidate_map.items():
        if not key or not levels:
            continue
        score = _score_level_match(norm, key)
        if score > best_score:
            best_score = score
            best = levels[0]
    return best


def _resolve_host_level(link_level, host_level_by_name):
    if not link_level or not host_level_by_name:
        return None
    name = getattr(link_level, "Name", None)
    if not name:
        return None
    return _resolve_host_level_name(name, host_level_by_name)


def _get_symbol(elem):
    symbol = getattr(elem, "Symbol", None)
    if symbol is not None:
        return symbol
    try:
        type_id = elem.GetTypeId()
    except Exception:
        type_id = None
    if type_id:
        doc = getattr(elem, "Document", None)
        if doc:
            try:
                return doc.GetElement(type_id)
            except Exception:
                return None
    return None


def _name_variants(elem):
    names = set()
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            names.add(raw_name)
    except Exception:
        pass
    if isinstance(elem, FamilyInstance):
        symbol = _get_symbol(elem)
        family = getattr(symbol, "Family", None) if symbol else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        if not family_name or not type_name:
            try:
                fam_param = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
            except Exception:
                fam_param = None
            if fam_param and not family_name:
                try:
                    fam_id = fam_param.AsElementId()
                except Exception:
                    fam_id = None
                if fam_id:
                    doc = getattr(elem, "Document", None)
                    if doc:
                        try:
                            fam_elem = doc.GetElement(fam_id)
                        except Exception:
                            fam_elem = None
                        if fam_elem is not None:
                            family_name = getattr(fam_elem, "Name", None)
                if not family_name:
                    try:
                        family_name = fam_param.AsValueString()
                    except Exception:
                        family_name = None
            try:
                type_param = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
            except Exception:
                type_param = None
            if type_param and not type_name:
                try:
                    type_id = type_param.AsElementId()
                except Exception:
                    type_id = None
                if type_id:
                    doc = getattr(elem, "Document", None)
                    if doc:
                        try:
                            type_elem = doc.GetElement(type_id)
                        except Exception:
                            type_elem = None
                        if type_elem is not None:
                            type_name = getattr(type_elem, "Name", None)
                if not type_name:
                    try:
                        type_name = type_param.AsValueString()
                    except Exception:
                        type_name = None
        if family_name and type_name:
            names.add(u"{} : {}".format(family_name, type_name))
            names.add(u"{} : {}".format(type_name, family_name))
        if type_name:
            names.add(type_name)
        if family_name:
            names.add(family_name)
    elif isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        group_name = getattr(gtype, "Name", None) if gtype else None
        if group_name:
            names.add(group_name)
    return {_normalize_name(name) for name in names if _normalize_name(name)}


def _parse_payload_pose(payload_text):
    if not payload_text:
        return None
    text = str(payload_text)
    entries = {}
    if "\n" in text:
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, remainder = line.partition(":")
            entries[key.strip().lower()] = remainder.strip()
    else:
        pattern = re.compile(
            r"(Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
            r"Parent ElementId|Parent Element ID|LevelId)\s*:\s*",
            re.IGNORECASE,
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1).strip().lower()
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")

    location = None
    rotation = None
    parent_rotation = None
    parent_element_id = None
    level_id = None
    for key, value in entries.items():
        if key.startswith("location xyz"):
            parts = [p.strip() for p in value.split(",")]
            if len(parts) == 3:
                try:
                    location = tuple(float(p) for p in parts)
                except Exception:
                    location = None
        elif key.startswith("parent rotation"):
            try:
                parent_rotation = float(value)
            except Exception:
                parent_rotation = None
        elif key.startswith("parent elementid") or key.startswith("parent element id"):
            try:
                parent_element_id = int(value)
            except Exception:
                parent_element_id = None
        elif key.startswith("levelid"):
            try:
                level_id = int(value)
            except Exception:
                try:
                    level_id = int(float(value))
                except Exception:
                    level_id = None
        elif key.startswith("rotation"):
            try:
                rotation = float(value)
            except Exception:
                rotation = None
    if not location:
        return None
    point = XYZ(location[0], location[1], location[2])
    final_rotation = parent_rotation if parent_rotation is not None else (rotation or 0.0)
    return {
        "point": point,
        "rotation": final_rotation,
        "parent_element_id": parent_element_id,
        "level_id": level_id,
    }


def _parent_id_from_params(params):
    if not params:
        return None
    for key in PARENT_ID_KEYS:
        if key in params:
            try:
                return int(params.get(key))
            except Exception:
                try:
                    return int(float(params.get(key)))
                except Exception:
                    return None
    for key in ("Element_Linker Parameter", "Element_Linker"):
        payload = params.get(key)
        if not payload:
            continue
        match = PARENT_ID_RE.search(str(payload))
        if match:
            try:
                return int(match.group(1))
            except Exception:
                return None
    return None


def _parent_id_from_defs(repo, cad_name, labels):
    for label in labels or []:
        linked_def = repo.definition_for_label(cad_name, label)
        if not linked_def:
            continue
        params = linked_def.get_static_params() or {}
        parent_id = _parent_id_from_params(params)
        if parent_id is not None:
            return parent_id
    return None


def _get_element_point(elem):
    location = getattr(elem, "Location", None)
    if not location:
        return None
    point = getattr(location, "Point", None)
    if point:
        return point
    curve = getattr(location, "Curve", None)
    if curve:
        try:
            return curve.Evaluate(0.5, True)
        except Exception:
            try:
                return curve.GetEndPoint(0)
            except Exception:
                return None
    return None


def _get_orientation_vector(elem):
    location = getattr(elem, "Location", None)
    if location and hasattr(location, "Rotation"):
        try:
            angle = float(location.Rotation)
            return XYZ(math.cos(angle), math.sin(angle), 0.0)
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


def _angle_from_vector(vec):
    if not vec:
        return 0.0
    try:
        return math.degrees(math.atan2(vec.Y, vec.X))
    except Exception:
        return 0.0


def _transform_point(transform, point):
    if point is None or transform is None:
        return point
    try:
        return transform.OfPoint(point)
    except Exception:
        return point


def _get_link_transform(link_inst):
    if link_inst is None:
        return None
    try:
        return link_inst.GetTotalTransform()
    except Exception:
        try:
            return link_inst.GetTransform()
        except Exception:
            return None


def _combine_transform(parent_transform, child_transform):
    if parent_transform is None:
        return child_transform
    if child_transform is None:
        return parent_transform
    try:
        return parent_transform.Multiply(child_transform)
    except Exception:
        return None


def _doc_key(doc):
    if doc is None:
        return None
    try:
        path = doc.PathName
    except Exception:
        path = None
    if not path:
        try:
            path = doc.Title
        except Exception:
            path = None
    if not path:
        try:
            path = str(doc.GetHashCode())
        except Exception:
            path = None
    return path


def _walk_link_documents(doc, parent_transform, doc_chain):
    if doc is None:
        return
    doc_key = _doc_key(doc)
    if doc_key and doc_key in doc_chain:
        return
    next_chain = set(doc_chain or set())
    if doc_key:
        next_chain.add(doc_key)
    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        transform = _combine_transform(parent_transform, _get_link_transform(link_inst))
        yield link_doc, transform
        for nested in _walk_link_documents(link_doc, transform, next_chain):
            yield nested


def _iter_link_documents(doc):
    for link_doc, transform in _walk_link_documents(doc, None, set()):
        yield link_doc, transform


def _collect_placeholders(doc, normalized_targets):
    placements = {}
    if not normalized_targets:
        return placements
    host_level_by_name = _collect_host_levels(doc)

    def _store(name, point, rotation, parent_element_id, level=None, level_name=None):
        if not name or point is None:
            return
        placements.setdefault(name, []).append({
            "point": point,
            "rotation": rotation,
            "parent_element_id": parent_element_id,
            "level": level,
            "level_name": level_name,
        })

    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for elem in collector:
        if not isinstance(elem, (FamilyInstance, Group)):
            continue
        variants = _name_variants(elem)
        if not variants:
            continue
        if not any(name in normalized_targets for name in variants):
            continue
        point = _get_element_point(elem)
        if point is None:
            continue
        rotation = _angle_from_vector(_get_orientation_vector(elem))
        level = _get_element_level(elem)
        level_name = getattr(level, "Name", None) if level else None
        try:
            parent_element_id = elem.Id.IntegerValue
        except Exception:
            parent_element_id = None
        for name in variants:
            if name in normalized_targets:
                _store(name, point, rotation, parent_element_id, level, level_name)

    for link_doc, transform in _iter_link_documents(doc):
        linked_instances = FilteredElementCollector(link_doc).OfClass(FamilyInstance)
        for inst in linked_instances:
            variants = _name_variants(inst)
            if not variants:
                continue
            if not any(name in normalized_targets for name in variants):
                continue
            point = _transform_point(transform, _get_element_point(inst))
            if point is None:
                continue
            vec = _get_orientation_vector(inst)
            if vec and transform is not None:
                try:
                    vec = transform.OfVector(vec)
                except Exception:
                    pass
            rotation = _angle_from_vector(vec)
            parent = _get_parent_element(inst)
            parent_level = _get_element_level(parent)
            link_level = _prefer_parent_level(_get_element_level(inst), parent_level)
            level_name = (
                (getattr(parent_level, "Name", None) if parent_level else None)
                or _get_level_name_from_param(parent)
                or _get_level_name_from_param(inst)
                or (getattr(link_level, "Name", None) if link_level else None)
            )
            host_level = _resolve_host_level_name(level_name, host_level_by_name) or _resolve_host_level(
                link_level, host_level_by_name
            )
            try:
                parent_element_id = inst.Id.IntegerValue
            except Exception:
                parent_element_id = None
            for name in variants:
                if name in normalized_targets:
                    _store(name, point, rotation, parent_element_id, host_level, level_name)

    return placements


def _anchor_rows_for_cad(repo, cad_name):
    anchors = []
    if hasattr(repo, "anchor_definitions_for_cad"):
        try:
            anchors = repo.anchor_definitions_for_cad(cad_name)
        except Exception:
            anchors = []
    rows = []
    for anchor in anchors or []:
        params = anchor.get_static_params() or {}
        payload = None
        for key in ("Element_Linker Parameter", "Element_Linker"):
            value = params.get(key)
            if value:
                payload = value
                break
        pose = _parse_payload_pose(payload)
        if pose:
            rows.append({
                "point": pose["point"],
                "rotation": pose["rotation"],
                "parent_element_id": pose.get("parent_element_id"),
                "level_id": pose.get("level_id"),
            })
    return rows


def _truth_group_maps(equipment_defs):
    child_to_group = {}
    group_display = {}
    for entry in equipment_defs or []:
        name = (entry.get("name") or entry.get("id") or "").strip()
        if not name:
            continue
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id:
            source_id = (entry.get("id") or name).strip()
        display = (entry.get(TRUTH_SOURCE_NAME_KEY) or name).strip()
        child_to_group[name] = source_id or name
        if source_id:
            group_display.setdefault(source_id, display or source_id)
    return child_to_group, group_display


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
    repo = _build_repository(data)

    equipment_names = repo.cad_names()
    if not equipment_names:
        forms.alert("No equipment definitions found in {}.".format(yaml_label), title=TITLE)
        return

    target_names = {_normalize_name(name) for name in equipment_names if _normalize_name(name)}
    placeholders = _collect_placeholders(doc, target_names)
    if not placeholders:
        forms.alert(
            "No linked elements were found whose names match equipment definitions. "
            "Verify the Revit model contains linked elements with matching names.",
            title=TITLE,
        )
        return

    # Build category mapping from equipment definitions
    # Map equipment name -> set of categories (parent + all child linked element categories)
    eq_defs = data.get("equipment_definitions") or []
    name_to_categories = {}  # Changed to plural - can have multiple categories
    for eq_def in eq_defs:
        name = eq_def.get("name") or eq_def.get("id")
        if not name:
            continue

        categories = set()

        # Add parent category
        parent_filter = eq_def.get("parent_filter")
        if parent_filter:
            parent_cat = parent_filter.get("category") or parent_filter.get("Category")
            if parent_cat:
                categories.add(parent_cat)

        # Add all linked element categories
        linked_sets = eq_def.get("linked_sets") or []
        for linked_set in linked_sets:
            linked_defs = linked_set.get("linked_element_definitions") or []
            for linked_def in linked_defs:
                child_cat = linked_def.get("category") or linked_def.get("Category")
                if child_cat:
                    categories.add(child_cat)

        if categories:
            name_to_categories[name] = categories

    # Show category selection dialog if we have category data
    selected_categories = None
    if name_to_categories:
        selected_categories = forms.SelectFromList.show(
            MEP_CATEGORIES,
            title="Select Categories to Place",
            multiselect=True,
            button_name="Place Selected"
        )
        if not selected_categories:
            return  # User cancelled

        # Convert to set for faster lookup
        selected_categories = set(selected_categories)

        LOG.info("="*60)
        LOG.info("CATEGORY FILTER ACTIVE")
        LOG.info("Selected categories: {}".format(selected_categories))
        LOG.info("Total equipment definitions: {}".format(len(equipment_names)))
        LOG.info("Equipment with category data: {}".format(len(name_to_categories)))
        LOG.info("="*60)

    child_to_group, _ = _truth_group_maps(data.get("equipment_definitions") or [])
    successful_groups = set()
    rows = []
    selection_map = {}
    placed_defs = set()
    missing_labels = []
    missing_levels = []
    fallback_levels = []
    nearest_level_fallbacks = []
    deduped_rows = 0
    seen_rows = set()
    default_level = _find_level_one(doc)
    default_level_id = None
    if default_level is not None:
        try:
            default_level_id = default_level.Id.IntegerValue
        except Exception:
            default_level_id = None
    try:
        host_levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        host_levels = []

    skipped_by_category = 0
    for cad_name in equipment_names:
        normalized = _normalize_name(cad_name)
        matches = placeholders.get(normalized)
        if not matches:
            matches = _anchor_rows_for_cad(repo, cad_name)
        if not matches:
            continue
        labels = repo.labels_for_cad(cad_name)
        if not labels:
            missing_labels.append(cad_name)
            continue

        # Filter labels by category if selected
        if selected_categories:
            filtered_labels = []
            for label in labels:
                linked_def = repo.definition_for_label(cad_name, label)
                yaml_cat = linked_def.get_category() if linked_def else None
                cat_name = yaml_cat or _get_category_name_from_label(doc, label)
                if cat_name and cat_name in selected_categories:
                    filtered_labels.append(label)
                    LOG.debug("Including label '{}' - category '{}' in selected".format(label, cat_name))
                else:
                    skipped_by_category += 1
                    LOG.debug("Skipping label '{}' - category '{}' not in selected {}".format(
                        label, cat_name, selected_categories
                    ))
            labels = filtered_labels

        if not labels:
            continue  # All labels filtered out

        selection_map[cad_name] = labels
        parent_id_from_yaml = _parent_id_from_defs(repo, cad_name, labels)
        group_key = child_to_group.get(cad_name) or cad_name
        any_row = False
        for match in matches:
            point = match.get("point")
            rotation = match.get("rotation")
            if point is None:
                continue
            level_id_val = match.get("level_id")
            if level_id_val is None:
                level = match.get("level")
                if level is not None:
                    try:
                        level_id_val = level.Id.IntegerValue
                    except Exception:
                        level_id_val = None
            if level_id_val is None:
                level_name = match.get("level_name")
                if default_level_id is not None:
                    level_id_val = default_level_id
                    if level_name:
                        fallback_levels.append((cad_name, level_name))
                else:
                    nearest = _find_closest_level(host_levels, point.Z if point else None)
                    if nearest is not None:
                        try:
                            level_id_val = nearest.Id.IntegerValue
                        except Exception:
                            level_id_val = None
                        if level_id_val is not None:
                            nearest_level_fallbacks.append((cad_name, getattr(nearest, "Name", None)))
                    if level_id_val is None:
                        if level_name:
                            missing_levels.append((cad_name, level_name))
                        else:
                            missing_levels.append((cad_name, "(no level info)"))
                        continue
            parent_id = parent_id_from_yaml if parent_id_from_yaml is not None else match.get("parent_element_id")
            row_key = _placement_key(point, rotation, level_id_val, parent_id)
            if row_key in seen_rows:
                deduped_rows += 1
                continue
            if row_key is not None:
                seen_rows.add(row_key)
            rows.append(_build_row(cad_name, point, rotation, parent_id, level_id_val))
            any_row = True
        if any_row:
            placed_defs.add(cad_name)
            if group_key:
                successful_groups.add(group_key)

    if not rows:
        summary = [
            "No placements were generated. Ensure equipment definitions include linked types "
            "and that matching linked elements exist in the model.",
        ]
        if selected_categories:
            summary.append("")
            summary.append("Selected categories: {}".format(", ".join(sorted(selected_categories))))
            summary.append("Equipment found in selected categories: {}".format(
                len([n for n in equipment_names if name_to_categories.get(n, set()).intersection(selected_categories)])
            ))
        if not default_level_id:
            summary.append("")
            summary.append("Level 1 not found in host model; could not apply default placement level.")
        if missing_levels:
            summary.append("")
            summary.append("Skipped (no matching host level):")
            seen = set()
            for name, level_name in missing_levels:
                label = "{} -> {}".format(name, level_name)
                seen.add(label)
            for item in sorted(seen):
                summary.append(" - {}".format(item))
        if deduped_rows:
            summary.append("")
            summary.append("Skipped {} duplicate placement row(s).".format(deduped_rows))
        forms.alert("\n".join(summary), title=TITLE)
        return

    results = _place_requests(doc, repo, selection_map, rows, default_level=None)
    placed = results.get("placed", 0)
    summary = [
        "Processed {} linked element(s).".format(len(rows)),
        "Placed {} element(s).".format(placed),
    ]
    if selected_categories:
        summary.append("")
        summary.append("Selected categories: {}".format(", ".join(sorted(selected_categories))))
        summary.append("Equipment skipped by category filter: {}".format(skipped_by_category))

    unmatched_defs = []
    for name in equipment_names:
        if name in placed_defs:
            continue
        group_key = child_to_group.get(name)
        if group_key and group_key in successful_groups:
            continue
        unmatched_defs.append(name)
    unmatched_defs.sort()
    if unmatched_defs:
        summary.append("")
        summary.append("No matching linked elements for:")
        summary.extend(" - {}".format(name) for name in unmatched_defs)

    if missing_labels:
        summary.append("")
        summary.append("Definitions missing linked types:")
        sample = missing_labels[:5]
        summary.extend(" - {}".format(name) for name in sample)
        if len(missing_labels) > len(sample):
            summary.append("   (+{} more)".format(len(missing_labels) - len(sample)))

    if missing_levels:
        summary.append("")
        summary.append("Skipped (no matching host level):")
        seen = set()
        for name, level_name in missing_levels:
            label = "{} -> {}".format(name, level_name)
            if label in seen:
                continue
            seen.add(label)
        for item in sorted(seen):
            summary.append(" - {}".format(item))

    if fallback_levels:
        summary.append("")
        summary.append("Defaulted to Level 1 (no matching host level): {}".format(len(fallback_levels)))
    if nearest_level_fallbacks:
        summary.append("")
        summary.append("Defaulted to nearest host level by elevation: {}".format(len(nearest_level_fallbacks)))

    if deduped_rows:
        summary.append("")
        summary.append("Skipped {} duplicate placement row(s).".format(deduped_rows))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
