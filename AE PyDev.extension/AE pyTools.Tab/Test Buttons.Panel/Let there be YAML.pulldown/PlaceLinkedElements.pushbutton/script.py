# -*- coding: utf-8 -*-
"""Place every equipment definition on matching linked elements automatically."""

import math
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import FamilyInstance, FilteredElementCollector, Group, RevitLinkInstance, XYZ

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from profile_schema import equipment_defs_to_legacy, load_data as load_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

TITLE = "Place Linked Elements"

try:
    basestring
except NameError:
    basestring = str


def _pick_profile_data_path():
    cached = get_cached_yaml_path()
    if cached and os.path.exists(cached):
        return cached
    path = forms.pick_file(file_ext="yaml", title="Select profileData YAML file", init_dir=os.path.dirname(os.path.join(LIB_ROOT, "profileData.yaml")))
    if path:
        set_cached_yaml_path(path)
    return path


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
                    result.append(remainder)
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
                    result[key] = remainder
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result[key] = value
        if result is None:
            result = {}
        return result, idx

    parsed, _ = parse_block(0, 0)
    return parsed if isinstance(parsed, dict) else {}


def _load_profile_store(data_path):
    data = load_profile_data(data_path)
    if data.get("equipment_definitions"):
        return data
    try:
        with open(data_path, "r", encoding="utf-8") as handle:
            fallback = _simple_yaml_parse(handle.read())
        if fallback.get("equipment_definitions"):
            return fallback
    except Exception:
        pass
    return data


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
    )
    return engine.place_from_csv(rows, selection_map)


def _build_row(name, point, rotation_deg):
    return {
        "Name": name,
        "Count": "1",
        "Position X": str(point.X * 12.0),
        "Position Y": str(point.Y * 12.0),
        "Position Z": str(point.Z * 12.0),
        "Rotation": str(rotation_deg or 0.0)
    }


def _normalize_name(value):
    if not value:
        return ""
    normalized = " ".join(str(value).strip().lower().split())
    return normalized


def _name_variants(elem):
    names = set()
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            names.add(raw_name)
    except Exception:
        pass
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        if family_name and type_name:
            names.add(u"{} : {}".format(family_name, type_name))
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
    try:
        facing = getattr(elem, "FacingOrientation", None)
        if facing:
            return facing
    except Exception:
        pass
    location = getattr(elem, "Location", None)
    if location and hasattr(location, "Rotation"):
        try:
            angle = float(location.Rotation)
            return XYZ(math.cos(angle), math.sin(angle), 0.0)
        except Exception:
            return None
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


def _collect_placeholders(doc, normalized_targets):
    placements = {}
    if not normalized_targets:
        return placements

    def _store(name, point, rotation):
        if not name or point is None:
            return
        placements.setdefault(name, []).append({
            "point": point,
            "rotation": rotation,
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
        for name in variants:
            if name in normalized_targets:
                _store(name, point, rotation)

    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        try:
            transform = link_inst.GetTotalTransform()
        except Exception:
            try:
                transform = link_inst.GetTransform()
            except Exception:
                transform = None
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
            for name in variants:
                if name in normalized_targets:
                    _store(name, point, rotation)

    return placements


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    data_path = _pick_profile_data_path()
    if not data_path:
        return
    data = _load_profile_store(data_path)
    repo = _build_repository(data)

    equipment_names = repo.cad_names()
    if not equipment_names:
        forms.alert("No equipment definitions found in the selected YAML.", title=TITLE)
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

    rows = []
    selection_map = {}
    placed_defs = set()
    missing_labels = []
    for cad_name in equipment_names:
        normalized = _normalize_name(cad_name)
        matches = placeholders.get(normalized)
        if not matches:
            continue
        labels = repo.labels_for_cad(cad_name)
        if not labels:
            missing_labels.append(cad_name)
            continue
        selection_map[cad_name] = labels
        placed_defs.add(cad_name)
        for match in matches:
            point = match.get("point")
            rotation = match.get("rotation")
            if point is None:
                continue
            rows.append(_build_row(cad_name, point, rotation))

    if not rows:
        forms.alert(
            "No placements were generated. Ensure equipment definitions include linked types "
            "and that matching linked elements exist in the model.",
            title=TITLE,
        )
        return

    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    elif level_sel:
        level = level_sel

    results = _place_requests(doc, repo, selection_map, rows, default_level=level)
    placed = results.get("placed", 0)
    summary = [
        "Processed {} linked element(s).".format(len(rows)),
        "Placed {} element(s).".format(placed),
    ]

    unmatched_defs = sorted(name for name in equipment_names if name not in placed_defs)
    if unmatched_defs:
        summary.append("")
        summary.append("No matching linked elements for:")
        sample = unmatched_defs[:5]
        summary.extend(" - {}".format(name) for name in sample)
        if len(unmatched_defs) > len(sample):
            summary.append("   (+{} more)".format(len(unmatched_defs) - len(sample)))

    if missing_labels:
        summary.append("")
        summary.append("Definitions missing linked types:")
        sample = missing_labels[:5]
        summary.extend(" - {}".format(name) for name in sample)
        if len(missing_labels) > len(sample):
            summary.append("   (+{} more)".format(len(missing_labels) - len(sample)))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
