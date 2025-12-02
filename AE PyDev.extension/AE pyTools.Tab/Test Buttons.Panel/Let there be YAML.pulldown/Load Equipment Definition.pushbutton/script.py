# -*- coding: utf-8 -*-
"""
Load Equipment Definition
-------------------------
Select a profileData YAML, choose an equipment definition, and place all of its
linked types at a picked point.
"""

from __future__ import print_function

import io
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import XYZ

LIB_ROOT = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "..",
        "..",
        "CEDLib.lib",
    )
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.placement_engine import PlaceElementsEngine  # noqa: E402
from LogicClasses.profile_repository import ProfileRepository  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402
from LogicClasses.linked_equipment import build_child_requests, find_equipment_by_name  # noqa: E402
try:
    from profile_schema import equipment_defs_to_legacy, load_data as load_profile_data  # noqa: E402
except Exception:
    from CEDLib.lib.profile_schema import equipment_defs_to_legacy, load_data as load_profile_data  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")

try:
    basestring
except NameError:
    basestring = str


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
            led_list = []
            for led in ls_copy.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_copy = dict(led)
                tags = led_copy.get("tags")
                if isinstance(tags, list):
                    led_copy["tags"] = [t if isinstance(t, dict) else {} for t in tags]
                else:
                    led_copy["tags"] = []
                offsets = led_copy.get("offsets")
                if isinstance(offsets, list):
                    led_copy["offsets"] = [o if isinstance(o, dict) else {} for o in offsets]
                else:
                    led_copy["offsets"] = [{}]
                led_list.append(led_copy)
            ls_copy["linked_element_definitions"] = led_list
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
        types = []
        for t in prof_copy.get("types") or []:
            if not isinstance(t, dict):
                continue
            t_copy = dict(t)
            inst_cfg = t_copy.get("instance_config")
            if not isinstance(inst_cfg, dict):
                inst_cfg = {}
            offsets = inst_cfg.get("offsets")
            if not isinstance(offsets, list) or not offsets:
                offsets = [{}]
            inst_cfg["offsets"] = [off if isinstance(off, dict) else {} for off in offsets]
            tags = inst_cfg.get("tags")
            if isinstance(tags, list):
                inst_cfg["tags"] = [tag if isinstance(tag, dict) else {} for tag in tags]
            else:
                inst_cfg["tags"] = []
            params = inst_cfg.get("parameters")
            if not isinstance(params, dict):
                params = {}
            inst_cfg["parameters"] = params
            t_copy["instance_config"] = inst_cfg
            types.append(t_copy)
        prof_copy["types"] = types
        cleaned.append(prof_copy)
    return cleaned


def _build_repository(data):
    cleaned_defs = _sanitize_equipment_definitions(data.get("equipment_definitions") or [])
    legacy_profiles = equipment_defs_to_legacy(cleaned_defs)
    cleaned_profiles = _sanitize_profiles(legacy_profiles)
    eq_defs = ProfileRepository._parse_profiles(cleaned_profiles)
    return ProfileRepository(eq_defs)


def _place_child_requests(repo, child_requests):
    selection_map = {}
    rows = []
    for request in child_requests or []:
        cad_name = request.get("name")
        labels = request.get("labels")
        point = request.get("target_point")
        rotation = request.get("rotation")
        if not cad_name or not labels or point is None:
            continue
        selection_map[cad_name] = labels
        rows.append({
            "Name": cad_name,
            "Count": "1",
            "Position X": str(point.X * 12.0),
            "Position Y": str(point.Y * 12.0),
            "Position Z": str(point.Z * 12.0),
            "Rotation": str(rotation or 0.0),
        })
    if not selection_map or not rows:
        return 0
    engine = PlaceElementsEngine(revit.doc, repo, allow_tags=False, transaction_name="Load Equipment Definition (Children)")
    try:
        results = engine.place_from_csv(rows, selection_map)
    except Exception as exc:
        forms.alert("Failed to place linked child equipment:\\n\\n{}".format(exc), title="Load Equipment Definition")
        return 0
    return results.get("placed", 0)


def _gather_child_requests(parent_def, base_point, base_rotation, repo, data):
    requests = []
    if not parent_def:
        return requests
    for linked_set in parent_def.get("linked_sets") or []:
        for led_entry in linked_set.get("linked_element_definitions") or []:
            led_id = (led_entry.get("id") or "").strip()
            if not led_id:
                continue
            reqs = build_child_requests(repo, data, parent_def, base_point, base_rotation, led_id)
            if reqs:
                requests.extend(reqs)
    return requests


def main():
    data_path = _pick_profile_data_path()
    if not data_path:
        return

    raw_data = _load_profile_store(data_path)
    repo = _build_repository(raw_data)
    cad_names = repo.cad_names()
    if not cad_names:
        forms.alert("No equipment definitions found in the selected YAML.", title="Load Equipment Definition")
        return

    cad_choice = forms.SelectFromList.show(
        cad_names,
        title="Select equipment definition to place",
        multiselect=False,
        button_name="Load",
    )
    if not cad_choice:
        return
    cad_choice = cad_choice if isinstance(cad_choice, basestring) else cad_choice[0]

    labels = repo.labels_for_cad(cad_choice)
    if not labels:
        forms.alert("Equipment definition '{}' has no linked types.".format(cad_choice), title="Load Equipment Definition")
        return

    try:
        base_pt = revit.pick_point(message="Pick base point for '{}'".format(cad_choice))
    except Exception:
        base_pt = None
    if not base_pt:
        return

    selection_map = {cad_choice: labels}
    rows = [{
        "Name": cad_choice,
        "Count": "1",
        "Position X": str(base_pt.X * 12.0),
        "Position Y": str(base_pt.Y * 12.0),
        "Position Z": str(base_pt.Z * 12.0),
        "Rotation": "0",
    }]

    parent_def = find_equipment_by_name(raw_data, cad_choice)
    if parent_def:
        child_requests = _gather_child_requests(parent_def, base_pt, 0.0, repo, raw_data)
        if child_requests:
            if forms.alert(
                "Load '{}' with {} linked child equipment definition(s)?".format(cad_choice, len(child_requests)),
                title="Load Equipment Definition",
                yes=True,
                no=True,
            ):
                for request in child_requests:
                    name = request.get("name")
                    labels = request.get("labels")
                    point = request.get("target_point")
                    rotation = request.get("rotation")
                    if not name or not labels or point is None:
                        continue
                    selection_map[name] = labels
                    rows.append({
                        "Name": name,
                        "Count": "1",
                        "Position X": str(point.X * 12.0),
                        "Position Y": str(point.Y * 12.0),
                        "Position Z": str(point.Z * 12.0),
                        "Rotation": str(rotation or 0.0),
                    })

    engine = PlaceElementsEngine(revit.doc, repo, allow_tags=False, transaction_name="Load Equipment Definition")
    try:
        results = engine.place_from_csv(rows, selection_map)
    except Exception as exc:
        forms.alert("Error during placement:\n\n{}".format(exc), title="Load Equipment Definition")
        return

    forms.alert("Placed {} element(s) for equipment definition '{}'.".format(results.get("placed", 0), cad_choice), title="Load Equipment Definition")


if __name__ == "__main__":
    main()
