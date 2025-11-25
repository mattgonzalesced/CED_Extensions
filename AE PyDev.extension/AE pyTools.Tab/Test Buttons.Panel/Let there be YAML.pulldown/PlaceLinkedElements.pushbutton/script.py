# -*- coding: utf-8 -*-
"""Place an equipment definition and all of its linked children at once."""

import os
import sys

from pyrevit import revit, forms

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository, tag_key_from_dict  # noqa: E402
from profile_schema import equipment_defs_to_legacy, load_data as load_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402
from LogicClasses.linked_equipment import build_child_requests, find_equipment_by_name  # noqa: E402

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


def _collect_tag_defs(repo, selection_map):
    tags = {}
    for cad_name, labels in selection_map.items():
        if isinstance(labels, basestring):
            labels_iter = [labels]
        else:
            labels_iter = list(labels)
        for label in labels_iter:
            linked_def = repo.definition_for_label(cad_name, label)
            if not linked_def:
                continue
            placement = linked_def.get_placement()
            if not placement:
                continue
            for tag in placement.get_tags() or []:
                key = tag_key_from_dict(tag)
                if key and key not in tags:
                    tags[key] = tag
    return tags


def _place_requests(doc, repo, selection_map, rows, default_level=None, view_id=None):
    if not selection_map or not rows:
        return {"placed": 0}
    tag_defs = _collect_tag_defs(repo, selection_map)
    tag_view_map = {}
    if view_id:
        for key in tag_defs:
            tag_view_map.setdefault(key, []).append(view_id)
    engine = PlaceElementsEngine(doc, repo, default_level=default_level, tag_view_map=tag_view_map)
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

    parent_choice = forms.SelectFromList.show(
        equipment_names,
        title="Select equipment definition to place",
        multiselect=False,
        button_name="Select")
    if not parent_choice:
        return
    parent_choice = parent_choice if isinstance(parent_choice, basestring) else parent_choice[0]

    parent_labels = repo.labels_for_cad(parent_choice)
    if not parent_labels:
        forms.alert("Equipment definition '{}' has no linked types.".format(parent_choice), title=TITLE)
        return

    try:
        base_point = revit.pick_point(message="Pick placement point for '{}'".format(parent_choice))
    except Exception:
        base_point = None
    if not base_point:
        return

    rotation_input = forms.ask_for_string(prompt="Enter rotation (degrees) for '{}':".format(parent_choice), default="0")
    try:
        base_rotation = float(rotation_input or 0.0)
    except Exception:
        base_rotation = 0.0

    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    elif level_sel:
        level = level_sel

    active_view = getattr(doc, "ActiveView", None)
    view_id = getattr(getattr(active_view, "Id", None), "IntegerValue", None)

    rows = [_build_row(parent_choice, base_point, base_rotation)]
    selection_map = {parent_choice: parent_labels}

    parent_def = find_equipment_by_name(data, parent_choice)
    if parent_def:
        child_requests = _gather_child_requests(parent_def, base_point, base_rotation, repo, data)
        if child_requests:
            prompt = "Place {} linked child equipment definition(s) as well?".format(len(child_requests))
            if forms.alert(prompt, title=TITLE, yes=True, no=True):
                for request in child_requests:
                    name = request.get("name")
                    labels = request.get("labels")
                    point = request.get("target_point")
                    rotation = request.get("rotation")
                    if not name or not labels or point is None:
                        continue
                    selection_map[name] = labels
                    rows.append(_build_row(name, point, rotation))

    results = _place_requests(doc, repo, selection_map, rows, default_level=level, view_id=view_id)
    placed = results.get("placed", 0)
    forms.alert("Placed {} element(s).".format(placed), title=TITLE)


if __name__ == "__main__":
    main()
