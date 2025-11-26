# -*- coding: utf-8 -*-
"""
PlaceElements
-------------
YAML-driven placement tool that mirrors the old Populate Elements flow but
reads profiles from CEDLib.lib/profileData.yaml and uses LogicClasses for all
logic. Supports both family instances and model groups.
"""

import io
import os
import sys

from pyrevit import revit, forms

# Add CEDLib.lib to sys.path so LogicClasses/UIClasses can be imported
# (need to climb out of pushbutton/pulldown/panel/tab/extension)
LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import (  # noqa: E402,I100
    ProfileRepository,
    PlaceElementsEngine,
    read_xyz_csv,
)
from UIClasses.PlaceElementsUI import PlaceElementsWindow  # noqa: E402
from profile_schema import equipment_defs_to_legacy, load_data as load_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")

try:
    basestring
except NameError:
    basestring = str

def _dedupe_preserve(seq):
    seen = set()
    result = []
    for item in seq:
        if item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


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


def _build_repository(data_path):
    data = _load_profile_store(data_path)
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def main():
    data_path = _pick_profile_data_path()
    if not data_path:
        return

    csv_path = forms.pick_file(file_ext="csv", title="Select XYZ / CAD CSV")
    if not csv_path:
        return

    csv_rows, cad_names = read_xyz_csv(csv_path)
    if not csv_rows:
        forms.alert("No rows in CSV.", title="Place Elements (YAML)")
        return

    repo = _build_repository(data_path)
    if not repo.cad_names():
        forms.alert("No profiles found in the selected YAML file.", title="Place Elements (YAML)")
        return

    initial_mapping = {}
    for cad in cad_names:
        labels = _dedupe_preserve(repo.labels_for_cad(cad))
        if labels:
            initial_mapping[cad] = labels

    xaml_path = os.path.join(LIB_ROOT, "UIClasses", "PlaceElementsUI.xaml")
    window = PlaceElementsWindow(
        xaml_path=xaml_path,
        cad_names=cad_names,
        profile_repo=repo,
        initial_mapping=initial_mapping,
    )
    window.show_dialog()
    if not getattr(window, "DialogResult", False):
        return

    selection_map = window.result_mapping
    if not selection_map:
        forms.alert("No selections to place.", title="Place Elements (YAML)")
        return

    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    else:
        level = level_sel

    engine = PlaceElementsEngine(revit.doc, repo, default_level=level, allow_tags=False)
    try:
        results = engine.place_from_csv(csv_rows, selection_map)
    except Exception as exc:
        forms.alert("Error during placement:\n\n{0}".format(exc), title="Place Elements (YAML)")
        return

    # Debug output when run with pyRevit Ctrl+Click (prints to console)
    # debug output removed

    msg_lines = [
        "Placement complete.",
        "",
        "Total CSV rows: {0}".format(results.get("total_rows", 0)),
        "Rows with a selected CAD mapping: {0}".format(results.get("rows_with_mapping", 0)),
        "Rows with valid coordinates: {0}".format(results.get("rows_with_coords", 0)),
        "Elements/Groups placed: {0}".format(results.get("placed", 0)),
    ]

    if results.get("placed", 0) == 0:
        msg_lines.append("")
        msg_lines.append("No elements were placed. Check that:")
        msg_lines.append(" - Profiles exist for the CAD names in profileData.yaml.")
        msg_lines.append(" - Labels match loaded Revit families/types or model groups.")

    forms.alert("\n".join(msg_lines), title="Place Elements (YAML)")


if __name__ == "__main__":
    main()
