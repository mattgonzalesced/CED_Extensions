# -*- coding: utf-8 -*-
"""
Place Elements (YAML)
---------------------
YAML-driven placement tool that mirrors the old Populate Elements flow but now
reads the active YAML stored in Extensible Storage. Supports both family
instances and model groups.
"""

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
from profile_schema import equipment_defs_to_legacy  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Place Elements (YAML)"

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


def _build_repository(data):
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def main():
    try:
        data_path, raw_data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)

    csv_path = forms.pick_file(file_ext="csv", title="Select XYZ / CAD CSV")
    if not csv_path:
        return

    csv_rows, cad_names = read_xyz_csv(csv_path)
    if not csv_rows:
        forms.alert("No rows in CSV.", title=TITLE)
        return

    repo = _build_repository(raw_data)
    if not repo.cad_names():
        forms.alert("No profiles found in {}.".format(yaml_label), title=TITLE)
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
        forms.alert("No selections to place.", title=TITLE)
        return

    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    else:
        level = level_sel

    engine = PlaceElementsEngine(
        revit.doc,
        repo,
        default_level=level,
        allow_tags=False,
        transaction_name="Place CAD Elements (YAML)",
    )
    try:
        results = engine.place_from_csv(csv_rows, selection_map)
    except Exception as exc:
        forms.alert("Error during placement:\n\n{0}".format(exc), title=TITLE)
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
        msg_lines.append(" - Profiles exist for the CAD names in {}.".format(yaml_label))
        msg_lines.append(" - Labels match loaded Revit families/types or model groups.")

    forms.alert("\n".join(msg_lines), title=TITLE)


if __name__ == "__main__":
    main()
