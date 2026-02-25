# -*- coding: utf-8 -*-
"""
Place Elements (YAML) - Category Filter
----------------------------------------
YAML-driven placement tool that mirrors the old Populate Elements flow but now
reads the active YAML stored in Extensible Storage. Supports both family
instances and model groups. Includes category filtering before placement.
"""

import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import FamilySymbol, FilteredElementCollector, GroupType, BuiltInParameter

# Add CEDLib.lib to sys.path so LogicClasses/UIClasses can be imported
# (relative to extension root: ...\AE PyDev.extension\AE pyTools.Tab\...)
LIB_ROOT = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import (  # noqa: E402,I100
    ProfileRepository,
    PlaceElementsEngine,
    read_xyz_csv,
)
from UIClasses.PlaceElementsUI import PlaceElementsWindow  # noqa: E402
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Place Elements (YAML) - Category Filter"

# MEP and Architecture categories for filtering
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
    "Generic Annotations",
    "Generic Models",
    "HVAC Zones",
    "Lighting Devices",
    "Lighting Fixtures",
    "Mechanical Control Devices",
    "Mechanical Equipment",
    "Model Groups",
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

    # Show category selection dialog
    selected_categories = forms.SelectFromList.show(
        MEP_CATEGORIES,
        title="Select Categories to Place",
        multiselect=True,
        button_name="Continue"
    )
    if not selected_categories:
        return  # User cancelled

    # Convert to set for faster lookup
    selected_categories = set(selected_categories)

    # Build initial mapping with category filtering
    initial_mapping = {}
    skipped_by_category = 0

    for cad in cad_names:
        labels = _dedupe_preserve(repo.labels_for_cad(cad))
        if not labels:
            continue

        # Filter labels by category (runtime only, not YAML)
        filtered_labels = []
        for label in labels:
            cat_name = _get_category_name_from_label(revit.doc, label)

            if cat_name and cat_name in selected_categories:
                filtered_labels.append(label)
            else:
                skipped_by_category += 1

        if filtered_labels:
            initial_mapping[cad] = filtered_labels

    if not initial_mapping:
        forms.alert(
            "No profiles found matching the selected categories.\n\n"
            "Selected categories: {}".format(", ".join(sorted(selected_categories))),
            title=TITLE
        )
        return

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

    # Re-filter selection_map to ensure only selected categories are placed
    filtered_selection_map = {}
    for cad, labels in selection_map.items():
        filtered_labels = []
        for label in labels:
            cat_name = _get_category_name_from_label(revit.doc, label)
            if cat_name and cat_name in selected_categories:
                filtered_labels.append(label)
            else:
                skipped_by_category += 1
        if filtered_labels:
            filtered_selection_map[cad] = filtered_labels

    selection_map = filtered_selection_map
    if not selection_map:
        forms.alert("No items match the selected categories after user selection.", title=TITLE)
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
        transaction_name="Place CAD Elements (YAML) - Category Filter",
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
        "Selected categories: {}".format(", ".join(sorted(selected_categories))),
        "Labels skipped by category filter: {}".format(skipped_by_category),
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
        msg_lines.append(" - Selected categories match the equipment definitions.")

    forms.alert("\n".join(msg_lines), title=TITLE)


if __name__ == "__main__":
    main()
