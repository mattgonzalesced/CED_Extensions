# -*- coding: utf-8 -*-
"""
Entry script for Element Linker.

Workflow:
1. Ask user for CSV of CAD points.
2. Collect CAD block names from CSV.
3. Auto-map each CAD name to ALL types in its Element_Linker profile (editable popup allows add/remove).
4. Ask user to pick a level.
5. Run ElementPlacementEngine to place everything.
"""

from pyrevit import revit, forms, script

from ElementLinkerUtils import read_xyz_csv
from Element_Linker import CAD_BLOCK_PROFILES
from ElementLinkerWindow import ElementLinkerWindow
from ElementLinkerEngine import ElementPlacementEngine


def main():
    doc = revit.doc

    # 1. Pick CSV file
    csv_path = forms.pick_file(file_ext='csv', title='Select XYZ / CAD CSV')
    if not csv_path:
        return

    csv_rows, cad_names = read_xyz_csv(csv_path)
    if not csv_rows:
        forms.alert("No rows in CSV.", title="Element Linker")
        return

    # 2. Build mapping automatically: place ALL types for each CAD block profile
    cad_selection_map = {}
    for cad_name in cad_names:
        profile = CAD_BLOCK_PROFILES.get(cad_name)
        if not profile:
            continue
        labels = [lbl for lbl in profile.get_type_labels() if lbl]
        if labels:
            # keep order but drop accidental duplicates
            seen = set()
            distinct_labels = []
            for lbl in labels:
                if lbl in seen:
                    continue
                seen.add(lbl)
                distinct_labels.append(lbl)
            cad_selection_map[cad_name] = distinct_labels

    if not cad_selection_map:
        forms.alert("No CAD names have profiles with types to place.", title="Element Linker")
        return

    # 3. Allow editing via WPF popup (pre-populated with all types)
    xaml_path = script.get_bundle_file('ElementLinkerWindow.xaml')
    if not xaml_path:
        forms.alert("ElementLinkerWindow.xaml not found in bundle.", title="Element Linker")
        return

    window = ElementLinkerWindow(
        xaml_path=xaml_path,
        cad_names=cad_names,
        cad_block_profiles=CAD_BLOCK_PROFILES,
        initial_mapping=cad_selection_map,
    )
    window.show_dialog()

    if not window.DialogResult:
        return

    cad_selection_map = window.result_mapping
    if not cad_selection_map:
        forms.alert("No selections to place.", title="Element Linker")
        return

    # 4. Pick placement level
    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    else:
        level = level_sel

    # 5. Run placement engine
    engine = ElementPlacementEngine(doc, default_level=level)
    engine.place_from_csv(csv_rows, cad_selection_map)


if __name__ == "__main__":
    main()

