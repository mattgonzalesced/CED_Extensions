# -*- coding: utf-8 -*-
"""
Entry script for Element Linker.

Workflow:
1. Ask user for CSV of CAD points.
2. Collect CAD block names from CSV.
3. Show WPF window (ElementLinkerWindow.xaml) to map each CAD name to an Element_Linker profile.
4. Ask user to pick a level.
5. Run ElementPlacementEngine to place everything.
"""

import os

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

    # 2. Show mapping UI
    xaml_path = script.get_bundle_file('ElementLinkerWindow.xaml')
    if not xaml_path or not os.path.exists(xaml_path):
        forms.alert("ElementLinkerWindow.xaml not found in bundle.", title="Element Linker")
        return

    window = ElementLinkerWindow(xaml_path, cad_names, CAD_BLOCK_PROFILES)
    window.show_dialog()

    if not window.DialogResult:
        return

    cad_selection_map = window.result_mapping
    if not cad_selection_map:
        forms.alert("No mappings selected.", title="Element Linker")
        return

    # 3. Pick placement level
    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    else:
        level = level_sel

    # 4. Run placement engine
    engine = ElementPlacementEngine(doc, default_level=level)
    engine.place_from_csv(csv_rows, cad_selection_map)


if __name__ == "__main__":
    main()

