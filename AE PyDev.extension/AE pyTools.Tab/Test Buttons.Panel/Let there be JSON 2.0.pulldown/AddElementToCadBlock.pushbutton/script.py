# -*- coding: utf-8 -*-
"""
AddElementToCadBlock (persistent)
---------------------------------
Pick a CAD block from a CSV (only needs a name column), then pick Family:Type
entries from the current document to add. The selections are appended to
element_data.yaml so they persist across runs and load through Element_Linker.
"""

import os
import io
import csv
import time
import json

from pyrevit import forms, revit, script
from Autodesk.Revit import DB
from CktParamWindow import CktParamWindow

ELEMENT_LINKER_PATH = script.get_bundle_file(os.path.join("..", "..", "..", "..", "lib", "Element_Linker.py"))
ELEMENT_DATA_PATH = script.get_bundle_file(os.path.join("..", "..", "..", "..", "lib", "element_data.yaml"))

# Fallback relative path if bundle lookup fails
if not ELEMENT_LINKER_PATH or not os.path.exists(ELEMENT_LINKER_PATH):
    ELEMENT_LINKER_PATH = os.path.abspath(os.path.join(script.get_script_path(), "..", "..", "..", "..", "lib", "Element_Linker.py"))
if not ELEMENT_DATA_PATH or not os.path.exists(ELEMENT_DATA_PATH):
    ELEMENT_DATA_PATH = os.path.abspath(os.path.join(script.get_script_path(), "..", "..", "..", "..", "lib", "element_data.yaml"))


def _read_rows(csv_path):
    rows = []
    with open(csv_path, "r") as f:
        reader = csv.DictReader(f)
        for row in reader:
            norm = {}
            for k, v in row.items():
                if k is None:
                    continue
                norm[k.strip().lower()] = v
            rows.append(norm)
    return rows


def _get_str(row, *keys):
    for k in keys:
        if not k:
            continue
        lk = k.strip().lower()
        if lk in row and row[lk]:
            return unicode(row[lk]).strip()
    return ""


ELECTRICAL_PARAMS = [
    "CKT_Panel_CEDT",
    "CKT_Circuit Number_CEDT",
    "CKT_Rating_CED",
    "CKT_Load Name_CEDT",
    "CKT_Schedule Notes_CEDT",
]


def _prompt_electrical_params(display_label):
    """One dialog to capture all CKT_* values for electrical/lighting fixtures."""
    xaml = script.get_bundle_file("CktParamWindow.xaml")
    if xaml:
        window = CktParamWindow(xaml, u"{} — CKT parameters".format(display_label))
        if window.show_dialog():
            return window.get_values()
        return None

    # Fallback: sequential prompts if XAML not found
    params = {}
    for pname in ELECTRICAL_PARAMS:
        value = forms.ask_for_string(
            prompt="Value for {} -> {}".format(display_label, pname),
            default="",
        )
        if value is None:
            value = ""
        params[pname] = value.strip()
    return params


def _prompt_params(display_label, category_name=None, known=None):
    """
    Interactive parameter entry.
    - For Electrical/Lighting Fixtures: prompt only known CKT_* params.
    - For everything else: free-form name/value pairs until blank.
    """
    if category_name in {"Electrical Fixtures", "Lighting Fixtures"}:
        params = _prompt_electrical_params(display_label)
        return params or {}

    # Default free-form parameters
    params = {}
    while True:
        name = forms.ask_for_string(
            prompt="Parameter name for {} (blank to finish)".format(display_label),
            default=""
        )
        if not name:
            break
        value = forms.ask_for_string(
            prompt="Value for {} -> {}".format(display_label, name),
            default=""
        )
        if value is None:
            value = ""
        params[name.strip()] = value.strip()
    return params


def _read_data_file(path):
    if not os.path.exists(path):
        return {"profiles": []}
    with io.open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_data_file(path, data):
    with io.open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def _profile_priority(profile_dict):
    """Return sort key based on category preference."""
    cats = set()
    for t in profile_dict.get("types", []):
        cat = t.get("category_name")
        if cat:
            cats.add(cat)

    if "Electrical Fixtures" in cats:
        order = 0
    elif cats and cats.issubset({"Data Devices"}):
        order = 1
    elif "Lighting Fixtures" in cats:
        order = 2
    elif "Plumbing Fixtures" in cats:
        order = 3
    else:
        order = 4

    return (order, profile_dict.get("cad_name", ""))


def _sort_profiles_in_place(data):
    profiles = data.get("profiles") or []
    profiles.sort(key=_profile_priority)
    data["profiles"] = profiles


def _append_user_block_entries(block_name, entries):
    """
    Append to element_data.yaml (JSON) so Element_Linker can load new profiles/types.
    entries: list of (category, family, type_name, is_group, ox, oy, oz, rot, params_dict)
    """
    if not ELEMENT_DATA_PATH or not os.path.exists(ELEMENT_DATA_PATH):
        forms.alert("element_data.yaml not found; cannot persist changes.", title="Add Element To CAD Block")
        return False

    data = _read_data_file(ELEMENT_DATA_PATH)
    profiles = data.get("profiles") or []

    # Find or create profile entry
    profile = None
    for prof in profiles:
        if prof.get("cad_name") == block_name:
            profile = prof
            break
    if profile is None:
        profile = {"cad_name": block_name, "types": []}
        profiles.append(profile)

    for cat, fam, typ, is_group, ox, oy, oz, rot, params in entries:
        profile["types"].append({
            "label": u"{} : {}".format(fam, typ),
            "category_name": cat,
            "is_group": bool(is_group),
            "instance_config": {
                "parameters": params or {},
                "offsets": [{
                    "x_inches": ox,
                    "y_inches": oy,
                    "z_inches": oz,
                    "rotation_deg": rot,
                }],
                "tags": [],
            },
        })

    data["profiles"] = profiles
    _sort_profiles_in_place(data)
    _write_data_file(ELEMENT_DATA_PATH, data)
    return True


def main():
    # 1) Pick CSV to get CAD block names
    csv_path = forms.pick_file(file_ext="csv", title="Select CAD Block CSV")
    if not csv_path:
        return

    rows = _read_rows(csv_path)
    if not rows:
        forms.alert("No rows found in CSV.", title="Add Element To CAD Block")
        return

    cad_names = sorted({
        _get_str(r, "name", "cad name", "cad_name", "cad block", "cadblock", "block")
        for r in rows if _get_str(r, "name", "cad name", "cad_name", "cad block", "cadblock", "block")
    })
    if not cad_names:
        forms.alert("CSV has no CAD block names (column like 'Name' or 'CAD Name').", title="Add Element To CAD Block")
        return

    cad_choice = forms.SelectFromList.show(
        cad_names,
        title="Select CAD Block to Add Types",
        multiselect=False,
        button_name="Select"
    )
    if not cad_choice:
        return

    # 2) Let user pick Family:Type from current document
    doc = revit.doc
    symbols = list(DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements())
    options = []
    option_map = {}
    for sym in symbols:
        try:
            fam = sym.Family.Name if sym.Family else None
            type_param = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = None
            if type_param:
                type_name = type_param.AsString()
            if not type_name:
                type_name = getattr(sym, "Name", None)
            category = sym.Category.Name if sym.Category else ""
        except Exception:
            continue
        if not fam or not type_name:
            continue
        display = u"{} : {} ({})".format(fam, type_name, category)
        options.append(display)
        # Gather a quick set of parameter names from the symbol only (fast)
        param_names = set()
        try:
            for p in sym.Parameters:
                try:
                    if p.Definition and p.Definition.Name:
                        param_names.add(p.Definition.Name)
                except Exception:
                    continue
        except Exception:
            pass
        option_map[display] = (category, fam, type_name, sym, param_names)

    if not options:
        forms.alert("No family symbols found in the document.", title="Add Element To CAD Block")
        return

    picked = forms.SelectFromList.show(
        sorted(options),
        title="Add Family:Type to '{}'".format(cad_choice),
        multiselect=True,
        button_name="Add"
    )
    if not picked:
        return

    entries = []
    for disp in picked:
        category, family, type_name, symbol, known_params = option_map.get(disp, ("", "", "", None, set()))
        if not category or not family or not type_name or symbol is None:
            continue
        # Prompt for parameters; electrical/lighting prompt fixed CKT_* set
        params = _prompt_params(disp, category_name=category, known=known_params)
        if not params:
            cancel_add = forms.alert(
                "No parameters were entered for:\n\n{}\n\nDo you want to cancel adding this element?".format(disp),
                yes=True,
                no=True,
                title="Add Element To CAD Block"
            )
            if cancel_add:
                continue

        entries.append((category, family, type_name, False, 0.0, 0.0, 0.0, 0.0, params))

    if not entries:
        forms.alert("No entries selected.", title="Add Element To CAD Block")
        return

    t = DB.Transaction(doc, "Add CAD Block Element Data")
    t.Start()
    try:
        ok = _append_user_block_entries(cad_choice, entries)
        if not ok:
            t.RollBack()
            return
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    forms.alert("Added {} type(s) to profile '{}' and persisted to element_data.yaml.\nReload Element_Linker to use them.".format(len(entries), cad_choice),
                title="Add Element To CAD Block")


if __name__ == "__main__":
    main()
