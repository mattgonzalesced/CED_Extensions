# -*- coding: utf-8 -*-
"""
EraseElementData
----------------
Pick a CAD block from a CSV, then select and remove Type entries from element_data.yaml.
Writes in a transaction so users can undo any related Revit changes (file changes persist).
"""

import os
import io
import csv
import json
import hashlib
import datetime

from pyrevit import forms, revit, script
from Autodesk.Revit import DB

ELEMENT_DATA_PATH = script.get_bundle_file(os.path.join("..", "..", "..", "..", "lib", "element_data.yaml"))
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


def _read_data_file(path):
    if not os.path.exists(path):
        return {"profiles": []}
    with io.open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_data_file(path, data):
    with io.open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def _file_hash(path):
    if not os.path.exists(path):
        return ""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def _append_log(action, cad_name, type_labels, before_hash, after_hash):
    log_path = os.path.join(os.path.dirname(ELEMENT_DATA_PATH), "element_data.log")
    entry = {
        "timestamp": datetime.datetime.utcnow().isoformat() + "Z",
        "user": os.getenv("USERNAME") or os.getenv("USER") or "unknown",
        "action": action,
        "cad_name": cad_name,
        "type_labels": list(type_labels or []),
        "before_hash": before_hash,
        "after_hash": after_hash,
    }
    with io.open(log_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=True) + "\n")


def _profile_priority(profile_dict):
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


def _erase_entries(profile_name, type_labels):
    data = _read_data_file(ELEMENT_DATA_PATH)
    profiles = data.get("profiles") or []

    target = None
    for prof in profiles:
        if prof.get("cad_name") == profile_name:
            target = prof
            break

    if not target:
        forms.alert("Profile '{}' not found in element_data.yaml.".format(profile_name), title="Erase Element Data")
        return False

    before = len(target.get("types") or [])
    target["types"] = [t for t in target.get("types") or [] if t.get("label") not in type_labels]

    # Drop profile entirely if no types remain
    if not target["types"]:
        profiles = [p for p in profiles if p.get("cad_name") != profile_name]
    data["profiles"] = profiles
    _sort_profiles_in_place(data)

    _write_data_file(ELEMENT_DATA_PATH, data)
    after = len(target.get("types") or []) if target in profiles else 0
    return before != after


def main():
    # 1) Pick CSV to get CAD block names
    csv_path = forms.pick_file(file_ext="csv", title="Select CAD Block CSV")
    if not csv_path:
        return

    rows = _read_rows(csv_path)
    cad_names = sorted({
        _get_str(r, "name", "cad name", "cad_name", "cad block", "cadblock", "block")
        for r in rows if _get_str(r, "name", "cad name", "cad_name", "cad block", "cadblock", "block")
    })
    if not cad_names:
        forms.alert("CSV has no CAD block names (column like 'Name' or 'CAD Name').", title="Erase Element Data")
        return

    # 2) Load existing data and filter available profiles
    data = _read_data_file(ELEMENT_DATA_PATH)
    profiles_by_name = {p.get("cad_name"): p for p in data.get("profiles") or []}
    available_names = [n for n in cad_names if n in profiles_by_name]
    if not available_names:
        forms.alert("None of the CSV CAD blocks exist in element_data.yaml.", title="Erase Element Data")
        return

    cad_choice = forms.SelectFromList.show(
        sorted(available_names),
        title="Select CAD Block to Erase Types",
        multiselect=False,
        button_name="Select"
    )
    if not cad_choice:
        return

    profile = profiles_by_name.get(cad_choice)
    type_labels = [t.get("label") for t in profile.get("types") or [] if t.get("label")]
    if not type_labels:
        forms.alert("Profile '{}' has no types to erase.".format(cad_choice), title="Erase Element Data")
        return

    picked = forms.SelectFromList.show(
        sorted(type_labels),
        title="Select Types to Erase from '{}'".format(cad_choice),
        multiselect=True,
        button_name="Erase"
    )
    if not picked:
        return

    before_hash = _file_hash(ELEMENT_DATA_PATH)
    doc = revit.doc
    t = DB.Transaction(doc, "Erase CAD Block Element Data")
    t.Start()
    try:
        changed = _erase_entries(cad_choice, set(picked))
        if not changed:
            t.RollBack()
            return
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    after_hash = _file_hash(ELEMENT_DATA_PATH)
    _append_log("erase", cad_choice, picked, before_hash, after_hash)

    forms.alert("Erased {} type(s) from profile '{}' and saved to element_data.yaml.\nReload Populate Elements to use updated data.".format(len(picked), cad_choice),
                title="Erase Element Data")


if __name__ == "__main__":
    main()
