# -*- coding: utf-8 -*-
"""
Merge Linked Elements
---------------------
Allows the user to designate one equipment definition as the source of truth and
copy its configuration to one or more other definitions.  This lets similar
linked element names (e.g. "GO Checkstand 1" vs "GO Checkstand 2") share the
exact same placement data without editing the placer.
"""

import copy
import os
import sys

from pyrevit import forms, script

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Merge Linked Elements"

TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"


def _build_definition_map(equipment_defs):
    mapping = {}
    ordered = []
    for entry in equipment_defs or []:
        if not isinstance(entry, dict):
            continue
        name = (entry.get("name") or entry.get("id") or "").strip()
        if not name:
            continue
        mapping[name] = entry
        ordered.append(name)
    ordered.sort(key=lambda val: val.lower())
    return mapping, ordered


def _truth_groups(equipment_defs):
    groups = {}
    id_to_entry = {}
    name_to_entry = {}
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        eq_name = (entry.get("name") or entry.get("id") or "").strip()
        if eq_id:
            id_to_entry[eq_id] = entry
        if eq_name:
            name_to_entry[eq_name] = entry
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        eq_name = (entry.get("name") or entry.get("id") or "").strip()
        if not eq_name:
            continue
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id:
            source_id = eq_id or eq_name
        display_name = (entry.get(TRUTH_SOURCE_NAME_KEY) or "").strip()
        if not display_name:
            display_name = eq_name
        group = groups.setdefault(source_id, {
            "display_name": display_name,
            "members": [],
            "source_entry": None,
            "source_profile_name": None,
        })
        group["members"].append(eq_name)
        if eq_id and eq_id == source_id:
            group["source_entry"] = entry
            group["source_profile_name"] = eq_name
    for source_id, data in groups.items():
        if not data.get("source_entry"):
            fallback = data["members"][0]
            entry = name_to_entry.get(fallback) or id_to_entry.get(source_id)
            data["source_entry"] = entry
            data["source_profile_name"] = fallback
        if not data.get("display_name"):
            data["display_name"] = data.get("source_profile_name") or source_id
    return groups


def _copy_fields(source_entry, target_entry):
    """Copy everything except identifying fields (name, id)."""
    keep_keys = {"name", "id"}
    for key, value in list(target_entry.items()):
        if key in keep_keys:
            continue
        target_entry.pop(key, None)
    for key, value in source_entry.items():
        if key in keep_keys:
            continue
        target_entry[key] = copy.deepcopy(value)


def _ensure_truth_source(entry):
    if not isinstance(entry, dict):
        return None, None
    eq_id = (entry.get("id") or entry.get("name") or "").strip()
    eq_name = (entry.get("name") or entry.get("id") or "").strip()
    if not eq_id:
        return None, None
    entry[TRUTH_SOURCE_ID_KEY] = eq_id
    if eq_name:
        entry[TRUTH_SOURCE_NAME_KEY] = eq_name
    return eq_id, eq_name


def _repoint_truth_children(equipment_defs, old_source_id, new_source_id, new_source_name):
    old_id = (old_source_id or "").strip()
    new_id = (new_source_id or "").strip()
    if not old_id or not new_id or old_id == new_id:
        return
    for entry in equipment_defs or []:
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id and entry.get("id"):
            source_id = str(entry.get("id") or "").strip()
        if source_id == old_id:
            entry[TRUTH_SOURCE_ID_KEY] = new_id
            if new_source_name:
                entry[TRUTH_SOURCE_NAME_KEY] = new_source_name


def main():
    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    equipment_defs = data.get("equipment_definitions") or []
    def_map, ordered_names = _build_definition_map(equipment_defs)
    truth_groups = _truth_groups(equipment_defs)
    if not def_map:
        forms.alert("No equipment definitions are available to merge.", title=TITLE)
        return

    yaml_label = get_yaml_display_name(yaml_path)

    if truth_groups:
        sorted_keys = sorted(truth_groups.keys(), key=lambda k: (truth_groups[k]["display_name"] or k).lower())
        label_counts = {}
        base_info = []
        for key in sorted_keys:
            group = truth_groups[key]
            base_label = group.get("display_name") or group.get("source_profile_name") or key
            member_count = len(group.get("members") or [])
            if member_count > 1:
                base_label = u"{} ({} profiles)".format(base_label, member_count)
            base_info.append((key, base_label))
            label_counts[base_label] = label_counts.get(base_label, 0) + 1
        source_items = []
        label_to_key = {}
        for key, base_label in base_info:
            label = base_label
            if label_counts.get(base_label, 0) > 1:
                label = u"{} [{}]".format(base_label, key)
            source_items.append(label)
            label_to_key[label] = key
        source_choice = forms.SelectFromList.show(
            source_items,
            title="Select source definition (truth)",
            multiselect=False,
            button_name="Select",
        )
        if not source_choice:
            return
        source_label = source_choice if isinstance(source_choice, basestring) else source_choice[0]
        source_key = label_to_key.get(source_label)
        group = truth_groups.get(source_key or "")
        source_entry = group.get("source_entry") if group else None
        source_name = group.get("source_profile_name") if group else None
        if not source_entry or not source_name:
            forms.alert("Could not resolve the selected source definition.", title=TITLE)
            return
        source_display = group.get("display_name") or source_name
    else:
        source_choice = forms.SelectFromList.show(
            ordered_names,
            title="Select source definition (truth)",
            multiselect=False,
            button_name="Select",
        )
        if not source_choice:
            return
        source_name = source_choice if isinstance(source_choice, basestring) else source_choice[0]
        source_entry = def_map.get(source_name)
        source_display = source_name
        if not source_entry:
            forms.alert("Could not resolve the selected source definition.", title=TITLE)
            return

    target_candidates = [name for name in ordered_names if name != source_name]
    if not target_candidates:
        forms.alert("There are no other definitions to merge into.", title=TITLE)
        return

    target_choices = forms.SelectFromList.show(
        target_candidates,
        title="Select definition(s) to merge into '{}'".format(source_name),
        multiselect=True,
        button_name="Merge",
    )
    if not target_choices:
        return

    root_id, root_name = _ensure_truth_source(source_entry)
    if not root_id:
        root_id = (source_entry.get("id") or source_entry.get("name") or source_name).strip()
    if not root_name:
        root_name = (source_entry.get("name") or source_display or source_name).strip()
    merged = []
    for target_name in target_choices:
        target_entry = def_map.get(target_name)
        if not target_entry or target_entry is source_entry:
            continue
        _copy_fields(source_entry, target_entry)
        target_entry[TRUTH_SOURCE_ID_KEY] = root_id
        if root_name:
            target_entry[TRUTH_SOURCE_NAME_KEY] = root_name
        _repoint_truth_children(
            equipment_defs,
            target_entry.get("id"),
            root_id,
            root_name,
        )
        merged.append(target_name)

    if not merged:
        forms.alert("No definitions were merged.", title=TITLE)
        return

    save_active_yaml_data(
        None,
        data,
        "Merge Linked Elements",
        "Merged definitions {} into '{}'".format(", ".join(merged), source_name),
    )

    lines = [
        "Merged linked elements successfully.",
        "Source definition: {}".format(source_display or source_name),
        "Merged into source: {}".format(", ".join(merged)),
        "",
        "Updated data saved back to {}.".format(yaml_label),
    ]
    forms.alert("\n".join(lines), title=TITLE)


if __name__ == "__main__":
    try:
        basestring
    except NameError:
        basestring = str
    main()
