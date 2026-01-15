# -*- coding: utf-8 -*-
"""
Unmerge YAML Profiles
---------------------
Detach a merged profile from a truth-source group so it becomes independent again.
"""

import os
import sys

from pyrevit import forms

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Unmerge YAML Profiles"

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


def main():
    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    equipment_defs = data.get("equipment_definitions") or []
    def_map, _ordered_names = _build_definition_map(equipment_defs)
    if not def_map:
        forms.alert("No equipment definitions are available to unmerge.", title=TITLE)
        return

    truth_groups = _truth_groups(equipment_defs)
    merged_groups = {
        key: group for key, group in truth_groups.items()
        if len(group.get("members") or []) > 1
    }
    if not merged_groups:
        forms.alert("No merged profiles found in the active YAML.", title=TITLE)
        return

    sorted_keys = sorted(
        merged_groups.keys(),
        key=lambda k: (merged_groups[k].get("display_name") or k).lower()
    )
    label_counts = {}
    base_info = []
    for key in sorted_keys:
        group = merged_groups[key]
        base_label = group.get("display_name") or group.get("source_profile_name") or key
        merged_count = max(len(group.get("members") or []) - 1, 0)
        if merged_count:
            base_label = u"{} ({} merged)".format(base_label, merged_count)
        base_info.append((key, base_label))
        label_counts[base_label] = label_counts.get(base_label, 0) + 1
    group_items = []
    label_to_key = {}
    for key, base_label in base_info:
        label = base_label
        if label_counts.get(base_label, 0) > 1:
            label = u"{} [{}]".format(base_label, key)
        group_items.append(label)
        label_to_key[label] = key

    group_choice = forms.SelectFromList.show(
        group_items,
        title="Select merged profile group to unmerge from",
        multiselect=False,
        button_name="Select",
    )
    if not group_choice:
        return
    group_label = group_choice if isinstance(group_choice, basestring) else group_choice[0]
    group_key = label_to_key.get(group_label)
    group = merged_groups.get(group_key or "")
    if not group:
        forms.alert("Could not resolve the selected merged group.", title=TITLE)
        return

    source_name = group.get("source_profile_name") or (group.get("members") or [None])[0]
    candidates = [name for name in (group.get("members") or []) if name and name != source_name]
    if not candidates:
        forms.alert("No merged profiles are available to unmerge from this group.", title=TITLE)
        return
    candidates.sort(key=lambda val: val.lower())

    unmerge_choice = forms.SelectFromList.show(
        candidates,
        title="Select merged profile to unmerge",
        multiselect=False,
        button_name="Unmerge",
    )
    if not unmerge_choice:
        return
    target_name = unmerge_choice if isinstance(unmerge_choice, basestring) else unmerge_choice[0]

    target_entry = def_map.get(target_name)
    if not target_entry:
        forms.alert("Could not resolve '{}' in the active YAML.".format(target_name), title=TITLE)
        return

    new_id = (target_entry.get("id") or target_entry.get("name") or target_name).strip()
    new_name = (target_entry.get("name") or target_entry.get("id") or target_name).strip()
    if not new_id:
        forms.alert("Unable to determine a new truth source id for '{}'.".format(target_name), title=TITLE)
        return

    target_entry[TRUTH_SOURCE_ID_KEY] = new_id
    if new_name:
        target_entry[TRUTH_SOURCE_NAME_KEY] = new_name

    yaml_label = get_yaml_display_name(yaml_path)
    save_active_yaml_data(
        None,
        data,
        TITLE,
        "Unmerged '{}' from '{}'".format(target_name, source_name or group.get("display_name") or group_key),
    )

    forms.alert(
        "Unmerged profile successfully.\n\n"
        "Profile: {}\n"
        "Former group: {}\n\n"
        "Updated data saved back to {}.".format(
            target_name,
            group.get("display_name") or source_name or group_key,
            yaml_label,
        ),
        title=TITLE,
    )


if __name__ == "__main__":
    try:
        basestring
    except NameError:
        basestring = str
    main()
