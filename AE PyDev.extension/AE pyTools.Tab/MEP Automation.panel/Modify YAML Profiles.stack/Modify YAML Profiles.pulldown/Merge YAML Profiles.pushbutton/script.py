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


def main():
    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    equipment_defs = data.get("equipment_definitions") or []
    def_map, ordered_names = _build_definition_map(equipment_defs)
    if not def_map:
        forms.alert("No equipment definitions are available to merge.", title=TITLE)
        return

    yaml_label = get_yaml_display_name(yaml_path)

    source_choice = forms.SelectFromList.show(
        ordered_names,
        title="Select source definition (truth)",
        multiselect=False,
        button_name="Select",
    )
    if not source_choice:
        return
    source_name = source_choice if isinstance(source_choice, basestring) else source_choice[0]
    if source_name not in def_map:
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

    source_entry = def_map[source_name]
    merged = []
    for target_name in target_choices:
        target_entry = def_map.get(target_name)
        if not target_entry or target_entry is source_entry:
            continue
        _copy_fields(source_entry, target_entry)
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
        "Source definition: {}".format(source_name),
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
