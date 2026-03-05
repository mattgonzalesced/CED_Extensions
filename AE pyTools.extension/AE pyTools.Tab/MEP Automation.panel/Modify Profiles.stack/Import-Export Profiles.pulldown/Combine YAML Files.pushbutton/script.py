# -*- coding: utf-8 -*-
"""
Combine two equipment-definition YAML files into a new dataset.
"""

import copy
import io
import os
import re
import sys
try:
    from collections.abc import Mapping
except ImportError:
    from collections import Mapping

from pyrevit import forms, script
output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.profile_schema import load_data_from_text, dump_data_to_string  # noqa: E402

TITLE = "Combine YAML Files"
ID_PATTERN = re.compile(r"^(.*?)(\d+)$")
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")

try:
    basestring
except NameError:
    basestring = str


def _read_text(path):
    with io.open(path, "r", encoding="utf-8") as handle:
        return handle.read()


def _load_yaml(path):
    raw = _read_text(path)
    return load_data_from_text(raw, path)


def _split_id(value):
    if value is None:
        return None, None, None
    text = value if isinstance(value, str) else str(value)
    match = ID_PATTERN.match(text.strip())
    if not match:
        return None, None, None
    prefix, digits = match.groups()
    try:
        number = int(digits)
    except Exception:
        return None, None, None
    return prefix, number, len(digits)


def _collect_id_stats(defs):
    max_num = 0
    prefix_counts = {}
    max_pad = 0
    for entry in defs:
        if not isinstance(entry, Mapping):
            continue
        prefix, number, pad = _split_id(entry.get("id"))
        if prefix is None or number is None:
            continue
        max_num = max(max_num, number)
        max_pad = max(max_pad, pad or 0)
        prefix_counts[prefix] = prefix_counts.get(prefix, 0) + 1
    return max_num, prefix_counts, max_pad


def _pick_default_prefix(prefix_counts):
    if not prefix_counts:
        return None
    return sorted(prefix_counts.items(), key=lambda item: (-item[1], item[0]))[0][0]


def _collect_set_id_stats(defs):
    max_num = 0
    prefix_counts = {}
    max_pad = 0
    for entry in defs:
        if not isinstance(entry, Mapping):
            continue
        for linked_set in entry.get("linked_sets") or []:
            if not isinstance(linked_set, Mapping):
                continue
            prefix, number, pad = _split_id(linked_set.get("id"))
            if prefix is None or number is None:
                continue
            max_num = max(max_num, number)
            max_pad = max(max_pad, pad or 0)
            prefix_counts[prefix] = prefix_counts.get(prefix, 0) + 1
    return max_num, prefix_counts, max_pad


def _rewrite_linker_payload(params, old_set_id, new_set_id, old_led_id, new_led_id):
    if not isinstance(params, Mapping):
        return
    for key in ELEMENT_LINKER_PARAM_NAMES:
        value = params.get(key)
        if not isinstance(value, basestring):
            continue
        updated = value
        if old_set_id:
            updated = updated.replace(old_set_id, new_set_id)
        if old_led_id:
            updated = updated.replace(old_led_id, new_led_id)
        if updated != value:
            params[key] = updated


def _renumber_defs(defs, start_num, default_prefix, default_pad, start_set_num, default_set_prefix, default_set_pad):
    updated = copy.deepcopy(defs)
    eq_current = start_num
    set_current = start_set_num
    eq_id_map = {}
    new_ids = []
    remapped_set_count = 0
    remapped_led_count = 0

    for entry in updated:
        if not isinstance(entry, Mapping):
            new_ids.append(None)
            continue
        old_id = (entry.get("id") or "").strip()
        prefix, _, pad = _split_id(old_id)
        if prefix is None:
            prefix = default_prefix
        if pad is None or pad <= 0:
            pad = default_pad
        new_id = "{}{}".format(prefix, str(eq_current).zfill(pad))
        new_ids.append(new_id)
        if old_id and old_id not in eq_id_map:
            eq_id_map[old_id] = new_id
        eq_current += 1

    for idx, entry in enumerate(updated):
        if not isinstance(entry, Mapping):
            continue
        new_eq_id = new_ids[idx]
        entry["id"] = new_eq_id

        old_truth_source = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        remapped_truth = eq_id_map.get(old_truth_source)
        if remapped_truth:
            entry[TRUTH_SOURCE_ID_KEY] = remapped_truth
        else:
            entry[TRUTH_SOURCE_ID_KEY] = new_eq_id
            if not (entry.get(TRUTH_SOURCE_NAME_KEY) or "").strip():
                entry[TRUTH_SOURCE_NAME_KEY] = (entry.get("name") or new_eq_id or "").strip()

        for linked_set in entry.get("linked_sets") or []:
            if not isinstance(linked_set, Mapping):
                continue
            old_set_id = (linked_set.get("id") or "").strip()
            set_prefix, _, set_pad = _split_id(old_set_id)
            if set_prefix is None:
                set_prefix = default_set_prefix
            if set_pad is None or set_pad <= 0:
                set_pad = default_set_pad
            new_set_id = "{}{}".format(set_prefix, str(set_current).zfill(set_pad))
            linked_set["id"] = new_set_id
            set_current += 1
            remapped_set_count += 1

            counter = 0
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, Mapping):
                    continue
                old_led_id = (led.get("id") or "").strip()
                if led.get("is_parent_anchor"):
                    new_led_id = "{}-LED-000".format(new_set_id)
                else:
                    counter += 1
                    new_led_id = "{}-LED-{:03d}".format(new_set_id, counter)
                led["id"] = new_led_id
                _rewrite_linker_payload(led.get("parameters"), old_set_id, new_set_id, old_led_id, new_led_id)
                remapped_led_count += 1

    return updated, eq_current, set_current, remapped_set_count, remapped_led_count


def main():
    init_dir = LIB_ROOT if os.path.isdir(LIB_ROOT) else None
    first_path = forms.pick_file(
        file_ext="yaml",
        title="Select the first YAML file",
        init_dir=init_dir,
    )
    if not first_path:
        return

    second_path = forms.pick_file(
        file_ext="yaml",
        title="Select the second YAML file",
        init_dir=os.path.dirname(first_path) or init_dir,
    )
    if not second_path:
        return

    try:
        first_data = _load_yaml(first_path)
        second_data = _load_yaml(second_path)
    except Exception as exc:
        forms.alert("Failed to read YAML:\n\n{}".format(exc), title=TITLE)
        return

    first_defs = list(first_data.get("equipment_definitions") or [])
    second_defs = list(second_data.get("equipment_definitions") or [])

    max_num, prefix_counts, max_pad = _collect_id_stats(first_defs)
    if max_pad <= 0:
        max_pad = 3
    default_prefix = _pick_default_prefix(prefix_counts)
    if not default_prefix:
        _, second_prefix_counts, second_pad = _collect_id_stats(second_defs)
        default_prefix = _pick_default_prefix(second_prefix_counts) or "EQ-"
        if max_pad <= 0 and second_pad:
            max_pad = second_pad

    max_set_num, set_prefix_counts, max_set_pad = _collect_set_id_stats(first_defs)
    if max_set_pad <= 0:
        max_set_pad = 3
    default_set_prefix = _pick_default_prefix(set_prefix_counts)
    if not default_set_prefix:
        _, second_set_prefix_counts, second_set_pad = _collect_set_id_stats(second_defs)
        default_set_prefix = _pick_default_prefix(second_set_prefix_counts) or "SET-"
        if max_set_pad <= 0 and second_set_pad:
            max_set_pad = second_set_pad

    start_num = max_num + 1
    start_set_num = max_set_num + 1
    renumbered_second, next_num, next_set_num, set_count, led_count = _renumber_defs(
        second_defs, start_num, default_prefix, max_pad, start_set_num, default_set_prefix, max_set_pad
    )

    combined = {"equipment_definitions": first_defs + renumbered_second}
    combined_text = dump_data_to_string(combined)

    default_name = "combined_profiles.yaml"
    save_path = forms.save_file(
        file_ext="yaml",
        title=TITLE,
        default_name=default_name,
    )
    if not save_path:
        return

    try:
        with io.open(save_path, "w", encoding="utf-8") as handle:
            handle.write(combined_text)
    except Exception as exc:
        forms.alert("Failed to save combined YAML:\n\n{}".format(exc), title=TITLE)
        return

    summary = [
        "Combined YAML saved to:",
        save_path,
        "",
        "First file entries: {}".format(len(first_defs)),
        "Second file entries: {}".format(len(second_defs)),
        "Renumbered profile IDs: {}{} to {}{}".format(
            default_prefix,
            str(start_num).zfill(max_pad),
            default_prefix,
            str(next_num - 1).zfill(max_pad),
        ),
        "Renumbered linked set IDs: {}{} to {}{}".format(
            default_set_prefix,
            str(start_set_num).zfill(max_set_pad),
            default_set_prefix,
            str(next_set_num - 1).zfill(max_set_pad),
        ),
        "Updated {} linked sets and {} linked element IDs from second file.".format(set_count, led_count),
        "Truth-source IDs in second file were remapped to the new profile IDs.",
    ]
    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
