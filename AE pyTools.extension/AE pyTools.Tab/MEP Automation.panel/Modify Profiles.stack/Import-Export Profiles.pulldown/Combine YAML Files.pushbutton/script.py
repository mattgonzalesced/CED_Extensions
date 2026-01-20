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

from pyrevit import forms

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.profile_schema import load_data_from_text, dump_data_to_string  # noqa: E402

TITLE = "Combine YAML Files"
ID_PATTERN = re.compile(r"^(.*?)(\d+)$")


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


def _renumber_defs(defs, start_num, default_prefix, default_pad):
    updated = copy.deepcopy(defs)
    current = start_num
    for entry in updated:
        if not isinstance(entry, Mapping):
            continue
        prefix, _, pad = _split_id(entry.get("id"))
        if prefix is None:
            prefix = default_prefix
        if pad is None or pad <= 0:
            pad = default_pad
        entry["id"] = "{}{}".format(prefix, str(current).zfill(pad))
        current += 1
    return updated, current


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

    start_num = max_num + 1
    renumbered_second, next_num = _renumber_defs(
        second_defs, start_num, default_prefix, max_pad
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
        "Renumbered IDs start: {}{}".format(default_prefix, str(start_num).zfill(max_pad)),
        "Renumbered IDs end: {}{}".format(default_prefix, str(next_num - 1).zfill(max_pad)),
    ]
    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
