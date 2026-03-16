# -*- coding: utf-8 -*-
"""
Audit CAD model 2.0 (batman)
----------------------------
Process a CSV file containing CAD block names and create or merge equipment profiles.
"""

import csv
import copy
import json
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

TITLE = "Audit CAD model 2.0 (batman)"
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
YAML_SCHEMA_VERSION = 4

try:
    basestring
except NameError:
    basestring = str


# --------------------------------------------------------------------------- #
# Ignore list management
# --------------------------------------------------------------------------- #


def _ignore_list_file():
    """Get path to the ignore list JSON file."""
    try:
        return script.get_appdata_file("cad_audit_ignore_list.json")
    except Exception:
        return os.path.join(os.path.expanduser("~"), "cad_audit_ignore_list.json")


def _load_ignore_list():
    """Load the ignore list from JSON file."""
    path = _ignore_list_file()
    if os.path.exists(path):
        try:
            with open(path, "r") as handle:
                data = json.load(handle)
                if isinstance(data, list):
                    return set(data)
        except Exception:
            return set()
    return set()


def _save_ignore_list(ignore_set):
    """Save the ignore list to JSON file."""
    try:
        directory = os.path.dirname(_ignore_list_file())
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        with open(_ignore_list_file(), "w") as handle:
            json.dump(sorted(list(ignore_set)), handle, indent=2)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Equipment definition helpers
# --------------------------------------------------------------------------- #


def _next_eq_number(equipment_defs):
    """Get next available equipment number."""
    max_id = 0
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        if not eq_id:
            continue
        suffix = eq_id.split("-")[-1]
        try:
            num = int(suffix)
        except Exception:
            continue
        if num > max_id:
            max_id = num
    return max_id + 1


def _create_empty_profile(cad_name, seq):
    """Create an empty equipment profile with default structure."""
    eq_id = "EQ-{:03d}".format(seq)
    set_id = "SET-{:03d}".format(seq)
    return {
        "id": eq_id,
        "name": cad_name,
        "version": 1,
        "schema_version": YAML_SCHEMA_VERSION,
        "allow_parentless": True,
        "allow_unmatched_parents": True,
        "prompt_on_parent_mismatch": False,
        "parent_filter": {
            "category": "Uncategorized",
            "family_name_pattern": "*",
            "type_name_pattern": "*",
            "parameter_filters": {},
        },
        "equipment_properties": {},
        "linked_sets": [
            {
                "id": set_id,
                "name": "{} Types".format(cad_name),
                "linked_element_definitions": [],
            }
        ],
    }


def _find_equipment_definition_by_name(equipment_defs, name):
    """Find equipment definition by name."""
    for eq in equipment_defs or []:
        if (eq.get("name") or eq.get("id")) == name:
            return eq
    return None


def _copy_fields(source_entry, target_entry):
    """Copy everything except identifying fields (name, id)."""
    keep_keys = {"name", "id"}
    for key in list(target_entry.keys()):
        if key in keep_keys:
            continue
        target_entry.pop(key, None)
    for key, value in source_entry.items():
        if key in keep_keys:
            continue
        target_entry[key] = copy.deepcopy(value)


def _create_merged_profile(cad_name, source_entry, seq):
    """Create a new profile merged from source."""
    eq_id = "EQ-{:03d}".format(seq)
    new_profile = _create_empty_profile(cad_name, seq)
    new_profile["id"] = eq_id
    _copy_fields(source_entry, new_profile)
    source_id = source_entry.get("id") or source_entry.get("name") or ""
    source_name = source_entry.get("name") or source_entry.get("id") or ""
    if source_id:
        new_profile[TRUTH_SOURCE_ID_KEY] = source_id
    if source_name:
        new_profile[TRUTH_SOURCE_NAME_KEY] = source_name
    return new_profile


# --------------------------------------------------------------------------- #
# CSV processing
# --------------------------------------------------------------------------- #


def _read_csv_names(csv_path):
    """Read unique names from CSV file (second column with header 'Name')."""
    names_set = set()
    names = []
    try:
        with open(csv_path, "r") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                name = row.get("Name", "").strip()
                if name and name not in names_set:
                    names_set.add(name)
                    names.append(name)
    except Exception as exc:
        raise RuntimeError("Failed to read CSV file: {}".format(str(exc)))
    return names


# --------------------------------------------------------------------------- #
# Main logic
# --------------------------------------------------------------------------- #


def _build_definition_map(equipment_defs):
    """Build a map of definition names to entries."""
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


def process_csv():
    """Main processing function."""
    # Load YAML data
    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    equipment_defs = data.setdefault("equipment_definitions", [])
    def_map, ordered_names = _build_definition_map(equipment_defs)

    # Load ignore list
    ignore_list = _load_ignore_list()

    # Select CSV file
    csv_path = forms.pick_file(
        file_ext="csv",
        title="Select CSV file with CAD block names"
    )
    if not csv_path:
        return

    # Read CSV names
    try:
        csv_names = _read_csv_names(csv_path)
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    if not csv_names:
        forms.alert("No names found in CSV file.", title=TITLE)
        return

    # Process each name
    created = []
    merged = []
    skipped = []
    ignored = []

    for cad_name in csv_names:
        # Check if in ignore list
        if cad_name in ignore_list:
            ignored.append(cad_name)
            continue

        # Check if already exists
        existing = _find_equipment_definition_by_name(equipment_defs, cad_name)
        if existing:
            skipped.append(cad_name)
            continue

        # Ask user what to do
        choices = [
            "Create new empty profile",
            "Merge into existing profile",
            "Skip this entry",
            "Add to ignore list (skip permanently)"
        ]
        choice = forms.CommandSwitchWindow.show(
            choices,
            message="Profile '{}' not found.\n\nWhat would you like to do?".format(cad_name)
        )

        if not choice or choice == "Skip this entry":
            skipped.append(cad_name)
            continue

        if choice == "Add to ignore list (skip permanently)":
            ignore_list.add(cad_name)
            ignored.append(cad_name)
            continue

        if choice == "Create new empty profile":
            # Create empty profile
            seq = _next_eq_number(equipment_defs)
            new_profile = _create_empty_profile(cad_name, seq)
            equipment_defs.append(new_profile)
            created.append(cad_name)
            continue

        if choice == "Merge into existing profile":
            # Select source profile
            source_choice = forms.SelectFromList.show(
                ordered_names,
                title="Select source profile (truth) for '{}'".format(cad_name),
                multiselect=False,
                button_name="Select"
            )
            if not source_choice:
                skipped.append(cad_name)
                continue

            source_name = source_choice if isinstance(source_choice, basestring) else source_choice[0]
            source_entry = def_map.get(source_name)
            if not source_entry:
                forms.alert(
                    "Could not resolve the selected source definition.",
                    title=TITLE
                )
                skipped.append(cad_name)
                continue

            # Create merged profile
            seq = _next_eq_number(equipment_defs)
            new_profile = _create_merged_profile(cad_name, source_entry, seq)
            equipment_defs.append(new_profile)
            merged.append("{} -> {}".format(cad_name, source_name))
            continue

    # Save ignore list
    _save_ignore_list(ignore_list)

    # Save YAML data if changes were made
    if created or merged:
        save_active_yaml_data(
            None,
            data,
            "Audit CAD model 2.0",
            "Processed CSV: {} created, {} merged, {} skipped, {} ignored".format(
                len(created), len(merged), len(skipped), len(ignored)
            ),
        )

    # Display summary
    yaml_label = get_yaml_display_name(yaml_path)
    summary_lines = [
        "CSV processing complete.",
        "",
        "Created: {}".format(len(created)),
        "Merged: {}".format(len(merged)),
        "Skipped (already exist): {}".format(len(skipped)),
        "Ignored (permanent): {}".format(len(ignored)),
        "",
    ]

    if created:
        summary_lines.append("Created profiles:")
        for name in created:
            summary_lines.append("  - {}".format(name))
        summary_lines.append("")

    if merged:
        summary_lines.append("Merged profiles:")
        for item in merged:
            summary_lines.append("  - {}".format(item))
        summary_lines.append("")

    if ignored:
        summary_lines.append("Added to ignore list:")
        for name in ignored:
            summary_lines.append("  - {}".format(name))
        summary_lines.append("")

    if created or merged:
        summary_lines.append("Updated data saved to {}.".format(yaml_label))

    forms.alert("\n".join(summary_lines), title=TITLE)


if __name__ == "__main__":
    process_csv()
