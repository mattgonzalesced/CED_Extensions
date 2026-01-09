# -*- coding: utf-8 -*-
"""
SAFE YAML Deduplication and Reordering Script
----------------------------------------------
1. Groups equipment definitions by NAME (not ID)
2. Merges duplicates by combining their linked_sets
3. Assigns unique IDs sequentially
4. Reorders keys to canonical format

SAFETY FEATURES:
- Reports what will be merged before doing it
- Writes to a NEW file
- Validates data integrity
"""

import io
import json
import os
import sys

# Add CEDLib to path
LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.insert(0, LIB_ROOT)

try:
    import yaml
except ImportError:
    print("ERROR: PyYAML is required. Install with: pip install pyyaml")
    sys.exit(1)

# Define canonical key order
EQUIPMENT_KEY_ORDER = [
    "id", "name", "version", "schema_version", "allow_parentless",
    "allow_unmatched_parents", "prompt_on_parent_mismatch",
    "parent_filter", "equipment_properties", "linked_sets",
]

PARENT_FILTER_KEY_ORDER = [
    "category", "family_name_pattern", "type_name_pattern", "parameter_filters",
]

LINKED_SET_KEY_ORDER = ["id", "name", "linked_element_definitions"]

LED_KEY_ORDER = [
    "id", "is_parent_anchor", "is_group", "label", "category",
    "parameters", "tags", "text_notes", "offsets",
]

OFFSET_KEY_ORDER = ["x_inches", "y_inches", "z_inches", "rotation_deg"]


def reorder_dict(data, key_order):
    """Reorder dictionary keys, preserving extra keys at end."""
    if not isinstance(data, dict):
        return data
    ordered = {}
    for key in key_order:
        if key in data:
            ordered[key] = data[key]
    for key in data:
        if key not in ordered:
            ordered[key] = data[key]
    return ordered


def reorder_offset(offset):
    return reorder_dict(offset, OFFSET_KEY_ORDER)


def reorder_led(led):
    reordered = reorder_dict(led, LED_KEY_ORDER)
    if "offsets" in reordered and isinstance(reordered["offsets"], list):
        reordered["offsets"] = [reorder_offset(o) for o in reordered["offsets"]]
    return reordered


def reorder_linked_set(linked_set):
    reordered = reorder_dict(linked_set, LINKED_SET_KEY_ORDER)
    if "linked_element_definitions" in reordered and isinstance(reordered["linked_element_definitions"], list):
        reordered["linked_element_definitions"] = [
            reorder_led(led) for led in reordered["linked_element_definitions"]
        ]
    return reordered


def reorder_parent_filter(parent_filter):
    return reorder_dict(parent_filter, PARENT_FILTER_KEY_ORDER)


def reorder_equipment_definition(eq_def):
    reordered = reorder_dict(eq_def, EQUIPMENT_KEY_ORDER)
    if "parent_filter" in reordered and isinstance(reordered["parent_filter"], dict):
        reordered["parent_filter"] = reorder_parent_filter(reordered["parent_filter"])
    if "linked_sets" in reordered and isinstance(reordered["linked_sets"], list):
        reordered["linked_sets"] = [reorder_linked_set(ls) for ls in reordered["linked_sets"]]
    return reordered


def normalize_name(name):
    """Normalize name for comparison."""
    return (name or "").strip().lower()


def merge_equipment_definitions(eq_defs):
    """
    Merge equipment definitions with the same name.

    Returns:
        - merged_defs: List of merged equipment definitions
        - merge_report: List of tuples (name, original_count, merged_count)
    """
    # Group by normalized name
    grouped = {}
    for eq_def in eq_defs:
        name = eq_def.get("name") or eq_def.get("id") or "Unknown"
        norm_name = normalize_name(name)
        if norm_name not in grouped:
            grouped[norm_name] = []
        grouped[norm_name].append(eq_def)

    # Merge groups
    merged_defs = []
    merge_report = []

    for norm_name, group in sorted(grouped.items()):
        if len(group) == 1:
            # No duplicates, just use as-is
            merged_defs.append(group[0])
        else:
            # Merge duplicates
            primary = group[0].copy()
            original_name = primary.get("name") or primary.get("id") or "Unknown"

            # Collect all linked_sets from all duplicates
            all_linked_sets = []
            for eq_def in group:
                for linked_set in eq_def.get("linked_sets", []):
                    all_linked_sets.append(linked_set)

            # Combine all LEDs into one linked_set
            all_leds = []
            for linked_set in all_linked_sets:
                for led in linked_set.get("linked_element_definitions", []):
                    all_leds.append(led)

            # Create merged linked_set
            primary["linked_sets"] = [{
                "id": "SET-001",  # Will be renumbered later
                "name": "{} Types".format(original_name),
                "linked_element_definitions": all_leds
            }]

            merged_defs.append(primary)
            merge_report.append((original_name, len(group), len(all_leds)))

    return merged_defs, merge_report


def renumber_ids(merged_defs):
    """Assign sequential IDs to equipment definitions and their children."""
    for idx, eq_def in enumerate(merged_defs, 1):
        eq_id = "EQ-{:03d}".format(idx)
        eq_def["id"] = eq_id

        # Renumber linked_sets
        for set_idx, linked_set in enumerate(eq_def.get("linked_sets", []), 1):
            set_id = "SET-{:03d}".format(idx)
            linked_set["id"] = set_id

            # Renumber LEDs
            led_counter = 1
            for led in linked_set.get("linked_element_definitions", []):
                if not led.get("is_parent_anchor"):
                    led["id"] = "{}-LED-{:03d}".format(set_id, led_counter)
                    led_counter += 1

    return merged_defs


def main():
    input_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved.yaml"
    output_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_DEDUPLICATED.yaml"

    print("=" * 80)
    print("SAFE YAML DEDUPLICATION AND REORDERING SCRIPT")
    print("=" * 80)
    print()
    print("Input file:  {}".format(input_file))
    print("Output file: {}".format(output_file))
    print()

    # Check if input file exists
    if not os.path.exists(input_file):
        print("ERROR: Input file does not exist!")
        return 1

    # Check if output file already exists
    if os.path.exists(output_file):
        print("WARNING: Output file already exists!")
        response = raw_input("Overwrite? (yes/no): ").strip().lower()
        if response != "yes":
            print("Aborted.")
            return 1
        print()

    # Step 1: Load the raw YAML text and parse it manually to catch all duplicates
    print("Step 1: Loading raw YAML data...")
    try:
        with io.open(input_file, "r", encoding="utf-8") as f:
            raw_text = f.read()

        # Parse YAML - this will be a list under equipment_definitions key
        # Even duplicates will be in the list
        data = yaml.safe_load(raw_text)
        eq_defs = data.get("equipment_definitions", [])

        print("  SUCCESS: Loaded {} equipment definitions from YAML".format(len(eq_defs)))
    except Exception as e:
        print("  ERROR: Failed to load data: {}".format(e))
        return 1

    # Step 2: Analyze duplicates
    print()
    print("Step 2: Analyzing duplicates by name...")
    name_counts = {}
    for eq_def in eq_defs:
        name = eq_def.get("name") or eq_def.get("id") or "Unknown"
        norm_name = normalize_name(name)
        name_counts[norm_name] = name_counts.get(norm_name, 0) + 1

    duplicates = [(name, count) for name, count in name_counts.items() if count > 1]
    if duplicates:
        print("  Found {} equipment names with duplicates:".format(len(duplicates)))
        for name, count in sorted(duplicates):
            # Find original name (not normalized)
            original_name = next(
                (eq.get("name") or eq.get("id") for eq in eq_defs
                 if normalize_name(eq.get("name") or eq.get("id") or "") == name),
                name
            )
            print("    '{}' appears {} times".format(original_name, count))
    else:
        print("  No duplicate names found")

    # Step 3: Merge duplicates
    print()
    print("Step 3: Merging equipment definitions by name...")
    try:
        merged_defs, merge_report = merge_equipment_definitions(eq_defs)
        print("  SUCCESS: Merged {} definitions into {} unique definitions".format(
            len(eq_defs), len(merged_defs)
        ))
        if merge_report:
            print()
            print("  Merge details:")
            for name, original_count, led_count in merge_report:
                print("    '{}': merged {} entries into {} LEDs".format(
                    name, original_count, led_count
                ))
    except Exception as e:
        print("  ERROR: Failed to merge: {}".format(e))
        import traceback
        traceback.print_exc()
        return 1

    # Step 4: Renumber IDs
    print()
    print("Step 4: Renumbering IDs sequentially...")
    try:
        merged_defs = renumber_ids(merged_defs)
        print("  SUCCESS: Assigned IDs EQ-001 through EQ-{:03d}".format(len(merged_defs)))
    except Exception as e:
        print("  ERROR: Failed to renumber: {}".format(e))
        return 1

    # Step 5: Reorder keys
    print()
    print("Step 5: Reordering keys to canonical format...")
    try:
        reordered_defs = [reorder_equipment_definition(eq) for eq in merged_defs]
        print("  SUCCESS: Reordered {} equipment definitions".format(len(reordered_defs)))
    except Exception as e:
        print("  ERROR: Failed to reorder: {}".format(e))
        return 1

    # Step 6: Write output
    print()
    print("Step 6: Writing output file...")
    try:
        output_data = {"equipment_definitions": reordered_defs}
        with io.open(output_file, "w", encoding="utf-8") as f:
            yaml.dump(output_data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
        print("  SUCCESS: Wrote output file")
    except Exception as e:
        print("  ERROR: Failed to write: {}".format(e))
        return 1

    # Step 7: Verify output
    print()
    print("Step 7: Verifying output file...")
    try:
        with io.open(output_file, "r", encoding="utf-8") as f:
            verify_data = yaml.safe_load(f)
        verify_defs = verify_data.get("equipment_definitions", [])
        print("  SUCCESS: Verified {} equipment definitions in output".format(len(verify_defs)))

        # Check for duplicate IDs
        ids = [eq.get("id") for eq in verify_defs]
        if len(ids) != len(set(ids)):
            print("  WARNING: Output has duplicate IDs!")
            return 1
        print("  SUCCESS: All IDs are unique")
    except Exception as e:
        print("  ERROR: Failed to verify: {}".format(e))
        return 1

    # Success!
    print()
    print("=" * 80)
    print("SUCCESS!")
    print("=" * 80)
    print()
    print("Summary:")
    print("  Original entries: {}".format(len(eq_defs)))
    print("  Merged entries:   {}".format(len(merged_defs)))
    print("  Duplicates removed: {}".format(len(eq_defs) - len(merged_defs)))
    print()
    print("Output written to:")
    print("  {}".format(output_file))
    print()
    print("NEXT STEPS:")
    print("1. Review the output file")
    print("2. Test loading it in the MEP Automation panel")
    print("3. If satisfied, replace the original file")
    print()

    return 0


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print()
        print("Aborted by user.")
        sys.exit(1)
    except Exception as e:
        print()
        print("UNEXPECTED ERROR: {}".format(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)
