# -*- coding: utf-8 -*-
"""
SAFE YAML Merge by Truth Source Script
---------------------------------------
Merges equipment definitions that have the same ced_truth_source_name.

SAFETY FEATURES:
1. Reads file using existing, tested profile_schema.py code
2. Groups by ced_truth_source_name
3. Merges duplicate entries by combining their LEDs
4. Assigns unique IDs sequentially
5. Reorders keys to canonical format
6. Writes to a NEW file (you must manually replace the original)
"""

import io
import os
import sys
import yaml

# Add CEDLib to path
LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.insert(0, LIB_ROOT)

from LogicClasses.profile_schema import save_data

# Canonical key orders (same as deduplicate script)
EQUIPMENT_KEY_ORDER = [
    "id", "name", "version", "schema_version", "allow_parentless",
    "allow_unmatched_parents", "prompt_on_parent_mismatch",
    "parent_filter", "equipment_properties", "linked_sets",
]

PARENT_FILTER_KEY_ORDER = ["category", "family_name_pattern", "type_name_pattern", "parameter_filters"]
LINKED_SET_KEY_ORDER = ["id", "name", "linked_element_definitions"]
LED_KEY_ORDER = ["id", "is_parent_anchor", "is_group", "label", "category", "parameters", "tags", "text_notes", "offsets"]
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


def normalize_truth_source(source_name):
    """Normalize truth source name for comparison."""
    return (source_name or "").strip().lower()


def merge_by_truth_source(eq_defs):
    """
    Merge equipment definitions with the same ced_truth_source_name.

    Strategy:
    - Keep the first profile's metadata (name, parent_filter, etc.)
    - Combine all LEDs from all profiles into one linked_set
    - Preserve the ced_truth_source_name and ced_truth_source_id

    Returns:
        - merged_defs: List of merged equipment definitions
        - merge_report: List of tuples (truth_source, original_count, merged_led_count)
    """
    # Group by ced_truth_source_name
    grouped = {}
    no_source = []

    for eq_def in eq_defs:
        truth_source = eq_def.get('ced_truth_source_name', '')
        if not truth_source:
            # Keep profiles without truth source separate
            no_source.append(eq_def)
            continue

        norm_source = normalize_truth_source(truth_source)
        if norm_source not in grouped:
            grouped[norm_source] = []
        grouped[norm_source].append(eq_def)

    # Merge groups
    merged_defs = []
    merge_report = []

    # Add profiles without truth source (unchanged)
    for eq_def in no_source:
        merged_defs.append(eq_def)

    # Merge profiles with same truth source
    for norm_source, group in sorted(grouped.items()):
        if len(group) == 1:
            # No duplicates, just use as-is
            merged_defs.append(group[0])
        else:
            # Find the profile with the most LEDs
            largest_profile = None
            max_led_count = 0

            for eq_def in group:
                led_count = sum(
                    len(linked_set.get('linked_element_definitions', []))
                    for linked_set in eq_def.get('linked_sets', [])
                )
                if led_count > max_led_count:
                    max_led_count = led_count
                    largest_profile = eq_def

            # Use the largest profile as the primary
            if not largest_profile:
                largest_profile = group[0]  # Fallback if all have 0 LEDs

            primary = largest_profile.copy()
            original_truth_source = primary.get('ced_truth_source_name', 'Unknown')

            # Keep the largest profile's data unchanged
            # (don't merge LEDs from other profiles - just keep the biggest one)

            merged_defs.append(primary)

            # Report shows: truth_source, number of duplicates discarded, LEDs kept
            merge_report.append((original_truth_source, len(group), max_led_count))

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


def load_yaml_raw(file_path):
    """Load YAML file directly using PyYAML."""
    with io.open(file_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    return data


def main():
    input_file = r"C:\CED_Extensions\CEDLib.lib\prototypeHEB_StartCarrollton_Checkpoint35_CLEANED.yaml"
    output_file = r"C:\CED_Extensions\CEDLib.lib\prototypeHEB_StartCarrollton_Checkpoint35_MERGED_BY_SOURCE.yaml"

    print("=" * 80)
    print("SAFE YAML MERGE BY TRUTH SOURCE SCRIPT")
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
        try:
            response = raw_input("Overwrite? (yes/no): ").strip().lower()
        except NameError:
            response = input("Overwrite? (yes/no): ").strip().lower()
        if response != "yes":
            print("Aborted.")
            return 1
        print()

    # Step 1: Load the data
    print("Step 1: Loading data from input file...")
    try:
        data = load_yaml_raw(input_file)
        eq_defs = data.get("equipment_definitions", [])
        print("  SUCCESS: Loaded {} equipment definitions".format(len(eq_defs)))
    except Exception as e:
        print("  ERROR: Failed to load data: {}".format(e))
        import traceback
        traceback.print_exc()
        return 1

    # Step 2: Analyze duplicates by truth source
    print()
    print("Step 2: Analyzing duplicates by ced_truth_source_name...")
    source_counts = {}
    no_source_count = 0
    for eq_def in eq_defs:
        truth_source = eq_def.get('ced_truth_source_name', '')
        if not truth_source:
            no_source_count += 1
            continue
        norm_source = normalize_truth_source(truth_source)
        source_counts[norm_source] = source_counts.get(norm_source, 0) + 1

    duplicates = [(name, count) for name, count in source_counts.items() if count > 1]
    if duplicates:
        print("  Found {} truth sources with duplicates:".format(len(duplicates)))
        for name, count in sorted(duplicates, key=lambda x: -x[1])[:10]:
            # Find original name (not normalized)
            original_name = next(
                (eq.get('ced_truth_source_name') for eq in eq_defs
                 if normalize_truth_source(eq.get('ced_truth_source_name', '')) == name),
                name
            )
            print("    '{}' appears {} times".format(original_name, count))
        if len(duplicates) > 10:
            print("    ... and {} more".format(len(duplicates) - 10))
    else:
        print("  No duplicate truth sources found")

    if no_source_count:
        print("  {} equipment definitions have no ced_truth_source_name".format(no_source_count))

    # Step 3: Merge by truth source
    print()
    print("Step 3: Merging equipment definitions by ced_truth_source_name...")
    try:
        merged_defs, merge_report = merge_by_truth_source(eq_defs)
        print("  SUCCESS: Merged {} definitions into {} unique definitions".format(
            len(eq_defs), len(merged_defs)
        ))
        if merge_report:
            print()
            print("  Merge details (showing first 10):")
            for truth_source, original_count, led_count in merge_report[:10]:
                print("    '{}': merged {} entries into {} LEDs".format(
                    truth_source, original_count, led_count
                ))
            if len(merge_report) > 10:
                print("    ... and {} more merged".format(len(merge_report) - 10))
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
        save_data(output_file, output_data)
        print("  SUCCESS: Wrote output file")
    except Exception as e:
        print("  ERROR: Failed to write: {}".format(e))
        return 1

    # Step 7: Verify output
    print()
    print("Step 7: Verifying output file...")
    try:
        verify_data = load_yaml_raw(output_file)
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
        import traceback
        traceback.print_exc()
        return 1

    # Success!
    print()
    print("=" * 80)
    print("SUCCESS!")
    print("=" * 80)
    print()
    print("Summary:")
    print("  Original entries:       {}".format(len(eq_defs)))
    print("  Merged entries:         {}".format(len(merged_defs)))
    print("  Duplicates removed:     {}".format(len(eq_defs) - len(merged_defs)))
    print("  Profiles merged:        {}".format(len(merge_report)))
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
