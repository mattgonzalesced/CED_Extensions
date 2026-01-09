# -*- coding: utf-8 -*-
"""
SAFE YAML Reordering Script
----------------------------
Reorders equipment definitions in a YAML file to match the canonical order
without losing any data.

SAFETY FEATURES:
1. Reads file using existing, tested profile_schema.py code
2. Only reorders keys, never modifies values
3. Preserves all keys, even non-canonical ones
4. Validates data before and after
5. Writes to a NEW file (you must manually replace the original)
"""

import io
import json
import os
import sys

# Add CEDLib to path
LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.insert(0, LIB_ROOT)

# Try to use profile_schema, fall back to PyYAML directly
try:
    from LogicClasses.profile_schema import load_data, save_data
    print("Using profile_schema for YAML I/O")
except ImportError:
    # Fallback: use PyYAML directly if profile_schema can't be imported
    try:
        import yaml

        def load_data(path):
            with io.open(path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f)
                return data if data else {"equipment_definitions": []}

        def save_data(path, data):
            with io.open(path, "w", encoding="utf-8") as f:
                yaml.dump(data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)

        print("Using PyYAML directly for YAML I/O")
    except ImportError:
        print("ERROR: Cannot import required YAML libraries")
        sys.exit(1)

# Define canonical key order for equipment definitions
EQUIPMENT_KEY_ORDER = [
    "id",
    "name",
    "version",
    "schema_version",
    "allow_parentless",
    "allow_unmatched_parents",
    "prompt_on_parent_mismatch",
    "parent_filter",
    "equipment_properties",
    "linked_sets",
]

# Define canonical key order for parent_filter
PARENT_FILTER_KEY_ORDER = [
    "category",
    "family_name_pattern",
    "type_name_pattern",
    "parameter_filters",
]

# Define canonical key order for linked_sets
LINKED_SET_KEY_ORDER = [
    "id",
    "name",
    "linked_element_definitions",
]

# Define canonical key order for linked_element_definitions
LED_KEY_ORDER = [
    "id",
    "is_parent_anchor",
    "is_group",
    "label",
    "category",
    "parameters",
    "tags",
    "text_notes",
    "offsets",
]

# Define canonical key order for offsets
OFFSET_KEY_ORDER = [
    "x_inches",
    "y_inches",
    "z_inches",
    "rotation_deg",
]


def reorder_dict(data, key_order):
    """
    Reorder dictionary keys according to key_order, preserving extra keys at the end.

    Args:
        data: Dictionary to reorder
        key_order: List of keys in desired order

    Returns:
        New dictionary with keys in order
    """
    if not isinstance(data, dict):
        return data

    ordered = {}

    # First, add keys in canonical order
    for key in key_order:
        if key in data:
            ordered[key] = data[key]

    # Then add any extra keys that weren't in the canonical order
    for key in data:
        if key not in ordered:
            ordered[key] = data[key]

    return ordered


def reorder_offset(offset):
    """Reorder an offset dictionary."""
    return reorder_dict(offset, OFFSET_KEY_ORDER)


def reorder_led(led):
    """Reorder a linked_element_definition dictionary."""
    reordered = reorder_dict(led, LED_KEY_ORDER)

    # Reorder offsets if present
    if "offsets" in reordered and isinstance(reordered["offsets"], list):
        reordered["offsets"] = [reorder_offset(o) for o in reordered["offsets"]]

    return reordered


def reorder_linked_set(linked_set):
    """Reorder a linked_set dictionary."""
    reordered = reorder_dict(linked_set, LINKED_SET_KEY_ORDER)

    # Reorder linked_element_definitions if present
    if "linked_element_definitions" in reordered and isinstance(reordered["linked_element_definitions"], list):
        reordered["linked_element_definitions"] = [
            reorder_led(led) for led in reordered["linked_element_definitions"]
        ]

    return reordered


def reorder_parent_filter(parent_filter):
    """Reorder a parent_filter dictionary."""
    return reorder_dict(parent_filter, PARENT_FILTER_KEY_ORDER)


def reorder_equipment_definition(eq_def):
    """Reorder an equipment definition dictionary."""
    reordered = reorder_dict(eq_def, EQUIPMENT_KEY_ORDER)

    # Reorder parent_filter if present
    if "parent_filter" in reordered and isinstance(reordered["parent_filter"], dict):
        reordered["parent_filter"] = reorder_parent_filter(reordered["parent_filter"])

    # Reorder linked_sets if present
    if "linked_sets" in reordered and isinstance(reordered["linked_sets"], list):
        reordered["linked_sets"] = [reorder_linked_set(ls) for ls in reordered["linked_sets"]]

    return reordered


def count_keys_recursive(data):
    """Count all keys recursively in a nested data structure."""
    if isinstance(data, dict):
        count = len(data)
        for value in data.values():
            count += count_keys_recursive(value)
        return count
    elif isinstance(data, list):
        count = 0
        for item in data:
            count += count_keys_recursive(item)
        return count
    return 0


def validate_data_integrity(original, reordered):
    """
    Validate that reordering didn't lose any data.

    Args:
        original: Original data structure
        reordered: Reordered data structure

    Returns:
        (is_valid, error_message)
    """
    # Convert both to JSON for comparison (order-independent)
    try:
        original_json = json.dumps(original, sort_keys=True)
        reordered_json = json.dumps(reordered, sort_keys=True)
    except Exception as e:
        return False, "Failed to serialize data for comparison: {}".format(e)

    if original_json != reordered_json:
        return False, "Data content changed during reordering"

    # Count keys to ensure nothing was lost
    original_count = count_keys_recursive(original)
    reordered_count = count_keys_recursive(reordered)

    if original_count != reordered_count:
        return False, "Key count mismatch: {} -> {}".format(original_count, reordered_count)

    return True, "Data integrity validated"


def main():
    input_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved.yaml"
    output_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved_REORDERED.yaml"

    print("=" * 80)
    print("SAFE YAML REORDERING SCRIPT")
    print("=" * 80)
    print("")
    print("Input file:  {}".format(input_file))
    print("Output file: {}".format(output_file))
    print("")

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
        print("")

    # Step 1: Load the data
    print("Step 1: Loading data from input file...")
    try:
        data = load_data(input_file)
        eq_defs = data.get("equipment_definitions", [])
        print("  SUCCESS: Loaded {} equipment definitions".format(len(eq_defs)))
    except Exception as e:
        print("  ERROR: Failed to load data: {}".format(e))
        return 1

    # Step 2: Reorder the data
    print("")
    print("Step 2: Reordering equipment definitions...")
    try:
        reordered_defs = [reorder_equipment_definition(eq_def) for eq_def in eq_defs]
        reordered_data = {"equipment_definitions": reordered_defs}
        print("  SUCCESS: Reordered {} equipment definitions".format(len(reordered_defs)))
    except Exception as e:
        print("  ERROR: Failed to reorder data: {}".format(e))
        return 1

    # Step 3: Validate data integrity
    print("")
    print("Step 3: Validating data integrity...")
    is_valid, message = validate_data_integrity(data, reordered_data)
    if not is_valid:
        print("  ERROR: Validation failed - {}".format(message))
        print("  ABORTED - No files were written")
        return 1
    print("  SUCCESS: {}".format(message))

    # Step 4: Write to output file
    print("")
    print("Step 4: Writing reordered data to output file...")
    try:
        save_data(output_file, reordered_data)
        print("  SUCCESS: Wrote reordered data to output file")
    except Exception as e:
        print("  ERROR: Failed to write output file: {}".format(e))
        return 1

    # Step 5: Verify output file can be loaded
    print("")
    print("Step 5: Verifying output file can be loaded...")
    try:
        verification_data = load_data(output_file)
        verification_defs = verification_data.get("equipment_definitions", [])
        print("  SUCCESS: Loaded {} equipment definitions from output file".format(len(verification_defs)))
    except Exception as e:
        print("  ERROR: Failed to load output file: {}".format(e))
        print("  WARNING: Output file may be corrupted!")
        return 1

    # Step 6: Final validation
    print("")
    print("Step 6: Final validation...")
    is_valid, message = validate_data_integrity(data, verification_data)
    if not is_valid:
        print("  ERROR: Final validation failed - {}".format(message))
        print("  WARNING: Output file may be corrupted!")
        return 1
    print("  SUCCESS: {}".format(message))

    # Success!
    print("")
    print("=" * 80)
    print("SUCCESS!")
    print("=" * 80)
    print("")
    print("The reordered YAML file has been written to:")
    print("  {}".format(output_file))
    print("")
    print("NEXT STEPS:")
    print("1. Review the output file to ensure it looks correct")
    print("2. If satisfied, manually rename/replace the original file:")
    print("   - Delete or rename the original: {}".format(input_file))
    print("   - Rename the output file to: {}".format(os.path.basename(input_file)))
    print("3. Test loading the file in the MEP Automation panel")
    print("")

    return 0


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("")
        print("Aborted by user.")
        sys.exit(1)
    except Exception as e:
        print("")
        print("UNEXPECTED ERROR: {}".format(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)
