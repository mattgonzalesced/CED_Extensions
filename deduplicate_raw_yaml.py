# -*- coding: utf-8 -*-
"""
SAFE YAML Deduplication Script - Raw Parser Version
----------------------------------------------------
Manually parses YAML to preserve ALL entries (even duplicates that PyYAML would drop).
Then deduplicates by NAME and assigns unique IDs.
"""

import io
import os
import sys
import re

try:
    import yaml
except ImportError:
    print("ERROR: PyYAML required. Install: pip install pyyaml")
    sys.exit(1)

# Canonical key orders
EQUIPMENT_KEY_ORDER = [
    "id", "name", "version", "schema_version", "allow_parentless",
    "allow_unmatched_parents", "prompt_on_parent_mismatch",
    "parent_filter", "equipment_properties", "linked_sets",
]

PARENT_FILTER_KEY_ORDER = ["category", "family_name_pattern", "type_name_pattern", "parameter_filters"]
LINKED_SET_KEY_ORDER = ["id", "name", "linked_element_definitions"]
LED_KEY_ORDER = ["id", "is_parent_anchor", "is_group", "label", "category", "parameters", "tags", "text_notes", "offsets"]
OFFSET_KEY_ORDER = ["x_inches", "y_inches", "z_inches", "rotation_deg"]


def parse_raw_yaml_equipment_defs(raw_text):
    """
    Manually split the YAML into individual equipment definition blocks.
    This preserves duplicates that PyYAML would drop.
    """
    lines = raw_text.splitlines()

    equipment_blocks = []
    current_block = []
    in_equipment_list = False

    for line in lines:
        # Check if we're starting the equipment_definitions list
        if line.strip() == "equipment_definitions:":
            in_equipment_list = True
            continue

        if not in_equipment_list:
            continue

        # Check for list item at indent level 0 (relative to equipment_definitions)
        if line.startswith("- "):
            # Save previous block if exists
            if current_block:
                equipment_blocks.append("\n".join(current_block))
            # Start new block
            current_block = [line]
        elif line and not line[0].isspace() and line.strip():
            # Hit a top-level key, stop processing
            break
        else:
            # Continue current block
            if current_block:
                current_block.append(line)

    # Save last block
    if current_block:
        equipment_blocks.append("\n".join(current_block))

    return equipment_blocks


def parse_equipment_block(block_text):
    """Parse a single equipment definition block into a dict."""
    # Remove the leading "- " and parse as YAML
    if block_text.strip().startswith("- "):
        yaml_text = block_text[block_text.index("- ") + 2:]
    else:
        yaml_text = block_text

    try:
        eq_def = yaml.safe_load(yaml_text)
        return eq_def if isinstance(eq_def, dict) else {}
    except Exception as e:
        print("WARNING: Failed to parse equipment block: {}".format(e))
        print("Block preview: {}...".format(yaml_text[:100]))
        return {}


def reorder_dict(data, key_order):
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


def reorder_equipment_definition(eq_def):
    reordered = reorder_dict(eq_def, EQUIPMENT_KEY_ORDER)
    if "parent_filter" in reordered and isinstance(reordered["parent_filter"], dict):
        reordered["parent_filter"] = reorder_dict(reordered["parent_filter"], PARENT_FILTER_KEY_ORDER)
    if "linked_sets" in reordered and isinstance(reordered["linked_sets"], list):
        for i, ls in enumerate(reordered["linked_sets"]):
            reordered["linked_sets"][i] = reorder_dict(ls, LINKED_SET_KEY_ORDER)
            if "linked_element_definitions" in ls and isinstance(ls["linked_element_definitions"], list):
                for j, led in enumerate(ls["linked_element_definitions"]):
                    reordered["linked_sets"][i]["linked_element_definitions"][j] = reorder_dict(led, LED_KEY_ORDER)
                    if "offsets" in led and isinstance(led["offsets"], list):
                        for k, off in enumerate(led["offsets"]):
                            reordered["linked_sets"][i]["linked_element_definitions"][j]["offsets"][k] = reorder_dict(off, OFFSET_KEY_ORDER)
    return reordered


def normalize_name(name):
    return (name or "").strip().lower()


def merge_by_name(eq_defs):
    """Group by name and merge."""
    grouped = {}
    for eq_def in eq_defs:
        name = eq_def.get("name") or eq_def.get("id") or "Unknown"
        norm_name = normalize_name(name)
        if norm_name not in grouped:
            grouped[norm_name] = []
        grouped[norm_name].append(eq_def)

    merged_defs = []
    merge_report = []

    for norm_name, group in sorted(grouped.items()):
        if len(group) == 1:
            merged_defs.append(group[0])
        else:
            # Merge
            primary = group[0].copy()
            original_name = primary.get("name") or primary.get("id") or "Unknown"

            all_leds = []
            for eq_def in group:
                for linked_set in eq_def.get("linked_sets", []):
                    for led in linked_set.get("linked_element_definitions", []):
                        all_leds.append(led)

            primary["linked_sets"] = [{
                "id": "SET-001",
                "name": "{} Types".format(original_name),
                "linked_element_definitions": all_leds
            }]

            merged_defs.append(primary)
            merge_report.append((original_name, len(group), len(all_leds)))

    return merged_defs, merge_report


def renumber_ids(merged_defs):
    for idx, eq_def in enumerate(merged_defs, 1):
        eq_id = "EQ-{:03d}".format(idx)
        eq_def["id"] = eq_id

        for set_idx, linked_set in enumerate(eq_def.get("linked_sets", []), 1):
            set_id = "SET-{:03d}".format(idx)
            linked_set["id"] = set_id

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
    print("YAML DEDUPLICATION - RAW PARSER")
    print("=" * 80)
    print()

    if not os.path.exists(input_file):
        print("ERROR: Input file not found")
        return 1

    if os.path.exists(output_file):
        response = raw_input("Output file exists. Overwrite? (yes/no): ").strip().lower()
        if response != "yes":
            print("Aborted.")
            return 1

    # Step 1: Parse raw YAML to extract all equipment definition blocks
    print("Step 1: Parsing raw YAML file...")
    with io.open(input_file, "r", encoding="utf-8") as f:
        raw_text = f.read()

    blocks = parse_raw_yaml_equipment_defs(raw_text)
    print("  Found {} equipment definition blocks in raw file".format(len(blocks)))

    # Step 2: Parse each block
    print()
    print("Step 2: Parsing individual equipment definitions...")
    eq_defs = []
    for i, block in enumerate(blocks, 1):
        eq_def = parse_equipment_block(block)
        if eq_def:
            eq_defs.append(eq_def)
    print("  Successfully parsed {} equipment definitions".format(len(eq_defs)))

    # Step 3: Analyze duplicates
    print()
    print("Step 3: Analyzing duplicates by name...")
    name_counts = {}
    for eq_def in eq_defs:
        name = eq_def.get("name") or eq_def.get("id") or "Unknown"
        norm_name = normalize_name(name)
        name_counts[norm_name] = name_counts.get(norm_name, 0) + 1

    duplicates = [(name, count) for name, count in name_counts.items() if count > 1]
    if duplicates:
        print("  Found {} names with duplicates:".format(len(duplicates)))
        for name, count in sorted(duplicates, key=lambda x: -x[1])[:10]:
            original_name = next(
                (eq.get("name") or eq.get("id") for eq in eq_defs
                 if normalize_name(eq.get("name") or eq.get("id") or "") == name),
                name
            )
            print("    '{}' x{}".format(original_name, count))
        if len(duplicates) > 10:
            print("    ... and {} more".format(len(duplicates) - 10))
    else:
        print("  No duplicates found")

    # Step 4: Merge by name
    print()
    print("Step 4: Merging duplicates by name...")
    merged_defs, merge_report = merge_by_name(eq_defs)
    print("  Merged {} entries into {} unique definitions".format(len(eq_defs), len(merged_defs)))

    if merge_report:
        print()
        print("  Merge details (showing first 10):")
        for name, orig_count, led_count in merge_report[:10]:
            print("    '{}': {} entries -> {} LEDs".format(name, orig_count, led_count))
        if len(merge_report) > 10:
            print("    ... and {} more merged".format(len(merge_report) - 10))

    # Step 5: Renumber IDs
    print()
    print("Step 5: Assigning sequential IDs...")
    merged_defs = renumber_ids(merged_defs)
    print("  Assigned EQ-001 through EQ-{:03d}".format(len(merged_defs)))

    # Step 6: Reorder keys
    print()
    print("Step 6: Reordering to canonical format...")
    final_defs = [reorder_equipment_definition(eq) for eq in merged_defs]
    print("  Reordered {} definitions".format(len(final_defs)))

    # Step 7: Write output
    print()
    print("Step 7: Writing output...")
    output_data = {"equipment_definitions": final_defs}
    with io.open(output_file, "w", encoding="utf-8") as f:
        yaml.dump(output_data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
    print("  Written to: {}".format(output_file))

    # Step 8: Verify
    print()
    print("Step 8: Verifying...")
    with io.open(output_file, "r", encoding="utf-8") as f:
        verify = yaml.safe_load(f)
    verify_defs = verify.get("equipment_definitions", [])
    ids = [eq.get("id") for eq in verify_defs]
    print("  Verified {} definitions".format(len(verify_defs)))
    print("  All IDs unique: {}".format(len(ids) == len(set(ids))))

    print()
    print("=" * 80)
    print("SUCCESS!")
    print("=" * 80)
    print()
    print("Summary:")
    print("  Original entries:     {}".format(len(eq_defs)))
    print("  After deduplication:  {}".format(len(merged_defs)))
    print("  Entries removed:      {}".format(len(eq_defs) - len(merged_defs)))
    print()

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nAborted")
        sys.exit(1)
    except Exception as e:
        print("\nERROR: {}".format(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)
