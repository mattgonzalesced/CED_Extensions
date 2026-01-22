# -*- coding: utf-8 -*-
"""
Place Spaces from CSV
---------------------
Creates unbounded spaces from a CSV file with the following columns:
- Position X: X coordinate in feet-inches format (e.g., "76'-2 7/8"")
- Position Y: Y coordinate in feet-inches format
- #: Space Number
- #(1): Second part of space name
- #(2): First part of space name
Space Name = #(2) + " " + #(1)
"""

import csv
import os
import re
from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    FamilySymbol,
    Level,
    Transaction,
    UV,
    XYZ,
)

TITLE = "Place Spaces from CSV"
LOG = script.get_logger()


def parse_feet_inches(value_str):
    """
    Parse Revit feet-inches format like "76'-2 7/8"" to decimal feet.

    Args:
        value_str: String like "76'-2 7/8"" or "0'-0"" or "-16'-10 15/16""

    Returns:
        float: Value in decimal feet
    """
    if not value_str:
        return 0.0

    # Remove extra quotes and whitespace (strip all quotes, not just one)
    value_str = value_str.strip()
    while value_str.startswith('"') and value_str.endswith('"'):
        value_str = value_str[1:-1].strip()

    # Check for negative sign at the start
    sign = 1.0
    if value_str.startswith('-'):
        sign = -1.0
        value_str = value_str[1:].strip()

    # Pattern: feet'-inches fractional"
    # Examples: 76'-2 7/8", 0'-0", 25'-1 15/16"
    # Note: No negative signs in pattern since we handle that above
    pattern = r"(\d+)'-(\d+(?:\s+\d+/\d+)?)\""
    match = re.match(pattern, value_str)

    if not match:
        # Try just feet
        feet_pattern = r"(\d+)'"
        feet_match = re.match(feet_pattern, value_str)
        if feet_match:
            return sign * float(feet_match.group(1))
        LOG.warning("Could not parse coordinate: {}".format(value_str))
        return 0.0

    feet = float(match.group(1))
    inches_str = match.group(2)

    # Parse inches which might be like "2 7/8" or just "0"
    # Use abs() to handle any remaining negative signs in the inches portion
    inches = 0.0
    if ' ' in inches_str:
        # Has fractional part
        parts = inches_str.split()
        inches = abs(float(parts[0]))
        if len(parts) > 1:
            # Parse fraction
            frac_parts = parts[1].split('/')
            if len(frac_parts) == 2:
                inches += float(frac_parts[0]) / float(frac_parts[1])
    else:
        inches = abs(float(inches_str))

    # Convert to feet (inches / 12) and apply sign
    total_feet = sign * (feet + (inches / 12.0))
    return total_feet


def select_level(doc):
    """Prompt user to select a level."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    level_dict = {level.Name: level for level in levels}

    if not level_dict:
        forms.alert("No levels found in document.", title=TITLE, exitscript=True)

    level_names = sorted(level_dict.keys())
    selected_name = forms.SelectFromList.show(
        level_names,
        title="Select Level for Spaces",
        button_name="Select",
        multiselect=False
    )

    if not selected_name:
        return None

    return level_dict[selected_name]


def select_csv_file():
    """Prompt user to select a CSV file."""
    from System.Windows.Forms import OpenFileDialog, DialogResult

    dialog = OpenFileDialog()
    dialog.Title = "Select CSV File"
    dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    dialog.Multiselect = False

    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None


def read_csv_spaces(csv_path):
    """
    Read spaces from CSV file.

    Supports two formats:
    1. Old format: '#' column for number, '#(1)' and '#(2)' for name parts
    2. New format: 'Room Name' column like "Room Name 107A ROOM ELECTRICAL"

    Returns:
        List of dicts with keys: number, name, x, y
    """
    spaces = []

    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)

        for row_num, row in enumerate(reader, start=2):  # Start at 2 (header is row 1)
            try:
                space_number = None
                space_name = None

                # Check for 'Room Name' column first (new format)
                room_name = (row.get('Room Name') or row.get('Name') or '').strip()

                # ONLY process rows that start with "Room Name"
                if room_name and room_name.startswith("Room Name"):
                    # Parse "Room Name 107A ROOM ELECTRICAL" format
                    # Pattern: "Room Name <number> <name>"
                    remaining = room_name.replace("Room Name", "", 1).strip()
                    # Split into number (first word) and rest
                    remaining_parts = remaining.split(None, 1)
                    if remaining_parts:
                        space_number = remaining_parts[0]
                        space_name = remaining_parts[1] if len(remaining_parts) > 1 else ""
                    else:
                        # Just "Room Name" with no number - skip
                        continue

                # Fall back to old format with '#', '#(1)', '#(2)' columns
                if not space_number:
                    space_number = (row.get('#') or '').strip()

                    if space_number:
                        # Get name parts from '#(2)' and '#(1)' columns
                        name_part2 = (row.get('#(2)') or '').strip()
                        name_part1 = (row.get('#(1)') or '').strip()

                        # Concatenate: #(2) + " " + #(1)
                        if name_part2 and name_part1:
                            space_name = "{} {}".format(name_part2, name_part1)
                        elif name_part2:
                            space_name = name_part2
                        elif name_part1:
                            space_name = name_part1
                        else:
                            space_name = ""

                # Skip rows without space number
                if not space_number:
                    continue

                # Parse coordinates
                x_str = row.get('Position X', '')
                y_str = row.get('Position Y', '')

                x_feet = parse_feet_inches(x_str)
                y_feet = parse_feet_inches(y_str)

                spaces.append({
                    'number': space_number,
                    'name': space_name or "",
                    'x': x_feet,
                    'y': y_feet,
                    'row': row_num
                })

            except Exception as e:
                LOG.warning("Error parsing row {}: {}".format(row_num, e))
                continue

    return spaces


def get_space_tag_type(doc):
    """Get the first available space tag type, preferring PR_Space Tag_CED."""
    # Collect all Space Tag FamilySymbols
    collector = FilteredElementCollector(doc)\
        .OfClass(FamilySymbol)\
        .OfCategory(BuiltInCategory.OST_MEPSpaceTags)\
        .ToElements()

    # First try to find PR_Space Tag_CED : Name / Number
    preferred_tag = None
    fallback_tag = None

    for tag_symbol in collector:
        if not isinstance(tag_symbol, FamilySymbol):
            continue

        # Get family name
        family_name = tag_symbol.FamilyName if hasattr(tag_symbol, 'FamilyName') else ""

        # Get type name
        type_param = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = type_param.AsString() if type_param else ""

        LOG.debug("Found space tag: {} : {}".format(family_name, type_name))

        # Check for preferred tag
        if "PR_Space Tag_CED" in family_name and "Name / Number" in type_name:
            LOG.info("Using preferred space tag: {} : {}".format(family_name, type_name))
            return tag_symbol

        # Check for any PR_Space Tag_CED
        if "PR_Space Tag_CED" in family_name:
            preferred_tag = tag_symbol

        # Keep any tag as fallback
        if not fallback_tag:
            fallback_tag = tag_symbol

    # Return in order of preference
    result = preferred_tag or fallback_tag
    if result:
        LOG.info("Using space tag: {} : {}".format(
            result.FamilyName if hasattr(result, 'FamilyName') else "Unknown",
            result.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if result.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
        ))
    return result


def create_spaces(doc, level, spaces, active_view, add_tags=True):
    """
    Create spaces in Revit.

    Args:
        doc: Revit document
        level: Level to place spaces on
        spaces: List of space dicts
        active_view: Active view for tagging
        add_tags: Whether to add tags

    Returns:
        dict: Statistics about placement
    """
    placed = 0
    failed = 0
    tagged = 0
    errors = []

    # Get space tag type if tagging is enabled
    tag_type = None
    if add_tags:
        tag_type = get_space_tag_type(doc)
        if not tag_type:
            LOG.warning("No space tag type found - tags will not be created")

    with Transaction(doc, "Place Spaces from CSV") as t:
        t.Start()

        try:
            for space_data in spaces:
                try:
                    # Create unbounded space at X,Y coordinate
                    uv = UV(space_data['x'], space_data['y'])
                    space = doc.Create.NewSpace(level, uv)

                    # Set space number
                    if space_data['number']:
                        number_param = space.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                        if number_param and not number_param.IsReadOnly:
                            number_param.Set(space_data['number'])

                    # Set space name
                    if space_data['name']:
                        name_param = space.get_Parameter(BuiltInParameter.ROOM_NAME)
                        if name_param and not name_param.IsReadOnly:
                            name_param.Set(space_data['name'])

                    placed += 1
                    LOG.debug("Created space: {} - {} at ({}, {})".format(
                        space_data['number'],
                        space_data['name'],
                        space_data['x'],
                        space_data['y']
                    ))

                    # Tag the space if enabled and in a valid view
                    if add_tags and tag_type and active_view:
                        try:
                            # Get space location point
                            location = space.Location
                            if location:
                                point = location.Point
                                # Create tag at space location
                                tag = doc.Create.NewSpaceTag(space, uv, active_view)
                                if tag:
                                    tagged += 1
                                    LOG.debug("Tagged space: {}".format(space_data['number']))
                        except Exception as tag_error:
                            LOG.warning("Failed to tag space {}: {}".format(
                                space_data['number'], tag_error
                            ))

                except Exception as e:
                    failed += 1
                    error_msg = "Row {}: {} - {}".format(
                        space_data['row'],
                        space_data['number'],
                        str(e)
                    )
                    errors.append(error_msg)
                    LOG.error("Failed to create space: {}".format(error_msg))

            t.Commit()

        except Exception as e:
            t.RollBack()
            LOG.error("Transaction failed: {}".format(e))
            raise

    return {
        'placed': placed,
        'failed': failed,
        'tagged': tagged,
        'errors': errors
    }


def main():
    doc = revit.doc

    if not doc:
        forms.alert("No active document.", title=TITLE, exitscript=True)

    # Step 1: Select CSV file
    csv_path = select_csv_file()
    if not csv_path:
        forms.alert("No CSV file selected.", title=TITLE, exitscript=True)

    if not os.path.exists(csv_path):
        forms.alert("CSV file not found:\n{}".format(csv_path), title=TITLE, exitscript=True)

    # Step 2: Read CSV
    LOG.info("Reading CSV: {}".format(csv_path))
    try:
        spaces = read_csv_spaces(csv_path)
    except Exception as e:
        forms.alert("Error reading CSV:\n{}".format(e), title=TITLE, exitscript=True)

    if not spaces:
        forms.alert("No spaces found in CSV.\n\nMake sure the CSV has:\n- '#' column with space numbers\n- 'Position X' and 'Position Y' columns\n- Optional '#(1)' and '#(2)' for space names", title=TITLE, exitscript=True)

    LOG.info("Found {} spaces in CSV".format(len(spaces)))

    # Step 3: Select level
    level = select_level(doc)
    if not level:
        forms.alert("No level selected.", title=TITLE, exitscript=True)

    # Step 4: Confirm
    message = "Ready to create {} spaces on level '{}'.\n\nContinue?".format(
        len(spaces),
        level.Name
    )
    if not forms.alert(message, title=TITLE, yes=True, no=True):
        script.exit()

    # Step 5: Create spaces
    LOG.info("Creating spaces...")
    active_view = doc.ActiveView
    try:
        results = create_spaces(doc, level, spaces, active_view, add_tags=True)
    except Exception as e:
        forms.alert("Error creating spaces:\n{}".format(e), title=TITLE, exitscript=True)

    # Step 6: Report results
    summary = []
    summary.append("Spaces Created: {}".format(results['placed']))
    summary.append("Spaces Tagged: {}".format(results['tagged']))

    if results['failed']:
        summary.append("Failed: {}".format(results['failed']))
        summary.append("")
        summary.append("Errors:")
        for error in results['errors'][:10]:  # Show first 10 errors
            summary.append("  - {}".format(error))
        if len(results['errors']) > 10:
            summary.append("  ... and {} more".format(len(results['errors']) - 10))

    forms.alert("\n".join(summary), title=TITLE)
    LOG.info("Complete: {} placed, {} tagged, {} failed".format(
        results['placed'], results['tagged'], results['failed']
    ))


if __name__ == '__main__':
    main()
