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

# Marker column added to the CSV after cleanup so we can detect future runs
# without re-scanning row contents. The presence of this column = already cleaned.
CLEANED_FLAG_COLUMN = 'CSV_CLEANED'
CLEANED_FLAG_VALUE = 'true'


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


def is_csv_cleaned(csv_path):
    """
    Determine if the CSV has already been cleaned up.

    Detection is based on the presence of the CLEANED_FLAG_COLUMN marker column
    that clean_csv() adds after it runs. If that column exists in the header,
    the CSV has already been processed and should NOT be cleaned again
    (running the cleanup twice would duplicate '#' values into Name).
    """
    try:
        with open(csv_path, 'r') as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or []

            if CLEANED_FLAG_COLUMN in fieldnames:
                return True

            # No marker present. If the '#' columns also aren't there, there's
            # nothing the cleanup would do anyway, so treat as cleaned.
            has_hash_cols = any(c in fieldnames for c in ('#', '#(1)', '#(2)'))
            if not has_hash_cols:
                return True

            return False
    except Exception as e:
        LOG.warning("Could not determine CSV cleanup state: {}".format(e))
        # Be conservative: if we can't tell, skip cleanup so we don't corrupt data.
        return True


def clean_csv(csv_path):
    """
    Clean the CSV in place by concatenating the Name column with the '#',
    '#(1)', '#(2)' columns for every row where any of those has a value,
    then append a CLEANED_FLAG_COLUMN marker column so future runs can detect
    that cleanup has already happened.

    This mirrors the exact logic from the standalone 'csv cleanup.py' script,
    but uses the stdlib csv module so it works under pyRevit (IronPython 2.7).

    Original pandas logic:
        mask = df[['#', '#(1)', '#(2)']].notna().any(axis=1)
        df.loc[mask, 'Name'] = (
            df.loc[mask, 'Name'].fillna('').astype(str) + ' ' +
            df.loc[mask, '#'].fillna('').astype(str) + ' ' +
            df.loc[mask, '#(1)'].fillna('').astype(str) + ' ' +
            df.loc[mask, '#(2)'].fillna('').astype(str)
        ).str.strip()
    """
    rows = []
    fieldnames = None

    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames) if reader.fieldnames else []
        for row in reader:
            hash_val = row.get('#') or ''
            h1_val = row.get('#(1)') or ''
            h2_val = row.get('#(2)') or ''

            # Mask equivalent: any of '#', '#(1)', '#(2)' has a value.
            if hash_val.strip() or h1_val.strip() or h2_val.strip():
                name_val = row.get('Name') or ''
                # Preserve the exact order and spacing from the pandas version:
                #   Name + ' ' + # + ' ' + #(1) + ' ' + #(2), then .strip()
                combined = '{} {} {} {}'.format(
                    name_val, hash_val, h1_val, h2_val
                ).strip()
                row['Name'] = combined

            rows.append(row)

    if not fieldnames:
        return

    # Append the marker column so subsequent runs can see this file was cleaned.
    if CLEANED_FLAG_COLUMN not in fieldnames:
        fieldnames.append(CLEANED_FLAG_COLUMN)
    for row in rows:
        row[CLEANED_FLAG_COLUMN] = CLEANED_FLAG_VALUE

    # Overwrite the same file. Binary mode keeps csv.DictWriter happy on IronPython.
    with open(csv_path, 'wb') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    LOG.info("CSV cleaned and flagged with '{}' column: {}".format(
        CLEANED_FLAG_COLUMN, csv_path
    ))


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


def get_construction_permit_views(doc, level):
    """
    Get all views with 'Construction' or 'Permit' in their view class parameter.

    Args:
        doc: Revit document
        level: Level to filter views by

    Returns:
        List of views
    """
    from Autodesk.Revit.DB import View, ViewPlan

    views = []
    all_views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType().ToElements()

    plan_views_on_level = 0
    views_with_param = 0

    for view in all_views:
        try:
            # Skip templates
            if view.IsTemplate:
                continue

            # Only check plan views that match the level
            if isinstance(view, ViewPlan):
                if hasattr(view, 'GenLevel') and view.GenLevel:
                    if view.GenLevel.Id == level.Id:
                        plan_views_on_level += 1
                        LOG.info("Checking plan view on level: '{}'".format(view.Name))
                    else:
                        continue
                else:
                    continue
            else:
                # Skip non-plan views
                continue

            # Look for View Classification parameter
            view_class_param = view.LookupParameter("View Classification")

            if view_class_param:
                views_with_param += 1
                view_class = view_class_param.AsString()
                LOG.debug("View '{}' has View Classification = '{}'".format(view.Name, view_class))
                # Check for "Construction / Permit" (exact match)
                if view_class and view_class.strip() == "Construction / Permit":
                    views.append(view)
                    LOG.info(">>> MATCHED: Found Construction/Permit view: {}".format(view.Name))
            else:
                LOG.debug("View '{}' has NO View Classification parameter".format(view.Name))
        except Exception as e:
            LOG.error("Error checking view {}: {}".format(view.Name if hasattr(view, 'Name') else "Unknown", e))
            continue

    LOG.info("Summary: {} plan views on level, {} with PR_View_Class param, {} matched".format(
        plan_views_on_level, views_with_param, len(views)))
    return views


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


def create_spaces(doc, level, spaces, views_to_tag, add_tags=True):
    """
    Create spaces in Revit.

    Args:
        doc: Revit document
        level: Level to place spaces on
        spaces: List of space dicts
        views_to_tag: List of views to create tags in
        add_tags: Whether to add tags

    Returns:
        dict: Statistics about placement
    """
    placed = 0
    failed = 0
    tagged = 0
    tagged_views = set()
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

                    # Tag the space in all Construction/Permit views
                    if add_tags and tag_type and views_to_tag:
                        for view in views_to_tag:
                            try:
                                # Create tag at space location
                                tag = doc.Create.NewSpaceTag(space, uv, view)
                                if tag:
                                    tagged += 1
                                    tagged_views.add(view.Id.IntegerValue)
                                    LOG.debug("Tagged space {} in view {}".format(
                                        space_data['number'], view.Name
                                    ))
                            except Exception as tag_error:
                                LOG.warning("Failed to tag space {} in view {}: {}".format(
                                    space_data['number'], view.Name, tag_error
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
        'tagged_views': len(tagged_views),
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

    # Step 1b: Clean the CSV if it hasn't been cleaned yet.
    if is_csv_cleaned(csv_path):
        LOG.info("CSV already cleaned - skipping cleanup step.")
    else:
        LOG.info("CSV not yet cleaned - running cleanup...")
        try:
            clean_csv(csv_path)
        except Exception as e:
            forms.alert("Error cleaning CSV:\n{}".format(e), title=TITLE, exitscript=True)

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

    # Step 4: Get Construction/Permit views
    views_to_tag = get_construction_permit_views(doc, level)
    if not views_to_tag:
        LOG.warning("No Construction/Permit views found on level {}".format(level.Name))
        if not forms.alert(
            "No Construction or Permit views found on level '{}'.\n\nSpaces will be created but not tagged.\n\nContinue?".format(level.Name),
            title=TITLE, yes=True, no=True
        ):
            script.exit()

    # Step 5: Confirm
    message = "Ready to create {} spaces on level '{}'".format(len(spaces), level.Name)
    if views_to_tag:
        message += "\n\nTags will be placed in {} Construction/Permit views.".format(len(views_to_tag))
    message += "\n\nContinue?"

    if not forms.alert(message, title=TITLE, yes=True, no=True):
        script.exit()

    # Step 6: Create spaces
    LOG.info("Creating spaces...")
    try:
        results = create_spaces(doc, level, spaces, views_to_tag, add_tags=True)
    except Exception as e:
        forms.alert("Error creating spaces:\n{}".format(e), title=TITLE, exitscript=True)

    # Step 7: Report results
    summary = []
    summary.append("Spaces Created: {}".format(results['placed']))
    summary.append("Space Tags Created: {}".format(results['tagged']))
    summary.append("Views Tagged: {}".format(results['tagged_views']))

    if results['failed']:
        summary.append("")
        summary.append("Failed: {}".format(results['failed']))
        summary.append("")
        summary.append("Errors:")
        for error in results['errors'][:10]:  # Show first 10 errors
            summary.append("  - {}".format(error))
        if len(results['errors']) > 10:
            summary.append("  ... and {} more".format(len(results['errors']) - 10))

    forms.alert("\n".join(summary), title=TITLE)
    LOG.info("Complete: {} placed, {} tags in {} views, {} failed".format(
        results['placed'], results['tagged'], results['tagged_views'], results['failed']
    ))


if __name__ == '__main__':
    main()
