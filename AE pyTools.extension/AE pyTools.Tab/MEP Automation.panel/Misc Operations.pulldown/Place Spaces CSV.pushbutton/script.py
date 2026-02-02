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

Tags are automatically created in all floor plan and ceiling plan views
with "Construction / Permit" classification.
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
    ViewPlan,
    ViewType,
    XYZ,
)

TITLE = "Place Spaces from CSV"
LOG = script.get_logger()


def parse_feet_inches(value_str):
    """
    Parse Revit feet-inches format like "76'-2 7/8"" to decimal feet.

    Args:
        value_str: String like "76'-2 7/8"" or "0'-0""

    Returns:
        float: Value in decimal feet
    """
    if not value_str:
        return 0.0

    # Remove extra quotes and whitespace
    value_str = value_str.strip().strip('"')

    # Pattern: feet'-inches fractional"
    # Examples: 76'-2 7/8", 0'-0", 25'-1 15/16"
    pattern = r"(-?\d+)'-(-?\d+(?:\s+\d+/\d+)?)\""
    match = re.match(pattern, value_str)

    if not match:
        # Try just feet
        feet_pattern = r"(-?\d+)'"
        feet_match = re.match(feet_pattern, value_str)
        if feet_match:
            return float(feet_match.group(1))
        LOG.warning("Could not parse coordinate: {}".format(value_str))
        return 0.0

    feet = float(match.group(1))
    inches_str = match.group(2)

    # Parse inches which might be like "2 7/8" or just "0"
    inches = 0.0
    if ' ' in inches_str:
        # Has fractional part
        parts = inches_str.split()
        inches = float(parts[0])
        if len(parts) > 1:
            # Parse fraction
            frac_parts = parts[1].split('/')
            if len(frac_parts) == 2:
                inches += float(frac_parts[0]) / float(frac_parts[1])
    else:
        inches = float(inches_str)

    # Convert to feet (inches / 12)
    total_feet = feet + (inches / 12.0)
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

    Returns:
        List of dicts with keys: number, name, x, y
    """
    spaces = []

    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)

        for row_num, row in enumerate(reader, start=2):  # Start at 2 (header is row 1)
            try:
                # Get space number from '#' column
                space_number = (row.get('#') or '').strip()

                # Skip rows without space number
                if not space_number:
                    continue

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

                # Parse coordinates
                x_str = row.get('Position X', '')
                y_str = row.get('Position Y', '')

                x_feet = parse_feet_inches(x_str)
                y_feet = parse_feet_inches(y_str)

                spaces.append({
                    'number': space_number,
                    'name': space_name,
                    'x': x_feet,
                    'y': y_feet,
                    'row': row_num
                })

            except Exception as e:
                LOG.warning("Error parsing row {}: {}".format(row_num, e))
                continue

    return spaces


def get_construction_permit_views(doc):
    """Get all floor plan and ceiling plan views in 'Construction / Permit' classification."""
    # Collect all floor plans and ceiling plans
    floor_plans = FilteredElementCollector(doc)\
        .OfClass(ViewPlan)\
        .ToElements()

    construction_permit_views = []

    for view in floor_plans:
        view_type = view.ViewType
        if view_type == ViewType.FloorPlan or view_type == ViewType.CeilingPlan:
            param = view.LookupParameter("View Classification")
            if param and param.HasValue:
                view_classification = param.AsString()
                if view_classification == "Construction / Permit":
                    construction_permit_views.append(view)
                    LOG.debug("Found Construction / Permit view: {} ({})".format(
                        view.Name,
                        "Floor Plan" if view_type == ViewType.FloorPlan else "Ceiling Plan"
                    ))

    LOG.info("Found {} Construction / Permit views".format(len(construction_permit_views)))
    return construction_permit_views


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


def create_spaces(doc, level, spaces, construction_permit_views, add_tags=True):
    """
    Create spaces in Revit.

    Args:
        doc: Revit document
        level: Level to place spaces on
        spaces: List of space dicts
        construction_permit_views: List of Construction/Permit views to tag in
        add_tags: Whether to add tags

    Returns:
        dict: Statistics about placement
    """
    placed = 0
    failed = 0
    tagged = 0
    tagged_views = {}
    errors = []

    # Get space tag type if tagging is enabled
    tag_type = None
    if add_tags:
        tag_type = get_space_tag_type(doc)
        if not tag_type:
            LOG.warning("No space tag type found - tags will not be created")

    # Store created spaces with their UV coordinates for tagging
    created_spaces = []

    with Transaction(doc, "Place Spaces from CSV") as t:
        t.Start()

        try:
            # Step 1: Create all spaces
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
                    created_spaces.append((space, uv, space_data['number']))
                    LOG.debug("Created space: {} - {} at ({}, {})".format(
                        space_data['number'],
                        space_data['name'],
                        space_data['x'],
                        space_data['y']
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

            # Step 2: Tag spaces in all Construction/Permit views
            if add_tags and tag_type and construction_permit_views and created_spaces:
                for view in construction_permit_views:
                    view_name = view.Name
                    view_tag_count = 0

                    for space, uv, space_number in created_spaces:
                        try:
                            tag = doc.Create.NewSpaceTag(space, uv, view)
                            if tag:
                                view_tag_count += 1
                                tagged += 1
                        except Exception as tag_error:
                            LOG.warning("Failed to tag space {} in view '{}': {}".format(
                                space_number, view_name, tag_error
                            ))

                    if view_tag_count > 0:
                        tagged_views[view_name] = view_tag_count
                        LOG.debug("Tagged {} spaces in view '{}'".format(view_tag_count, view_name))

            t.Commit()

        except Exception as e:
            t.RollBack()
            LOG.error("Transaction failed: {}".format(e))
            raise

    return {
        'placed': placed,
        'failed': failed,
        'tagged': tagged,
        'tagged_views': tagged_views,
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

    # Step 4: Get Construction/Permit views for tagging
    construction_permit_views = get_construction_permit_views(doc)
    if not construction_permit_views:
        LOG.warning("No Construction / Permit views found - tags will not be created")
        if not forms.alert(
            "No 'Construction / Permit' floor or ceiling plans found.\n\n"
            "Spaces will be created but NOT tagged.\n\n"
            "Continue?",
            title=TITLE,
            yes=True,
            no=True
        ):
            script.exit()
    else:
        # Show confirmation with view count
        message = "Ready to create {} spaces on level '{}'.\n\n" \
                  "Tags will be created in {} Construction / Permit views.\n\n" \
                  "Continue?".format(
            len(spaces),
            level.Name,
            len(construction_permit_views)
        )
        if not forms.alert(message, title=TITLE, yes=True, no=True):
            script.exit()

    # Step 5: Create spaces and tags
    LOG.info("Creating spaces...")
    try:
        results = create_spaces(doc, level, spaces, construction_permit_views, add_tags=True)
    except Exception as e:
        forms.alert("Error creating spaces:\n{}".format(e), title=TITLE, exitscript=True)

    # Step 6: Report results
    summary = []
    summary.append("Spaces Created: {}".format(results['placed']))
    summary.append("Total Tags Created: {}".format(results['tagged']))

    # Show per-view tag counts
    tagged_views = results.get('tagged_views', {})
    if tagged_views:
        summary.append("")
        summary.append("Tags per view:")
        for view_name, count in sorted(tagged_views.items()):
            summary.append("  - {}: {}".format(view_name, count))

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
    LOG.info("Complete: {} placed, {} tagged across {} views, {} failed".format(
        results['placed'], results['tagged'], len(tagged_views), results['failed']
    ))


if __name__ == '__main__':
    main()
