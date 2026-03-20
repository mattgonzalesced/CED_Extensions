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
<<<<<<< HEAD

Tags are automatically created in all floor plan and ceiling plan views
with "Construction / Permit" classification.
=======
>>>>>>> origin/develop
"""

import csv
import os
import re
from pyrevit import forms, revit, script
<<<<<<< HEAD
output = script.get_output()
output.close_others()
=======
>>>>>>> origin/develop
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    FamilySymbol,
    Level,
    Transaction,
    UV,
<<<<<<< HEAD
    ViewPlan,
    ViewType,
=======
>>>>>>> origin/develop
    XYZ,
)

TITLE = "Place Spaces from CSV"
LOG = script.get_logger()


def parse_feet_inches(value_str):
    """
    Parse Revit feet-inches format like "76'-2 7/8"" to decimal feet.

    Args:
<<<<<<< HEAD
        value_str: String like "76'-2 7/8"" or "0'-0""
=======
        value_str: String like "76'-2 7/8"" or "0'-0"" or "-16'-10 15/16""
>>>>>>> origin/develop

    Returns:
        float: Value in decimal feet
    """
    if not value_str:
        return 0.0

<<<<<<< HEAD
    # Remove whitespace and all quote characters for easier parsing
    value_str = value_str.strip().replace('"', '')

    # Pattern: feet'-inches fractional (quotes already removed)
    # Examples: 76'-2 7/8, 0'-0, 25'-1 15/16, -167'-5 1/16
    pattern = r"(-?\d+)'-(\d+(?:\s+\d+/\d+)?)"
=======
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
>>>>>>> origin/develop
    match = re.match(pattern, value_str)

    if not match:
        # Try just feet
<<<<<<< HEAD
        feet_pattern = r"(-?\d+)'"
        feet_match = re.match(feet_pattern, value_str)
        if feet_match:
            return float(feet_match.group(1))
=======
        feet_pattern = r"(\d+)'"
        feet_match = re.match(feet_pattern, value_str)
        if feet_match:
            return sign * float(feet_match.group(1))
>>>>>>> origin/develop
        LOG.warning("Could not parse coordinate: {}".format(value_str))
        return 0.0

    feet = float(match.group(1))
    inches_str = match.group(2)

    # Parse inches which might be like "2 7/8" or just "0"
<<<<<<< HEAD
=======
    # Use abs() to handle any remaining negative signs in the inches portion
>>>>>>> origin/develop
    inches = 0.0
    if ' ' in inches_str:
        # Has fractional part
        parts = inches_str.split()
<<<<<<< HEAD
        inches = float(parts[0])
=======
        inches = abs(float(parts[0]))
>>>>>>> origin/develop
        if len(parts) > 1:
            # Parse fraction
            frac_parts = parts[1].split('/')
            if len(frac_parts) == 2:
                inches += float(frac_parts[0]) / float(frac_parts[1])
    else:
<<<<<<< HEAD
        inches = float(inches_str)

    # Convert to feet (inches / 12)
    # If feet is negative, inches should also be negative
    inches_in_feet = inches / 12.0
    if feet < 0:
        total_feet = feet - inches_in_feet
    else:
        total_feet = feet + inches_in_feet
=======
        inches = abs(float(inches_str))

    # Convert to feet (inches / 12) and apply sign
    total_feet = sign * (feet + (inches / 12.0))
>>>>>>> origin/develop
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

<<<<<<< HEAD
=======
    Supports two formats:
    1. Old format: '#' column for number, '#(1)' and '#(2)' for name parts
    2. New format: 'Room Name' column like "Room Name 107A ROOM ELECTRICAL"

>>>>>>> origin/develop
    Returns:
        List of dicts with keys: number, name, x, y
    """
    spaces = []

    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)

        for row_num, row in enumerate(reader, start=2):  # Start at 2 (header is row 1)
            try:
<<<<<<< HEAD
                # Get space number from '#' column
                space_number = (row.get('#') or '').strip()
=======
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
>>>>>>> origin/develop

                # Skip rows without space number
                if not space_number:
                    continue

<<<<<<< HEAD
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

=======
>>>>>>> origin/develop
                # Parse coordinates
                x_str = row.get('Position X', '')
                y_str = row.get('Position Y', '')

                x_feet = parse_feet_inches(x_str)
                y_feet = parse_feet_inches(y_str)

                spaces.append({
                    'number': space_number,
<<<<<<< HEAD
                    'name': space_name,
=======
                    'name': space_name or "",
>>>>>>> origin/develop
                    'x': x_feet,
                    'y': y_feet,
                    'row': row_num
                })

            except Exception as e:
                LOG.warning("Error parsing row {}: {}".format(row_num, e))
                continue

    return spaces


<<<<<<< HEAD
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
=======
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
>>>>>>> origin/develop


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


<<<<<<< HEAD
def create_spaces(doc, level, spaces, construction_permit_views, add_tags=True):
=======
def create_spaces(doc, level, spaces, views_to_tag, add_tags=True):
>>>>>>> origin/develop
    """
    Create spaces in Revit.

    Args:
        doc: Revit document
        level: Level to place spaces on
        spaces: List of space dicts
<<<<<<< HEAD
        construction_permit_views: List of Construction/Permit views to tag in
=======
        views_to_tag: List of views to create tags in
>>>>>>> origin/develop
        add_tags: Whether to add tags

    Returns:
        dict: Statistics about placement
    """
    placed = 0
    failed = 0
    tagged = 0
<<<<<<< HEAD
    tagged_views = {}
=======
    tagged_views = set()
>>>>>>> origin/develop
    errors = []

    # Get space tag type if tagging is enabled
    tag_type = None
    if add_tags:
        tag_type = get_space_tag_type(doc)
        if not tag_type:
            LOG.warning("No space tag type found - tags will not be created")

<<<<<<< HEAD
    # Store created spaces with their UV coordinates for tagging
    created_spaces = []

=======
>>>>>>> origin/develop
    with Transaction(doc, "Place Spaces from CSV") as t:
        t.Start()

        try:
<<<<<<< HEAD
            # Step 1: Create all spaces
=======
>>>>>>> origin/develop
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
<<<<<<< HEAD
                    created_spaces.append((space, uv, space_data['number']))
=======
>>>>>>> origin/develop
                    LOG.debug("Created space: {} - {} at ({}, {})".format(
                        space_data['number'],
                        space_data['name'],
                        space_data['x'],
                        space_data['y']
                    ))

<<<<<<< HEAD
=======
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

>>>>>>> origin/develop
                except Exception as e:
                    failed += 1
                    error_msg = "Row {}: {} - {}".format(
                        space_data['row'],
                        space_data['number'],
                        str(e)
                    )
                    errors.append(error_msg)
                    LOG.error("Failed to create space: {}".format(error_msg))

<<<<<<< HEAD
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

=======
>>>>>>> origin/develop
            t.Commit()

        except Exception as e:
            t.RollBack()
            LOG.error("Transaction failed: {}".format(e))
            raise

    return {
        'placed': placed,
        'failed': failed,
        'tagged': tagged,
<<<<<<< HEAD
        'tagged_views': tagged_views,
=======
        'tagged_views': len(tagged_views),
>>>>>>> origin/develop
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

<<<<<<< HEAD
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
=======
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
>>>>>>> origin/develop

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
<<<<<<< HEAD
    LOG.info("Complete: {} placed, {} tagged across {} views, {} failed".format(
        results['placed'], results['tagged'], len(tagged_views), results['failed']
=======
    LOG.info("Complete: {} placed, {} tags in {} views, {} failed".format(
        results['placed'], results['tagged'], results['tagged_views'], results['failed']
>>>>>>> origin/develop
    ))


if __name__ == '__main__':
    main()
