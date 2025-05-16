# -*- coding: utf-8 -*-

import clr
import math
from pyrevit import script, forms
from pyrevit import revit, DB
from System.Collections.Generic import List
# Access the current Revit document
doc = revit.doc
output = script.get_output()
output.close_others()
output.set_width(800)

# Define the categories using BuiltInCategory enum with readable names
category_enum = {
    "Multi-Category Tags": DB.BuiltInCategory.OST_MultiCategoryTags,
    "Keynote Tags": DB.BuiltInCategory.OST_KeynoteTags,
    "Electrical Fixture Tags": DB.BuiltInCategory.OST_ElectricalFixtureTags,
    "Electrical Equipment Tags": DB.BuiltInCategory.OST_ElectricalEquipmentTags,
    "Lighting Fixture Tags": DB.BuiltInCategory.OST_LightingFixtureTags,
    "Lighting Device Tags": DB.BuiltInCategory.OST_LightingDeviceTags,
    "Data Device Tags": DB.BuiltInCategory.OST_DataDeviceTags,
    "Fire Alarm Device Tags": DB.BuiltInCategory.OST_FireAlarmDeviceTags,
    "Security Device Tags": DB.BuiltInCategory.OST_SecurityDeviceTags,
    "Wire Tags": DB.BuiltInCategory.OST_WireTags
}

# Initialize dictionaries to store the count of elements per category and model orientation
category_counts = {}
model_orientation_counts = {}

selection = revit.get_selection()
if selection:
    selected_ids = List[DB.ElementId](selection.element_ids)

elements = []

for name, cat in category_enum.items():
    if selection:
        collector = DB.FilteredElementCollector(doc, selected_ids) \
            .OfCategory(cat) \
            .WhereElementIsNotElementType() \
            .ToElements()
    else:
        collector = DB.FilteredElementCollector(doc) \
            .OfCategory(cat) \
            .WhereElementIsNotElementType() \
            .ToElements()

    count = len(collector)
    category_counts[name] = count
    elements.extend(collector)

# Pre-filter elements based on their orientation value before committing a transaction
filtered_elements = []
for element in elements:
    orientation_param = element.LookupParameter("Orientation")
    if orientation_param:
        original_orientation = orientation_param.AsInteger()  # Get the current orientation value
        if original_orientation == 2:  # Assuming "2" is "Model"
            category_name = element.Category.Name
            if category_name not in model_orientation_counts:
                model_orientation_counts[category_name] = 0
            model_orientation_counts[category_name] += 1
            filtered_elements.append(element)

# Further process the filtered elements
affected_elements = []

if filtered_elements:
    transaction = DB.Transaction(doc, "Update Tag Orientation")

    try:
        transaction.Start()

        for element in filtered_elements:
            angle_param = element.LookupParameter("Angle")
            orientation_param = element.LookupParameter("Orientation")

            if angle_param and orientation_param:
                angle_radians = angle_param.AsDouble()
                angle_degrees = math.degrees(angle_radians)  # Convert radians to degrees

                # Determine the new orientation
                if round(angle_degrees) in [0, 180, 360]:
                    new_orientation = "Horizontal"
                    orientation_param.Set(0)  # Set to horizontal (e.g., 0)
                elif round(angle_degrees) in [90, 270]:
                    new_orientation = "Vertical"
                    orientation_param.Set(1)  # Set to vertical (e.g., 1)
                else:
                    continue

                # Get the element type name and category name
                element_type = "Unknown Type"
                try:
                    element_type_element = doc.GetElement(element.GetTypeId())
                    if element_type_element:
                        element_type = element_type_element.get_Parameter(
                            DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                except Exception as e:
                    output.print_md("Error retrieving type for Element ID {}: {}".format(element.Id, str(e)))

                element_category = element.Category.Name

                # Add the element to the list of affected elements
                affected_elements.append((element.Id, new_orientation, element_type, element_category))

        transaction.Commit()

    except Exception as e:
        transaction.RollBack()
        raise e  # Re-raise the exception to know what went wrong

# Shift-click to print the detailed table, regular click to show an alert
if __shiftclick__:
    # Print counts of elements with "Model" orientation, organized by category
    output.print_md("\n\n## Elements with 'Model' orientation before adjustment:\n")
    for category, count in model_orientation_counts.items():
        output.print_md("**{}**: {}".format(category, count))

    output.print_md("\n\n---\n\n")

    # Print details of affected elements (only those that were originally "Model")
    if affected_elements:
        output.print_md("## Affected Elements (originally 'Model' orientation):")
        table_data = []
        for elem_id, new_orientation, element_type, element_category in affected_elements:
            clickable_id = output.linkify(elem_id)
            table_data.append([clickable_id, element_type, element_category, new_orientation])
        output.print_table(table_data, columns=["Element ID", "Type Name", "Category", "New Orientation"])
    else:
        output.print_md("No elements were affected.")
else:
    # Regular click: Show alert with summary
    if affected_elements:
        forms.alert('{} tag(s) were updated.'.format(len(affected_elements)))
    else:
        forms.alert('No tags needed updating.')
