# -*- coding: utf-8 -*-
import math
import os

# pyRevit imports
from pyrevit import revit, forms, script

# Local imports
from utils import (feet_inch_to_inches, create_safe_control_name,
                   read_xyz_csv, read_matchings_json, organize_symbols_by_category,
                   determine_family_category, organize_model_groups)

# Revit API imports
from Autodesk.Revit.DB import (FilteredElementCollector, FamilySymbol, Structure, XYZ, Transaction, ElementTransformUtils, Line, ForgeTypeId, UnitUtils, GlobalParametersManager, ViewSchedule, BuiltInParameter)
from pyrevit import DB

# Try to import annotation-specific types (may not exist in older Revit versions)
try:
    from Autodesk.Revit.DB import FamilyPlacementType, ViewType
    SUPPORTS_ANNOTATION_PLACEMENT = True
except ImportError:
    SUPPORTS_ANNOTATION_PLACEMENT = False

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

# Import required WPF types.
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType, TextWrapping, FontWeights, Visibility
from System.Windows.Controls import StackPanel, ComboBox, Button, Border, Grid, ColumnDefinition, TextBlock
import System.Windows.Controls as swc  # For Orientation
from System.Windows.Media import Brushes
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

# Get active Revit document.
doc = revit.doc

#------------------------------------------------------------------------------
# 1. Prompt the user to select XYZ Locations CSV file.
xyz_csv_path = forms.pick_file(file_ext='csv', title='Select the XYZ Locations CSV')
if not xyz_csv_path:
    script.exit()  # Exit if no file is selected

# 2. Prompt the user to select Structured Matchings JSON file.
matchings_json_path = forms.pick_file(file_ext='json', title='Select the Structured Matchings JSON')
if not matchings_json_path:
    script.exit()  # Exit if no file is selected

#------------------------------------------------------------------------------
# 3. Read XYZ Locations CSV data.
xyz_rows, unique_names = read_xyz_csv(xyz_csv_path)

if not unique_names:
    script.exit("No valid 'Name' column found in XYZ CSV or CSV is empty.")

#------------------------------------------------------------------------------
# 4. Read Structured Matchings JSON to create automatic mappings.
matchings_dict, groups_dict, parameters_dict, offsets_dict = read_matchings_json(matchings_json_path)

#------------------------------------------------------------------------------
# 5. Collect all FamilySymbols from the project and organize by category.
fixture_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()


# print("FIXTURE_SYMBOLS {}".format(fixture_symbols))


# Organize symbols by category using utility function
(symbol_label_map, symbols_by_category, families_by_category,
 types_by_family, family_symbols) = organize_symbols_by_category(fixture_symbols)

symbol_labels_sorted = sorted(symbol_label_map.keys())

#------------------------------------------------------------------------------
# 5b. Collect and organize Model Groups
group_label_map, details_by_model_group, model_group_names = organize_model_groups(doc)

# Add "Model Groups" as a pseudo-category to integrate with existing UI
if model_group_names:
    symbols_by_category["Model Groups"] = sorted(group_label_map.keys())
    families_by_category["Model Groups"] = model_group_names
    # For model groups, "types" are the detail groups
    for model_group_name in model_group_names:
        types_by_family[model_group_name] = details_by_model_group[model_group_name]

#------------------------------------------------------------------------------
# 6. Load the XAML UI from a separate file.
xaml_path = os.path.join(os.path.dirname(__file__), "MappingWindow.XAML")
fs = FileStream(xaml_path, FileMode.Open)
window = XamlReader.Load(fs)
fs.Close()

# Get references to the named controls.
mapping_stack_panel = window.FindName("MappingStackPanel")
divisor_textbox = window.FindName("DivisorTextBox")
ok_button = window.FindName("OkButton")
cancel_button = window.FindName("CancelButton")

# Learning stuff
import webbrowser
def on_secret_click(sender, args):
    webbrowser.open("https://www.youtube.com/watch?v=dQw4w9WgXcQ")

ruler_icon = window.FindName("Rulericon")
Lightbulb_icon = window.FindName("Lightbulbicon")

ruler_icon.MouseLeftButtonDown += on_secret_click
Lightbulb_icon.MouseLeftButtonDown += on_secret_click
#------------------------------------------------------------------------------
# Add a header row for the columns.
headerGrid = Grid()
headerGrid.Margin = Thickness(5)

headerCol1 = ColumnDefinition()
headerCol1.Width = GridLength(1, GridUnitType.Star)
headerCol2 = ColumnDefinition()
headerCol2.Width = GridLength(2, GridUnitType.Star)  # Give more space to dropdown area
headerGrid.ColumnDefinitions.Add(headerCol1)
headerGrid.ColumnDefinitions.Add(headerCol2)

headerText1 = TextBlock()
headerText1.Text = "CAD Block"
headerText1.FontSize = 16
headerText1.FontWeight = FontWeights.Bold
headerText1.HorizontalAlignment = HorizontalAlignment.Stretch


from System.Windows.Controls import Grid as WpfGrid
WpfGrid.SetColumn(headerText1, 0)
headerGrid.Children.Add(headerText1)

# Create sub-header for the dropdown area
subHeaderGrid = Grid()
subHeaderGrid.Background = Brushes.WhiteSmoke  # Light gray, almost white
subHeaderCol1 = ColumnDefinition()
subHeaderCol1.Width = GridLength(120, GridUnitType.Pixel)  # Category column
subHeaderCol2 = ColumnDefinition()
subHeaderCol2.Width = GridLength(2, GridUnitType.Star)  # Family dropdown column
subHeaderCol3 = ColumnDefinition()
subHeaderCol3.Width = GridLength(1.5, GridUnitType.Star)  # Type dropdown column
subHeaderCol4 = ColumnDefinition()
subHeaderCol4.Width = GridLength(60, GridUnitType.Pixel)  # Buttons column
subHeaderGrid.ColumnDefinitions.Add(subHeaderCol1)
subHeaderGrid.ColumnDefinitions.Add(subHeaderCol2)
subHeaderGrid.ColumnDefinitions.Add(subHeaderCol3)
subHeaderGrid.ColumnDefinitions.Add(subHeaderCol4)

subHeaderText1 = TextBlock()
subHeaderText1.Text = "Category"
subHeaderText1.FontSize = 14
subHeaderText1.FontWeight = FontWeights.Bold
subHeaderText1.HorizontalAlignment = HorizontalAlignment.Left
subHeaderText1.Margin = Thickness(5, 0, 0, 0)
WpfGrid.SetColumn(subHeaderText1, 0)
subHeaderGrid.Children.Add(subHeaderText1)

subHeaderText2 = TextBlock()
subHeaderText2.Text = "Family"
subHeaderText2.FontSize = 14
subHeaderText2.FontWeight = FontWeights.Bold
subHeaderText2.HorizontalAlignment = HorizontalAlignment.Left
subHeaderText2.Margin = Thickness(5, 0, 0, 0)
WpfGrid.SetColumn(subHeaderText2, 1)
subHeaderGrid.Children.Add(subHeaderText2)

subHeaderText3 = TextBlock()
subHeaderText3.Text = "Type"
subHeaderText3.FontSize = 14
subHeaderText3.FontWeight = FontWeights.Bold
subHeaderText3.HorizontalAlignment = HorizontalAlignment.Left
subHeaderText3.Margin = Thickness(5, 0, 0, 0)
WpfGrid.SetColumn(subHeaderText3, 2)
subHeaderGrid.Children.Add(subHeaderText3)

subHeaderText4 = TextBlock()
subHeaderText4.Text = "Actions"
subHeaderText4.FontSize = 14
subHeaderText4.FontWeight = FontWeights.Bold
subHeaderText4.HorizontalAlignment = HorizontalAlignment.Center
WpfGrid.SetColumn(subHeaderText4, 3)
subHeaderGrid.Children.Add(subHeaderText4)

WpfGrid.SetColumn(subHeaderGrid, 1)
headerGrid.Children.Add(subHeaderGrid)

headerBorder = Border()
headerBorder.BorderBrush = Brushes.Gray
headerBorder.BorderThickness = Thickness(0, 0, 0, 1)
headerBorder.Child = headerGrid

mapping_stack_panel.Children.Insert(0, headerBorder)

#------------------------------------------------------------------------------
# Dictionary to hold dropdown containers keyed by CAD block name.
mapping_dict = {}
# Counter for unique control IDs
control_counter = 0

# Helper function to create a dropdown row
def create_dropdown_row(cad_name, preselected_family=None, is_first=True, category=None):
    global control_counter
    control_counter += 1

    rowGrid = Grid()
    rowGrid.Margin = Thickness(2)

    # Five columns: Category, Family dropdown, Type dropdown, X button, + button
    col0 = ColumnDefinition()
    col0.Width = GridLength(120, GridUnitType.Pixel)  # Category column
    col1 = ColumnDefinition()
    col1.Width = GridLength(2, GridUnitType.Star)  # Family dropdown
    col2 = ColumnDefinition()
    col2.Width = GridLength(1.5, GridUnitType.Star)  # Type dropdown
    col3 = ColumnDefinition()
    col3.Width = GridLength(30, GridUnitType.Pixel)  # X button
    col4 = ColumnDefinition()
    col4.Width = GridLength(30, GridUnitType.Pixel)  # + button
    rowGrid.ColumnDefinitions.Add(col0)
    rowGrid.ColumnDefinitions.Add(col1)
    rowGrid.ColumnDefinitions.Add(col2)
    rowGrid.ColumnDefinitions.Add(col3)
    rowGrid.ColumnDefinitions.Add(col4)

    # Create Category label
    categoryLabel = TextBlock()
    categoryLabel.Text = category if category else "All Categories"
    categoryLabel.VerticalAlignment = VerticalAlignment.Center
    categoryLabel.HorizontalAlignment = HorizontalAlignment.Left
    categoryLabel.Margin = Thickness(5, 0, 5, 0)
    categoryLabel.FontSize = 11
    categoryLabel.Foreground = Brushes.DarkGray
    WpfGrid.SetColumn(categoryLabel, 0)
    rowGrid.Children.Add(categoryLabel)

    # Create Family ComboBox
    familyCmb = ComboBox()
    familyCmb.IsEditable = True
    familyCmb.IsTextSearchEnabled = True
    familyCmb.HorizontalAlignment = HorizontalAlignment.Stretch
    familyCmb.Margin = Thickness(0, 0, 5, 0)
    # Create a valid WPF control name by removing invalid characters
    safe_name = create_safe_control_name(cad_name, control_counter)
    familyCmb.Name = "familyCmb_{}_{}".format(safe_name, control_counter)

    # Create Type ComboBox
    typeCmb = ComboBox()
    typeCmb.IsEditable = True
    typeCmb.IsTextSearchEnabled = True
    typeCmb.HorizontalAlignment = HorizontalAlignment.Stretch
    typeCmb.Margin = Thickness(0, 0, 5, 0)
    typeCmb.Name = "typeCmb_{}_{}".format(safe_name, control_counter)
    typeCmb.IsEnabled = False  # Disabled until family is selected

    # Populate the Family ComboBox with families filtered by category
    if category and category in families_by_category:
        # Use category-filtered family list
        ## print("DEBUG: Family dropdown - Filtering by category '{}', found {} families".format(category, len(families_by_category[category])))
        for family in families_by_category[category]:
            familyCmb.Items.Add(family)
    else:
        # Use all families if no category specified
        all_families = sorted(types_by_family.keys())
        ## print("DEBUG: Family dropdown - No category filter, showing all {} families".format(len(all_families)))
        for family in all_families:
            familyCmb.Items.Add(family)

    # Pre-select family and type if provided
    preselected_family_name = None
    preselected_type_name = None
    if preselected_family and " : " in preselected_family:
        parts = preselected_family.split(" : ")
        if len(parts) == 2:
            preselected_type_name = parts[0]
            preselected_family_name = parts[1]

    if preselected_family_name:
        # Check if family exists in the current dropdown
        family_found = False
        for i in range(familyCmb.Items.Count):
            if str(familyCmb.Items[i]) == preselected_family_name:
                familyCmb.SelectedIndex = i
                family_found = True
                break

        if family_found and preselected_type_name:
            # Populate and select the type
            if preselected_family_name in types_by_family:
                typeCmb.IsEnabled = True
                for type_name in types_by_family[preselected_family_name]:
                    typeCmb.Items.Add(type_name)
                # Try to select the preselected type
                for i in range(typeCmb.Items.Count):
                    if str(typeCmb.Items[i]) == preselected_type_name:
                        typeCmb.SelectedIndex = i
                        break

    # Add event handler for family selection change
    def family_selection_changed(sender, args):
        selected_family = familyCmb.SelectedItem
        if selected_family:
            family_name = str(selected_family)
            # Clear and populate type dropdown
            typeCmb.Items.Clear()
            typeCmb.IsEnabled = True
            if family_name in types_by_family:
                for type_name in types_by_family[family_name]:
                    typeCmb.Items.Add(type_name)
            typeCmb.SelectedIndex = -1
        else:
            # Disable type dropdown if no family selected
            typeCmb.Items.Clear()
            typeCmb.IsEnabled = False

    familyCmb.SelectionChanged += family_selection_changed

    WpfGrid.SetColumn(familyCmb, 1)
    rowGrid.Children.Add(familyCmb)

    WpfGrid.SetColumn(typeCmb, 2)
    rowGrid.Children.Add(typeCmb)

    # Create X button (for removing this dropdown)
    x_btn = Button()
    x_btn.Content = "X"
    x_btn.Width = 25
    x_btn.Height = 25
    x_btn.Margin = Thickness(2)
    x_btn.IsEnabled = not is_first  # First dropdown cannot be removed
    x_btn.Name = "x_btn_{}_{}".format(safe_name, control_counter)
    WpfGrid.SetColumn(x_btn, 3)
    rowGrid.Children.Add(x_btn)

    # Create + button (for adding new dropdown)
    plus_btn = Button()
    plus_btn.Content = "+"
    plus_btn.Width = 25
    plus_btn.Height = 25
    plus_btn.Margin = Thickness(2)
    plus_btn.Name = "plus_btn_{}_{}".format(safe_name, control_counter)
    WpfGrid.SetColumn(plus_btn, 4)
    rowGrid.Children.Add(plus_btn)

    return rowGrid, familyCmb, typeCmb, x_btn, plus_btn

# Helper function to add a new dropdown to a CAD block container
def add_dropdown_to_container(cad_name, container, preselected_family=None, category=None):
    # First, hide all existing + buttons in this container
    for child in container.Children:
        if hasattr(child, 'Children'):
            for grandchild in child.Children:
                if hasattr(grandchild, 'Content') and grandchild.Content == "+":
                    grandchild.Visibility = Visibility.Hidden

    # Create new dropdown row
    is_first = container.Children.Count == 0
    new_row, familyCmb, typeCmb, x_btn, plus_btn = create_dropdown_row(cad_name, preselected_family, is_first, category)

    # Add event handlers
    def x_clicked(sender, args):
        container.Children.Remove(new_row)
        # Show the + button on the last remaining row
        if container.Children.Count > 0:
            last_child = container.Children[container.Children.Count - 1]
            if hasattr(last_child, 'Children'):
                for grandchild in last_child.Children:
                    if hasattr(grandchild, 'Content') and grandchild.Content == "+":
                        grandchild.Visibility = Visibility.Visible
                        break

    def plus_clicked(sender, args):
        # Show category selection popup with actual Revit categories
        category_options = sorted(symbols_by_category.keys())
        selected_category = forms.SelectFromList.show(
            category_options,
            title="Select Revit Category",
            multiselect=False,
            button_name="OK"
        )

        if selected_category:
            # SelectFromList returns a list, so get the first item
            selected_cat = selected_category[0] if isinstance(selected_category, list) else str(selected_category)
            ## print("DEBUG: Plus button clicked, selected category: '{}'".format(selected_cat))
            add_dropdown_to_container(cad_name, container, category=selected_cat)

    x_btn.Click += x_clicked
    plus_btn.Click += plus_clicked

    container.Children.Add(new_row)
    return familyCmb, typeCmb

# Dynamically create a mapping section for each unique CAD block name
for name in sorted(unique_names):
    # Main container for this CAD block
    mainGrid = Grid()
    mainGrid.Margin = Thickness(5)

    col1 = ColumnDefinition()
    col1.Width = GridLength(1, GridUnitType.Star)
    col2 = ColumnDefinition()
    col2.Width = GridLength(2, GridUnitType.Star)  # Give more space to dropdown area
    mainGrid.ColumnDefinitions.Add(col1)
    mainGrid.ColumnDefinitions.Add(col2)

    # CAD block name
    textBlock = TextBlock()
    textBlock.Text = name
    textBlock.TextWrapping = TextWrapping.Wrap
    textBlock.VerticalAlignment = VerticalAlignment.Top
    textBlock.HorizontalAlignment = HorizontalAlignment.Stretch
    textBlock.Margin = Thickness(0, 5, 10, 0)
    WpfGrid.SetColumn(textBlock, 0)
    mainGrid.Children.Add(textBlock)

    # Container for dropdown rows
    dropdownContainer = StackPanel()
    dropdownContainer.Orientation = swc.Orientation.Vertical
    dropdownContainer.HorizontalAlignment = HorizontalAlignment.Stretch
    WpfGrid.SetColumn(dropdownContainer, 1)
    mainGrid.Children.Add(dropdownContainer)

    # Get pre-mapped families and groups for this CAD block
    auto_families = matchings_dict.get(name, [])
    auto_groups = groups_dict.get(name, [])
    dropdown_controls = []

    # Process families first
    if auto_families:
        # Create one dropdown for each pre-mapped family
        for i, family in enumerate(auto_families):
            # Determine which category this family belongs to using utility function
            family_category = determine_family_category(family, symbols_by_category)

            ## print("DEBUG: CSV family '{}' assigned to category '{}'".format(family, family_category))
            familyCmb, typeCmb = add_dropdown_to_container(name, dropdownContainer, family, family_category)
            dropdown_controls.append((familyCmb, typeCmb))

    # Process model groups
    if auto_groups:
        # Create one dropdown for each pre-mapped group (model group + detail group combo)
        for i, group in enumerate(auto_groups):
            # Model groups are in the "Model Groups" category
            ## print("DEBUG: CSV group '{}' assigned to category 'Model Groups'".format(group))
            familyCmb, typeCmb = add_dropdown_to_container(name, dropdownContainer, group, "Model Groups")
            dropdown_controls.append((familyCmb, typeCmb))

    # If no pre-mappings, create one empty dropdown
    if not auto_families and not auto_groups:
        # Create one empty dropdown with all categories (pass None for category to show all)
        ## print("DEBUG: CAD block '{}' has no pre-mapped families or groups, creating dropdown with all options".format(name))
        familyCmb, typeCmb = add_dropdown_to_container(name, dropdownContainer, category=None)
        dropdown_controls.append((familyCmb, typeCmb))

    # Store the dropdown container and controls for this CAD block
    mapping_dict[name] = {
        'container': dropdownContainer,
        'controls': dropdown_controls
    }

    # Wrap in a border
    border = Border()
    border.BorderBrush = Brushes.Gray
    border.BorderThickness = Thickness(0, 0, 0, 1)
    border.Margin = Thickness(10)
    border.Child = mainGrid

    mapping_stack_panel.Children.Add(border)


#------------------------------------------------------------------------------
# Prepare a result container.
result = {}

# Define event handlers for OK and Cancel.
def ok_clicked(sender, args):
    selections = {}
    for name, mapping_info in mapping_dict.items():
        container = mapping_info['container']
        selected_families = []

        # Go through all dropdown rows in the container
        for child in container.Children:
            if hasattr(child, 'Children'):
                family_name = None
                type_name = None

                # Find the family and type ComboBoxes in this row
                for grandchild in child.Children:
                    if hasattr(grandchild, 'Name') and grandchild.Name and grandchild.Name.startswith('familyCmb_'):
                        if grandchild.SelectedItem is not None:
                            family_name = str(grandchild.SelectedItem)
                    elif hasattr(grandchild, 'Name') and grandchild.Name and grandchild.Name.startswith('typeCmb_'):
                        if grandchild.SelectedItem is not None:
                            type_name = str(grandchild.SelectedItem)

                # Only add if both family and type are selected
                if family_name and type_name:
                    # Reconstruct the "Type : Family" format
                    family_label = "{} : {}".format(type_name, family_name)
                    selected_families.append(family_label)

        if selected_families:
            selections[name] = selected_families

    result["selections"] = selections
    result["divisor"] = divisor_textbox.Text
    window.DialogResult = True
    window.Close()

def cancel_clicked(sender, args):
    window.DialogResult = False
    window.Close()

ok_button.Click += ok_clicked
cancel_button.Click += cancel_clicked

#------------------------------------------------------------------------------
# Show the window as a dialog.

selected_level = forms.select_levels(multiple=False)
dialog_result = window.ShowDialog()

if not dialog_result:
    # print("User canceled fixture mapping.")
    pass
else:
    try:
        divisor = float(result["divisor"])
    except ValueError:
        # print("Invalid divisor value entered.")
        pass
    else:
        # Build a mapping for CAD blocks to lists of items to place (family symbols OR group tuples)
        name_to_items_map = {}
        for name, item_labels in result["selections"].items():
            items_to_place = []
            # Track occurrence count for each label to handle duplicates
            label_occurrence_count = {}

            for label in item_labels:
                # Determine occurrence index for this label
                occurrence_index = label_occurrence_count.get(label, 0)
                label_occurrence_count[label] = occurrence_index + 1

                # Check if it's a family symbol
                if label in symbol_label_map:
                    items_to_place.append((symbol_label_map[label], label, occurrence_index))
                # Check if it's a model group
                elif label in group_label_map:
                    items_to_place.append((group_label_map[label], label, occurrence_index))
                else:
                    # print("WARNING: Item '{}' not found in project. Skipping.".format(label))
                    pass
            if items_to_place:
                name_to_items_map[name] = items_to_place

        t = Transaction(doc, "Place Families from CSV")
        t.Start()
        try:
            for row in xyz_rows:
                cad_name = row.get("Name", "").strip()
                if not cad_name:
                    continue

                x_str = row.get("Position X", "").strip()
                y_str = row.get("Position Y", "").strip()
                z_str = row.get("Position Z", "").strip()

                if not x_str or not y_str or not z_str:
                    # print("Skipping row due to missing coordinate:", row)
                    continue

                x_inches = feet_inch_to_inches(x_str)
                y_inches = feet_inch_to_inches(y_str)
                z_inches = feet_inch_to_inches(z_str)
                if x_inches is None or y_inches is None or z_inches is None:
                    # print("Skipping row due to conversion error:", row)
                    continue

                try:
                    rot_deg = float(row.get("Rotation", 0.0))
                except ValueError:
                    ## print("Skipping row with invalid rotation:", row)
                    continue

                x = x_inches / divisor
                y = y_inches / divisor
                z = z_inches / divisor
                loc = XYZ(x, y, z)

                # Get all items mapped to this CAD block name (families OR groups)
                items = name_to_items_map.get(cad_name, [])
                if not items:
                    ## print("No items mapped for '{}'. Skipping row.".format(cad_name))
                    continue

                # Place each item at the same location
                for item_wrapper in items:
                    # Safely unpack the 3-tuple (item, label, occurrence_index)
                    if isinstance(item_wrapper, tuple) and len(item_wrapper) == 3:
                        actual_item, label_from_ui, occurrence_index = item_wrapper
                    else:
                        # Backward compatibility: old format without occurrence tracking
                        actual_item = item_wrapper
                        label_from_ui = None
                        occurrence_index = 0

                    # Check if actual_item is a FamilySymbol or a Group tuple
                    if isinstance(actual_item, tuple) and len(actual_item) == 2:
                        # This is a model group: (model_group_type, detail_group_type or None)
                        from Autodesk.Revit.DB import Element
                        model_group_type, detail_group_type = actual_item

                        print("\n=== PROCESSING MODEL GROUP PLACEMENT ===")
                        print("CAD block: '{}'".format(cad_name))
                        print("Label from UI: '{}'".format(label_from_ui if label_from_ui else "None"))
                        print("Model group type ID: {}".format(model_group_type.Id if model_group_type else "None"))
                        print("Detail group type ID: {}".format(detail_group_type.Id if detail_group_type else "None"))

                        # Use label from UI if available, otherwise find it
                        if label_from_ui:
                            group_label = label_from_ui
                        else:
                            # Fallback: reverse lookup in group_label_map
                            group_label = None
                            for label, grp_tuple in group_label_map.items():
                                if grp_tuple[0].Id == model_group_type.Id:
                                    group_label = label
                                    break

                        # Get offset from JSON (default to 0)
                        offset_x_inches = 0.0
                        offset_y_inches = 0.0
                        offset_z_inches = 0.0
                        offset_rotation_deg = 0.0
                        if group_label and cad_name in offsets_dict:
                            label_offsets_array = offsets_dict[cad_name].get(group_label, [])

                            # Use occurrence index to get the correct offset
                            if isinstance(label_offsets_array, list) and occurrence_index < len(label_offsets_array):
                                label_offsets = label_offsets_array[occurrence_index]
                            elif isinstance(label_offsets_array, dict):
                                # Backward compatibility: old single-offset format
                                label_offsets = label_offsets_array
                            else:
                                label_offsets = {}

                            offset_x_inches = label_offsets.get("x", 0.0)
                            offset_y_inches = label_offsets.get("y", 0.0)
                            offset_z_inches = label_offsets.get("z", 0.0)
                            offset_rotation_deg = label_offsets.get("r", 0.0)

                        # Convert offset from inches to feet
                        offset_x_feet = offset_x_inches / 12.0
                        offset_y_feet = offset_y_inches / 12.0
                        offset_z_feet = offset_z_inches / 12.0

                        # Apply offset to base location
                        offset_loc = XYZ(x + offset_x_feet, y + offset_y_feet, z + offset_z_feet)

                        # Place model group at offset location
                        model_group_name = Element.Name.__get__(model_group_type)
                        ## print("DEBUG: Placing model group '{}' at offset location {}".format(model_group_name, offset_loc))
                        model_group_instance = doc.Create.PlaceGroup(offset_loc, model_group_type)

                        # DEBUG: Check what detail groups are actually available for this placed instance
                        print("\n=== CHECKING AVAILABLE DETAIL GROUPS ===")
                        print("Model group: '{}' (Type ID: {})".format(model_group_name, model_group_type.Id))
                        try:
                            available_detail_ids = model_group_instance.GetAvailableAttachedDetailGroupTypeIds()
                            print("GetAvailableAttachedDetailGroupTypeIds returned {} detail groups".format(len(available_detail_ids)))

                            if available_detail_ids:
                                print("Available attached detail groups:")
                                for detail_id in available_detail_ids:
                                    detail_type = doc.GetElement(detail_id)
                                    if detail_type:
                                        detail_name = Element.Name.__get__(detail_type)
                                        print("  - '{}' (ID: {})".format(detail_name, detail_id))
                            else:
                                print("NO attached detail groups available for this model group")
                        except Exception as ex:
                            print("ERROR getting available detail groups: {}".format(ex))

                        # Apply rotation to model group around BASE location
                        total_rotation = rot_deg + offset_rotation_deg
                        if abs(total_rotation) > 1e-6:
                            angle_radians = math.radians(total_rotation)
                            axis = Line.CreateBound(loc, loc + XYZ(0, 0, 1))  # loc is BASE
                            ElementTransformUtils.RotateElement(doc, model_group_instance.Id, axis, angle_radians)

                        # Show the detail group in the active view if specified
                        if detail_group_type:
                            detail_group_name = Element.Name.__get__(detail_group_type)
                            print("\n=== DETAIL GROUP ATTACHMENT ===")
                            print("Model group: '{}' (ID: {})".format(model_group_name, model_group_instance.Id))
                            print("Detail group to attach: '{}' (ID: {})".format(detail_group_name, detail_group_type.Id))

                            try:
                                # Direct approach like wmlib.py - just try to show the detail group
                                # It will only work if the detail group is pre-attached in the Group Editor
                                model_group_instance.ShowAttachedDetailGroups(doc.ActiveView, detail_group_type.Id)
                                print("SUCCESS: Attached detail group '{}' to model group '{}'".format(
                                    detail_group_name, model_group_name))
                            except Exception as ex:
                                # This is expected if the detail group isn't pre-attached in the Group Editor
                                if "is not attached" in str(ex) or "not available" in str(ex):
                                    print("INFO: Detail group '{}' is not attached to model group '{}' in the Group Editor".format(
                                        detail_group_name, model_group_name))
                                    print("      To fix: Edit the model group and attach the detail group in the Group Editor")
                                else:
                                    print("ERROR: Failed to attach detail group: {}".format(ex))

                    else:
                        # This is a FamilySymbol (existing logic)
                        symbol = actual_item
                        if not symbol.IsActive:
                            symbol.Activate()
                            doc.Regenerate()

                        # Use label from UI if available, otherwise find it
                        if label_from_ui:
                            symbol_label = label_from_ui
                        else:
                            # Fallback: reverse lookup in symbol_label_map
                            symbol_label = None
                            for label, sym in symbol_label_map.items():
                                if sym.Id == symbol.Id:
                                    symbol_label = label
                                    break

                        # Get offset from JSON (default to 0)
                        offset_x_inches = 0.0
                        offset_y_inches = 0.0
                        offset_z_inches = 0.0
                        offset_rotation_deg = 0.0

                        if symbol_label and cad_name in offsets_dict:
                            label_offsets_array = offsets_dict[cad_name].get(symbol_label, [])

                            # Use occurrence index to get the correct offset
                            if isinstance(label_offsets_array, list) and occurrence_index < len(label_offsets_array):
                                label_offsets = label_offsets_array[occurrence_index]
                            elif isinstance(label_offsets_array, dict):
                                # Backward compatibility: old single-offset format
                                label_offsets = label_offsets_array
                            else:
                                label_offsets = {}
                            offset_x_inches = label_offsets.get("x", 0.0)
                            offset_y_inches = label_offsets.get("y", 0.0)
                            offset_z_inches = label_offsets.get("z", 0.0)
                            offset_rotation_deg = label_offsets.get("r", 0.0)

                        # Convert offset from inches to feet
                        offset_x_feet = offset_x_inches / 12.0
                        offset_y_feet = offset_y_inches / 12.0
                        offset_z_feet = offset_z_inches / 12.0

                        # Apply offset to base location
                        # Z coordinate represents "Elevation from Level" for level-based families
                        offset_loc = XYZ(x + offset_x_feet, y + offset_y_feet, z + offset_z_feet)

                        # Place instance at offset location
                        instance = None
                        placement_succeeded = False

                        try:
                            # Check if this is an annotation family that requires view-based placement
                            requires_view_placement = False

                            if SUPPORTS_ANNOTATION_PLACEMENT:
                                try:
                                    if hasattr(symbol, 'Family') and hasattr(symbol.Family, 'FamilyPlacementType'):
                                        if symbol.Family.FamilyPlacementType == FamilyPlacementType.ViewBased:
                                            requires_view_placement = True
                                except:
                                    # If check fails, assume it's a model family
                                    pass

                            if requires_view_placement:
                                # Annotation families need to be placed in a view
                                current_view = doc.ActiveView
                                if not current_view:
                                    # print("ERROR: No active view available to place annotation '{}'. Skipping.".format(symbol_label))
                                    pass
                                elif current_view.ViewType == ViewType.ThreeD:
                                    # print("WARNING: Cannot place annotation '{}' at ({}, {}) - active view is 3D. Switch to a 2D view.".format(symbol_label, x, y))
                                    pass
                                else:
                                    instance = doc.Create.NewFamilyInstance(offset_loc, symbol, current_view)
                                    placement_succeeded = True
                            else:
                                # Model families use level-based placement
                                instance = doc.Create.NewFamilyInstance(offset_loc, symbol, selected_level, Structure.StructuralType.NonStructural)
                                placement_succeeded = True

                                # Override elevation if there's a Z offset (family types may have default elevations)
                                if offset_z_feet != 0.0:
                                    elev_param = instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                                    if elev_param and not elev_param.IsReadOnly:
                                        elev_param.Set(offset_z_feet)

                        except Exception as ex:
                            # print("ERROR: Failed to place '{}' at ({}, {}): {}".format(symbol_label, x, y, ex))
                            pass

                        # Only set parameters and apply rotation if placement succeeded
                        if placement_succeeded and instance is not None:
                            # Set parameters if they exist
                            if symbol_label and cad_name in parameters_dict:
                                params_array = parameters_dict.get(cad_name, {}).get(symbol_label, [])

                                # Use occurrence index to get the correct parameters
                                if isinstance(params_array, list) and occurrence_index < len(params_array):
                                    params = params_array[occurrence_index]
                                elif isinstance(params_array, dict):
                                    # Backward compatibility: old single-parameter format
                                    params = params_array
                                else:
                                    params = {}

                                # Apply parameters directly to instance
                                if params:
                                    print("DEBUG: Found {} parameters to set for symbol '{}'".format(len(params), symbol_label))
                                    for param_name, param_value in params.items():
                                        try:
                                            print("DEBUG: Attempting to set parameter '{}' = '{}'".format(param_name, param_value))
                                            param = instance.LookupParameter(param_name) or instance.get_Parameter(param_name)

                                            if not param:
                                                print("DEBUG: Parameter '{}' NOT FOUND on instance".format(param_name))
                                                continue

                                            if param.IsReadOnly:
                                                print("DEBUG: Parameter '{}' is READ-ONLY, skipping".format(param_name))
                                                continue

                                            storage_type = param.StorageType.ToString()
                                            print("DEBUG: Parameter '{}' storage type: {}".format(param_name, storage_type))

                                            # Type conversion handlers
                                            if storage_type == "Integer":
                                                param.Set(int(param_value))
                                                print("DEBUG: Set '{}' as INTEGER: {}".format(param_name, int(param_value)))
                                            elif storage_type == "Double":
                                                # Special handling for electrical unit parameters
                                                if "Apparent Load" in param_name:
                                                    forge_type_va = ForgeTypeId("autodesk.unit.unit:voltAmperes-1.0.1")
                                                    converted = UnitUtils.ConvertToInternalUnits(float(param_value), forge_type_va)
                                                    param.Set(converted)
                                                    print("DEBUG: Set '{}' as APPARENT LOAD: {} (converted to {})".format(param_name, param_value, converted))
                                                elif "Voltage_CED" in param_name:
                                                    forge_type_volts = ForgeTypeId("autodesk.unit.unit:volts-1.0.1")
                                                    converted = UnitUtils.ConvertToInternalUnits(float(param_value), forge_type_volts)
                                                    param.Set(converted)
                                                    print("DEBUG: Set '{}' as VOLTAGE: {} (converted to {})".format(param_name, param_value, converted))
                                                else:
                                                    param.Set(float(param_value))
                                                    print("DEBUG: Set '{}' as DOUBLE: {}".format(param_name, float(param_value)))
                                            elif storage_type == "ElementId":
                                                # ElementId parameters - likely from key schedule
                                                # Find all key schedules in the project
                                                key_schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
                                                key_schedules = [ks for ks in key_schedules if ks.Definition.IsKeySchedule]

                                                found_element_id = None

                                                # Search through all key schedules to find matching value
                                                for key_schedule in key_schedules:
                                                    # Get all rows (elements) in this key schedule
                                                    schedule_elements = FilteredElementCollector(doc, key_schedule.Id).ToElements()

                                                    for schedule_elem in schedule_elements:
                                                        # Check all parameters on this schedule element
                                                        for schedule_param in schedule_elem.Parameters:
                                                            if schedule_param.HasValue:
                                                                param_val = schedule_param.AsString() or str(schedule_param.AsValueString())
                                                                if param_val == str(param_value):
                                                                    found_element_id = schedule_elem.Id
                                                                    print("DEBUG: Found key schedule row with value '{}' in schedule '{}'".format(param_value, key_schedule.Name))
                                                                    break
                                                        if found_element_id:
                                                            break
                                                    if found_element_id:
                                                        break

                                                if found_element_id:
                                                    param.Set(found_element_id)
                                                    print("DEBUG: Set '{}' to key schedule ElementId for value: {}".format(param_name, param_value))
                                                else:
                                                    print("ERROR: Could not find key schedule entry with value '{}' for parameter '{}'".format(param_value, param_name))
                                            else:
                                                # For string/text parameters
                                                param.Set(str(param_value))
                                                print("DEBUG: Set '{}' as STRING: '{}'".format(param_name, str(param_value)))

                                        except Exception as ex:
                                            print("ERROR: Failed to set parameter '{}' = '{}': {}".format(param_name, param_value, ex))
                                else:
                                    print("DEBUG: NO parameters to set for symbol '{}'".format(symbol_label))

                            # Apply rotation around BASE location (not offset location)
                            total_rotation = rot_deg + offset_rotation_deg
                            if abs(total_rotation) > 1e-6:
                                try:
                                    angle_radians = math.radians(total_rotation)
                                    axis = Line.CreateBound(loc, loc + XYZ(0, 0, 1))  # loc is BASE
                                    ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle_radians)
                                except Exception as ex:
                                    # print("ERROR: Failed to rotate '{}': {}".format(symbol_label, ex))
                                    pass
            t.Commit()
            # print("Finished placing families from CSV.")
        except InvalidOperationException as ex:
            t.RollBack()
            # print("Transaction rolled back. Error:", ex)
