# -*- coding: utf-8 -*-
import codecs
import csv
import math
import os

import clr
# Revit API imports
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilySymbol,
    Structure,
    XYZ,
    Transaction,
    BuiltInCategory,
    BuiltInParameter,
    ElementTransformUtils,
    Line
)
from Autodesk.Revit.Exceptions import InvalidOperationException
# pyRevit imports
from pyrevit import revit, forms, script

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

# Import required WPF types.
from System.Windows import Thickness, HorizontalAlignment, GridLength, GridUnitType, TextWrapping, FontWeights
from System.Windows.Controls import ComboBox, Border, Grid, ColumnDefinition, TextBlock
from System.Windows.Media import Brushes
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

#------------------------------------------------------------------------------
# Helper function to convert a feet-inches string to total inches.
def feet_inch_to_inches(value):
    """
    Converts a string like "139'-10 3/16"" into total inches (float).
    If the feet part is negative, the inches are subtracted.
    For example, "-5'6"" returns -66.
    Returns None if conversion fails.
    """
    try:
        value = value.strip()
        if not value:
            return None
        parts = value.split("'")
        if len(parts) < 2:
            return float(value)
        # Convert the feet part.
        feet = float(parts[0])
        # Process the inches part.
        inch_part = parts[1].replace('"', '').strip()
        # Remove any negative sign from inches (we rely on feet sign)
        if inch_part.startswith("-"):
            inch_part = inch_part[1:]
        # Convert the inches part, handling fractions if present.
        if " " in inch_part:
            inch_parts = inch_part.split(" ")
            inches = float(inch_parts[0])
            if len(inch_parts) > 1:
                fraction = inch_parts[1]
                num, denom = fraction.split("/")
                inches += float(num) / float(denom)
        else:
            if inch_part == "":
                inches = 0.0
            else:
                inches = float(inch_part)
        # If feet is negative, subtract the inches instead of adding.
        if feet < 0:
            return feet * 12 - inches
        else:
            return feet * 12 + inches
    except Exception as ex:
        print("Error converting '{0}' to inches: {1}".format(value, ex))
        return None

#------------------------------------------------------------------------------
# Get active Revit document.
doc = revit.doc

#------------------------------------------------------------------------------
# 1. Prompt the user to select a CSV file.
csv_path = forms.pick_file(file_ext='csv', title='Select the Lighting Fixture CSV')
if not csv_path:
    script.exit()  # Exit if no file is selected

#------------------------------------------------------------------------------
# 2. Read CSV data and collect unique fixture names.
fixture_rows = []
unique_names = set()

with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f, delimiter=',')
    if reader.fieldnames:
        reader.fieldnames = [h.strip() for h in reader.fieldnames if h is not None]
    for row in reader:
        # Skip rows where the "Count" field is not "1".
        count_val = row.get("Count")
        if count_val is None or count_val.strip() != "1":
            continue
        # Skip rows with no Position X value.
        posX = row.get("Position X")
        if posX is None or posX.strip() == "":
            continue
        fixture_rows.append(row)
        fixture_name = row.get("Name")
        if fixture_name is None:
            print("WARNING: Row missing 'Name':", row)
        else:
            fixture_name = fixture_name.strip()
            if fixture_name:
                unique_names.add(fixture_name)
            else:
                print("WARNING: Row missing 'Name':", row)

if not unique_names:
    script.exit("No valid 'Name' column found or CSV is empty.")

def extract_fixture_type(csv_name):
    # Look for the substring between 'LUMINAIRE-' and '_Symbol'
    start_marker = "LUMINAIRE-"
    end_marker = "_Symbol"
    start_index = csv_name.find(start_marker)
    end_index = csv_name.find(end_marker, start_index)
    if start_index != -1 and end_index != -1:
        # Extract the substring after the start marker and before the end marker.
        # This automatically trims out any extra parts.
        return csv_name[start_index+len(start_marker):end_index].strip()
    return None

#------------------------------------------------------------------------------
# 3. Collect lighting fixture FamilySymbols from the project.
fixture_symbols = FilteredElementCollector(doc) \
    .OfClass(FamilySymbol) \
    .OfCategory(BuiltInCategory.OST_LightingFixtures) \
    .ToElements()

# Build a mapping of label -> symbol.
# Format the label as "Type : Family".
symbol_label_map = {}
for sym in fixture_symbols:
    try:
        family_name = sym.Family.Name
    except Exception:
        family_name = "UnknownFamily"
    try:
        param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = param.AsString() if param else "UnknownType"
    except Exception:
        type_name = "UnknownType"
    label = "{} : {}".format(type_name, family_name)
    symbol_label_map[label] = sym

symbol_labels_sorted = sorted(symbol_label_map.keys())
if not symbol_labels_sorted:
    script.exit("No valid lighting fixture FamilySymbols were found in the project.")

#------------------------------------------------------------------------------
# 4. Load the XAML UI from a separate file.
xaml_path = os.path.join(os.path.dirname(__file__), "MappingWindow.XAML")
fs = FileStream(xaml_path, FileMode.Open)
window = XamlReader.Load(fs)
fs.Close()

# Get references to the named controls.
mapping_stack_panel = window.FindName("MappingStackPanel")
divisor_textbox = window.FindName("DivisorTextBox")
ok_button = window.FindName("OkButton")
cancel_button = window.FindName("CancelButton")

#------------------------------------------------------------------------------
# Add a header row for the columns.
headerGrid = Grid()
headerGrid.Margin = Thickness(5)

headerCol1 = ColumnDefinition()
headerCol1.Width = GridLength(1, GridUnitType.Star)
headerCol2 = ColumnDefinition()
headerCol2.Width = GridLength(1, GridUnitType.Star)
headerGrid.ColumnDefinitions.Add(headerCol1)
headerGrid.ColumnDefinitions.Add(headerCol2)

headerText1 = TextBlock()
headerText1.Text = "CSV Fixture Name"
headerText1.FontSize = 16
headerText1.FontWeight = FontWeights.Bold
headerText1.HorizontalAlignment = HorizontalAlignment.Stretch


from System.Windows.Controls import Grid as WpfGrid
WpfGrid.SetColumn(headerText1, 0)
headerGrid.Children.Add(headerText1)

headerText2 = TextBlock()
headerText2.Text = "Revit Family (Type : Family)"
headerText2.FontSize = 16
headerText2.FontWeight = FontWeights.Bold
headerText2.HorizontalAlignment = HorizontalAlignment.Stretch
WpfGrid.SetColumn(headerText2, 1)
headerGrid.Children.Add(headerText2)

headerBorder = Border()
headerBorder.BorderBrush = Brushes.Gray
headerBorder.BorderThickness = Thickness(0, 0, 0, 1)
headerBorder.Child = headerGrid

mapping_stack_panel.Children.Insert(0, headerBorder)

#------------------------------------------------------------------------------
# Dictionary to hold ComboBox controls keyed by fixture name.
mapping_dict = {}

# Dynamically create a mapping row for each unique fixture name using a Grid.
for name in sorted(unique_names):
    rowGrid = Grid()
    rowGrid.Margin = Thickness(5)
    
    col1 = ColumnDefinition()
    col1.Width = GridLength(1, GridUnitType.Star)
    col2 = ColumnDefinition()
    col2.Width = GridLength(1, GridUnitType.Star)
    rowGrid.ColumnDefinitions.Add(col1)
    rowGrid.ColumnDefinitions.Add(col2)
    
    textBlock = TextBlock()
    textBlock.Text = name
    textBlock.TextWrapping = TextWrapping.Wrap
    textBlock.HorizontalAlignment = HorizontalAlignment.Stretch
    WpfGrid.SetColumn(textBlock, 0)
    rowGrid.Children.Add(textBlock)
    
    cmb = ComboBox()
    cmb.IsEditable = True
    cmb.IsTextSearchEnabled = True
    cmb.HorizontalAlignment = HorizontalAlignment.Stretch
    cmb.Margin = Thickness(5, 0, 0, 0)
    
    # Populate the ComboBox with available symbol labels.
    for item in symbol_labels_sorted:
        cmb.Items.Add(item)
    
    # Extract the fixture type code from the CSV fixture name.
    fixture_type_code = extract_fixture_type(name)
    preselected = False
    if fixture_type_code:
        # Try to match this code with the Fixture Type_CEDT parameter for each symbol.
        for index, label in enumerate(symbol_labels_sorted):
            symbol = symbol_label_map[label]
            # Use LookupParameter to get the type parameter "Fixture Type_CEDT"
            param = symbol.LookupParameter("Identity Type Mark")
            if param:
                symbol_type_value = param.AsString()  # or param.AsValueString() if needed
                # Compare case-insensitively
                if symbol_type_value and symbol_type_value.strip().lower() == fixture_type_code.lower():
                    cmb.SelectedIndex = index  # preselect this matching item
                    preselected = True
                    break
    # If no match was found, do not set any default, leaving the ComboBox blank.
    if not preselected:
        cmb.SelectedIndex = -1  # Ensure it's blank; alternatively, just do nothing if default is None.
    
    WpfGrid.SetColumn(cmb, 1)
    rowGrid.Children.Add(cmb)
    
    # Wrap the rowGrid in a Border to add a divider.
    border = Border()
    border.BorderBrush = Brushes.Gray
    border.BorderThickness = Thickness(0, 0, 0, 1)
    border.Margin = Thickness(10)
    border.Child = rowGrid
    
    mapping_dict[name] = cmb
    mapping_stack_panel.Children.Add(border)


#------------------------------------------------------------------------------
# Prepare a result container.
result = {}

# Define event handlers for OK and Cancel.
def ok_clicked(sender, args):
    selections = {}
    for name, cmb in mapping_dict.items():
        # If no selection was made, skip that fixture.
        if cmb.SelectedItem is None:
            continue
        selections[name] = cmb.SelectedItem
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
    print("User canceled fixture mapping.")
else:
    try:
        divisor = float(result["divisor"])
    except ValueError:
        print("Invalid divisor value entered.")
    else:
        # Build a mapping only for fixtures where the user made a selection.
        name_to_symbol_map = {name: symbol_label_map[label]
                              for name, label in result["selections"].items()}
        t = Transaction(doc, "Place CSV Lighting Fixtures")
        t.Start()        
        try:
            for row in fixture_rows:
                fixture_name = row.get("Name", "").strip()
                if not fixture_name:
                    continue

                x_str = row.get("Position X", "").strip()
                y_str = row.get("Position Y", "").strip()
                z_str = row.get("Position Z", "").strip()

                if not x_str or not y_str or not z_str:
                    print("Skipping row due to missing coordinate:", row)
                    continue

                x_inches = feet_inch_to_inches(x_str)
                y_inches = feet_inch_to_inches(y_str)
                z_inches = feet_inch_to_inches(z_str)
                if x_inches is None or y_inches is None or z_inches is None:
                    print("Skipping row due to conversion error:", row)
                    continue

                try:
                    rot_deg = float(row.get("Rotation", 0.0))
                except ValueError:
                    print("Skipping row with invalid rotation:", row)
                    continue

                x = x_inches / divisor
                y = y_inches / divisor
                z = z_inches / divisor

                symbol = name_to_symbol_map.get(fixture_name)
                if not symbol:
                    print("No FamilySymbol mapped for '{}'. Skipping row.".format(fixture_name))
                    continue

                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()

                loc = XYZ(x, y, z)
                inst = doc.Create.NewFamilyInstance(loc, symbol, selected_level, Structure.StructuralType.NonStructural)

                if abs(rot_deg) > 1e-6:
                    angle_radians = math.radians(rot_deg)
                    axis = Line.CreateBound(loc, loc + XYZ(0, 0, 1))
                    ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle_radians)
            t.Commit()
            print("Finished placing fixtures from CSV.")
        except InvalidOperationException as ex:
            t.RollBack()
            print("Transaction rolled back. Error:", ex)
