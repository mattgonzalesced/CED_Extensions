# -*- coding: utf-8 -*-
import math
import csv
import codecs

# pyRevit imports
from pyrevit import revit, forms, script

# Revit API imports
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilySymbol,
    Structure,
    XYZ,
    Line,
    ElementTransformUtils,
    Transaction,
    BuiltInCategory,
    BuiltInParameter
)
from Autodesk.Revit.Exceptions import InvalidOperationException

# Get the active document
doc = revit.doc

# ---------------------------------------------------------------------
# 1) Prompt the user to select a CSV file
# ---------------------------------------------------------------------
csv_path = forms.pick_file(file_ext='csv', title='Select the Lighting Fixture CSV')
if not csv_path:
    script.exit()  # Exit if no file is selected

# ---------------------------------------------------------------------
# 2) Read CSV data using codecs (to handle BOM) and clean keys.
# ---------------------------------------------------------------------
fixture_rows = []
unique_names = set()

with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
    # Change delimiter if your CSV is not comma-delimited (e.g. delimiter='\t' for tab)
    reader = csv.DictReader(f, delimiter=',')
    if reader.fieldnames:
        reader.fieldnames = [h.strip() for h in reader.fieldnames]
    #print("DEBUG: CSV Columns found => {}".format(reader.fieldnames))
    for raw_row in reader:
        clean_row = {k.strip(): v for k, v in raw_row.items() if k is not None}
        fixture_rows.append(clean_row)
        fixture_name = clean_row.get("Name", "").strip()
        if fixture_name:
            unique_names.add(fixture_name)
        else:
            print("WARNING: Row missing 'Name' or it is empty:", clean_row)

if not unique_names:
    print("No valid 'Name' column found or CSV is empty. Exiting.")
    script.exit()

# ---------------------------------------------------------------------
# 3) Collect lighting fixture FamilySymbols from the project.
#     (We assume they are in the OST_LightingFixtures category.)
# ---------------------------------------------------------------------
fixture_symbols = FilteredElementCollector(doc) \
                    .OfClass(FamilySymbol) \
                    .OfCategory(BuiltInCategory.OST_LightingFixtures) \
                    .ToElements()

# Build a mapping of label -> symbol using safe attribute access.
# Here we use the built-in parameter SYMBOL_NAME_PARAM to get the type name.
symbol_label_map = {}
for sym in fixture_symbols:
    try:
        family_obj = sym.Family  # Should be available
        family_name = family_obj.Name
    except Exception as e:
        family_name = "UnknownFamily"
    # Get the type name from the SYMBOL_NAME_PARAM
    try:
        param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if param:
            type_name = param.AsString()
        else:
            type_name = "UnknownType"
    except Exception as e:
        type_name = "UnknownType"
    
    label = "{} : {}".format(family_name, type_name)
    symbol_label_map[label] = sym

symbol_labels_sorted = sorted(symbol_label_map.keys())
if not symbol_labels_sorted:
    print("No valid lighting fixture FamilySymbols were found in the project.")
    script.exit()

# ---------------------------------------------------------------------
# 4) For each unique CSV fixture name, prompt the user to pick a FamilySymbol.
# ---------------------------------------------------------------------
name_to_symbol_map = {}
for uname in sorted(unique_names):
    prompt_title = "Match CSV Fixture: {} with a lighting fixture type:".format(uname)
    chosen_label = forms.SelectFromList.show(
        symbol_labels_sorted,
        title=prompt_title,
        width=600,
        button_name='Select Family Type'
    )
    if not chosen_label:
        print("User canceled selection for '{}'. Exiting.".format(uname))
        script.exit()
    name_to_symbol_map[uname] = symbol_label_map[chosen_label]

# ---------------------------------------------------------------------
# 5) Ask the user for a divisor value for the XYZ coordinates.
# ---------------------------------------------------------------------
divisor_str = forms.ask_for_string(
    prompt="Enter coordinate divisor (all XYZ values will be divided by this):",
    default="12"
)
try:
    divisor = float(divisor_str)
except Exception as e:
    script.exit("Invalid divisor provided: {}".format(divisor_str))

# ---------------------------------------------------------------------
# 6) Place each fixture instance from the CSV in one transaction.
# ---------------------------------------------------------------------
t = Transaction(doc, "Place CSV Lighting Fixtures")
t.Start()

try:
    for row in fixture_rows:
        fixture_name = row.get("Name", "").strip()
        if not fixture_name:
            continue
        
        try:
            # Divide coordinates by the provided divisor.
            x = float(row.get("Position X", 0.0)) / divisor
            y = float(row.get("Position Y", 0.0)) / divisor
            z = float(row.get("Position Z", 0.0)) / divisor
            rot_deg = float(row.get("Rotation", 0.0))
        except ValueError:
            print("Skipping row with invalid numeric data:", row)
            continue
        
        symbol = name_to_symbol_map.get(fixture_name)
        if not symbol:
            print("No FamilySymbol mapped for '{}'. Skipping row.".format(fixture_name))
            continue
        
        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()
        
        loc = XYZ(x, y, z)
        inst = doc.Create.NewFamilyInstance(
            loc,
            symbol,
            Structure.StructuralType.NonStructural
        )
        
        if abs(rot_deg) > 1e-6:
            angle_radians = math.radians(rot_deg)
            axis = Line.CreateBound(loc, loc + XYZ(0, 0, 1))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle_radians)
    
    t.Commit()
    print("Finished placing fixtures from CSV.")
except InvalidOperationException as ex:
    t.RollBack()
    print("Transaction rolled back. Error:", ex)
