from pyrevit import forms
from Autodesk.Revit.DB import FilteredElementCollector, Grid, UnitUtils
from Autodesk.Revit.DB import UnitTypeId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import UIApplication

# Get the current Revit document and application
uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document

# Ask user to select a grid
selection = uidoc.Selection
reference = selection.PickObject(ObjectType.Element, "Select a Grid to Measure its Length")
selected_grid = doc.GetElement(reference.ElementId)

# Check if a grid was selected
if selected_grid:
    grid = selected_grid
    grid_curve = grid.Curve

    if grid_curve:
        # Get the length of the grid curve
        length = grid_curve.Length

        # Convert the length to feet (assuming internal units are in feet)
        length_in_feet = UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Feet)

        # Display the length to the user
        forms.alert("The length of the selected grid is {:.2f} feet".format(length_in_feet), exitscript=True)
    else:
        forms.alert("Could not determine the grid curve.", exitscript=True)
else:
    forms.alert("No grid selected.", exitscript=True)