from pyrevit import forms
from Autodesk.Revit.DB import FilteredElementCollector, Grid, UnitUtils, Line, XYZ, Transaction
from Autodesk.Revit.DB import UnitTypeId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import UIApplication

# Get the current Revit document and application
uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document

# Ask user to select a grid
selection = uidoc.Selection
reference = selection.PickObject(ObjectType.Element, "Select a Grid to Shift")
selected_grid = doc.GetElement(reference.ElementId)

# Check if a grid was selected
if selected_grid:
    grid = selected_grid
    grid_curve = grid.Curve

    if grid_curve:
        # Get the start and end points of the grid curve
        start_point = grid_curve.GetEndPoint(0)
        end_point = grid_curve.GetEndPoint(1)

        # Shift the Y value of the end point down by 20 units
        new_end_point = XYZ(end_point.X, end_point.Y - 20, end_point.Z)

        # Create a new line with the updated end point
        new_curve = Line.CreateBound(start_point, new_end_point)

        # Start a transaction to modify the grid
        with Transaction(doc, 'Shift Grid End Point') as t:
            t.Start()
            # Delete the old grid
            doc.Delete(grid.Id)
            # Create a new grid with the updated curve
            new_grid = Grid.Create(doc, new_curve)
            new_grid.Name = grid.Name  # Retain the original grid name
            t.Commit()

        # Display a message to the user
        forms.alert(
            "The Y value of the end point of the selected grid has been shifted down by 20 units.",
            exitscript=True
        )
    else:
        forms.alert("Could not determine the grid curve.", exitscript=True)
else:
    forms.alert("No grid selected.", exitscript=True)