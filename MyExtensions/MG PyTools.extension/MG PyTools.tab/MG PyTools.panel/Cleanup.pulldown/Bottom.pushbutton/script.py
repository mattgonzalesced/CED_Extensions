from pyrevit import script
from Autodesk.Revit import DB
##TEST
# Access the current document and application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView

# Start a transaction to modify the grids
with DB.Transaction(doc, "Hide Bottom Bubbles on All Vertical Grids in Current View") as t:
    t.Start()
    
    # Collect all grids in the active view
    collector = DB.FilteredElementCollector(doc, active_view.Id)
    grids = collector.OfClass(DB.Grid).ToElements()
    
    while True:
        vertical_grid = None
        for grid in grids:
            if grid.Curve is not None and isinstance(grid.Curve, DB.Line):
                direction = grid.Curve.Direction
                # Check if the grid is vertical (direction close to the Y-axis)
                if abs(direction.X) < 0.01 and abs(direction.Y) > 0.99:
                    # If the bottom bubble is still visible, select this grid
                    if grid.IsBubbleVisibleInView(DB.DatumEnds.End1, active_view):
                        vertical_grid = grid
                        break
        
        if vertical_grid:
            # Hide the bottom bubble in the current view using the correct parameter
            vertical_grid.HideBubbleInView(DB.DatumEnds.End1, active_view)
            script.get_logger().info("Bottom bubble hidden successfully on the first vertical grid in the current view.")
        else:
            script.get_logger().warning("No more vertical grids with visible bottom bubbles found in the current view.")
            break
    
    t.Commit()
