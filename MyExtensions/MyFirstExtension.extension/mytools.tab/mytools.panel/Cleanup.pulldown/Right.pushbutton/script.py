from pyrevit import script
from Autodesk.Revit import DB

# Access the current document and application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView

# Start a transaction to modify the grids
with DB.Transaction(doc, "Hide Right Bubbles on All Horizontal Grids in Current View") as t:
    t.Start()
    
    # Collect all grids in the active view
    collector = DB.FilteredElementCollector(doc, active_view.Id)
    grids = collector.OfClass(DB.Grid).ToElements()
    
    while True:
        horizontal_grid = None
        for grid in grids:
            if grid.Curve is not None and isinstance(grid.Curve, DB.Line):
                direction = grid.Curve.Direction
                # Check if the grid is horizontal (direction close to the X-axis)
                if abs(direction.X) > 0.99 and abs(direction.Y) < 0.01:
                    # Determine which bubble is on the right by comparing the start and end points
                    curve = grid.Curve
                    start_point = curve.GetEndPoint(0)
                    end_point = curve.GetEndPoint(1)
                    # West is further right than East
                    right_bubble = DB.DatumEnds.End0 if start_point.X < end_point.X else DB.DatumEnds.End1
                    
                    # If the right bubble is still visible, select this grid
                    if grid.IsBubbleVisibleInView(right_bubble, active_view):
                        horizontal_grid = grid
                        break
        
        if horizontal_grid:
            # Hide the right bubble in the current view using the correct parameter
            horizontal_grid.HideBubbleInView(right_bubble, active_view)
            script.get_logger().info("Right bubble hidden successfully on the first horizontal grid in the current view.")
        else:
            script.get_logger().warning("No more horizontal grids with visible right bubbles found in the current view.")
            break
    
    t.Commit()
