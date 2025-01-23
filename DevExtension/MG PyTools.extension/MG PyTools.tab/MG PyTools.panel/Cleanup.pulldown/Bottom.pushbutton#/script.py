from pyrevit import script
from Autodesk.Revit import DB

# Access the current document and application
doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView

# Start a transaction to modify the grids
transaction = DB.Transaction(doc, "Hide Bottom Bubbles on All Vertical Grids in Current View")
transaction.Start()

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
                if direction.Y > 0:
                    # If the bottom bubble is still visible, select this grid
                    if grid.IsBubbleVisibleInView(DB.DatumEnds.End1, active_view):
                        vertical_grid = (grid, DB.DatumEnds.End1)
                        break
                elif direction.Y < 0:
                    # If the top bubble is still visible, select this grid
                    if grid.IsBubbleVisibleInView(DB.DatumEnds.End0, active_view):
                        vertical_grid = (grid, DB.DatumEnds.End0)
                        break

    if vertical_grid:
        grid, bubble_end = vertical_grid
        # Hide the appropriate bubble in the current view
        grid.HideBubbleInView(bubble_end, active_view)
        # script.get_logger().info("Bubble hidden successfully on a vertical grid in the current view.")
    else:
        # script.get_logger().warning("No more vertical grids with visible bubbles found in the current view.")
        break

# Commit the transaction after all changes
transaction.Commit()
