from pyrevit import script
from Autodesk.Revit import DB, UI

# Access the current document and application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView

# Get the crop box of the active view
crop_box = active_view.CropBox

if crop_box is None:
    script.get_logger().warning("The active view does not have a crop region.")
else:
    # Get the minimum and maximum points of the crop box
    min_point = crop_box.Min
    max_point = crop_box.Max
    
    # Create a message with the XYZ positions of the crop box extents
    message = "Crop Region Extents:\nMin Point: X={}, Y={}, Z={}\nMax Point: X={}, Y={}, Z={}".format(
        min_point.X, min_point.Y, min_point.Z,
        max_point.X, max_point.Y, max_point.Z
    )
    
    # Show the message in a task dialog
    UI.TaskDialog.Show("Crop Region Extents", message)
