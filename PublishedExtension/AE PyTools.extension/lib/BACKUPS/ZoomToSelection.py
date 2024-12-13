from pyrevit import revit, DB, forms

# Check if there is a selection
if forms.check_selection():
    # Get the current document and selection
    doc = revit.doc
    uidoc = revit.uidoc
    selection = revit.get_selection()

    # Initialize variables for the bounding box
    min_point = None
    max_point = None

    # Iterate through the selected elements to calculate the bounding box
    for elem in selection:
        # Get the bounding box of the element
        bbox = elem.get_BoundingBox(doc.ActiveView)
        if bbox:
            # Update the bounding box extents
            if min_point is None:
                min_point = bbox.Min
                max_point = bbox.Max
            else:
                min_point = DB.XYZ(
                    min(min_point.X, bbox.Min.X),
                    min(min_point.Y, bbox.Min.Y),
                    min(min_point.Z, bbox.Min.Z)
                )
                max_point = DB.XYZ(
                    max(max_point.X, bbox.Max.X),
                    max(max_point.Y, bbox.Max.Y),
                    max(max_point.Z, bbox.Max.Z)
                )

    # Check if a valid bounding box was found
    if min_point and max_point:
        # Get the active UIView
        active_ui_view = None
        for ui_view in uidoc.GetOpenUIViews():
            if ui_view.ViewId == doc.ActiveView.Id:
                active_ui_view = ui_view
                break

        # Use ZoomAndCenterRectangle if the active UIView is found
        if active_ui_view:
            active_ui_view.ZoomAndCenterRectangle(min_point, max_point)
