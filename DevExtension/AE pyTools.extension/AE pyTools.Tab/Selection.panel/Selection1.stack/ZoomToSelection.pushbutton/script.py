from pyrevit import revit, DB, script

# Initialize configuration
config = script.get_config("zoom_selection_config")

# TODO test aspec ratio functionality. potentially scale down the zoom_factor by some constant
def expand_bounding_box(min_point, max_point, zoom_factor):
    """
    Expand the bounding box by a zoom factor.
    """

    width = max_point.X - min_point.X
    height = max_point.Y - min_point.Y
    aspect_ratio = max(width, height) / min(width, height)

    if width > height:
        x_factor = zoom_factor
        y_factor = zoom_factor / aspect_ratio
    else:
        x_factor = zoom_factor / aspect_ratio
        y_factor = zoom_factor

    x_expansion = (max_point.X - min_point.X) * (x_factor - 1) / 2
    y_expansion = (max_point.Y - min_point.Y) * (y_factor - 1) / 2
    z_expansion = (max_point.Z - min_point.Z) * (zoom_factor - 1) / 2  # Uniform for Z-axis

    expanded_min = DB.XYZ(min_point.X - x_expansion, min_point.Y - y_expansion, min_point.Z - z_expansion)
    expanded_max = DB.XYZ(max_point.X + x_expansion, max_point.Y + y_expansion, max_point.Z + z_expansion)

    return expanded_min, expanded_max


def calculate_bounding_box(selection, active_view):
    """
    Calculate the bounding box for the given selection in the active view.
    """
    min_point = None
    max_point = None

    for elem in selection:
        bbox = elem.get_BoundingBox(active_view)
        if bbox:
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

    return min_point, max_point


def main():
    """
    Main function to handle zooming to selection with configurable zoom factors.
    """
    # Get the document, UI document, and selection
    doc = revit.doc
    uidoc = revit.uidoc
    selection = revit.get_selection()

    if not selection:
        return

    # Retrieve the zoom factor from configuration
    zoom_factor = config.get_option("zoom_factor", 1.0)  # Default to 1.0 if not set

    # Calculate bounding box
    min_point, max_point = calculate_bounding_box(selection, doc.ActiveView)

    if min_point and max_point:
        if zoom_factor == 1:
            expanded_min, expanded_max = min_point, max_point
        else:
            expanded_min, expanded_max = expand_bounding_box(
                min_point,
                max_point,
                zoom_factor)
#TODO adjust script to work correctly with sheets as the open view.
        # Get the active UIView
        active_ui_view = None
        for ui_view in uidoc.GetOpenUIViews():
            if ui_view.ViewId == doc.ActiveView.Id:
                active_ui_view = ui_view
                break

        # Zoom and center the view
        if active_ui_view:
            active_ui_view.ZoomAndCenterRectangle(expanded_min, expanded_max)


# Execute the script
if __name__ == "__main__":
    main()
