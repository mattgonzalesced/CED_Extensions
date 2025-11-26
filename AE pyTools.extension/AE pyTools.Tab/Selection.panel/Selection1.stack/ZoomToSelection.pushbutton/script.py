# -*- coding: utf-8 -*-
from pyrevit import revit, DB, script


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



def find_viewport_offset(view):
    # Get the model → projection transform
    twb_list = view.GetModelToProjectionTransforms()
    if not twb_list or twb_list.Count == 0:
        return DB.XYZ(0,0,0)

    T_mp = twb_list[0].GetModelToProjectionTransform()
    T_pm = T_mp.Inverse  # projection → model

    # Pick any known model coordinate anchor, e.g. (0,0,0)
    anchor_model = DB.XYZ(0, 0, 0)

    # Convert model → projection → model
    proj = T_mp.OfPoint(anchor_model)
    anchored_back = T_pm.OfPoint(proj)

    # Offset = what the view THINKS (0,0,0) is in model space
    offset = anchored_back - anchor_model
    return offset

def calculate_sheet_space_bbox(doc, uidoc, selection, sheet_view):
    """
    Returns the bounding box of selected model elements transformed
    into sheet space using each viewport's projection transform.
    """
    min_sheet = None
    max_sheet = None

    # Collect all viewports on the sheet
    viewports = DB.FilteredElementCollector(doc, sheet_view.Id)\
                  .OfClass(DB.Viewport)\
                  .ToElements()

    for elem in selection:
        elem_id = elem.Id

        # Find viewport where this element appears
        for vp in viewports:
            vp_view = doc.GetElement(vp.ViewId)
            if not vp_view:
                continue

            # Try to get bounding box *in the viewport's view*
            bbox = elem.get_BoundingBox(vp_view)
            if not bbox:
                continue

            # Transform from projection space -> sheet space
            tform = vp.GetProjectionToSheetTransform()

            # Transform corners
            sheet_min = tform.OfPoint(bbox.Min)
            sheet_max = tform.OfPoint(bbox.Max)

            # Aggregate global min/max
            if min_sheet is None:
                min_sheet = sheet_min
                max_sheet = sheet_max
            else:
                min_sheet = DB.XYZ(
                    min(min_sheet.X, sheet_min.X),
                    min(min_sheet.Y, sheet_min.Y),
                    min(min_sheet.Z, sheet_min.Z)
                )
                max_sheet = DB.XYZ(
                    max(max_sheet.X, sheet_max.X),
                    max(max_sheet.Y, sheet_max.Y),
                    max(max_sheet.Z, sheet_max.Z)
                )

    return min_sheet, max_sheet

def main():
    """
    Main function to handle zooming to selection with configurable zoom factors.
    """
    doc = revit.doc
    uidoc = revit.uidoc
    selection = revit.get_selection()

    if not selection:
        return

    config = script.get_config("zoom_selection_config")
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
                zoom_factor/2)
        # TODO adjust script to work correctly with sheets as the open view.
        # Get the active UIView
        active_ui_view = None
        for ui_view in uidoc.GetOpenUIViews():
            if ui_view.ViewId == doc.ActiveView.Id:
                active_ui_view = ui_view
                break

        # Zoom and center the view
        if active_ui_view:
            active_ui_view.ZoomAndCenterRectangle(expanded_min, expanded_max)

def main_sheet():
    doc = revit.doc
    uidoc = revit.uidoc
    selection = revit.get_selection()

    if not selection:
        return

    view = doc.ActiveView

    # --- Get model→projection transform ---
    twb_list = view.GetModelToProjectionTransforms()
    if not twb_list or twb_list.Count == 0:
        return

    t_model_to_proj = twb_list[0].GetModelToProjectionTransform()
    t_proj_to_model = t_model_to_proj.Inverse  # property

    min_x = None
    min_y = None
    max_x = None
    max_y = None

    for elem in selection:
        bbox = elem.get_BoundingBox(view)
        if not bbox:
            continue

        for pt in [bbox.Min, bbox.Max]:
            proj_pt = t_model_to_proj.OfPoint(pt)

            px = proj_pt.X
            py = proj_pt.Y

            if min_x is None:
                min_x = px
                min_y = py
                max_x = px
                max_y = py
            else:
                min_x = min(min_x, px)
                min_y = min(min_y, py)
                max_x = max(max_x, px)
                max_y = max(max_y, py)

    # projection-space center
    cx_proj = (min_x + max_x) / 2.0
    cy_proj = (min_y + max_y) / 2.0

    # convert projection center back to MODEL SPACE
    center_model = t_proj_to_model.OfPoint(DB.XYZ(cx_proj, cy_proj, 0))

    # --- APPLY THE OFFSET FIX HERE ---
    offset = find_viewport_offset(view)

    corrected_center = DB.XYZ(
        center_model.X - offset.X,
        center_model.Y - offset.Y,
        center_model.Z
    )

    # tiny rectangle for centering
    delta = 2
    pmin = DB.XYZ(corrected_center.X - delta, corrected_center.Y - delta, corrected_center.Z)
    pmax = DB.XYZ(corrected_center.X + delta, corrected_center.Y + delta, corrected_center.Z)

    # zoom active view
    active_ui = None
    for uv in uidoc.GetOpenUIViews():
        if uv.ViewId == view.Id:
            active_ui = uv
            break

    if active_ui:
        active_ui.ZoomAndCenterRectangle(pmin, pmax)





# Execute the script
if __name__ == "__main__":
    main()
