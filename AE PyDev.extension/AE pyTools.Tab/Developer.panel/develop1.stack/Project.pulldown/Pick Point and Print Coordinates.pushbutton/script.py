# -*- coding: utf-8 -*-
__title__ = "Pick Points with Temporary Graphics"

import os

from pyrevit import revit, DB, script, forms


def pick_points_with_temp_graphics():
    """Prompts the user to pick points and visualizes them with temporary graphics."""
    doc = revit.doc
    uidoc = revit.uidoc
    picked_points = []
    output = script.get_output()
    wb = forms.WarningBar(title="Click to Pick Points. Press ESC to Finish.")

    # Get the TemporaryGraphicsManager for the active view
    temp_graphics_manager = DB.TemporaryGraphicsManager.GetTemporaryGraphicsManager(doc)

    try:
        with wb:
            user_cancelled = False
            while not user_cancelled:
                try:
                    # Let the user pick a point
                    picked_point = uidoc.Selection.PickPoint("Select a point or press ESC to display results.")
                    if picked_point:
                        xyz_point = DB.XYZ(picked_point.X, picked_point.Y, picked_point.Z)
                        picked_points.append(xyz_point)

                        # Create a control with an image at the picked point
                        create_temp_image(temp_graphics_manager, xyz_point, doc.ActiveView.Id)
                except Exception:
                    # User pressed ESC or cancelled
                    user_cancelled = True

            # Show the output window with results and clear graphics
            output.show()
            output.log_success("Selection Complete")
            output.print_md("### Picked Points")
            for idx, point in enumerate(picked_points, start=1):
                output.print_md("- **Point {}**: X={:.2f}, Y={:.2f}, Z={:.2f}".format(idx, point.X, point.Y, point.Z))
            temp_graphics_manager.Clear()

    except Exception as e:
        script.get_logger().error("An error occurred: {}".format(str(e)))

def create_temp_image(temp_graphics_manager, point, view_id):
    """Creates a temporary image at the specified point using TemporaryGraphicsManager."""
    try:
        # Path to the image file
        script_dir = os.path.dirname(__file__)
        image_path = os.path.join(script_dir, "point.bmp")

        # Verify the image file exists
        if not os.path.exists(image_path):
            script.get_logger().error("Image file not found at: {}".format(image_path))
            return

        # Create InCanvasControlData for the image
        control_data = DB.InCanvasControlData(image_path, point)

        # Add the control to the TemporaryGraphicsManager
        temp_graphics_manager.AddControl(control_data, view_id)

    except Exception as e:
        script.get_logger().error("Error creating temporary image: {}".format(str(e)))

# Run the function
pick_points_with_temp_graphics()
