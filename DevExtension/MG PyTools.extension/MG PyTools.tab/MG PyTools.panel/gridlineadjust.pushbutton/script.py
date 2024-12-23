from pyrevit import revit, DB

def set_wall_length(selected_elements):
    """
    Updates the length of selected walls to 40 feet by modifying their location curve.
    """
    # Convert 40 feet to internal units (Revit uses feet internally)
    target_length = 40.0

    # Start a transaction to modify Revit elements
    with revit.Transaction("Set Wall Length"):
        for element in selected_elements:
            if isinstance(element, DB.Wall):
                # Get the wall's location curve
                location_curve = element.Location
                if isinstance(location_curve, DB.LocationCurve):
                    curve = location_curve.Curve

                    # Get the start and end points of the curve
                    start_point = curve.GetEndPoint(0)
                    end_point = curve.GetEndPoint(1)

                    # Calculate the new end point to set the desired length
                    direction = (end_point - start_point).Normalize()
                    new_end_point = start_point + direction * target_length

                    # Create a new line with the updated length
                    new_curve = DB.Line.CreateBound(start_point, new_end_point)
                    location_curve.Curve = new_curve

                    print("Updated wall ID {} to Length = 40 feet".format(element.Id))
                else:
                    print("Wall ID {} does not have a valid location curve.".format(element.Id))

# Get the selected elements in the active Revit document
selected_elements = revit.get_selection()

if not selected_elements:
    print("No elements selected. Please select walls and try again.")
else:
    set_wall_length(selected_elements)
