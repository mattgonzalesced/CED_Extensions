# -*- coding: utf-8 -*-
__title__ = "Pick Point and Print Coordinates"

from pyrevit import revit, DB, forms, script

# Define the room prompts
# prompts = [
#     "BACK OFFICE", "BREAK ROOM", "DAIRY/ PRODUCE COOLER", "EQUIPMENT PLATFORM",
#     "FREEZER", "FRONT OFFICE", "HALLWAY", "MEAT COOLER", "RECEIVING AREA"
# ]

prompts = ["point"]
def pick_points_with_prompts():
    picked_points = []
    output = script.get_output()

    # Define the total number of steps for the progress bar
    max_steps = len(prompts)

    # Initialize the progress bar
    with forms.ProgressBar(title="Initializing...", max_value=max_steps, steps=1) as progress:
        # Iterate over each room prompt
        for i, room in enumerate(prompts):
            try:
                # Update the progress bar title and progress count
                progress.title = "Pick point for: {} ({}/{})".format(room, i, max_steps)
                progress.update_progress(i, max_steps)

                # Prompt user to pick a point for the current room
                point_ref = revit.uidoc.Selection.PickPoint("Select a point for {}".format(room))

                if point_ref:
                    # Ensure the picked point is a DB.XYZ object
                    picked_point = DB.XYZ(point_ref.X, point_ref.Y, point_ref.Z)
                    picked_points.append((room, picked_point))
                else:
                    forms.alert("No point selected for {}. Exiting script.".format(room), exitscript=True)

                # Check if the user cancelled the progress bar
                if progress.cancelled:
                    forms.alert("Operation cancelled by user.", exitscript=True)
                    return

            except Exception as e:
                forms.alert("Error picking point for {}: {}".format(room, str(e)), exitscript=True)
                return

    # Print the results in the output window
    output.print_md("### Picked Points Coordinates")
    for room, point in picked_points:
        output.print_md("- **{}**: X={:.2f}, Y={:.2f}, Z={:.2f}".format(room, point.X, point.Y, point.Z))


# Run the function
pick_points_with_prompts()
