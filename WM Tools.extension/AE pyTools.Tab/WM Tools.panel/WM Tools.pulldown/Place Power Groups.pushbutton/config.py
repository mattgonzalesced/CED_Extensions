# -*- coding: utf-8 -*-
from pyrevit import revit, script, forms

doc = revit.doc
logger = script.get_logger()
config = script.get_config("WM_power_group_offset")
offset_distance = config.get_option("group_placement_offset",0)

while True:
    user_input = forms.ask_for_string(
        title="Specify Offset Distance",
        prompt="Enter a distance (in decimal feet) to offset the placed group from \n"
               "the reference element.\n\n"
               "The offset is relative to the rotation:\n"
               "Positive = 'Above' or 'in front' of reference \n"
               "Negative = 'Below' or 'behind' reference",
        default=str(offset_distance)
    )

    if user_input is None:
        # User cancelled the input
        break

    try:
        offset_distance = float(user_input)
        break  # valid input, exit loop
    except:
        forms.alert("Invalid input. Please enter a numeric value in decimal feet.")

config.group_placement_offset = offset_distance
script.save_config()