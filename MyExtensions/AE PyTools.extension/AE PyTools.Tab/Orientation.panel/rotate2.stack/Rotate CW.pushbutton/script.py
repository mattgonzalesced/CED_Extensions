# -*- coding: utf-8 -*-
__title__   = "Rotate CW"
__doc__     = """Version = 1.0
Date    = 09.04.2024
________________________________________________________________
Description:
Rotates selected model elements 90 degrees clockwise while
maintaining their position. It is Similar to pressing the
spacebar, but works on elements connected with wires!

Note: 
This ignores Annotations, Pinned Elements, System Families, 
and Face-Based Elements hosted to vertical walls. 
________________________________________________________________
How-To:
Select Elements & push the button :)

________________________________________________________________
TODO:
[FEATURE] - 
________________________________________________________________
Last Updates:
- [09.04.2024] v1.0 
________________________________________________________________
Author: AEvelina"""


# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================

import clr
clr.AddReference('System')

from Snippets._rotateutils import collect_data_for_rotation_or_orientation, rotate_elements_group
from pyrevit import revit, DB, script
import math

# Get the active document
doc = revit.doc

config = script.get_config("orientation_config")

adjust_tag_position = getattr(config, "tag_position", True)
adjust_tag_angle = getattr(config, "tag_angle", False)

# Step 1: Get the selected elements and filter out pinned ones
selection = revit.get_selection()
filtered_selection = [el for el in selection if isinstance(el, DB.FamilyInstance) and not el.Pinned]

# Step 2: Pre-collect all necessary data before starting the transaction
element_data = collect_data_for_rotation_or_orientation(doc, filtered_selection,adjust_tag_position)

# Step 3: Define the rotation angle (90 degrees clockwise)
fixed_angle = -math.pi / 2

# Step 4: Rotate Elements and Adjust Tags in a Single Transaction
with DB.Transaction(doc, "Rotate Elements and Adjust Tags") as trans:
    trans.Start()

    for orientation_key, grouped_data in element_data.items():
        rotate_elements_group(doc, grouped_data, fixed_angle,adjust_tag_position)

    trans.Commit()
