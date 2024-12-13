# -*- coding: utf-8 -*-
__title__   = "Orient Down"
__doc__     = """Version = 1.0
Date    = 09.04.2024
________________________________________________________________
Description:
Orients selected model elements to face 'plan south' while 
maintaining their position. Useful if many elements need to
face the same direction. 
Also works on elements connected with wires! 

Note: 
This ignores Annotations, Pinned Elements, System Families, 
and Face-Based Elements hosted to vertical walls. 

Another Note: 
"Orientation" depends on how the family is built.
This assumes 'front' of the family points 'plan north'.
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

from Snippets._rotateutils import collect_data_for_rotation_or_orientation, orient_elements_group
from pyrevit import revit, DB, script

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

# Step 3: Define the target orientation (Y-axis for 0 degrees)
target_orientation = DB.XYZ(0, -1, 0)

# Step 4: Orient Elements and Adjust Tags in a Single Transaction
with DB.Transaction(doc, "Orient Elements and Adjust Tags") as trans:
    trans.Start()

    for orientation_key, grouped_data in element_data.items():
        orient_elements_group(doc, grouped_data, target_orientation,adjust_tag_position)

    trans.Commit()
