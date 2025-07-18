# -*- coding: utf-8 -*-

import clr
clr.AddReference('System')
from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script
from Snippets._elecutils import pick_circuits_from_list, get_circuits_from_selection

# Access the current Revit document
doc = revit.doc
uidoc = revit.uidoc
# Set up the output window
output = script.get_output()
output.close_others()
output.set_width(800)

output = script.get_output()
logger = script.get_logger()


# Step 1: Pick circuits
selection = revit.get_selection().elements
if selection:
    selected_circuits = get_circuits_from_selection(selection)
else:
    selected_circuits = pick_circuits_from_list(doc, select_multiple=True,include_spares_and_spaces=True)
if not selected_circuits:
    forms.alert("No circuits selected.", exitscript=True)

# Step 2: Choose scope of selection
action = forms.CommandSwitchWindow.show(
    ["Only Circuits", "Connected Elements", "Both"],
    message="What do you want to work with?",
    default="Connected Elements"
)

if not action:
    logger.info("No action selected. Exiting.")
    script.exit()

# Step 3: Get results based on choice
result_elements = set()

if action in ["Only Circuits", "Both"]:
    result_elements.update(selected_circuits)

if action in ["Connected Elements", "Both"]:
    for ckt in selected_circuits:
        elements = ckt.Elements
        if elements:
            result_elements.update([e for e in elements if isinstance(e, DB.Element)])

# Step 4: Output result
element_ids = List[DB.ElementId]([e.Id for e in result_elements])
uidoc.Selection.SetElementIds(element_ids)

