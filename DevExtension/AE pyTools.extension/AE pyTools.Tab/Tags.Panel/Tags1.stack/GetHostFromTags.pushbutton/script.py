# -*- coding: utf-8 -*-
__title__ = "Get Hosts From Tags"
__doc__ = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton
You can use it to start your pyRevit Add-In

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO: testing 
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""
from pyrevit import revit, DB, HOST_APP, forms
from System.Collections.Generic import List

# Get current selection
selection = revit.get_selection()

# List to store tagged elements
tagged_elements = []

# Process each selected element
for el in selection:
    if HOST_APP.is_newer_than(2022, or_equal=True):
        if isinstance(el, DB.IndependentTag):
            element_ids = el.GetTaggedLocalElementIds()
            if element_ids:
                tagged_elements.append(List[DB.ElementId](element_ids)[0])
        elif isinstance(el, DB.Architecture.RoomTag):
            tagged_elements.append(el.TaggedLocalRoomId)
        elif isinstance(el, DB.Mechanical.SpaceTag):
            tagged_elements.append(el.Space.Id)
        elif isinstance(el, DB.AreaTag):
            tagged_elements.append(el.Area.Id)
    else:
        if isinstance(el, DB.IndependentTag):
            tagged_elements.append(el.TaggedLocalElementId)
        elif isinstance(el, DB.Architecture.RoomTag):
            tagged_elements.append(el.TaggedLocalRoomId)
        elif isinstance(el, DB.Mechanical.SpaceTag):
            tagged_elements.append(el.Space.Id)
        elif isinstance(el, DB.AreaTag):
            tagged_elements.append(el.Area.Id)

# If tagged elements found
if tagged_elements:
    if __shiftclick__:
        # Append hosts to the current selection
        selection.append(tagged_elements)
    else:
        # Replace selection with only the hosts
        selection.set_to(tagged_elements)
else:
    # Notify user if no tags are selected
    forms.alert("Please select at least one tag to get its host.")
