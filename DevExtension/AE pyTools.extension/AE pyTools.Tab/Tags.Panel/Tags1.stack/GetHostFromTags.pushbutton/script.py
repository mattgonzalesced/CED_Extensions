# -*- coding: utf-8 -*-

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
