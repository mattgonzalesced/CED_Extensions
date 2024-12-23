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

import clr

clr.AddReference('System')

from Autodesk.Revit.DB import IndependentTag, FilteredElementCollector, ElementId, BuiltInCategory, Transaction
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, DB, UI, HOST_APP, script, output, forms
from System.Collections.Generic import List

# Get the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument




def get_host_from_tags(selection):
    tagged_elements = []
    for el in selection:
        if HOST_APP.is_newer_than(2022, or_equal=True):
            if isinstance(el, DB.IndependentTag):
                if el.GetTaggedLocalElementIds() is not None:
                    tagged_elements.append(List[DB.ElementId](el.GetTaggedLocalElementIds())[0])
            elif isinstance(el, DB.Architecture.RoomTag):
                tagged_elements.append(el.TaggedLocalRoomId)
            elif isinstance(el, DB.Mechanical.SpaceTag):
                tagged_elements.append(el.Space.Id)
            elif isinstance(el, DB.AreaTag):
                tagged_elements.append(el.Area.Id)
        else:

            if isinstance(el, DB.IndependentTag):
                if el.GetTaggedLocalElementIds() is not None:
                    tagged_elements.append(List[DB.ElementId](el.GetTaggedLocalElementIds())[0])
            elif isinstance(el, DB.Architecture.RoomTag):
                tagged_elements.append(el.TaggedLocalRoomId)
            elif isinstance(el, DB.Mechanical.SpaceTag):
                tagged_elements.append(el.Space.Id)
            elif isinstance(el, DB.AreaTag):
                tagged_elements.append(el.Area.Id)

    return tagged_elements



def main():
    # Get the current selection of element IDs from the user
    selection = revit.get_selection()
    tag_hosts = get_host_from_tags(selection)

    # Check if there are selected elements
    if not selection:
        TaskDialog.Show("Error", "No elements selected. Please select elements.")
        script.exit()

    if not len(tag_hosts) > 0:
        script.exit()
    else:
        if __shiftclick__:
            selection.append(tag_hosts)
        else:
            selection.set_to(tag_hosts)


