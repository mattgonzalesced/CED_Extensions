# -*- coding: utf-8 -*-

import clr

clr.AddReference('System')

from Autodesk.Revit.DB import IndependentTag, FilteredElementCollector, ElementId, BuiltInCategory, Transaction
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, DB, UI, HOST_APP, script, output, forms
from System.Collections.Generic import List

# Get the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def get_tags_from_host(selection):
    selection_ids = selection.element_ids
    tag_collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(IndependentTag)
    tag_iterator = tag_collector.GetElementIterator()
    tags = []
    while tag_iterator.MoveNext():
        tag = tag_iterator.Current
        tag_host_ids = tag.GetTaggedLocalElementIds()
        if any(sel_id in tag_host_ids for sel_id in selection_ids):
            tags.append(tag.Id)

    return tags


def main():
    # Get the current selection of element IDs from the user
    selection = revit.get_selection()

    model_elements = []
    hosted_tags = []
    # Check if there are selected elements
    if not selection:
        TaskDialog.Show("Error", "No elements selected. Please select elements.")
        script.exit()

    tags = get_tags_from_host(selection)

    if tags:
        if __shiftclick__:
            selection.append(tags)

        else:
            selection.set_to(tags)


if __name__ == "__main__":
    main()


