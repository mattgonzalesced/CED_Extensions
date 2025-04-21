# -*- coding: utf-8 -*-

from pyrevit import revit, forms, script
from pyrevit import DB as DB
from pyrevit import UI as UI
from System.Collections.Generic import List

output = script.get_output()
output.close_others()

doc = revit.doc
uidoc = UI.UIDocument(doc)

#MAIN
# ==================================================
if __name__ == '__main__':

    all_elements = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    unhide_elements = List[DB.ElementId]()

    for element in all_elements:
        if element.IsHidden(doc.ActiveView):
            unhide_elements.Add(element.Id)

    if unhide_elements.Count == 0:
        forms.alert("No Elements are manually hidden in the active view. Nice! \nScript Exiting.",
                    title="Unhide All Elements",
                    exitscript= True)


    with revit.Transaction("Unhide All Elements in Active View", doc) as t:
        doc.ActiveView.UnhideElements(unhide_elements)


    selection = uidoc.Selection.SetElementIds(unhide_elements)