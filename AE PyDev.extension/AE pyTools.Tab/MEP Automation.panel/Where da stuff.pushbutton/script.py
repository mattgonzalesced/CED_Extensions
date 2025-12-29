from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

linked_elem_id_int = 123456  # must be ElementId from LINKED model

# ---- validation ----
if not isinstance(linked_elem_id_int, int) or linked_elem_id_int <= 0:
    forms.alert("Invalid ElementId provided.", exitscript=True)

target_id = ElementId(linked_elem_id_int)
found_refs = []

links = FilteredElementCollector(doc).OfClass(RevitLinkInstance)

for link in links:
    try:
        link_doc = link.GetLinkDocument()
        if not link_doc:
            continue

        linked_elem = link_doc.GetElement(target_id)
        if not linked_elem:
            continue

        ref = Reference(linked_elem).CreateLinkReference(link)
        found_refs.append(ref)

    except Exception:
        continue

# ---- result handling ----
if not found_refs:
    forms.alert(
        "ElementId {} not found in any loaded link."
        .format(linked_elem_id_int),
        exitscript=True
    )

uidoc.Selection.SetReferences(found_refs)
