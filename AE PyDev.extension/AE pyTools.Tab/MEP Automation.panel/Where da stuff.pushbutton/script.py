from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# ---- prompt user ----
user_input = forms.ask_for_string(
    default="",
    prompt="Enter ElementId from the LINKED model:",
    title="Select Linked Element by ID"
)

if not user_input:
    forms.alert("No ElementId provided.", exitscript=True)

try:
    linked_elem_id_int = int(user_input)
    if linked_elem_id_int <= 0:
        raise ValueError
except ValueError:
    forms.alert("ElementId must be a positive integer.", exitscript=True)

target_id = ElementId(linked_elem_id_int)

found_refs = []
link_instance_ids = set()
found_link_elems = []

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
        link_instance_ids.add(link.Id)
        found_link_elems.append((link, linked_elem))

    except Exception:
        continue

if not found_refs:
    forms.alert(
        "ElementId {} was not found in any loaded link."
        .format(linked_elem_id_int),
        exitscript=True
    )

# ---- select + zoom ----
uidoc.Selection.SetReferences(found_refs)

def _get_active_ui_view(uidoc):
    active_view_id = uidoc.ActiveView.Id
    for ui_view in uidoc.GetOpenUIViews():
        if ui_view.ViewId == active_view_id:
            return ui_view
    return None

def _transform_bbox_to_host(link, bbox):
    t = link.GetTotalTransform()
    # Transform all 8 corners to handle rotated links.
    corners = [
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
    ]
    transformed = [t.OfPoint(pt) for pt in corners]
    min_x = min(pt.X for pt in transformed)
    min_y = min(pt.Y for pt in transformed)
    min_z = min(pt.Z for pt in transformed)
    max_x = max(pt.X for pt in transformed)
    max_y = max(pt.Y for pt in transformed)
    max_z = max(pt.Z for pt in transformed)
    return XYZ(min_x, min_y, min_z), XYZ(max_x, max_y, max_z)

ui_view = _get_active_ui_view(uidoc)
overall_min = None
overall_max = None

for link, linked_elem in found_link_elems:
    bbox = linked_elem.get_BoundingBox(None)
    if not bbox:
        continue
    host_min, host_max = _transform_bbox_to_host(link, bbox)
    if overall_min is None:
        overall_min = host_min
        overall_max = host_max
    else:
        overall_min = XYZ(
            min(overall_min.X, host_min.X),
            min(overall_min.Y, host_min.Y),
            min(overall_min.Z, host_min.Z),
        )
        overall_max = XYZ(
            max(overall_max.X, host_max.X),
            max(overall_max.Y, host_max.Y),
            max(overall_max.Z, host_max.Z),
        )

if ui_view and overall_min and overall_max:
    ui_view.ZoomAndCenterRectangle(overall_min, overall_max)
else:
    # Fallback: zoom to link instances when bbox is unavailable.
    uidoc.ShowElements(List[ElementId](link_instance_ids))
