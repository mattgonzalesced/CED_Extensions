# -*- coding: utf-8 -*-

from pyrevit import revit, DB, script, forms

doc = revit.doc
uidoc = revit.uidoc

output = script.get_output()


# --- Step 1: Collect Spaces ---
selection_ids = uidoc.Selection.GetElementIds()

if not selection_ids:
    forms.alert("Please select at least one Space.", exitscript=True)

spaces = []
for elid in selection_ids:
    el = doc.GetElement(elid)
    if isinstance(el, DB.SpatialElement):
        spaces.append(el)
    else:
        output.print_md("⚠️ Skipped element {} not a Space.".format(el.Id))

if not spaces:
    forms.alert("No Spaces selected.", exitscript=True)

# --- Step 2: Set Boundary Options ---
boundary_options = DB.SpatialElementBoundaryOptions()
boundary_options.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Center  # <- choose here

# --- Step 3: Extract Boundary Curves ---
space_boundaries = {}

for space in spaces:
    boundaries = space.GetBoundarySegments(boundary_options)

    if not boundaries:
        output.print_md("⚠️ No boundaries found for space {}.".format(space.Id))
        continue

    curve_list = []

    for loop in boundaries:  # Each loop is a list of BoundarySegment
        for segment in loop:
            curve = segment.GetCurve()
            curve_list.append(curve)

    space_boundaries[space.Id] = curve_list

# --- Step 2: Validate Active View ---
active_view = doc.ActiveView

if not isinstance(active_view, DB.ViewPlan):
    forms.alert("Active view must be a ViewPlan to create Area Boundary Lines.", exitscript=True)

# --- Step 3: Create SketchPlane ---


# --- Step 4: Start Transaction ---
t = DB.Transaction(doc, "Create Area Boundary Lines")
t.Start()
sketch_plane = DB.SketchPlane.Create(doc, active_view.GenLevel.Id)
created_lines = []

for curve in curve_list:
    try:
        new_line = doc.Create.NewAreaBoundaryLine(sketch_plane, curve, active_view)
        created_lines.append(new_line)
    except Exception as e:
        output.print_md("❌ Failed to create line from curve: {}".format(e))

t.Commit()

# --- Step 5: Done ---
forms.alert("Created {} Area Boundary Lines.".format(len(created_lines)))
