# -*- coding: utf-8 -*-
"""
Change Level by Selection with Relative Offset Preservation and Elevation Update – pyRevit Script

Functionality:
  1. Prompts the user to select one or more elements.
  2. Displays a popup with friendly names for each level (name, id, elevation) so you can pick one.
  3. For each selected element:
      - Determines its current host level using the “Level” or “Reference Level” parameters (or a fallback LevelId).
      - Reads the element’s current “Elevation” parameter (if available) to preserve its vertical offset.
      - Computes the offset as the difference between the element’s location and the old level’s elevation.
      - Updates the element’s host level parameters to the new target level.
      - Moves the element’s location (if possible) to achieve:
             new_z = target_level.Elevation + (old_location_z – old_level.Elevation)
      - If the element has an “Elevation” parameter (commonly used by hosted families), resets it to its original value.
  4. Reports how many elements were updated.

Note: For many hosted families (doors, windows, etc.) the “Elevation” parameter governs the vertical offset.
If your element still isn’t updating as expected, you may need to verify which parameters actually control the hosted offset for that category.
"""

from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    Level,
    LocationPoint,
    LocationCurve,
    XYZ,
    Transform
)
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

# 1. Get user-selected elements.
sel_ids = uidoc.Selection.GetElementIds()
if not sel_ids:
    forms.alert("Please select one or more elements.", exitscript=True)
elements = [doc.GetElement(eid) for eid in sel_ids]

# 2. Retrieve levels and build a friendly popup dictionary.
levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
if not levels:
    forms.alert("No levels found in the project.", exitscript=True)
levels_sorted = sorted(levels, key=lambda lvl: lvl.Elevation)
levels_dict = {}
for lvl in levels_sorted:
    friendly = "{} (ID: {}) | Elev: {:.2f}".format(lvl.Name, lvl.Id.IntegerValue, lvl.Elevation)
    levels_dict[friendly] = lvl

sel_str = forms.SelectFromList.show(
    sorted(levels_dict.keys()),
    title="Choose a Target Level",
    button_name="Select Level",
    multiselect=False
)
if not sel_str:
    forms.alert("No level selected.", exitscript=True)
target_level = levels_dict[sel_str]

# 3. Begin a transaction.
t = Transaction(doc, "Change Level for Selected Elements")
t.Start()
changed_count = 0

for elem in elements:
    updated = False

    # --- Determine the element's current host level ---
    old_level = None
    level_param = elem.LookupParameter("Level")
    if level_param and not level_param.IsReadOnly:
        old_id = level_param.AsElementId()
        if old_id and old_id.IntegerValue >= 0:
            old_level = doc.GetElement(old_id)
    if not old_level:
        ref_param = elem.LookupParameter("Reference Level")
        if ref_param and not ref_param.IsReadOnly:
            old_id = ref_param.AsElementId()
            if old_id and old_id.IntegerValue >= 0:
                old_level = doc.GetElement(old_id)
    # Fallback for FamilyInstances
    if not old_level and hasattr(elem, "LevelId"):
        old_level = doc.GetElement(elem.LevelId)
    
    # If no old level or location can be determined, skip.
    if not old_level or not elem.Location:
        continue

    # --- Read current element location and compute offset ---
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        current_z = loc.Point.Z
    elif isinstance(loc, LocationCurve):
        current_z = loc.Curve.GetEndPoint(0).Z
    else:
        current_z = 0.0

    offset = current_z - old_level.Elevation

    # --- If available, read the "Elevation" parameter value (used by many hosted families) ---
    elev_value = None
    elev_param = elem.LookupParameter("Elevation")
    if elev_param and not elev_param.IsReadOnly:
        elev_value = elev_param.AsDouble()

    # --- Update the element's level parameters ---
    param_changed = False
    if level_param and not level_param.IsReadOnly:
        if level_param.AsElementId() != target_level.Id:
            level_param.Set(target_level.Id)
            param_changed = True
    if elem.LookupParameter("Reference Level"):
        ref_param = elem.LookupParameter("Reference Level")
        if ref_param and not ref_param.IsReadOnly:
            if ref_param.AsElementId() != target_level.Id:
                ref_param.Set(target_level.Id)
                param_changed = True

    # --- Move the element's location to preserve offset ---
    new_z = target_level.Elevation + offset
    moved = False
    if isinstance(loc, LocationPoint):
        pt = loc.Point
        if abs(pt.Z - new_z) > 0.001:
            new_pt = XYZ(pt.X, pt.Y, new_z)
            loc.Point = new_pt
            moved = True
    elif isinstance(loc, LocationCurve):
        curve = loc.Curve
        start_pt = curve.GetEndPoint(0)
        if abs(start_pt.Z - new_z) > 0.001:
            delta_z = new_z - start_pt.Z
            translation = XYZ(0, 0, delta_z)
            try:
                loc.Move(translation)
                moved = True
            except Exception as ex:
                # Fallback: reassign the curve
                tf = Transform.CreateTranslation(XYZ(0, 0, delta_z))
                loc.Curve = curve.CreateTransformed(tf)
                moved = True

    # --- Reset the "Elevation" parameter to preserve internal offset (if applicable) ---
    if elev_value is not None and elev_param and not elev_param.IsReadOnly:
        # Set it back to its original value.
        elev_param.Set(elev_value)

    if param_changed or moved:
        changed_count += 1

t.Commit()
forms.alert("Changed level for {} element(s).".format(changed_count))
