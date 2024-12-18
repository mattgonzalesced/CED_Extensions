# -*- coding: utf-8 -*-
__doc__ = "Filter views with enabled crop box and find elements outside crop box"
__title__ = "Elements Outside Crop Box"
__author__ = "Your Name"

from pyrevit import DB
from pyrevit import script

output = script.get_output()

# Get the active Revit document
doc = __revit__.ActiveUIDocument.Document

# Collect all views in the document
views_collector = DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType()
output.close_others()

# Filter views with an enabled crop box, excluding specific view types
analyzable_views = [
    view for view in views_collector
    if not view.IsTemplate and
    view.CropBoxActive and
    view.ViewType not in (
        DB.ViewType.DraftingView,
        DB.ViewType.Section,
        DB.ViewType.ThreeD,
        DB.ViewType.Elevation
    )
]

output.print_md("## Views with Enabled Crop Box")
outside_elements_global = []

if not analyzable_views:
    output.print_md("### No views found with an active crop box.")
else:
    for view in analyzable_views:
        view_name = view.Name

        # Get crop box and its dimensions
        crop_box = view.CropBox
        min_point = crop_box.Min
        max_point = crop_box.Max

        # Collect all elements owned by the view
        collector = DB.FilteredElementCollector(doc, view.Id) \
            .WhereElementIsNotElementType() \
            .OwnedByView(view.Id)

        # Count elements outside the crop box
        outside_count = 0
        for elem in collector:
            location = elem.Location
            if isinstance(location, DB.LocationPoint):
                point = location.Point
                # Check if point is outside the crop box
                if not (min_point.X <= point.X <= max_point.X and
                        min_point.Y <= point.Y <= max_point.Y and
                        min_point.Z <= point.Z <= max_point.Z):
                    outside_count += 1
                    outside_elements_global.append(str(elem.Id.IntegerValue))

        # Output results for this view only if there are elements outside the crop box
        if outside_count > 0:
            output.print_md("{}: {} elements outside crop box".format(view_name, outside_count))

# Print all element IDs outside crop boxes, separated by commas
if outside_elements_global:
    output.print_md("### All Elements Outside Crop Boxes:")
    output.print_md(", ".join(outside_elements_global))
else:
    output.print_md("### No elements found outside any crop boxes.")
