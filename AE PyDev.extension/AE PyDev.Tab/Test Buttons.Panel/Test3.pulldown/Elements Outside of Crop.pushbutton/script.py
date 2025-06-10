# -*- coding: utf-8 -*-
# IRONPYTHON 2.7 COMPATIBLE (no f-strings, no .format usage)
from Autodesk.Revit.DB import FilteredElementCollector, GraphicsStyle, Line, XYZ
from pyrevit import revit, script

logger = script.get_logger()

doc = revit.doc
uidoc = revit.uidoc

# Use selection
selection_ids = uidoc.Selection.GetElementIds()
if not selection_ids:
    logger.error("No elements selected.")
    script.exit()

element = doc.GetElement(list(selection_ids)[0])

# Get the bounding box in the active view
bbox = element.get_BoundingBox(doc.ActiveView)
if not bbox:
    logger.error("Element has no bounding box.")
    script.exit()

min_pt = bbox.Min
max_pt = bbox.Max
centroid = XYZ((min_pt.X + max_pt.X) / 2, (min_pt.Y + max_pt.Y) / 2, (min_pt.Z + max_pt.Z) / 2)

# Create 2D corner points
p1 = XYZ(min_pt.X, min_pt.Y, 0)
p2 = XYZ(max_pt.X, min_pt.Y, 0)
p3 = XYZ(max_pt.X, max_pt.Y, 0)
p4 = XYZ(min_pt.X, max_pt.Y, 0)

# Find the line style "Solid_05 (Blue)"
target_style = None
styles = FilteredElementCollector(doc).OfClass(GraphicsStyle)
for style in styles:
    if style.Name == "Solid_05 (Blue)":
        target_style = style
        break

if not target_style:
    logger.error("Line style 'Solid_05 (Blue)' not found.")
    script.exit()

# Create detail lines in active view
with revit.Transaction("Draw Bounding Box Lines"):
    view = doc.ActiveView
    lines = [
        Line.CreateBound(p1, p2),
        Line.CreateBound(p2, p3),
        Line.CreateBound(p3, p4),
        Line.CreateBound(p4, p1),
        Line.CreateBound(centroid, min_pt)
    ]
    for ln in lines:
        detail_line = doc.Create.NewDetailCurve(view, ln)
        detail_line.LineStyle = target_style

logger.info("Detail lines created around element.")
