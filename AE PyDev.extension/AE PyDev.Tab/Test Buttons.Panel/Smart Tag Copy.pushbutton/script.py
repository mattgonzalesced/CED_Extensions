# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms, script
from pyrevit.revit import Transaction
import math

output = script.get_output()
logger = script.get_logger()

# ==============================
# Functions
# ==============================

def get_space_boundary_points(space):
    options = DB.SpatialElementBoundaryOptions()
    options.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
    segments = space.GetBoundarySegments(options)

    if not segments:
        return None

    points = []
    for seg_list in segments:
        for seg in seg_list:
            curve = seg.GetCurve()
            points.append(curve.GetEndPoint(0))
            points.append(curve.GetEndPoint(1))

    return points

def parse_manual_layout_input(input_str):
    try:
        parts = input_str.lower().replace(' ', '').split('x')
        cols = int(parts[0])
        rows = int(parts[1])
        return cols, rows
    except Exception as e:
        logger.warning("Failed to parse manual input: {}".format(e))
        return None, None

def auto_solve_layout(fixtures, tolerance=0.5):
    x_positions = [fixture.Location.Point.X for fixture in fixtures]
    y_positions = [fixture.Location.Point.Y for fixture in fixtures]

    x_clusters = cluster_positions(x_positions, tolerance)
    y_clusters = cluster_positions(y_positions, tolerance)

    cols = len(x_clusters)
    rows = len(y_clusters)

    return cols, rows

def cluster_positions(positions, tolerance):
    positions = sorted(positions)
    clusters = []
    cluster = [positions[0]]

    for pos in positions[1:]:
        if abs(pos - cluster[-1]) <= tolerance:
            cluster.append(pos)
        else:
            clusters.append(cluster)
            cluster = [pos]
    clusters.append(cluster)

    return clusters

def calculate_fixture_positions(min_x, min_y, width, length, cols, rows):
    # Correct D calculation so edge spacing = D, fixture to fixture = 2D
    edge_spacing_x = width / (2.0 * cols)
    edge_spacing_y = length / (2.0 * rows)

    positions = []
    for row in range(rows):
        for col in range(cols):
            x = min_x + edge_spacing_x * (1 + 2 * col)
            y = min_y + edge_spacing_y * (1 + 2 * row)
            positions.append((x, y))
    return positions

def move_fixtures_to_positions(fixtures, positions):
    for fixture, (x, y) in zip(fixtures, positions):
        loc = fixture.Location
        if hasattr(loc, 'Point'):
            old_z = loc.Point.Z
            new_point = DB.XYZ(x, y, old_z)
            loc.Point = new_point

# ==============================
# Main
# ==============================

selection = revit.get_selection()

spaces = [el for el in selection if isinstance(el, DB.SpatialElement)]
fixtures = [el for el in selection if isinstance(el, DB.FamilyInstance)]

if not spaces or not fixtures:
    forms.alert("Please select at least 1 Space and 1 or more Lighting Fixtures.", exitscript=True)

space = spaces[0]  # Assume only one space for now

# Ask user for layout method
# layout_method = forms.CommandSwitchWindow.show(
#     ["Manual Entry", "Auto Solve"],
#     message="Select layout method."
# )
#
# if not layout_method:
#     script.exit()
layout_method = "Auto Solve"
points = get_space_boundary_points(space)

if not points:
    forms.alert("No boundary points found for the selected Space.", exitscript=True)

x_vals = [pt.X for pt in points]
y_vals = [pt.Y for pt in points]

min_x = min(x_vals)
max_x = max(x_vals)
min_y = min(y_vals)
max_y = max(y_vals)

width = max_x - min_x
length = max_y - min_y

if layout_method == "Manual Entry":
    layout_string = forms.ask_for_string(default="2x3", prompt="Enter layout as ColumnsxRows (e.g., 2x3):")
    if not layout_string:
        script.exit()
    cols, rows = parse_manual_layout_input(layout_string)
    if not cols or not rows:
        forms.alert("Invalid layout input.", exitscript=True)

elif layout_method == "Auto Solve":
    cols, rows = auto_solve_layout(fixtures)

# Check fixture count vs available slots
if len(fixtures) != cols * rows:
    forms.alert("Number of fixtures ({}) does not match layout slots ({}x{}={}).".format(len(fixtures), cols, rows, cols*rows), exitscript=True)

positions = calculate_fixture_positions(min_x, min_y, width, length, cols, rows)

with Transaction("Place Fixtures Evenly", revit.doc):
    move_fixtures_to_positions(fixtures, positions)

forms.alert("Fixtures placed successfully!")

