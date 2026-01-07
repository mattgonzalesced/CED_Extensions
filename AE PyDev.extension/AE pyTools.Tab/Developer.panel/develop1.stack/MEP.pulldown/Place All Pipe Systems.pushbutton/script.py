# -*- coding: utf-8 -*-


__title__ = "Draw Unique Piping System Segments"
__author__ = "Your Name"

from Autodesk.Revit.DB import (BuiltInCategory, XYZ, FilteredElementCollector, Level)
from Autodesk.Revit.DB.Plumbing import Pipe, PipeType, PipingSystemType
from pyrevit import revit, forms, script
from pyrevit.revit import query

doc = revit.doc

# Retrieve the default pipe type using FilteredElementCollector.
pipe_types = FilteredElementCollector(doc).OfClass(PipeType).ToElements()
default_pipe_type = None
for current_pipe_type in pipe_types:
    current_type_name = query.get_name(current_pipe_type, title_on_sheet=False)
    if "default" in current_type_name.lower():
        default_pipe_type = current_pipe_type
        break
if not default_pipe_type and pipe_types:
    default_pipe_type = pipe_types[0]
if not default_pipe_type:
    forms.alert("No Pipe Type found in the project.", exitscript=True)

# Retrieve the active level.
active_level = None
if hasattr(revit.active_view, "GenLevel"):
    active_level = revit.active_view.GenLevel
else:
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    active_level = levels[0] if levels else None
if not active_level:
    forms.alert("No Level found in the project.", exitscript=True)

# Retrieve all piping system instances using FilteredElementCollector.
all_piping_system_instances = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()
if not all_piping_system_instances:
    forms.alert("No piping systems found in the project.", exitscript=True)

# Build a dictionary of unique piping system names (using query.get_name).
unique_systems = {}
for piping_system_instance in all_piping_system_instances:
    system_name = query.get_name(piping_system_instance, title_on_sheet=False)
    if system_name not in unique_systems:
        unique_systems[system_name] = piping_system_instance

# Retrieve all piping system types.
all_system_types = FilteredElementCollector(doc).OfClass(PipingSystemType).ToElements()
if not all_system_types:
    forms.alert("No Piping System Types found in the project.", exitscript=True)

# Define dimensions in feet.
pipe_segment_length = 10.0  # Pipe segment length in feet.
vertical_spacing = 5.0  # Y-offset between segments in feet.

# List to store the names of systems for which segments were created.
placed_system_names = []

# Create a context-managed transaction.
with revit.Transaction("Create Unique Pipe Segments"):
    for index, (system_name, system_instance) in enumerate(unique_systems.items()):
        # Attempt to find a piping system type that matches the system name.
        matching_system_type = None
        for system_type in all_system_types:
            system_type_name = query.get_name(system_type, title_on_sheet=False)
            if system_type_name.lower() == system_name.lower():
                matching_system_type = system_type
                break
        if not matching_system_type:
            matching_system_type = all_system_types[0]
        current_y_offset = index * vertical_spacing
        start_point = XYZ(0, current_y_offset, 0)
        end_point = XYZ(pipe_segment_length, current_y_offset, 0)
        Pipe.Create(doc, matching_system_type.Id, default_pipe_type.Id, active_level.Id, start_point, end_point)
        placed_system_names.append(system_name)

# Print the names of the piping systems where pipe segments were placed.
output = script.get_output()
output.print_md("Pipe segments were placed for the following unique piping systems:")
for name in placed_system_names:
    output.print_md(" - " + name)

forms.alert("Created {} pipe segment(s) for unique piping systems spaced 5 ft apart."
            .format(len(unique_systems)))
