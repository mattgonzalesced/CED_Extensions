# -*- coding: utf-8 -*-
__title__ = "Select Homeruns"

__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

________________________________________________________________
How-To:

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description

________________________________________________________________
Author: AEvelina """
import clr
clr.AddReference('System')

from pyrevit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Electrical
from System.Collections.Generic import List

# Get the active document and view
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Step 1: Collect all wires in the active view
wire_collector = FilteredElementCollector(doc, active_view.Id) \
    .OfClass(DB.Electrical.Wire) \
    .ToElements()

# Step 2: Initialize an empty list to store the home run wires
home_run_wires = []

# Step 3: Iterate through each wire, check its connectors
for wire in wire_collector:
    connector_manager = wire.ConnectorManager
    connector_set = connector_manager.Connectors

    # Check for connectors where one is connected and one is not connected
    connected_count = 0
    not_connected_count = 0

    for connector in connector_set:
        if connector.IsConnected:
            connected_count += 1
        else:
            not_connected_count += 1

    # Step 4: If the wire has 1 connected connector and 1 not connected, it's a home run
    if connected_count == 1 and not_connected_count == 1:
        home_run_wires.append(wire.Id)  # Add wire's ElementId to the selection list

# Step 5: Set the active selection to the home run wires
if home_run_wires:
    element_ids = List[DB.ElementId](home_run_wires)
    uidoc.Selection.SetElementIds(element_ids)

else:
    print("No home run wires found in the active view.")
