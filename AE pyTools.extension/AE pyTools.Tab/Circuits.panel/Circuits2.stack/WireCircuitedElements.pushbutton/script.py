# -*- coding: utf-8 -*-
__title__ = "Wire Circuited Elements"

from pyrevit import script, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    XYZ,
    ElementId,
    Transaction,
    BuiltInCategory,
    Electrical
)

# Get Revit document and selection
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
home_run_length = 4


# Initialize config to retrieve default wire type
config = script.get_config("wire_type_config")

#TODO need to adjust so that if the wire_type_config does not have an existing wire, then the config.py script runs, avoiding errors.

# Retrieve all wire types in the document
wire_types = FilteredElementCollector(doc).OfClass(Electrical.WireType).ToElements()
wire_type_options = {wire_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(): wire_type.Id
                     for wire_type in wire_types}

# Use default wire type from config, if available
default_wire_type_name = config.get_option("default_wire_type", None)
if default_wire_type_name and default_wire_type_name in wire_type_options:
    wire_type_id = wire_type_options[default_wire_type_name]
    script.get_logger().info("Using saved default wire type: " + default_wire_type_name)
else:
    script.get_logger().warning("No default wire type set or available.")
    script.exit()


# Helper function to get the first connector of a power circuit from an element
def get_first_connector_of_power_circuit(element):
    connector_manager = element.MEPModel.ConnectorManager
    connector_set = connector_manager.Connectors
    connector_iterator = connector_set.ForwardIterator()

    while connector_iterator.MoveNext():
        connector = connector_iterator.Current
        if connector.ElectricalSystemType == Electrical.ElectricalSystemType.PowerCircuit:
            return connector

    return None


# Nearest neighbor connection calculation based on element distances
def find_nearest_neighbor(element, remaining_elements):
    element_connector = get_first_connector_of_power_circuit(element)
    element_location = element_connector.Origin if element_connector else None
    nearest_element = None
    shortest_distance = float('inf')

    if element_location is not None:
        for other_element in remaining_elements:
            other_connector = get_first_connector_of_power_circuit(other_element)
            other_location = other_connector.Origin if other_connector else None

            if other_location:
                # Calculate the distance between the element and other element's connectors
                distance = element_location.DistanceTo(other_location)
                if distance < shortest_distance:
                    shortest_distance = distance
                    nearest_element = other_element

    return nearest_element


# Function to check for coincident points (same X and Y)
def are_points_coincident(point1, point2):
    return abs(point1.X - point2.X) < 0.001 and abs(point1.Y - point2.Y) < 0.001


# Function to create right-angled "chamfered" wires for connections
def create_right_angle_wire(start_point, end_point):
    # Determine the right-angle route (move along X first, then Y or vice versa)
    if abs(start_point.X - end_point.X) > abs(start_point.Y - end_point.Y):
        # Move along X-axis first, then Y
        vertex = XYZ(end_point.X, start_point.Y, start_point.Z)
    else:
        # Move along Y-axis first, then X
        vertex = XYZ(start_point.X, end_point.Y, start_point.Z)

    # Check for coincident points before returning
    if are_points_coincident(start_point, vertex):
        return [start_point, end_point]  # Skip vertex if coincident
    elif are_points_coincident(vertex, end_point):
        return [start_point, end_point]  # Skip vertex if coincident
    else:
        return [start_point, vertex, end_point]


# Function to create a right-angled home run of fixed length (3ft)
def create_right_angle_home_run(start_connector, last_connector=None, length=4):
    # Get the initial start point and determine the correct direction from the last connected wire
    start_point = start_connector.Origin

    if last_connector:
        last_point = last_connector.Origin
        # Calculate the direction vector of the last connected wire (last segment)
        direction_vector = XYZ(last_point.X - start_point.X, last_point.Y - start_point.Y, 0).Normalize()
    else:
        # If no previous connector, default to an X direction
        direction_vector = XYZ(0, 1, 0)  # Default to X direction if no last connector

    # Find the perpendicular vector (rotate 90 degrees around the Z-axis)
    if abs(direction_vector.X) > abs(direction_vector.Y):
        # If the last connecting wire runs along the X direction
        perp_vector = XYZ(0, 1, 0)  # Perpendicular in the Y direction
    else:
        # If the last connecting wire runs along the Y direction
        perp_vector = XYZ(1, 0, 0)  # Perpendicular in the X direction

    # Define the start -> vertex -> end right-angle points for the home run
    vertex_point = XYZ(start_point.X + perp_vector.X * (length / 2), start_point.Y + perp_vector.Y * (length / 2),
                       start_point.Z)
    end_point = XYZ(vertex_point.X + direction_vector.X * (length / 2),
                    vertex_point.Y + direction_vector.Y * (length / 2), vertex_point.Z)

    # Ensure that none of the points are coincident
    if are_points_coincident(start_point, vertex_point) or are_points_coincident(vertex_point, end_point):
        vertex_point = XYZ(start_point.X + perp_vector.X * 1.0, start_point.Y + perp_vector.Y * 1.0, start_point.Z)
        end_point = XYZ(vertex_point.X + direction_vector.X * 1.5, vertex_point.Y + direction_vector.Y * 1.5,
                        vertex_point.Z)

    return [start_point, vertex_point, end_point]




# Dictionary to store elements by circuit
elements_by_circuit = {}

# Iterate through selected elements
for sel_id in selected_ids:
    element = doc.GetElement(sel_id)

    # Get electrical systems (circuits) for the element
    if hasattr(element,'MEPModel') and element.MEPModel:
        electrical_systems = element.MEPModel.GetElectricalSystems()
        # Group elements by circuit
        for system in electrical_systems:
            if system.SystemType == Electrical.ElectricalSystemType.PowerCircuit:
                circuit_id = system.Id
                if circuit_id not in elements_by_circuit:
                    elements_by_circuit[circuit_id] = []
                elements_by_circuit[circuit_id].append(element)

# List to store wiring operations to be created in the transaction
wiring_operations = []


# Prepare wiring operations before the transaction
for circuit_id, elements in elements_by_circuit.items():
    remaining_elements = elements[:]  # Copy the list using slicing for IronPython

    # For each element, find the nearest neighbor to connect to
    while remaining_elements:
        element_start = remaining_elements.pop(0)

        # If no more elements, we will add the home run for the last one
        if not remaining_elements:
            connector = get_first_connector_of_power_circuit(element_start)

            if connector:
                # Find the last connector for the home run
                if wiring_operations:
                    last_connector = wiring_operations[-1][1]  # Get the last connector used
                    points = create_right_angle_home_run(connector,
                                                         last_connector,
                                                         home_run_length)

                    # Use WiringType.Arc for home run
                    wiring_operations.append((points, connector, None, DB.Electrical.WiringType.Arc))
                else:
                    # If there's only one element, create a default right-angled home run
                    start_point = connector.Origin
                    points = create_right_angle_home_run(connector)  # No previous connector, so just use one
                    wiring_operations.append((points, connector, None, DB.Electrical.WiringType.Arc))
        else:
            # Find the nearest neighbor element
            nearest_element = find_nearest_neighbor(element_start, remaining_elements)

            # Get connectors for both start and end elements
            start_connector = get_first_connector_of_power_circuit(element_start)
            end_connector = get_first_connector_of_power_circuit(nearest_element)

            # Ensure both connectors are valid PowerCircuit connectors
            if start_connector and end_connector:
                # Create a right-angle chamfered wire between elements
                points = create_right_angle_wire(start_connector.Origin, end_connector.Origin)

                # Use WiringType.Chamfer for connecting wires
                wiring_operations.append((points, start_connector, end_connector, DB.Electrical.WiringType.Chamfer))

# Start transaction and execute all wiring operations
t = Transaction(doc, "Create wires between circuit elements")
try:
    t.Start()

    for points, start_connector, end_connector, wiring_type in wiring_operations:
        # Create wires using the provided wire_type_id and the calculated points
        DB.Electrical.Wire.Create(doc, wire_type_id, doc.ActiveView.Id, wiring_type, points, start_connector,
                                  end_connector)

    # Commit the transaction if everything is successful
    t.Commit()

except Exception as e:
    # If there is an error, roll back the transaction to avoid leaving it open
    script.get_logger().error("An error occurred: {}".format(str(e)))
    t.RollBack()

# Notify the user
script.get_logger().info("Wires and home runs created successfully.")
