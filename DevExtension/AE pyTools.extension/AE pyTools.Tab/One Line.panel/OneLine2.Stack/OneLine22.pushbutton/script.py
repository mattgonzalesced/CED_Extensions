import clr
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilyInstance, ElementId,
    Connector, BuiltInParameter
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem
import pyrevit.revit
from pyrevit import script, revit

# Utility function to get circuit number
def get_circuit_number(system):
    """Retrieve the circuit number from an ElectricalSystem."""
    try:
        return system.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
    except Exception:
        return None

class ElectricalNode:
    """Class to represent a node in the electrical hierarchy."""
    def __init__(self, element, level=0, circuit_number=None, is_leaf=False, equipment_name=None):
        self.element = element
        self.children = []
        self.level = level
        self.circuit_number = circuit_number
        self.is_leaf = is_leaf
        self.equipment_name = equipment_name

    def add_child(self, child):
        self.children.append(child)

    def get_name(self):
        """Generate name for printing based on type."""
        if isinstance(self.element, ElectricalSystem):
            if self.is_leaf:
                return "{} LOAD".format(self.circuit_number) if self.circuit_number else "LOAD"
            return "{}".format(self.circuit_number)
        else:
            return "{}".format(self.equipment_name)

    def print_tree(self):
        """Print the tree with proper indentation based on level."""
        indent = '  ' * self.level
        print(indent + self.get_name())
        for child in self.children:
            child.print_tree()

class ElectricalSystemTree:
    """Class to build and manage the electrical system hierarchy."""
    def __init__(self, doc):
        self.doc = doc
        self.visited = set()

    def build_tree(self, root_element):
        """Initialize the tree from the root element."""
        root_node = ElectricalNode(root_element, level=0, equipment_name=root_element.Name)
        self.traverse_node(root_element, root_node, 0)
        root_node.print_tree()

    def traverse_node(self, element, parent_node, level):
        """Traverse nodes and their connected electrical systems."""
        if element.Id in self.visited:
            return
        self.visited.add(element.Id)

        # Fetch all electrical systems connected to this equipment
        all_systems = element.MEPModel.GetElectricalSystems() if hasattr(element, 'MEPModel') else []
        assigned_systems = element.MEPModel.GetAssignedElectricalSystems() if hasattr(element, 'MEPModel') else []

        upstream_systems = [sys for sys in all_systems if sys not in assigned_systems]
        downstream_systems = assigned_systems

        # Handle upstream systems by moving up one level
        for system in upstream_systems:
            self.process_upstream_system(system, parent_node, level - 1)

        # Process downstream systems by moving down one level
        for system in downstream_systems:
            self.process_downstream_system(system, parent_node, level + 1)

    def process_upstream_system(self, system, parent_node, level):
        """Process upstream systems and their base equipment."""
        if system.BaseEquipment and system.BaseEquipment.Id not in self.visited:
            circuit_number = get_circuit_number(system)
            equipment_name = system.BaseEquipment.Name

            # Create a node for the upstream equipment
            upstream_node = ElectricalNode(system.BaseEquipment, level, equipment_name=equipment_name)
            parent_node.add_child(upstream_node)

            # Traverse upstream equipment
            self.traverse_node(system.BaseEquipment, upstream_node, level)

    def process_downstream_system(self, system, parent_node, level):
        """Process downstream systems."""
        circuit_number = get_circuit_number(system)
        connectors = system.ConnectorManager.Connectors if system else None
        is_leaf = not any(connector.AllRefs for connector in connectors) if connectors else True

        # Create node for the system
        child_node = ElectricalNode(system, level, circuit_number=circuit_number, is_leaf=is_leaf)
        parent_node.add_child(child_node)

        # If it's not a leaf, continue traversing downstream
        if not is_leaf:
            self.traverse_connectors(system, child_node, level + 1)

    def traverse_connectors(self, system, parent_node, level):
        """Traverse connectors of an ElectricalSystem to find downstream nodes."""
        connectors = system.ConnectorManager.Connectors if system else None
        if not connectors:
            parent_node.is_leaf = True
            return

        for connector in connectors:
            for ref_connector in connector.AllRefs:
                connected_elem = ref_connector.Owner
                if connected_elem and connected_elem.Id not in self.visited:
                    if isinstance(connected_elem, FamilyInstance):
                        self.traverse_node(connected_elem, parent_node, level)

# Main execution
if __name__ == "__main__":
    output = script.get_output()
    doc = revit.doc

    # Prompt the user to select a root element
    selected_element = pyrevit.revit.pick_element("Select the root element for the electrical hierarchy")
    if not selected_element:
        print("No element selected. Exiting...")
    else:
        # Build and print the electrical system tree
        tree = ElectricalSystemTree(doc)
        tree.build_tree(selected_element)
