# -*- coding: utf-8 -*-
"""One-line system tree helpers.

Extracted from the OneLine22 testing script so it can be reused for
list boxes and branch calculations.
"""

from pyrevit import DB
from pyrevit.compat import get_elementid_value_func

get_id_value = get_elementid_value_func()


class TreeBranch(object):
    def __init__(self, system):
        self.element_id = system.Id
        self.base_eq = system.BaseEquipment
        self.base_eq_id = self.base_eq.Id if self.base_eq else DB.ElementId.InvalidElementId
        self.base_eq_name = DB.Element.Name.__get__(self.base_eq) if self.base_eq else "Unknown"
        self.circuit_number = system.CircuitNumber
        self.load_name = system.LoadName
        self.system = system
        self.is_feeder = False  # Set True if it connects two equipment nodes

    def __str__(self):
        return "- Circuit `{}` | Load `{}` | Feeder `{}`".format(
            self.circuit_number, self.load_name, self.is_feeder
        )

    def to_dict(self, parent_node=None):
        return {
            "Parent Panel": parent_node.panel_name if parent_node else "",
            "Parent ID": parent_node.element_id if parent_node else "",
            "Circuit Number": self.circuit_number,
            "Load Name": self.load_name,
            "Branch ID": self.element_id,
            "From Panel": self.base_eq_name,
            "Feeder": self.is_feeder
        }


class TreeNode(object):
    PART_TYPE_MAP = {
        14: "Panelboard",
        15: "Transformer",
        16: "Switchboard",
        17: "Other Panel",
        18: "Equipment Switch"
    }

    def __init__(self, element):
        self.element = element
        self.element_id = element.Id
        self.panel_name = DB.Element.Name.__get__(element)
        self.upstream = []
        self.downstream = []
        self.is_leaf = False
        self._part_type = self.get_family_part_type()
        self.equipment_type = self.PART_TYPE_MAP.get(self._part_type, "Unknown")

    def to_dict(self):
        return {
            "Panel Name": self.panel_name,
            "Element ID": self.element_id,
            "Equipment Type": self.equipment_type,
            "Is Leaf": self.is_leaf,
            "Upstream Count": len(self.upstream),
            "Downstream Count": len(self.downstream)
        }

    def get_family_part_type(self):
        if not self.element or not isinstance(self.element, DB.FamilyInstance):
            return None

        symbol = self.element.Symbol
        if not symbol:
            return None

        family = symbol.Family
        if not family:
            return None

        param = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
        if param and param.HasValue:
            return param.AsInteger()
        return None

    def collect_branches(self):
        mep = self.element.MEPModel
        if not mep:
            return

        all_systems = mep.GetElectricalSystems()
        assigned = mep.GetAssignedElectricalSystems()
        assigned_ids = set([sys.Id for sys in assigned]) if assigned else set()

        for sys in all_systems:
            br = TreeBranch(sys)
            if br.base_eq_id == self.element_id:
                self.downstream.append(br)
            else:
                self.upstream.append(br)

        # Leaf check
        self.is_leaf = not assigned or len(assigned) == 0


class SystemTree(object):
    def __init__(self):
        self.nodes = {}      # {element_id: TreeNode}
        self.root_nodes = [] # list of TreeNode

    def add_node(self, node):
        self.nodes[node.element_id] = node
        if not node.upstream:
            self.root_nodes.append(node)

    def get_node(self, element_id):
        return self.nodes.get(element_id)

    def to_list(self):
        data = []

        for node in self.nodes.values():
            node_record = node.to_dict()
            for branch in node.downstream:
                branch_record = branch.to_dict(parent_node=node)
                combined = dict(node_record)
                combined.update(branch_record)
                data.append(combined)

        return data

    def walk_tree(self, node, visitor, visited=None, level=0):
        if visited is None:
            visited = set()

        visitor(node, level)
        visited.add(node.element_id)

        for branch in node.downstream:
            visitor(branch, level + 1)

            system = branch.system
            if hasattr(system, "Elements"):
                for elem in list(system.Elements):
                    if elem.Category and int(get_id_value(elem.Category.Id)) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                        if elem.Id not in visited:
                            branch.is_feeder = True
                            child_node = self.nodes.get(elem.Id)
                            if child_node:
                                self.walk_tree(child_node, visitor, visited, level + 2)


def get_all_equipment_nodes(doc):
    nodes = {}
    collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType()

    for equip in collector:
        node = TreeNode(equip)
        node.collect_branches()
        nodes[equip.Id] = node
    return nodes


def build_system_tree(doc):
    equipment_map = get_all_equipment_nodes(doc)
    tree = SystemTree()
    for node in equipment_map.values():
        tree.add_node(node)
    return tree
