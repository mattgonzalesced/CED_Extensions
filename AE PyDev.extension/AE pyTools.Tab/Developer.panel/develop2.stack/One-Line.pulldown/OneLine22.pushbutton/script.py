# -*- coding: utf-8 -*-
from pyrevit import script, revit, DB, forms

logger = script.get_logger()
import csv


# -------------------------------
# Data Classes
# -------------------------------

class TreeBranch(object):
    def __init__(self, system):
        self.element_id = system.Id
        self.base_eq = system.BaseEquipment
        self.base_eq_id = self.base_eq.Id if self.base_eq else DB.ElementId.InvalidElementId
        self.base_eq_name = DB.Element.Name.__get__(self.base_eq) if self.base_eq else "Unknown"
        self.circuit_number = system.CircuitNumber
        self.load_name = system.LoadName
        self.system = system
        self.is_feeder = False  # Will be set True if it connects two equipment nodes

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
        self.nodes = {}  # {element_id: EquipmentNode}
        self.root_nodes = []  # list of EquipmentNode

    def add_node(self, node):
        self.nodes[node.element_id] = node
        if not node.upstream:
            self.root_nodes.append(node)

    def get_node(self, element_id):
        return self.nodes.get(element_id)

        # ... existing methods ...

    def walk_all(self):
        for root in self.root_nodes:
            output.print_md("---\n### Root: **{}**".format(root.panel_name))
            self.walk_tree(root, visited=set())

    def walk_tree(self, node, visited=None, level=0):
        if visited is None:
            visited = set()

        indent = "    " * level
        output.print_md("{}**{}** (ID: {}) _({})_".format(
            indent, node.panel_name, node.element_id, node.equipment_type
        ))

        visited.add(node.element_id)

        for branch in node.downstream:
            output.print_md("{}- Circuit `{}` | Load: `{}`".format(
                "    " * (level + 1),
                branch.circuit_number,
                branch.load_name
            ))

            system = branch.system
            non_nodes = []  # For collecting leafs

            if hasattr(system, "Elements"):
                for e in list(system.Elements):
                    if e.Category and int(e.Category.Id.IntegerValue) == int(
                            DB.BuiltInCategory.OST_ElectricalEquipment):
                        if e.Id not in visited:
                            branch.is_feeder = True
                            child_node = self.nodes.get(e.Id)
                            if child_node:
                                self.walk_tree(child_node, visited, level + 2)
                    else:
                        non_nodes.append(e)

            # LEAF SUMMARY
            leaf_counter = {}
            for e in non_nodes:
                cat = e.Category
                if not cat:
                    continue
                cat_name = cat.Name
                leaf_counter[cat_name] = leaf_counter.get(cat_name, 0) + 1

            for cat_name, count in leaf_counter.items():
                output.print_md("{}- `{}` leaf(s): **{}**".format(
                    "    " * (level + 2),
                    cat_name,
                    count
                ))

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

    def export_to_csv(self, filepath):
        data = self.to_list()
        if not data:
            logger.warning("No data to export.")
            return

        fieldnames = list(data[0].keys())

        try:
            with open(filepath, 'w') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(data)

            logger.info("Exported tree to CSV: {}".format(filepath))
        except Exception as e:
            logger.error("Failed to export CSV: {}".format(str(e)))

    def _request_path(self):
        path = forms.save_file(file_ext='csv',default_name="system_tree_export")
        return path


# ----------------------------
# Build the Equipment Map
# ----------------------------

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


# ----------------------------
# Recursively Build the Tree
# ----------------------------


# -------------------------------
# Main Execution
# -------------------------------
if __name__ == "__main__":
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    logger = script.get_logger()
    output.close_others()

    # Build node map
    equipment_map = get_all_equipment_nodes(doc)

    # Create and populate tree
    tree = SystemTree()
    for node in equipment_map.values():
        tree.add_node(node)

    tree.walk_all()
