# -*- coding: utf-8 -*-
from pyrevit import script, forms, revit, DB

logger = script.get_logger()
output = script.get_output()

# ------------------------------------------------------------
# CONSTANTS / CONFIG
# ------------------------------------------------------------

# Equipment part type → (Detail Family Name, Type Name)
PART_TYPE_MAPPING = {
    16: ("SLD-EQU-Switchboard-Top_CED", "Solid"),                     # Switchboard
    14: ("SLD-EQU-Panel-Top_CED", "Solid"),                           # Panelboard
    15: ("SLD-EQU-Transformer-Box-Ground-Top_CED", "Solid"),          # Transformer
    18: ("SLD-DEV-Motor Disconnect Combo-Top_CED", "Motor w/ Disconnect"),  # Equipment Switch / Disconnect
}

# Circuit symbol for feeders
CIRCUIT_SYMBOL_FAMILY = "SLD-FDR-Circuit Breaker_CED"
CIRCUIT_SYMBOL_TYPE = "Circuit Breaker"

# Feeder (conductor) symbol under breaker
FDR_SYMBOL_FAMILY = "SLD-FDR_CED"
FDR_SYMBOL_TYPE = "Medium Solid"
FDR_SYMBOL_LENGTH_PARAM = "Symbol Length_CED"

# Parameters to sync between model & detail items
PARAMETER_MAP = {
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Mains Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Main Breaker Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Short Circuit Rating_CEDT": DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING,
    "Mounting_CEDT": DB.BuiltInParameter.RBS_ELEC_MOUNTING,
    "Panel Modifications_CEDT": DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS,
    "Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM,
    "Total Connected Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM,
    "Total Demand Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Total Connected Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM,
    "Total Demand Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,

    "Voltage_CED": DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "CKT_Wire Size_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM
}

# Layout constants (Revit internal units = feet)
EQUIP_VERTICAL_SPACING = 2.5          # vertical distance between equipment levels
CIRCUIT_HORIZONTAL_SPACING = 1.5      # spacing between breakers
CIRCUIT_VERTICAL_OFFSET = 1.35        # distance below equipment for breaker symbols
SWITCHBOARD_BREAKER_OFFSET = 0.5      # 6" to the right of switchboard insertion
BREAKER_LENGTH_FEET = 0.75            # 9" breaker length
ROOT_HORIZONTAL_MARGIN = 5.0          # extra spacing between root subtrees

# Transformers: raise breaker insertion by 10.5" (0.875 ft)
TRANSFORMER_BREAKER_Y_ADJUST = 0.875

# ------------------------------------------------------------
# DATA CLASSES FOR TREE
# ------------------------------------------------------------

class TreeBranch(object):
    def __init__(self, system):
        self.element_id = system.Id
        self.base_eq = system.BaseEquipment
        self.base_eq_id = self.base_eq.Id if self.base_eq else DB.ElementId.InvalidElementId
        self.base_eq_name = DB.Element.Name.__get__(self.base_eq) if self.base_eq else "Unknown"
        self.circuit_number = system.CircuitNumber
        self.load_name = system.LoadName
        self.system = system
        self.is_feeder = False  # will be set True if it connects two equipment nodes

    def __str__(self):
        return "- Circuit `{}` | Load `{}` | Feeder `{}`".format(
            self.circuit_number, self.load_name, self.is_feeder
        )


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

        for sys in all_systems:
            br = TreeBranch(sys)
            # classify up/downstream relative to this element
            if br.base_eq_id == self.element_id:
                self.downstream.append(br)
            else:
                self.upstream.append(br)

            # mark as feeder if this circuit feeds other electrical equipment
            if hasattr(sys, "Elements"):
                for e in list(sys.Elements):
                    if e.Category and int(e.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                        br.is_feeder = True
                        break

        # Leaf check (no assigned downstream systems)
        self.is_leaf = not assigned or assigned.Count == 0


class SystemTree(object):
    def __init__(self):
        self.nodes = {}       # {element_id: TreeNode}
        self.root_nodes = []  # list of TreeNode

    def add_node(self, node):
        self.nodes[node.element_id] = node

    def build_roots(self):
        self.root_nodes = []
        for node in self.nodes.values():
            if not node.upstream:
                self.root_nodes.append(node)

    def get_node(self, element_id):
        return self.nodes.get(element_id)

    def _should_include_branch(self, node, branch):
        """
        Returns True if the branch should be included in output,
        based on equipment type rules.

        - Transformers & Switchboards: include ALL downstream circuits
        - Panelboards: include only circuits feeding other equipment
        """
        eq_type = (node.equipment_type or "").lower()

        # Transformers and Switchboards: include all circuits
        if "transformer" in eq_type or "switchboard" in eq_type:
            return True

        # Panelboards and others: only include feeders to equipment
        system = branch.system
        if hasattr(system, "Elements"):
            for e in list(system.Elements):
                if e.Category and int(e.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                    return True

        return False


# ------------------------------------------------------------
# DETAIL / SYMBOL CLASSES
# ------------------------------------------------------------

class PropKeyValue(object):
    """Storage class for matched property info and value."""
    def __init__(self, name, datatype, value, istype):
        self.name = name
        self.datatype = datatype
        self.value = value
        self.istype = istype

    def __repr__(self):
        return str(self.__dict__)


class ComponentSymbol(object):
    """Represents a detail item symbol (FamilySymbol)."""
    def __init__(self, symbol):
        self.symbol = symbol

    @staticmethod
    def get_symbol(doc, family_name, type_name):
        """Retrieve a symbol by family name and type name."""
        collector = DB.FilteredElementCollector(doc) \
            .OfCategory(DB.BuiltInCategory.OST_DetailComponents) \
            .WhereElementIsElementType()
        for symbol in collector:
            if isinstance(symbol, DB.FamilySymbol):
                family_name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                symbol_family_name = family_name_param.AsString() if family_name_param else None
                type_param = symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                symbol_type_name = type_param.AsString() if type_param else None
                if symbol_family_name == family_name and symbol_type_name == type_name:
                    return ComponentSymbol(symbol)
        return None

    def activate(self, doc):
        """Activate the symbol if not already active."""
        if not self.symbol.IsActive:
            self.symbol.Activate()
            doc.Regenerate()

    def place(self, doc, view, location, rotation=0):
        """Place the symbol at a location in the view."""
        try:
            detail_item = doc.Create.NewFamilyInstance(location, self.symbol, view)
            if rotation:
                axis = DB.Line.CreateBound(location, location + DB.XYZ(0, 0, 1))
                detail_item.Location.Rotate(axis, rotation)
            return detail_item
        except Exception as ex:
            logger.error("Failed to place symbol: {0}".format(str(ex)))
            return None


class ElectricalComponent(object):
    """Base class for electrical components."""
    def __init__(self, element):
        self.element = element

    def get_param_value(self, param):
        """Retrieve the value of a parameter."""
        param_value = self.element.get_Parameter(param)
        if param_value:
            if param_value.StorageType == DB.StorageType.String:
                return param_value.AsString()
            elif param_value.StorageType == DB.StorageType.Integer:
                return param_value.AsInteger()
            elif param_value.StorageType == DB.StorageType.Double:
                return param_value.AsDouble()
        return None

    def set_parameters(self, detail_item, param_map, component_id, symbol_id):
        """Set parameters on a detail item and the physical component."""
        # Copy standard mapped parameters
        for detail_param_name, source_param in param_map.items():
            value = self.get_param_value(source_param)
            if value is not None:
                detail_param = detail_item.LookupParameter(detail_param_name)
                if detail_param:
                    if isinstance(value, str):
                        detail_param.Set(value)
                    else:
                        detail_param.Set(value)

        # Copy Component ID and Symbol ID
        self._set_id(detail_item, "SLD_Component ID_CED", component_id)
        self._set_id(detail_item, "SLD_Symbol ID_CED", symbol_id)
        self._set_id(self.element, "SLD_Component ID_CED", component_id)
        self._set_id(self.element, "SLD_Symbol ID_CED", symbol_id)

    @staticmethod
    def _set_id(target, param_name, value):
        """Helper method to set ID values."""
        if not target:
            return
        param = target.LookupParameter(param_name)
        if param:
            param.Set(str(value))


class Equipment(ElectricalComponent):
    """Represents electrical equipment."""
    def get_part_type(self):
        """Retrieve the FAMILY_CONTENT_PART_TYPE value."""
        symbol = self.element.Symbol
        if not symbol:
            return None
        family = symbol.Family
        if not family:
            return None
        part_type_param = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
        if part_type_param and part_type_param.HasValue:
            return part_type_param.AsInteger()
        return None

    def get_panel_name(self):
        """Retrieve the panel name."""
        return self.get_param_value(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)


class Circuit(ElectricalComponent):
    """Represents an electrical circuit."""
    pass


# ------------------------------------------------------------
# TREE BUILDING & SELECTION
# ------------------------------------------------------------

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
    nodes = get_all_equipment_nodes(doc)
    tree = SystemTree()
    for node in nodes.values():
        tree.add_node(node)
    tree.build_roots()
    return tree


def format_panel_display(element):
    return "{0} (ID: {1})".format(DB.Element.Name.__get__(element), element.Id.IntegerValue)


def select_root_nodes(doc, tree):
    """Let user choose which root equipment to use as SLD start."""
    if not tree.root_nodes:
        forms.alert("No root electrical equipment found.", title="System Tree")
        return []

    display_to_node = {}
    options = []
    for node in tree.root_nodes:
        desc = format_panel_display(node.element)
        display_to_node[desc] = node
        options.append(desc)

    selected = forms.SelectFromList.show(
        options,
        title="Select Root Equipment for SLD",
        multiselect=True
    )
    if not selected:
        return []

    return [display_to_node[s] for s in selected]


def estimate_subtree_width(node, tree):
    """
    Rough width estimate (in feet) for this node's subtree,
    used to space equipment so they don't overlap.
    """
    eq_type_lower = (node.equipment_type or "").lower()
    entries = []

    for branch in node.downstream:
        if not tree._should_include_branch(node, branch):
            continue

        system = branch.system
        child_nodes = []

        if hasattr(system, "Elements"):
            for e in list(system.Elements):
                if e.Category and int(e.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                    cn = tree.get_node(e.Id)
                    if cn:
                        child_nodes.append(cn)

        if child_nodes:
            for cn in child_nodes:
                entries.append({"child": cn})
        else:
            entries.append({"child": None})

    count = len(entries)

    if count <= 0:
        base_width = 4.0
    else:
        if "switchboard" in eq_type_lower:
            # switchboard width ~ breakers span + 1.0ft (two 6" gaps) + margin
            base_width = (count - 1) * CIRCUIT_HORIZONTAL_SPACING + 1.0 + 2.0
        else:
            # centered breakers + small margins
            base_width = (count - 1) * CIRCUIT_HORIZONTAL_SPACING + 2.0

        if base_width < 4.0:
            base_width = 4.0

    max_child_width = 0.0
    for entry in entries:
        cn = entry["child"]
        if cn is not None:
            child_w = estimate_subtree_width(cn, tree)
            if child_w > max_child_width:
                max_child_width = child_w

    if max_child_width > base_width:
        base_width = max_child_width

    return base_width


def allocate_x_for_level(level_spans, level, desired_x, width):
    """
    Adjusts X position for a node at a given level so that its
    approximate [min_x, max_x] span does not overlap existing spans
    on that level. Returns the adjusted X.
    """
    spans = level_spans.get(level, [])
    half = width / 2.0
    x = desired_x

    while True:
        min_x = x - half
        max_x = x + half
        overlap = False

        for (smin, smax) in spans:
            # overlap if ranges intersect
            if not (max_x <= smin or min_x >= smax):
                overlap = True
                break

        if not overlap:
            spans.append((min_x, max_x))
            level_spans[level] = spans
            return x

        # bump to the right by half width + 1.0ft margin
        x += half + 1.0


# ------------------------------------------------------------
# SLD PLACEMENT LOGIC (TOP-DOWN)
# ------------------------------------------------------------

def place_equipment_recursive(doc, view, tree, node, origin, placed_equipment,
                              placed_circuits, visited, subtree_widths,
                              level_spans, level):
    """
    Recursively place equipment and feeder circuits in a top-down layout.

    origin: DB.XYZ location of this equipment symbol
    placed_equipment: dict(ElementId -> FamilyInstance)
    placed_circuits: set(ElementId)
    visited: set(ElementId) to avoid infinite loops
    subtree_widths: dict(ElementId -> float)
    level_spans: dict(level_index -> [(min_x, max_x)])
    level: current tree depth (root = 0)
    """
    if node.element_id in visited:
        return
    visited.add(node.element_id)

    element = node.element
    equip = Equipment(element)
    eq_type_lower = (node.equipment_type or "").lower()

    # Place equipment detail symbol (if not already placed)
    if node.element_id in placed_equipment:
        equip_detail = placed_equipment[node.element_id]
    else:
        part_type = equip.get_part_type()
        equip_detail = None
        if part_type in PART_TYPE_MAPPING:
            family_name, type_name = PART_TYPE_MAPPING[part_type]
            symbol = ComponentSymbol.get_symbol(doc, family_name, type_name)
            if symbol:
                symbol.activate(doc)
                equip_detail = symbol.place(doc, view, origin)
                if equip_detail:
                    equip.set_parameters(
                        equip_detail,
                        PARAMETER_MAP,
                        element.Id.IntegerValue,
                        equip_detail.Id.IntegerValue
                    )
        else:
            logger.warning("No part-type mapping for equipment ID {0}".format(element.Id.IntegerValue))

        placed_equipment[node.element_id] = equip_detail

    # Build list of branches (circuits) to include from this node
    entries = []  # each: {"branch": branch, "child": TreeNode or None}
    for branch in node.downstream:
        if not tree._should_include_branch(node, branch):
            continue

        system = branch.system
        child_nodes = []

        if hasattr(system, "Elements"):
            for e in list(system.Elements):
                if e.Category and int(e.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                    child_node = tree.get_node(e.Id)
                    if child_node:
                        child_nodes.append(child_node)

        if child_nodes:
            for cn in child_nodes:
                entries.append({"branch": branch, "child": cn})
        else:
            # terminal branch (to devices/fixtures)
            entries.append({"branch": branch, "child": None})

    if not entries:
        return

    count = len(entries)

    # For switchboards, set Symbol Width_CED to encompass all breakers
    if "switchboard" in eq_type_lower and equip_detail is not None and count > 0:
        try:
            width_feet = (count - 1) * CIRCUIT_HORIZONTAL_SPACING + 1.0  # 6" left + span + 6" right
            width_param = equip_detail.LookupParameter("Symbol Width_CED")
            if width_param:
                width_param.Set(width_feet)
        except Exception as ex:
            logger.warning(
                "Failed to set Symbol Width_CED on switchboard ID {0}: {1}".format(
                    element.Id.IntegerValue, str(ex)
                )
            )

    # Place circuit symbols and recurse for child equipment
    rotation = 3.0 * 3.14159 / 2.0  # 270 degrees

    for idx in range(count):
        entry = entries[idx]
        branch = entry["branch"]
        child_node = entry["child"]

        # Horizontal offset to spread circuits
        if "switchboard" in eq_type_lower:
            # first breaker 6" to the right, then spaced
            offset_x = SWITCHBOARD_BREAKER_OFFSET + idx * CIRCUIT_HORIZONTAL_SPACING
        else:
            # centered around equipment
            offset_x = (idx - (count - 1) / 2.0) * CIRCUIT_HORIZONTAL_SPACING

        circuit_x = origin.X + offset_x
        circuit_y = origin.Y - CIRCUIT_VERTICAL_OFFSET

        # Transformers: raise breaker insertion by 10.5"
        if "transformer" in eq_type_lower:
            circuit_y += TRANSFORMER_BREAKER_Y_ADJUST

        circuit_location = DB.XYZ(circuit_x, circuit_y, 0)

        circuit_id = branch.element_id
        circuit_element = doc.GetElement(circuit_id)
        circuit_wrapper = Circuit(circuit_element) if circuit_element is not None else None

        circuit_detail = None
        if circuit_id not in placed_circuits:
            circuit_symbol = ComponentSymbol.get_symbol(doc, CIRCUIT_SYMBOL_FAMILY, CIRCUIT_SYMBOL_TYPE)
            if circuit_symbol:
                circuit_symbol.activate(doc)
                circuit_detail = circuit_symbol.place(doc, view, circuit_location, rotation=rotation)
                if circuit_detail and circuit_wrapper is not None:
                    try:
                        circuit_wrapper.set_parameters(
                            circuit_detail,
                            PARAMETER_MAP,
                            circuit_element.Id.IntegerValue,
                            circuit_detail.Id.IntegerValue
                        )
                    except Exception as ex:
                        logger.warning(
                            "Failed to set parameters on breaker for circuit {0}: {1}".format(
                                circuit_element.Id.IntegerValue, str(ex)
                            )
                        )
            placed_circuits.add(circuit_id)

        # If this branch feeds another equipment, place FDR symbol and child equipment
        child_origin = None
        if child_node is not None:
            child_origin_y = origin.Y - EQUIP_VERTICAL_SPACING

            # FDR placement: insertion at breaker endpoint, grows downward
            if circuit_wrapper is not None and circuit_detail is not None:
                try:
                    # breaker endpoint (9" = 0.75 ft below breaker location, with 270° rotation)
                    breaker_end_y = circuit_y - BREAKER_LENGTH_FEET

                    # downstream equipment vertical target
                    target_y = child_origin_y

                    # total vertical length required for feeder
                    fdr_length = abs(target_y - breaker_end_y)

                    fdr_location = DB.XYZ(circuit_x, breaker_end_y, 0)

                    fdr_symbol = ComponentSymbol.get_symbol(doc, FDR_SYMBOL_FAMILY, FDR_SYMBOL_TYPE)
                    if fdr_symbol:
                        fdr_symbol.activate(doc)
                        fdr_detail = fdr_symbol.place(doc, view, fdr_location, rotation=rotation)
                        if fdr_detail:
                            # set the feeder length
                            length_param = fdr_detail.LookupParameter(FDR_SYMBOL_LENGTH_PARAM)
                            if length_param:
                                length_param.Set(fdr_length)

                            # copy all mapped parameters like breaker
                            circuit_wrapper.set_parameters(
                                fdr_detail,
                                PARAMETER_MAP,
                                circuit_element.Id.IntegerValue,
                                fdr_detail.Id.IntegerValue
                            )

                except Exception as ex:
                    logger.warning(
                        "Error placing feeder symbol for circuit {0}: {1}".format(
                            circuit_id.IntegerValue, str(ex)
                        )
                    )

            # Now place child equipment, with horizontal spacing to avoid overlap on this level
            # Desired X is directly under breaker
            desired_child_x = circuit_x
            child_width = subtree_widths.get(child_node.element_id, 4.0)
            adjusted_child_x = allocate_x_for_level(level_spans, level + 1, desired_child_x, child_width)
            child_origin = DB.XYZ(adjusted_child_x, child_origin_y, 0)

            place_equipment_recursive(
                doc,
                view,
                tree,
                child_node,
                child_origin,
                placed_equipment,
                placed_circuits,
                visited,
                subtree_widths,
                level_spans,
                level + 1
            )


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    doc = revit.doc
    view = revit.active_view

    if not isinstance(view, DB.ViewDrafting):
        forms.alert("Please run this script in a Drafting View.", title="SLD Generator")
        return

    output.close_others()
    output.print_md("## Single-Line Diagram Generator\n")

    # Build electrical system tree
    tree = build_system_tree(doc)

    # Let user select root equipment
    root_nodes = select_root_nodes(doc, tree)
    if not root_nodes:
        return

    # Precompute subtree widths for ALL nodes
    subtree_widths = {}
    for n in tree.nodes.values():
        try:
            subtree_widths[n.element_id] = estimate_subtree_width(n, tree)
        except Exception as ex:
            logger.warning(
                "Failed to estimate subtree width for node ID {0}: {1}".format(
                    n.element_id.IntegerValue, str(ex)
                )
            )
            subtree_widths[n.element_id] = 4.0

    # Compute root-level spacing based on subtree widths
    start_y = 0.0
    current_x = 0.0
    level_spans = {}  # level_index -> list of (min_x, max_x)

    placed_equipment = {}
    placed_circuits = set()

    with revit.Transaction("Generate SLD from Electrical System Tree"):
        for root_node in root_nodes:
            width = subtree_widths.get(root_node.element_id, 4.0)

            # center root within its subtree span
            root_x = current_x + width / 2.0
            root_origin = DB.XYZ(root_x, start_y, 0)

            # record span at level 0
            spans0 = level_spans.get(0, [])
            spans0.append((root_x - width / 2.0, root_x + width / 2.0))
            level_spans[0] = spans0

            visited = set()
            place_equipment_recursive(
                doc,
                view,
                tree,
                root_node,
                root_origin,
                placed_equipment,
                placed_circuits,
                visited,
                subtree_widths,
                level_spans,
                0
            )

            current_x += width + ROOT_HORIZONTAL_MARGIN

    output.print_md("Done. Placed equipment, breakers, and feeders based on system tree.")


if __name__ == "__main__":
    main()
