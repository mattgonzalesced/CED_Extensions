# -*- coding: utf-8 -*-
from pyrevit import script, forms, revit, DB

# Constants
PART_TYPE_MAPPING = {
    16: ("DME-EQU-Switchboard-Top_CED", "Solid"),
    14: ("DME-EQU-Panel-Top_CED", "Solid"),
    15: ("DME-EQU-Transformer-Box-Ground-Top_CED", "Solid"),
    18: ("DME-EQU-Disconnect-Top_CED", "Fused Disconnect")
}

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


class PropKeyValue(object):
    """Storage class for matched property info and value."""

    def __init__(self, name, datatype, value, istype):
        self.name = name
        self.datatype = datatype
        self.value = value
        self.istype = istype

    def __repr__(self):
        return str(self.__dict__)

class ElectricalComponent:
    """Base class for electrical components."""
    def __init__(self, element):
        self.element = element
        self.ComponentSymbol = ComponentSymbol

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
                    detail_param.Set(str(value) if isinstance(value, str) else value)

        # Copy Component ID and Symbol ID
        self._set_id(detail_item, "SLD_Component ID_CED", component_id)
        self._set_id(detail_item, "SLD_Symbol ID_CED", symbol_id)
        self._set_id(self.element, "SLD_Component ID_CED", component_id)
        self._set_id(self.element, "SLD_Symbol ID_CED", symbol_id)


    @staticmethod
    def _set_id(target, param_name, value):
        """Helper method to set ID values."""
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
        return part_type_param.AsInteger() if part_type_param and part_type_param.HasValue else None

    def get_panel_name(self):
        """Retrieve the panel name."""
        return self.get_param_value(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)

class Circuit(ElectricalComponent):
    """Represents an electrical circuit."""
    pass

class ComponentSymbol:
    """Represents a detail item symbol (FamilySymbol)."""
    def __init__(self, symbol):
        self.symbol = symbol

    @staticmethod
    def get_symbol(doc, family_name, type_name):
        """Retrieve a symbol by family name and type name."""
        collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DetailComponents).WhereElementIsElementType()
        for symbol in collector:
            if isinstance(symbol, DB.FamilySymbol):
                family_name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                symbol_family_name = family_name_param.AsString() if family_name_param else None
                if symbol_family_name == family_name and symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == type_name:
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
                detail_item.Location.Rotate(DB.Line.CreateBound(location, location + DB.XYZ(0, 0, 1)), rotation)
            return detail_item
        except:
            return None

def format_panel_display(panel):
    """Returns a string with panel name and element ID for display."""
    return "{} (ID: {})".format(panel.Name, panel.Id)

def collect_electrical_equipment(doc):
    """Collect all Electrical Equipment elements in the project."""
    panel_collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType()
    panel_options = [format_panel_display(panel) for panel in panel_collector]

    selected_panel_options = forms.SelectFromList.show(
        panel_options,
        title="Select Panels",
        multiselect=True
    )
    if not selected_panel_options:
        script.exit()

    # Map selected panel display names back to panel objects
    selected_panels = [panel for panel in panel_collector if format_panel_display(panel) in selected_panel_options]
    return selected_panels

def collect_circuits(doc):
    """Collect all Electrical Circuits in the project."""
    return DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit) \
        .WhereElementIsNotElementType() \
        .ToElements()

def main():
    doc = revit.doc
    view = revit.active_view

    if not isinstance(view, DB.ViewDrafting):
        forms.alert("Please run this script in a Drafting View.")
        return

    equipment_elements = collect_electrical_equipment(doc)
    circuits = collect_circuits(doc)

    start_point = DB.XYZ(0, 0, 0)
    vertical_offset = DB.XYZ(0, -10, 0)
    circuit_horizontal_spacing = 1.5
    circuit_vertical_offset = DB.XYZ(0, -1.35, 0)
    circuit_base_offset = DB.XYZ(0.55, 0, 0)

    with revit.Transaction("Place Detail Items for Electrical Equipment and Circuits"):
        for idx, equipment in enumerate(equipment_elements):
            component = Equipment(equipment)
            location = start_point + (idx * vertical_offset)

            # Place detail item for equipment
            part_type = component.get_part_type()
            if part_type in PART_TYPE_MAPPING:
                family_name, type_name = PART_TYPE_MAPPING[part_type]
                symbol = ComponentSymbol.get_symbol(doc, family_name, type_name)
                if symbol:
                    symbol.activate(doc)
                    detail_item = symbol.place(doc, view, location)
                    if detail_item:
                        component.set_parameters(detail_item, PARAMETER_MAP, equipment.Id.IntegerValue, detail_item.Id.IntegerValue)

            # Handle circuits for switchboards
            if part_type == 16:  # Switchboard
                switchboard_origin = location + circuit_base_offset + circuit_vertical_offset
                switchboard_circuits = [
                    Circuit(c) for c in circuits if c.BaseEquipment and c.BaseEquipment.Id == equipment.Id
                ]
                sorted_circuits = sorted(
                    switchboard_circuits, key=lambda x: x.get_param_value(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT)
                )
                for c_idx, circuit in enumerate(sorted_circuits):
                    circuit_location = switchboard_origin + DB.XYZ(c_idx * circuit_horizontal_spacing, 0, 0)
                    circuit_symbol = ComponentSymbol.get_symbol(doc, "DME-FDR-Circuit Breaker_CED", "Medium Solid")
                    if circuit_symbol:
                        circuit_symbol.activate(doc)
                        circuit_detail = circuit_symbol.place(doc, view, circuit_location, rotation=(3 * 3.14159 / 2))
                        if circuit_detail:
                            circuit.set_parameters(circuit_detail, PARAMETER_MAP, circuit.element.Id.IntegerValue, circuit_detail.Id.IntegerValue)

# Run the script
if __name__ == "__main__":
    main()
