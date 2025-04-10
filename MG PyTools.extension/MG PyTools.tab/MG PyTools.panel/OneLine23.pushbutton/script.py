# -*- coding: utf-8 -*-
import re
from pyrevit import forms, revit, DB
from place_tags import tag_family_instances

PART_TYPE_MAPPING = {
    16: ("DME-EQU-Switchboard-Top_CED", "Solid"),
    14: ("DME-EQU-Panel-Top_CED", "Solid"),
    15: ("DME-EQU-Transformer-Box-Ground-Top_CED", "Solid"),
    18: ("DME-EQU-Disconnect-Top_CED", "Fused Disconnect")
}

PARAMETER_MAP = {
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "EQU_Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "Voltage_CED": DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "CKT_Wire Size_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM,
    "Wire size_t": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM,
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Conduit and Wire size_CEDT": "WIRE_SIZE_STRING",
    "Mains Rating_CED": "MAIN_RATING",
    "Mains Type_CEDTMG": "MAINS_TYPE",
    "Distribution Systems_CEDT": DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM
}

def get_converted_value(param_obj):
    """
    Convert a parameter object to its Python value.
    If the parameter is the Distribution System parameter,
    use AsValueString() directly.
    For strings, if AsString() is empty, fall back to AsValueString().
    """
    if not param_obj or not param_obj.HasValue:
        return None

    if param_obj.Definition and param_obj.Definition.BuiltInParameter == DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM:
        return param_obj.AsValueString()

    st = param_obj.StorageType
    if st == DB.StorageType.String:
        value = param_obj.AsString()
        if not value or value.strip() == "":
            value = param_obj.AsValueString()
        return value
    elif st == DB.StorageType.Integer:
        return param_obj.AsInteger()
    elif st == DB.StorageType.Double:
        return param_obj.AsDouble()
    return None

def copy_param_value(source, detail_item, detail_param_name):
    """Copies a parameter's value from the source to the detail item."""
    value = get_converted_value(source)
    target = detail_item.LookupParameter(detail_param_name)
    target.Set(str(value) if isinstance(value, str) else value)

def copy_custom_param(element, detail_item, detail_param_name, element_param_name):
    """Copies a custom parameter from the element to the detail item."""
    source = element.LookupParameter(element_param_name)
    if source and source.HasValue:
        copy_param_value(source, detail_item, detail_param_name)

def set_ids(targets, param_name, value):
    """Sets the same ID parameter on multiple targets."""
    for target in targets:
        param = target.LookupParameter(param_name)
        param.Set(str(value))

class ElectricalComponent:
    def __init__(self, element):
        self.element = element

    def get_param_value(self, param):
        param_obj = self.element.get_Parameter(param)
        if param_obj and param_obj.HasValue:
            return get_converted_value(param_obj)
        return None

    def set_parameters(self, detail_item, param_map, component_id, symbol_id):
        for detail_param_name, source_param in param_map.items():
            if isinstance(source_param, DB.BuiltInParameter):
                value = self.get_param_value(source_param)
                if value is not None:
                    target = detail_item.LookupParameter(detail_param_name)
                    if target:
                        target.Set(str(value) if isinstance(value, str) else value)
            elif source_param == "MAIN_RATING":
                copy_custom_param(self.element, detail_item, detail_param_name, "Main Rating")
            elif source_param == "MAINS_TYPE":
                copy_custom_param(self.element, detail_item, detail_param_name, "Mains Type")
        set_ids([detail_item, self.element], "SLD_Component ID_CED", component_id)
        set_ids([detail_item, self.element], "SLD_Symbol ID_CED", symbol_id)

class Equipment(ElectricalComponent):
    def get_part_type(self):
        symbol = self.element.Symbol
        family = symbol.Family
        part_type_param = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
        return part_type_param.AsInteger() if part_type_param and part_type_param.HasValue else None

    def get_panel_name(self):
        return self.get_param_value(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)

class Circuit(ElectricalComponent):
    def set_parameters(self, detail_item, param_map, component_id, symbol_id):
        for detail_param_name, source_param in param_map.items():
            if source_param == "WIRE_SIZE_STRING":
                wire_size_str = getattr(self.element, "WireSizeString", None)
                target = detail_item.LookupParameter(detail_param_name)
                if target and not target.IsReadOnly:
                    target.Set(wire_size_str)
            elif source_param in ["MAIN_RATING", "MAINS_TYPE"]:
                element_param_name = "Main Rating" if source_param == "MAIN_RATING" else "Mains Type"
                copy_custom_param(self.element, detail_item, detail_param_name, element_param_name)
            elif isinstance(source_param, DB.BuiltInParameter):
                value = self.get_param_value(source_param)
                if value is not None:
                    target = detail_item.LookupParameter(detail_param_name)
                    if target:
                        target.Set(str(value) if isinstance(value, str) else value)
        set_ids([detail_item, self.element], "SLD_Component ID_CED", component_id)
        set_ids([detail_item, self.element], "SLD_Symbol ID_CED", symbol_id)

class ComponentSymbol:
    def __init__(self, symbol):
        self.symbol = symbol

    @staticmethod
    def get_symbol(doc, family_name, type_name):
        collector = DB.FilteredElementCollector(doc)\
                     .OfCategory(DB.BuiltInCategory.OST_DetailComponents)\
                     .WhereElementIsElementType()
        for symbol in collector:
            if isinstance(symbol, DB.FamilySymbol):
                family_name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                symbol_family = family_name_param.AsString() if family_name_param else None
                type_name_param = symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                type_name_str = type_name_param.AsString() if type_name_param else None
                if symbol_family == family_name and type_name_str == type_name:
                    return ComponentSymbol(symbol)
        return None

    def activate(self, doc):
        if not self.symbol.IsActive:
            self.symbol.Activate()
            doc.Regenerate()

    def place(self, doc, view, location, rotation=0):
        try:
            detail_item = doc.Create.NewFamilyInstance(location, self.symbol, view)
            if rotation:
                detail_item.Location.Rotate(
                    DB.Line.CreateBound(location, location + DB.XYZ(0, 0, 1)),
                    rotation)
            return detail_item
        except Exception:
            return None

def collect_electrical_equipment(doc):
    return DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)\
            .WhereElementIsNotElementType()\
            .ToElements()

def collect_circuits(doc):
    return DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit)\
            .WhereElementIsNotElementType()\
            .ToElements()

# ---------------------------------------------------------------------
# Modified main() to create hierarchical groups based on switchboards.
# Panels and transformers are now moved an additional 3.5 feet higher (Z axis)
# relative to the circuit breakers.
#
# Additionally:
# 1. Switchboards without a value in "CKT_Panel_CEDT" are placed along the top (y = 0)
#    with a 10-foot offset along the x axis.
# 2. If a switchboard’s "CKT_Panel_CEDT" matches a panel’s or transformer’s
#    "Panel Name_CEDT", then that entire switchboard group is moved 5 feet underneath
#    the correlated element.
# 3. For every switchboard group that is moved, the next one is placed 3.5 feet lower
#    on the y axis to prevent overlap.
# ---------------------------------------------------------------------
def main():
    doc = revit.doc
    view = revit.active_view
    if not isinstance(view, DB.ViewDrafting):
        forms.alert("Please run this script in a Drafting View.")
        return

    equipment_elements = collect_electrical_equipment(doc)
    circuits = collect_circuits(doc)

    # Separate equipment by part type.
    switchboard_components_with_panel = []
    switchboard_components_without_panel = []
    panel_components = []
    transformer_components = []
    other_components = []

    for equipment in equipment_elements:
        comp = Equipment(equipment)
        part_type = comp.get_part_type()
        if part_type == 16:
            # Check if the switchboard has a value for CKT_Panel_CEDT.
            panel_val = comp.get_param_value(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
            if panel_val:
                switchboard_components_with_panel.append(comp)
            else:
                switchboard_components_without_panel.append(comp)
        elif part_type == 14:
            panel_components.append(comp)
        elif part_type == 15:
            transformer_components.append(comp)
        else:
            other_components.append(comp)

    # Build a mapping for circuits keyed by their BaseEquipment Id.
    circuits_by_equipment = {}
    for c in circuits:
        if c.BaseEquipment:
            eq_id = c.BaseEquipment.Id.IntegerValue
            circuits_by_equipment.setdefault(eq_id, []).append(Circuit(c))

    # Define base positions and offsets.
    base_point = DB.XYZ(0, 0, 0)
    row_offset = DB.XYZ(0, -20, 0)         # vertical spacing between groups (for switchboards with panel value)
    group_offset = DB.XYZ(0, -5, 0)         # original horizontal offset for panels/transformers relative to breakers
    panel_offset = DB.XYZ(0, -1.5, 0)        # adjusted offset for panels (relative to breakers)
    transformer_offset = DB.XYZ(0, -1.75, 0) # adjusted offset for transformers (relative to breakers)
    circuit_offset = DB.XYZ(1.1, -1.35, 0)   # base offset for circuits relative to switchboard
    circuit_horizontal_spacing = 1.333334
    height_offset = DB.XYZ(0, 0, 3.5)        # additional 3.5 feet higher for panels/transformers

    # List to store switchboard groups.
    switchboard_groups = []

    with revit.Transaction("Place Hierarchical Switchboard Groups"):
        # 1. Place switchboards with a panel value using the original row spacing.
        for idx, sb in enumerate(switchboard_components_with_panel):
            group_base = base_point + idx * row_offset
            family_name, type_name = PART_TYPE_MAPPING.get(16, (None, None))
            if family_name:
                symbol = ComponentSymbol.get_symbol(doc, family_name, type_name)
                if symbol:
                    symbol.activate(doc)
                    sb_detail = symbol.place(doc, view, group_base)
                    if sb_detail:
                        sb.set_parameters(sb_detail, PARAMETER_MAP,
                                          sb.element.Id.IntegerValue,
                                          sb_detail.Id.IntegerValue)
                        group = {
                            "switchboard": sb,
                            "sb_detail": sb_detail,
                            "circuits": [],
                            "panels": [],
                            "transformers": [],
                            # Retrieve the switchboard's CKT_Panel_CEDT value for hierarchy matching.
                            "panel_key": sb.get_param_value(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
                        }
                        eq_id = sb.element.Id.IntegerValue
                        sb_circuits = circuits_by_equipment.get(eq_id, [])
                        if sb_circuits:
                            circuit_base = group_base + circuit_offset
                            for c_idx, circuit in enumerate(
                                    sorted(sb_circuits,
                                           key=lambda x: x.get_param_value(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT))):
                                circ_loc = circuit_base + DB.XYZ(c_idx * circuit_horizontal_spacing, 0, 0)
                                circuit_symbol = ComponentSymbol.get_symbol(doc, "DME-FDR-Circuit Breaker_CED", "Medium Solid")
                                if circuit_symbol:
                                    circuit_symbol.activate(doc)
                                    circ_detail = circuit_symbol.place(doc, view, circ_loc,
                                                                       rotation=(3 * 3.14159 / 2))
                                    if circ_detail:
                                        circuit.set_parameters(circ_detail, PARAMETER_MAP,
                                                               circuit.element.Id.IntegerValue,
                                                               circ_detail.Id.IntegerValue)
                                        group["circuits"].append(circ_detail)
                            count_breakers = len(group["circuits"])
                            width_value = max(3, count_breakers * 1.45)
                            width_param = sb_detail.LookupParameter("DME_Width")
                            if width_param and not width_param.IsReadOnly:
                                width_param.Set(width_value)
                        switchboard_groups.append(group)

        # 2. Place switchboards without a panel value along the top (y = 0) with a 10-foot x offset.
        for idx, sb in enumerate(switchboard_components_without_panel):
            group_base = DB.XYZ(idx * 10, 0, 0)
            family_name, type_name = PART_TYPE_MAPPING.get(16, (None, None))
            if family_name:
                symbol = ComponentSymbol.get_symbol(doc, family_name, type_name)
                if symbol:
                    symbol.activate(doc)
                    sb_detail = symbol.place(doc, view, group_base)
                    if sb_detail:
                        sb.set_parameters(sb_detail, PARAMETER_MAP,
                                          sb.element.Id.IntegerValue,
                                          sb_detail.Id.IntegerValue)
                        group = {
                            "switchboard": sb,
                            "sb_detail": sb_detail,
                            "circuits": [],
                            "panels": [],
                            "transformers": [],
                            "panel_key": sb.get_param_value(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
                        }
                        eq_id = sb.element.Id.IntegerValue
                        sb_circuits = circuits_by_equipment.get(eq_id, [])
                        if sb_circuits:
                            circuit_base = group_base + circuit_offset
                            for c_idx, circuit in enumerate(
                                    sorted(sb_circuits,
                                           key=lambda x: x.get_param_value(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT))):
                                circ_loc = circuit_base + DB.XYZ(c_idx * circuit_horizontal_spacing, 0, 0)
                                circuit_symbol = ComponentSymbol.get_symbol(doc, "DME-FDR-Circuit Breaker_CED", "Medium Solid")
                                if circuit_symbol:
                                    circuit_symbol.activate(doc)
                                    circ_detail = circuit_symbol.place(doc, view, circ_loc,
                                                                       rotation=(3 * 3.14159 / 2))
                                    if circ_detail:
                                        circuit.set_parameters(circ_detail, PARAMETER_MAP,
                                                               circuit.element.Id.IntegerValue,
                                                               circ_detail.Id.IntegerValue)
                                        group["circuits"].append(circ_detail)
                            count_breakers = len(group["circuits"])
                            width_value = max(3, count_breakers * 1.45)
                            width_param = sb_detail.LookupParameter("DME_Width")
                            if width_param and not width_param.IsReadOnly:
                                width_param.Set(width_value)
                        switchboard_groups.append(group)

        # 3. Place panels and assign them to groups based on matching breaker values.
        unmatched_panels = []
        for panel in panel_components:
            family_name, type_name = PART_TYPE_MAPPING.get(14, (None, None))
            if family_name:
                symbol = ComponentSymbol.get_symbol(doc, family_name, type_name)
                if symbol:
                    symbol.activate(doc)
                    # Place at a temporary location; will be repositioned relative to its group.
                    temp_loc = base_point
                    panel_detail = symbol.place(doc, view, temp_loc)
                    if panel_detail:
                        panel.set_parameters(panel_detail, PARAMETER_MAP,
                                             panel.element.Id.IntegerValue,
                                             panel_detail.Id.IntegerValue)
                        panel_name_param = panel_detail.LookupParameter("Panel Name_CEDT")
                        if panel_name_param and panel_name_param.HasValue:
                            panel_name = panel_name_param.AsString()
                            matched = False
                            # Check each group’s circuits for a matching CKT_Load Name_CEDT.
                            for group in switchboard_groups:
                                for breaker in group["circuits"]:
                                    breaker_param = breaker.LookupParameter("CKT_Load Name_CEDT")
                                    if breaker_param and breaker_param.HasValue:
                                        breaker_load = breaker_param.AsString()
                                        m = re.search(r'"([^"]+)"', breaker_load)
                                        if m:
                                            breaker_load = m.group(1)
                                        if breaker_load == panel_name:
                                            if hasattr(breaker, "Location") and breaker.Location:
                                                brk_loc = breaker.Location.Point
                                                new_loc = brk_loc + panel_offset + height_offset
                                                current_loc = panel_detail.Location.Point
                                                translation = new_loc - current_loc
                                                DB.ElementTransformUtils.MoveElement(doc, panel_detail.Id, translation)
                                                group["panels"].append(panel_detail)
                                                matched = True
                                                break
                                if matched:
                                    break
                            if not matched:
                                unmatched_panels.append(panel_detail)

        # 4. Place transformers and assign them similarly.
        unmatched_transformers = []
        for transformer in transformer_components:
            family_name, type_name = PART_TYPE_MAPPING.get(15, (None, None))
            if family_name:
                symbol = ComponentSymbol.get_symbol(doc, family_name, type_name)
                if symbol:
                    symbol.activate(doc)
                    temp_loc = base_point
                    transformer_detail = symbol.place(doc, view, temp_loc)
                    if transformer_detail:
                        transformer.set_parameters(transformer_detail, PARAMETER_MAP,
                                                   transformer.element.Id.IntegerValue,
                                                   transformer_detail.Id.IntegerValue)
                        transformer_name_param = transformer_detail.LookupParameter("Panel Name_CEDT")
                        if transformer_name_param and transformer_name_param.HasValue:
                            transformer_name = transformer_name_param.AsString()
                            matched = False
                            for group in switchboard_groups:
                                for breaker in group["circuits"]:
                                    breaker_param = breaker.LookupParameter("CKT_Load Name_CEDT")
                                    if breaker_param and breaker_param.HasValue:
                                        breaker_load = breaker_param.AsString()
                                        m = re.search(r'"([^"]+)"', breaker_load)
                                        if m:
                                            breaker_load = m.group(1)
                                        if transformer_name in breaker_load:
                                            if hasattr(breaker, "Location") and breaker.Location:
                                                brk_loc = breaker.Location.Point
                                                new_loc = brk_loc + transformer_offset + height_offset
                                                current_loc = transformer_detail.Location.Point
                                                translation = new_loc - current_loc
                                                DB.ElementTransformUtils.MoveElement(doc, transformer_detail.Id, translation)
                                                group["transformers"].append(transformer_detail)
                                                matched = True
                                                break
                                if matched:
                                    break
                            if not matched:
                                unmatched_transformers.append(transformer_detail)

        # 5. Reposition any unmatched panels/transformers to a fixed column.
        if unmatched_panels:
            fixed_x = -10
            base_y = 0
            spacing_y = -2.5
            for i, panel in enumerate(unmatched_panels):
                new_loc = DB.XYZ(fixed_x, base_y + i * spacing_y, panel.Location.Point.Z + 3.5)
                current_loc = panel.Location.Point
                translation = new_loc - current_loc
                DB.ElementTransformUtils.MoveElement(doc, panel.Id, translation)

        if unmatched_transformers:
            fixed_x = -20
            base_y = 0
            spacing_y = -2.5
            for i, transformer in enumerate(unmatched_transformers):
                new_loc = DB.XYZ(fixed_x, base_y + i * spacing_y, transformer.Location.Point.Z + 3.5)
                current_loc = transformer.Location.Point
                translation = new_loc - current_loc
                DB.ElementTransformUtils.MoveElement(doc, transformer.Id, translation)

        # 6. Handle nested groups based on correlated panel/transformer.
        # If a switchboard's "CKT_Panel_CEDT" (panel_key) matches a panel's or transformer's
        # "Panel Name_CEDT", then move that entire switchboard group beneath the correlated element.
        # Additionally, for every group moved, add an extra 3.5 feet offset on the y axis.
        nested_move_offset = 0.0
        for group in switchboard_groups:
            panel_key = group.get("panel_key")
            if not panel_key:
                continue

            correlated_item = None
            # Search in all other groups.
            for other_group in switchboard_groups:
                if other_group == group:
                    continue
                # Check panels.
                for panel in other_group.get("panels", []):
                    panel_param = panel.LookupParameter("Panel Name_CEDT")
                    if panel_param and panel_param.HasValue and panel_param.AsString() == panel_key:
                        correlated_item = panel
                        break
                if correlated_item:
                    break
                # Check transformers.
                for transformer in other_group.get("transformers", []):
                    trans_param = transformer.LookupParameter("Panel Name_CEDT")
                    if trans_param and trans_param.HasValue and trans_param.AsString() == panel_key:
                        correlated_item = transformer
                        break
                if correlated_item:
                    break

            if correlated_item:
                correlated_loc = correlated_item.Location.Point
                group_loc = group["sb_detail"].Location.Point
                # Move the group 5 feet below the correlated element plus an extra offset
                desired_loc = correlated_loc + DB.XYZ(0, -(5 + nested_move_offset), 0)
                translation = desired_loc - group_loc
                DB.ElementTransformUtils.MoveElement(doc, group["sb_detail"].Id, translation)
                for item in group["circuits"] + group["panels"] + group["transformers"]:
                    DB.ElementTransformUtils.MoveElement(doc, item.Id, translation)
                # Increment the cumulative offset by 3.5 feet for the next moved group.
                nested_move_offset += 5

    tag_family_instances(doc, view)

    with revit.Transaction("Set DME_SegmentLength for Circuit Breaker Instances"):
        breaker_instances = DB.FilteredElementCollector(doc)\
                             .OfCategory(DB.BuiltInCategory.OST_DetailComponents)\
                             .WhereElementIsNotElementType()\
                             .ToElements()
        for instance in breaker_instances:
            if hasattr(instance, "Symbol"):
                symbol = instance.Symbol
                if symbol:
                    family_name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                    if family_name_param and family_name_param.HasValue:
                        family_name = family_name_param.AsString()
                        if family_name == "DME-FDR-Circuit Breaker_CED":
                            seg_param = instance.LookupParameter("DME_SegmentLength")
                            if seg_param and not seg_param.IsReadOnly:
                                seg_param.Set(1.5)
                                
if __name__ == "__main__":
    main()
