# -*- coding: utf-8 -*-
from pyrevit import forms, revit, DB

# Mapping for Electrical Equipment Part Type values to Detail Items
part_type_mapping = {
    14: ("DME-EQU-Panel-Top_CED", "Solid"),
    15: ("DME-EQU-Transformer-Box-Ground-Top_CED", "Solid"),
    16: ("DME-EQU-Switchboard-Top_CED", "Solid"),
    17: ("DME-EQU-Panel-Top_CED", "Solid"),
    18: ("DME-EQU-Disconnect-Top_CED", "Fused Disconnect")
}
def get_family_part_type(family_instance):
    if not family_instance or not isinstance(family_instance, DB.FamilyInstance):
        return None

    symbol = family_instance.Symbol
    if not symbol:
        return None

    family = symbol.Family
    if not family:
        return None

    part_type_param = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
    if part_type_param and part_type_param.HasValue:
        return part_type_param.AsInteger()
    return None

def create_detail_item(doc, view, family_name, type_name, location):
    collector = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_DetailComponents).WhereElementIsElementType()

    detail_symbol = None
    for symbol in collector:
        if isinstance(symbol, DB.FamilySymbol):
            family_name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            symbol_family_name = family_name_param.AsString() if family_name_param else None

            if symbol_family_name == family_name and symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == type_name:
                detail_symbol = symbol
                break

    if not detail_symbol:
        return None

    if not detail_symbol.IsActive:
        try:
            detail_symbol.Activate()
            doc.Regenerate()
        except:
            return None

    try:
        return doc.Create.NewFamilyInstance(location, detail_symbol, view)
    except:
        return None

def collect_electrical_equipment(doc):
    return DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType() \
        .ToElements()

def set_parameters(equipment, detail_item, equipment_id, detail_id):
    component_id_param = equipment.LookupParameter("SLD_Component ID_CED")
    symbol_id_param = equipment.LookupParameter("SLD_Symbol ID_CED")
    if component_id_param:
        component_id_param.Set(str(equipment_id))
    if symbol_id_param:
        symbol_id_param.Set(str(detail_id))

    if detail_item:
        component_id_param = detail_item.LookupParameter("SLD_Component ID_CED")
        symbol_id_param = detail_item.LookupParameter("SLD_Symbol ID_CED")
        if component_id_param:
            component_id_param.Set(str(equipment_id))
        if symbol_id_param:
            symbol_id_param.Set(str(detail_id))

def main():
    doc = revit.doc
    view = revit.active_view

    if not isinstance(view, DB.ViewDrafting):
        forms.alert("Please run this script in a Drafting View.")
        return

    equipment_elements = collect_electrical_equipment(doc)

    start_point = DB.XYZ(0, 0, 0)
    offset = DB.XYZ(10, 0, 0)

    with revit.Transaction("Place Detail Items for Electrical Equipment"):
        for idx, equipment in enumerate(equipment_elements):
            part_type = get_family_part_type(equipment)
            panel_name = equipment.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)

            if part_type and part_type in part_type_mapping:
                family_name, type_name = part_type_mapping[part_type]
                location = start_point + (idx * offset)

                detail_item = create_detail_item(doc, view, family_name, type_name, location)

                if detail_item:
                    set_parameters(equipment, detail_item, equipment.Id.IntegerValue, detail_item.Id.IntegerValue)

# Run the script
if __name__ == "__main__":
    main()
