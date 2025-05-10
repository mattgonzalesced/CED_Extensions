# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script, output
from pyrevit.revit import query
from collections import defaultdict

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()
FAMILY_NAME = "EF-U_Refrig Power Connector-Balanced_CED-WM"
CIRCUIT_PARAM = "Refrigeration Circuit Number_CEDT"
TYPE_PARAM = "Type Name"  # Assuming query.get_name(instance.Symbol) gets the type name

def collect_connectors():
    collector = DB.FilteredElementCollector(doc) \
        .OfClass(DB.FamilyInstance) \
        .WhereElementIsNotElementType()

    grouped_connectors = defaultdict(list)
    for inst in collector:
        if inst.Symbol.Family.Name != FAMILY_NAME:
            continue
        ckt_number = inst.LookupParameter(CIRCUIT_PARAM)
        if not ckt_number or not ckt_number.HasValue:
            continue
        key = ckt_number.AsString().strip()
        grouped_connectors[key].append(inst)
    return grouped_connectors

def select_connector_ids(grouped_connectors):
    circuit_options = sorted(grouped_connectors.keys())
    selected_circuits = forms.SelectFromList.show(
        circuit_options,
        title="Select Refrigeration Circuits",
        multiselect=True
    )
    if not selected_circuits:
        return []

    # Build type map from selected circuits
    type_map = defaultdict(list)
    for ckt in selected_circuits:
        for inst in grouped_connectors[ckt]:
            typename = query.get_name(inst.Symbol)
            type_map[typename].append(inst)

    type_choices = sorted(type_map.keys())
    selected_types = forms.SelectFromList.show(
        type_choices,
        title="Select Connector Types",
        multiselect=True
    )
    if not selected_types:
        return []

    return [inst.Id for tname in selected_types for inst in type_map[tname]]


def main():
    grouped_connectors = collect_connectors()
    if not grouped_connectors:
        forms.alert("No matching family instances found.")
        return

    ids = select_connector_ids(grouped_connectors)

    if not ids:
        logger.info("No Ids Selected")
        script.exit()

    revit.get_selection().set_to(ids)
    logger.info("✅ Found {} matching connectors:".format(len(ids)))
    for i in ids:
        logger.info(" - ElementId: `{}`".format(i.IntegerValue))


if __name__ == "__main__":
    main()
