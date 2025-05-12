# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script, output
from pyrevit.revit import query
from collections import defaultdict
import Autodesk.Revit.DB.Electrical as DBE

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

# TODO: FIGURE OUT HOW TO CLEANLY LET IT USE OTHER FAMILIES.  
FAMILY_NAME_PLACEHOLDER = "EF-F_Existing Ckt Placeholder-Unbalanced_CED"
CONNECTOR_FAMILY_NAME = "EF-U_Refrig Power Connector-Balanced_CED-WM"
CIRCUIT_PARAM = "Refrigeration Circuit Number_CEDT"
TYPE_PARAM = "Type Name"  # Assuming query.get_name(instance.Symbol) gets the type name

def collect_connectors():
    collector = DB.FilteredElementCollector(doc) \
        .OfClass(DB.FamilyInstance) \
        .WhereElementIsNotElementType()

    grouped_connectors = defaultdict(list)
    for inst in collector:
        if inst.Symbol.Family.Name != CONNECTOR_FAMILY_NAME:
            continue
        ckt_number = inst.LookupParameter(CIRCUIT_PARAM)
        if not ckt_number or not ckt_number.HasValue:
            continue
        key = ckt_number.AsString().strip()
        grouped_connectors[key].append(inst)
    return grouped_connectors

def filter_and_select_connectors_by_type(grouped_connectors, circuit):
    # Get voltage and poles from the original circuit
    circuit_voltage = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE).AsDouble()
    circuit_poles = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES).AsInteger()
    voltage_unit = DB.UnitTypeId.Volts  # Use Revit-native unit type ID

    def is_compatible(inst):
        try:

            inst_voltage = inst.LookupParameter("Voltage_CED").AsDouble()
            inst_poles = inst.LookupParameter("Number of Poles_CED").AsInteger()

            if abs(inst_voltage - circuit_voltage) < 1e-4 and inst_poles == circuit_poles:
                return True

        except:
            return False

    # Prompt user to select circuit numbers
    circuit_keys = sorted(grouped_connectors.keys())
    selected_circuits = forms.SelectFromList.show(
        circuit_keys,
        title="Select Refrigeration Circuits",
        multiselect=True
    )
    if not selected_circuits:
        return []

    # Group matching instances by type name
    type_group_map = defaultdict(list)

    for ckt in selected_circuits:
        for inst in grouped_connectors[ckt]:
            if not is_compatible(inst):
                continue
            type_name = query.get_name(inst.Symbol)
            circuit_tag = inst.LookupParameter("Refrigeration Circuit Number_CEDT").AsString().strip()
            type_group_map[type_name].append((inst, circuit_tag))

    if not type_group_map:
        forms.alert("No compatible connector types found.")
        return []

    # Build labeled options
    label_map = {}
    for type_name, inst_data in type_group_map.items():
        sample = inst_data[0][0]
        inst_voltage = sample.LookupParameter("Voltage_CED").AsValueString()
        inst_poles = sample.LookupParameter("Number of Poles_CED").AsInteger()
        tags = sorted(set(c for _, c in inst_data))
        label = "{} ({}/{}P) ({})".format(type_name, inst_voltage, inst_poles, ", ".join(tags))
        label_map[label] = [inst.Id for inst, _ in inst_data]

    # Prompt user to select connector types
    selected_types = forms.SelectFromList.show(
        sorted(label_map.keys()),
        title="Select Connector Types ({}V / {}P)".format(
            int(DB.UnitUtils.ConvertFromInternalUnits(circuit_voltage, voltage_unit)),
            circuit_poles
        ),
        multiselect=True
    )
    if not selected_types:
        return []

    return [eid for label in selected_types for eid in label_map[label]]





def main():
    # Validate view
    if not isinstance(doc.ActiveView, DBE.PanelScheduleView):
        forms.alert("This tool only works in a Panel Schedule View.")
        return

    selection = revit.get_selection().elements
    if not selection:
        forms.alert("Please select a circuit cell in the panel schedule.")
        return

    # Filter for electrical circuits
    circuits = []

    for el in selection:
        if isinstance(el, DB.Electrical.ElectricalSystem):
            circuits.append(el)

    if len(circuits) != 1:
        forms.alert("You must select exactly one circuit.")
        return

    circuit = circuits[0]
    connected_elements = list(circuit.Elements)
    placeholder = None

    for el in connected_elements:

        if isinstance(el, DB.FamilyInstance) and el.Symbol.Family.Name == FAMILY_NAME_PLACEHOLDER:
            placeholder = el
            break

    if not placeholder:
        forms.alert("No placeholder family instance found in selected circuit.")
        return

    grouped_connectors = collect_connectors()
    if not grouped_connectors:
        forms.alert("No matching connectors found in model.")
        return

    selected_ids = filter_and_select_connectors_by_type(grouped_connectors, circuit)

    if not selected_ids:
        return

    # Try to add to circuit
    try:
        with revit.Transaction("Replace Existing Circuit"):
            component_set = DB.ElementSet()
            for eid in selected_ids:
                component_set.Insert(doc.GetElement(eid))
            success = circuit.AddToCircuit(component_set)

            if not success:
                raise Exception("AddToCircuit returned False.")

            # Delete placeholder
            doc.Delete(placeholder.Id)

    except Exception as e:
        forms.alert("Error adding to circuit: {}".format(str(e)))
        return

    output.print_md("✅ Placeholder circuit replaced successfully.")
    output.print_md("🧩 Added {} connectors.".format(len(selected_ids)))
    output.print_md("🗑️ Placeholder instance deleted.")

if __name__ == "__main__":
    main()

