# -*- coding: utf-8 -*-
from pyrevit import revit, UI, DB

from Autodesk.Revit.DB import Electrical
from pyrevit import forms
from pyrevit import script
from pyrevit.revit import query
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

# Import reusable utilities
from Snippets._elecutils import get_panel_dist_system

# Get the current document
doc = __revit__.ActiveUIDocument.Document

logger = script.get_logger()


def get_circuits_from_selection(include_electrical_equipment=True):
    """
    Retrieve electrical circuits from the UI selection.

    Args:
        include_electrical_equipment (bool): If False, excludes circuits from families in the electrical equipment category.

    Returns:
        List of selected circuits, optionally filtered based on the category of the families.

    Note: If more than one connector, it only includes the circuit on the primary connector.
    """
    selection = __revit__.ActiveUIDocument.Selection.GetElementIds()
    active_view = HOST_APP.active_view
    circuits = []
    discarded_elements = []

    if isinstance(active_view, Electrical.PanelScheduleView):
        for element_id in selection:
            element = doc.GetElement(element_id)

            if isinstance(element, Electrical.ElectricalSystem):
                circuits.append(element)
            else:
                logger.info("Discarding non-circuit element in PanelScheduleView: %s", element_id)
                discarded_elements.append(element_id)
    else:
        for element_id in selection:
            element = doc.GetElement(element_id)
            logger.info("Processing element: %s", element_id)

            if element.ViewSpecific:
                logger.info("Removing annotation %s", element_id)
                discarded_elements.append(element_id)

            elif isinstance(element, Electrical.ElectricalSystem):
                logger.info("Found electrical system: %s", element_id)
                circuits.append(element)

            elif isinstance(element, DB.FamilyInstance):
                if element.Category.Id == DB.ElementId(
                        DB.BuiltInCategory.OST_ElectricalEquipment) and not include_electrical_equipment:
                    logger.info("Skipping electrical equipment: %s", element_id)
                    discarded_elements.append(element_id)
                    continue

                mep_model = element.MEPModel
                if element.MEPModel:
                    connector_manager = mep_model.ConnectorManager
                    if connector_manager is None:
                        logger.info("No connectors found for MEP model: %s", element_id)
                        discarded_elements.append(element_id)
                    else:
                        connector_iterator = connector_manager.Connectors.ForwardIterator()
                        connector_iterator.Reset()
                        found_primary_circuit = False
                        while connector_iterator.MoveNext():
                            connector = connector_iterator.Current
                            connector_info = connector.GetMEPConnectorInfo()

                            if connector_info and connector_info.IsPrimary:
                                found_primary_circuit = True
                                logger.info("Primary connector found on family instance: %s", element_id)

                                refs = connector.AllRefs
                                if refs.IsEmpty:
                                    logger.info("Element not connected to any circuit: %s", element_id)
                                    discarded_elements.append(element_id)
                                else:
                                    for ref in refs:
                                        ref_owner = ref.Owner
                                        if ref_owner and isinstance(ref_owner, Electrical.ElectricalSystem):
                                            logger.info("Adding circuit from primary connector's owner: %s",
                                                        ref_owner.Id)
                                            circuits.append(ref_owner)
                                            break

                        if not found_primary_circuit:
                            electrical_systems = mep_model.GetElectricalSystems()
                            if electrical_systems:
                                for circuit in electrical_systems:
                                    if element.Category.Id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment):
                                        if circuit.BaseEquipment is None or circuit.BaseEquipment.Id != element_id:
                                            logger.info("Adding feeder circuit: %s", circuit.Id)
                                            circuits.append(circuit)
                                        else:
                                            logger.info("Omitting branch circuit: %s", circuit.Id)
                                    else:
                                        logger.info("Adding circuit from family: %s", circuit.Id)
                                        circuits.append(circuit)

    if not circuits:
        logger.info("No circuits found. Exiting script.")
        forms.alert(title="No Circuits Found", msg="No Circuits found from selection. Click OK to exit script.")
        script.exit()

    if discarded_elements:
        logger.info("{} incompatible element(s) have been discarded from selection:{}".format(len(discarded_elements),
                                                                                              discarded_elements))
    return circuits


# Helper function to get circuit data (Voltage and Number of Poles)
def get_circuit_data(circuit):
    """Returns a dictionary containing the number of poles and voltage for the circuit."""
    circuit_data = {
        'poles': None,
        'voltage': None
    }

    poles_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
    if poles_param and poles_param.HasValue:
        circuit_data['poles'] = poles_param.AsInteger()

    voltage_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
    if voltage_param and voltage_param.HasValue:
        circuit_data['voltage'] = voltage_param.AsDouble()

    return circuit_data


# Helper function to get panel's distribution system and voltage capacity
def get_panel_data(panel):
    """Returns a dictionary with the panel's distribution system voltage and phase."""
    panel_data = {
        'dist_system_name': None,
        'phase': None,
        'lg_voltage': None,
        'll_voltage': None
    }

    secondary_dist_system_param = panel.get_Parameter(DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS)
    if secondary_dist_system_param and secondary_dist_system_param.HasValue:
        dist_system_id = secondary_dist_system_param.AsElementId()
    else:
        dist_system_param = panel.get_Parameter(DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
        if dist_system_param and dist_system_param.HasValue:
            dist_system_id = dist_system_param.AsElementId()
        else:
            return panel_data

    dist_system_type = doc.GetElement(dist_system_id)

    if dist_system_type and isinstance(dist_system_type, Electrical.DistributionSysType):
        if hasattr(dist_system_type, 'Name'):
            panel_data['dist_system_name'] = dist_system_type.Name
        else:
            panel_data['dist_system_name'] = "Unnamed Distribution System"

        panel_data['phase'] = dist_system_type.ElectricalPhase

        lg_voltage = dist_system_type.VoltageLineToGround
        ll_voltage = dist_system_type.VoltageLineToLine

        if lg_voltage:
            lg_voltage_param = lg_voltage.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            panel_data['lg_voltage'] = lg_voltage_param.AsDouble() if lg_voltage_param else None

        if ll_voltage:
            ll_voltage_param = ll_voltage.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            panel_data['ll_voltage'] = ll_voltage_param.AsDouble() if ll_voltage_param else None

    return panel_data

def format_panel_display(panel, doc):
    """Returns a string with the distribution system name, panel name, and element ID for display."""
    panel_data = get_panel_dist_system(panel, doc)
    dist_system_name = panel_data['dist_system_name'] if panel_data['dist_system_name'] else "Unknown Dist. System"
    return "{} - {} (ID: {})".format(panel.Name,dist_system_name, panel.Id)

# Get a list of panels compatible with the selected circuit's poles and voltage requirements
def get_compatible_panels(selected_circuit):
    """Returns a list of compatible panels based on the selected circuit's poles and voltage."""
    circuit_data = get_circuit_data(selected_circuit)
    circuit_poles = circuit_data['poles']
    circuit_voltage = circuit_data['voltage']

    all_panels = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    compatible_panels = []
    for panel in all_panels:
        panel_data = get_panel_dist_system(panel, doc)
        panel_lg_voltage = panel_data['lg_voltage']
        panel_ll_voltage = panel_data['ll_voltage']
        phase = panel_data['phase']

        if circuit_poles == 1 and panel_lg_voltage and abs(panel_lg_voltage - circuit_voltage) < 1.0:
            compatible_panels.append(panel)
        elif circuit_poles >= 2 and panel_ll_voltage and abs(panel_ll_voltage - circuit_voltage) < 1.0:
            compatible_panels.append(panel)

    return compatible_panels


# Function to find open slots in the target panel
def find_open_slots(target_panel):
    """Find available slots in the target panel, prioritizing odd-numbered slots."""
    available_slots = list(range(1, 43))
    odd_slots = [slot for slot in available_slots if slot % 2 == 1]
    even_slots = [slot for slot in available_slots if slot % 2 == 0]
    return odd_slots + even_slots


# Move circuits to the selected target panel using available slots
def move_circuits_to_panel(circuits, target_panel):
    """Move circuits to the target panel using available slots."""
    available_slots = find_open_slots(target_panel)

    if not available_slots:
        forms.alert("No available slots in the target panel.", exitscript=True)

    with DB.Transaction(doc, "Move Circuits to New Panel") as trans:
        trans.Start()
        for i, circuit in enumerate(circuits):
            if i < len(available_slots):
                slot = available_slots[i]
                circuit.SelectPanel(target_panel)
            else:
                forms.alert("Not enough slots in the target panel.", exitscript=True)

        doc.Regenerate()
        trans.Commit()


# Main script logic
def main():
    selected_circuits = get_circuits_from_selection()

    compatible_panels = get_compatible_panels(selected_circuits[0])

    if not compatible_panels:
        forms.alert("No compatible panels found.", exitscript=True)

    panel_options = [format_panel_display(panel,doc) for panel in compatible_panels]

    target_panel_name = forms.SelectFromList.show(
        panel_options,
        title="Select Target Panel",
        prompt="Choose a panel to move the circuits to:",
        multiselect=False
    )

    if not target_panel_name:
        script.exit()

    target_panel_id_str = target_panel_name.split("(ID: ")[-1].rstrip(")")
    target_panel_id = DB.ElementId(int(target_panel_id_str))

    # Find the selected target panel by its ElementId
    target_panel = doc.GetElement(target_panel_id)

    if not target_panel:
        forms.alert("Panel not found.", exitscript=True)

    move_circuits_to_panel(selected_circuits, target_panel)


if __name__ == '__main__':
    main()
