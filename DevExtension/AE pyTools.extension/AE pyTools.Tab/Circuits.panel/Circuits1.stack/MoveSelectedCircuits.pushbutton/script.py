# -*- coding: utf-8 -*-
__title__ = "Move Selected Circuits"
__doc__ = """Version = 1.5
Date    = 10.10.2024
________________________________________________________________
Description:
Transfer circuits from one panel to another

Note: 
This ignores feeder circuits to panels, and circuits on 
Non-Primary connectors. 

________________________________________________________________
How-To:
Select circuits or elements connected to circuits
Choose a target panel from the drop down list

________________________________________________________________
TODO:
[FEATURE] - Add output to show results 
[FEATURE] - Provide option to include feeder circuits
________________________________________________________________
Last Updates:
- [09.04.2024] v1.0
- [10.10.2024] v1.5
________________________________________________________________
Author: AEvelina"""
#TODO fix issue with selecting circuits instead of families.

# Import required modules
from pyrevit import script, revit, forms, output
from Autodesk.Revit.DB import FilteredElementCollector, Electrical, Transaction, BuiltInCategory, BuiltInParameter, \
    FamilyInstance, ElementId
# Import reusable utilities
from Snippets._elecutils import get_panel_dist_system, get_compatible_panels

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
    circuits = []
    discarded_elements = []
    for element_id in selection:
        element = doc.GetElement(element_id)
        logger.info("Processing element: %s", element_id)

        if element.ViewSpecific:
            logger.info("removing annotation %s", element_id)
            discarded_elements.append(element_id)

        # Check if the element is an electrical system
        elif isinstance(element, Electrical.ElectricalSystem):
                logger.info("Found electrical system: %s", element_id)
                circuits.append(element)

            # If the element is a FamilyInstance
        elif isinstance(element, FamilyInstance):
            # Check if the family is in the electrical equipment category and the toggle is off
            if element.Category.Id == ElementId(
                    BuiltInCategory.OST_ElectricalEquipment) and not include_electrical_equipment:
                logger.info("Skipping electrical equipment: %s", element_id)
                discarded_elements.append(element_id)
                continue  # Skip this electrical equipment family

            mep_model = element.MEPModel
            if element.MEPModel:
                # Check if the MEP model has a valid ConnectorManager and connectors
                connector_manager = mep_model.ConnectorManager
                if connector_manager is None:
                    logger.info("No connectors found for MEP model: %s", element_id)
                    discarded_elements.append(element_id)
                else:
                    # Handle connectors for primary circuits first
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
                        if refs.IsEmpty is True:  # Check if there are no references
                            logger.info("Element not connected to any circuit: %s", element_id)
                            discarded_elements.append(element_id)
                        else:
                            for ref in connector.AllRefs:
                                ref_owner = ref.Owner
                                if ref_owner and isinstance(ref_owner, Electrical.ElectricalSystem):
                                    logger.info("Adding circuit from primary connector's owner: %s", ref_owner.Id)
                                    circuits.append(ref_owner)
                                else:
                                    logger.info("Reference does not belong to an electrical system: %s", element_id)
                                    discarded_elements.append(element_id)

                            break  # Once a primary connector is processed, stop checking other connectors

                    # If no primary connector was found, add circuits directly from the family instance
                    if not found_primary_circuit:
                        electrical_systems = mep_model.GetElectricalSystems()
                        if electrical_systems:
                            for circuit in electrical_systems:
                                # For electrical equipment families, check feeder/branch circuits
                                if element.Category.Id == ElementId(BuiltInCategory.OST_ElectricalEquipment):
                                    if circuit.BaseEquipment is None or circuit.BaseEquipment.Id != element_id:
                                        logger.info("Adding feeder circuit: %s", circuit.Id)
                                        circuits.append(circuit)
                                    else:
                                        logger.info("Omitting branch circuit: %s", circuit.Id)
                                else:
                                    logger.info("Adding circuit from family: %s", circuit.Id)
                                    circuits.append(circuit)

    # If no circuits are found, log and exit
    if not circuits:
        logger.info("No circuits found. Exiting script.")
        forms.alert(title="No Circuits Found", msg="No Circuits found from selection. Click OK to exit script.")
        script.exit()

    if len(discarded_elements) > 0:
        forms.alert(title="Alert",
                    msg="{} incompatible element(s) have been discarded from selection".format(len(discarded_elements)),
                    sub_msg='Click OK to continue, Cancel to exit script.',
                    cancel=True,
                    exitscript=True)
    return circuits


# Helper function to get circuit data (Voltage and Number of Poles)
def get_circuit_data(circuit):
    """Returns a dictionary containing the number of poles and voltage for the circuit."""
    circuit_data = {
        'poles': None,
        'voltage': None
    }

    # Get the number of poles (RBS_ELEC_NUMBER_OF_POLES)
    poles_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
    if poles_param and poles_param.HasValue:
        circuit_data['poles'] = poles_param.AsInteger()

    # Get the voltage (RBS_ELEC_VOLTAGE)
    voltage_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
    if voltage_param and voltage_param.HasValue:
        circuit_data['voltage'] = voltage_param.AsDouble()  # Stored in internal units (e.g., volts)

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

    # Check if the equipment is a transformer by looking for the RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS parameter
    secondary_dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS)
    if secondary_dist_system_param and secondary_dist_system_param.HasValue:
        # If the parameter exists, it's a transformer, so we use the secondary distribution system
        dist_system_id = secondary_dist_system_param.AsElementId()
    else:
        # Otherwise, it's a panel or switchboard, so we use the regular distribution system
        dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
        if dist_system_param and dist_system_param.HasValue:
            dist_system_id = dist_system_param.AsElementId()
        else:
            return panel_data  # No distribution system, return empty panel_data

    dist_system_type = doc.GetElement(dist_system_id)

    # Ensure the retrieved element is a valid DistributionSysType object
    if dist_system_type and isinstance(dist_system_type, Electrical.DistributionSysType):
        # Safely retrieve the distribution system's Name
        if hasattr(dist_system_type, 'Name'):
            panel_data['dist_system_name'] = dist_system_type.Name
        else:
            panel_data['dist_system_name'] = "Unnamed Distribution System"

        # Get the electrical phase (if available)
        panel_data['phase'] = dist_system_type.ElectricalPhase

        # Retrieve voltages if available
        lg_voltage = dist_system_type.VoltageLineToGround
        ll_voltage = dist_system_type.VoltageLineToLine

        # Get the actual voltage values
        if lg_voltage:
            lg_voltage_param = lg_voltage.get_Parameter(BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            panel_data['lg_voltage'] = lg_voltage_param.AsDouble() if lg_voltage_param else None

        if ll_voltage:
            ll_voltage_param = ll_voltage.get_Parameter(BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            panel_data['ll_voltage'] = ll_voltage_param.AsDouble() if ll_voltage_param else None

    return panel_data


# Get a list of panels compatible with the selected circuit's poles and voltage requirements
def get_compatible_panels(selected_circuit):
    """Returns a list of compatible panels based on the selected circuit's poles and voltage."""
    circuit_data = get_circuit_data(selected_circuit)
    circuit_poles = circuit_data['poles']
    circuit_voltage = circuit_data['voltage']

    # Collect all panels in the project
    all_panels = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    # Filter for panels that can accept circuits with the same voltage and poles
    compatible_panels = []
    for panel in all_panels:
        panel_data = get_panel_dist_system(panel, doc)
        panel_lg_voltage = panel_data['lg_voltage']
        panel_ll_voltage = panel_data['ll_voltage']
        phase = panel_data['phase']

        # Check for voltage compatibility
        if circuit_poles == 1 and panel_lg_voltage and abs(panel_lg_voltage - circuit_voltage) < 1.0:
            compatible_panels.append(panel)  # Compatible for single-pole (L-G voltage match)
        elif circuit_poles >= 2 and panel_ll_voltage and abs(panel_ll_voltage - circuit_voltage) < 1.0:
            compatible_panels.append(panel)  # Compatible for multi-pole (L-L voltage match)

    return compatible_panels


# Function to find open slots in the target panel
def find_open_slots(target_panel):
    """Find available slots in the target panel, prioritizing odd-numbered slots."""
    available_slots = list(range(1, 43))  # Example slots (replace with actual logic)
    odd_slots = [slot for slot in available_slots if slot % 2 == 1]
    even_slots = [slot for slot in available_slots if slot % 2 == 0]
    return odd_slots + even_slots


# Move circuits to the selected target panel using available slots
def move_circuits_to_panel(circuits, target_panel):
    """Move circuits to the target panel using available slots."""
    available_slots = find_open_slots(target_panel)

    if not available_slots:
        forms.alert("No available slots in the target panel.", exitscript=True)

    with Transaction(doc, "Move Circuits to New Panel") as trans:
        trans.Start()
        for i, circuit in enumerate(circuits):
            if i < len(available_slots):
                slot = available_slots[i]
                circuit.SelectPanel(target_panel)  # Revit API logic to select panel for the circuit
                # Assign the circuit to the slot (actual Revit logic needed here)
            else:
                forms.alert("Not enough slots in the target panel.", exitscript=True)

        doc.Regenerate()
        trans.Commit()


# Main script logic
def main():
    # Step 1: Get selected circuits
    selected_circuits = get_circuits_from_selection()

    # Step 2: Get a list of compatible panels for the first circuit (for simplicity)
    compatible_panels = get_compatible_panels(selected_circuits[0])

    if not compatible_panels:
        forms.alert("No compatible panels found.", exitscript=True)

    # Step 3: Show a dropdown list to let the user pick a target panel
    panel_names = [panel.Name for panel in compatible_panels]
    target_panel_name = forms.ask_for_one_item(panel_names, title="Select Target Panel",
                                               prompt="Choose a panel to move the circuits to:")

    if not target_panel_name:
        script.exit()

    # Find the selected panel by name
    target_panel = next((panel for panel in compatible_panels if panel.Name == target_panel_name), None)

    if not target_panel:
        forms.alert("Panel not found.", exitscript=True)

    # Step 4: Move the selected circuits to the new panel
    move_circuits_to_panel(selected_circuits, target_panel)


# Execute the main function
if __name__ == '__main__':
    main()
