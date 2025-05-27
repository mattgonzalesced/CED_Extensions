# -*- coding: utf-8 -*-
from pyrevit import revit, UI, DB

from Autodesk.Revit.DB import Electrical
from pyrevit import forms
from pyrevit import script
from pyrevit.revit import query
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.compat import get_elementid_value_func

# Import reusable utilities
from Snippets._elecutils import get_panel_dist_system, get_compatible_panels, move_circuits_to_panel, \
    get_circuits_from_panel, get_all_panels

# Get the current document
doc = __revit__.ActiveUIDocument.Document

logger = script.get_logger()
output = script.get_output()
output.close_others()


def get_sorted_filtered_panels(all_panels, doc):
    """Filters out panels with unknown distribution systems and sorts them by dist system name and panel name."""
    valid_panels = []

    # Filter out panels with "Unknown Dist. System"
    for panel in all_panels:
        panel_data = get_panel_dist_system(panel, doc)
        if panel_data['dist_system_name'] and panel_data['dist_system_name'] != "Unnamed Distribution System":
            valid_panels.append((panel.Name, panel_data['dist_system_name'], panel))

    # Sort panels first by distribution system name, then by panel name
    sorted_panels = sorted(valid_panels, key=lambda x: (x[0], x[1]))

    return [panel for _, _, panel in sorted_panels]


def prompt_for_panel(doc, panels, title, prompt_msg):
    """Prompts the user to select a panel from the given list."""
    sorted_panels = get_sorted_filtered_panels(panels, doc)

    panel_map = {}
    for panel in sorted_panels:
        display_str = format_panel_display(panel, doc)
        panel_map[display_str] = panel

    selected_display = forms.SelectFromList.show(
        panel_map.keys(),
        title=title,
        prompt=prompt_msg,
        multiselect=False
    )

    if not selected_display:
        script.exit()

    return panel_map[selected_display]


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
    seen_ids = set()

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
        logger.info("{} incompatible element(s) have been discarded from selection:{}".format(
            len(discarded_elements),
            discarded_elements)
        )
    unique_circuits = []
    seen_ids = set()
    for circuit in circuits:
        get_id_val = get_elementid_value_func()
        cid = get_id_val(circuit.Id)
        if cid not in seen_ids:
            unique_circuits.append(circuit)
            seen_ids.add(cid)

    return unique_circuits


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


def format_panel_display(panel, doc):
    """Returns a string with the distribution system name, panel name, and element ID for display."""
    panel_data = get_panel_dist_system(panel, doc)
    dist_system_name = panel_data['dist_system_name'] if panel_data['dist_system_name'] else "Unknown Dist. System"
    return "{} - {} (ID: {})".format(panel.Name, dist_system_name, panel.Id)


# Function to find open slots in the target panel
def find_open_slots(target_panel):
    """Find available slots in the target panel, prioritizing odd-numbered slots."""
    available_slots = list(range(1, 43))
    odd_slots = [slot for slot in available_slots if slot % 2 == 1]
    even_slots = [slot for slot in available_slots if slot % 2 == 0]
    return odd_slots + even_slots


# Main script logic
def main():
    selected_circuits = get_circuits_from_selection()
    all_panels = get_all_panels(doc)

    compatible_panels = []
    for circuit in selected_circuits:
        compatible_panels.extend(get_compatible_panels(circuit, all_panels, doc))

    if not compatible_panels:
        forms.alert("No compatible panels found.", exitscript=True)

    compatible_panels = list(set(compatible_panels))

    # Sort and filter compatible panels
    # Prompt the user to select a target panel
    target_panel = prompt_for_panel(
        doc,
        compatible_panels,
        title="Select Target Panel",
        prompt_msg="Choose the target panel to move the circuits to:"
    )

    if not target_panel:
        script.exit()

    if not target_panel:
        forms.alert("Panel not found.", exitscript=True)

    try:
        circuit_data = move_circuits_to_panel(selected_circuits, target_panel, doc, output)
    except Exception as e:
        output.print_md("**Error occurred while transferring circuits: {}**".format(str(e)))
        return

    # Step 6: Output success message and table
    output.print_md("**Circuits transferred successfully.**")
    output.print_table(circuit_data, ["Circuit ID", "Previous Circuit", "New Circuit"])


if __name__ == '__main__':
    main()
