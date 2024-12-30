# -*- coding: utf-8 -*-
__title__ = "Sync Ckt Data to Elements"

from pyrevit import DB, revit, script,forms
from pyrevit.revit.db import query
from Snippets._elecutils import get_all_light_devices, get_all_panels, get_all_elec_fixtures

doc = revit.doc

# Get the output window to log messages
output = script.get_output()

def collect_circuit_data():
    """
    Collect circuit data and determine which elements need to be updated.

    Returns:
        List of dictionaries with elements and their new parameter values.
    """
    # Get all circuits in the project
    circuit_collector = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem).ToElements()

    # Prepare data for updating
    elements_to_update = []

    # Loop through each circuit and get connected elements
    for circuit in circuit_collector:
        # Retrieve necessary circuit parameters using BuiltInParameters
        rating_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
        load_name_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
        ckt_notes_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
        panel_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        circuit_number_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        wire_size_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM)

        # Get the parameter values
        rating = query.get_param_value(rating_param) if rating_param else None
        load_name = query.get_param_value(load_name_param) if load_name_param else None
        circuit_notes = query.get_param_value(ckt_notes_param) if ckt_notes_param else None
        panel = query.get_param_value(panel_param) if panel_param else None
        circuit_number = query.get_param_value(circuit_number_param) if circuit_number_param else None
        wire_size = query.get_param_value(wire_size_param) if wire_size_param else None

        # Get the elements connected to the circuit
        connected_elements = circuit.Elements
        if not connected_elements:
            continue

        # Check if connected elements need to be updated
        for element in connected_elements:
            # Define the parameters to be set on the elements
            element_rating = query.get_param_value(query.get_param(element, "CKT_Rating_CED"))
            element_load_name = query.get_param_value(query.get_param(element, "CKT_Load Name_CEDT"))
            element_notes = query.get_param_value(query.get_param(element, "CKT_Schedule Notes_CEDT"))
            element_panel = query.get_param_value(query.get_param(element, "CKT_Panel_CEDT"))
            element_circuit_number = query.get_param_value(query.get_param(element, "CKT_Circuit Number_CEDT"))
            element_wire_size   =   query.get_param_value((query.get_param(element, "CKT_Wire Size_CEDT")))
            # Prepare data for elements that need to be updated
            updates_needed = {
                'element': element,
                'new_rating': None,
                'new_load_name': None,
                'new_notes': None,
                'new_panel': None,
                'new_circuit_number': None,
                'new_wire_size' : None,
                'update': False
            }

            # Determine if updates are needed based on parameter differences
            if rating and element_rating != rating:
                updates_needed['new_rating'] = rating
                updates_needed['update'] = True

            if load_name and element_load_name != load_name:
                updates_needed['new_load_name'] = load_name
                updates_needed['update'] = True

            if circuit_notes and element_notes != circuit_notes:
                updates_needed['new_notes'] = circuit_notes
                updates_needed['update'] = True

            if panel and element_panel != panel:
                updates_needed['new_panel'] = panel
                updates_needed['update'] = True

            if circuit_number and element_circuit_number != circuit_number:
                updates_needed['new_circuit_number'] = circuit_number
                updates_needed['update'] = True

            if wire_size and element_wire_size != circuit_number:
                updates_needed['new_wire_size'] = wire_size
                updates_needed['update'] = True

            # Add to the list if an update is required
            if updates_needed['update']:
                elements_to_update.append(updates_needed)

    return elements_to_update

def apply_circuit_data_updates(elements_to_update):
    """
    Apply the collected circuit data to elements using a transaction.

    Args:
        elements_to_update (list): List of dictionaries with elements and new parameter values.
    """
    # Execute updates within a transaction
    with revit.Transaction("Sync Circuit Data with Connected Elements"):
        for update in elements_to_update:
            element = update['element']

            # Update the element parameters if new values are present
            if update['new_rating']:
                element_rating_param = query.get_param(element, "CKT_Rating_CED")
                element_rating_param.Set(update['new_rating'])

            if update['new_load_name']:
                element_load_name_param = query.get_param(element, "CKT_Load Name_CEDT")
                element_load_name_param.Set(update['new_load_name'])

            if update['new_notes']:
                element_notes_param = query.get_param(element, "CKT_Schedule Notes_CEDT")
                element_notes_param.Set(update['new_notes'])

            if update['new_panel']:
                element_panel_param = query.get_param(element, "CKT_Panel_CEDT")
                element_panel_param.Set(update['new_panel'])

            if update['new_circuit_number']:
                element_circuit_number_param = query.get_param(element, "CKT_Circuit Number_CEDT")
                element_circuit_number_param.Set(update['new_circuit_number'])

            if update['new_wire_size']:
                element_circuit_number_param = query.get_param(element, "CKT_Wire Size_CEDT")
                element_circuit_number_param.Set(update['new_wire_size'])

def main():
    # Collect data that needs to be updated outside of the transaction
    elements_to_update = collect_circuit_data()
    apply_circuit_data_updates(elements_to_update)

    if __shiftclick__:  # Shift-click to display table
        output.print_table(
            table_data=[
                [
                    output.linkify(update['element'].Id),
                    update['new_rating'] if update['new_rating'] else "No Change",
                    update['new_load_name'] if update['new_load_name'] else "No Change",
                    update['new_notes'] if update['new_notes'] else "No Change",
                    update['new_panel'] if update['new_panel'] else "No Change",
                    update['new_circuit_number'] if update['new_circuit_number'] else "No Change"
                ] for update in elements_to_update
            ],
            title="Circuit Data Sync: Updates Summary",
            columns=["Element ID", "New Rating", "New Load Name", "New Notes", "New Panel", "New Circuit Number"]
        )
    else:  # Standard click to show alert
        forms.alert("{} elements updated.".format(len(elements_to_update)), title="Sync Complete")



if __name__ == "__main__":
    main()
