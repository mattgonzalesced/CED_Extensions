# -*- coding: utf-8 -*-
__doc__ = """Version = 1.0"""

from pyrevit import script, forms
from Snippets._elecutils import get_panel_dist_system, get_compatible_panels, move_circuits_to_panel, \
    get_circuits_from_panel, get_all_panels
from Autodesk.Revit.DB import Electrical, BuiltInCategory, Transaction, ElementId, FilteredElementCollector, \
    BuiltInParameter

import re

# Get the current document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Create output linkable table
output = script.get_output()


def parse_wire_size_string(wire_size_string):
    """Parse the wire size string and extract values based on given rules."""
    parsed_values = {
        'CKT_Wire Size_CEDT': wire_size_string,
        'CKT_Number of Sets_CED': 0,
        'CKT_Number of Wires_CED': 0,
        'CKT_Wire Hot Quantity_CED': 0,
        'CKT_Wire Hot Size_CEDT': '',
        'CKT_Wire Neutral Quantity_CED': 0,
        'CKT_Wire Neutral Size_CEDT': 'n/a',
        'CKT_Wire Ground Quantity_CED': 0,
        'CKT_Wire Ground Size_CEDT': ''
    }

    # Extract the number of sets if "runs" is present
    runs_match = re.search(r'(\d+)\s*runs', wire_size_string)
    if runs_match:
        parsed_values['CKT_Number of Sets_CED'] = int(runs_match.group(1))
    else:
        parsed_values['CKT_Number of Sets_CED'] = 1
    # Split the wire size string by commas to handle multiple sections
    wire_parts = wire_size_string.split(',')

    # Extract the hot wire quantity and size
    if len(wire_parts) > 0:
        hot_match = re.search(r'(\d+)-#([^,]+)', wire_parts[0])
        if hot_match:
            parsed_values['CKT_Wire Hot Quantity_CED'] = int(hot_match.group(1))
            parsed_values['CKT_Wire Hot Size_CEDT'] = hot_match.group(2)

    num_hashes = wire_size_string.count('#')

    if num_hashes == 2:
        # Case with 2 '#': No neutral wire, only ground
        if len(wire_parts) > 1:
            ground_match = re.search(r'(\d+)-#([^,]+)', wire_parts[1])
            if ground_match:
                parsed_values['CKT_Wire Ground Quantity_CED'] = int(ground_match.group(1))
                parsed_values['CKT_Wire Ground Size_CEDT'] = ground_match.group(2)

    elif num_hashes == 3:
        # Case with 3 '#': Includes neutral wire
        if len(wire_parts) > 1:
            neutral_match = re.search(r'(\d+)-#([^,]+)', wire_parts[1])
            if neutral_match:
                parsed_values['CKT_Wire Neutral Quantity_CED'] = int(neutral_match.group(1))
                parsed_values['CKT_Wire Neutral Size_CEDT'] = neutral_match.group(2)

        if len(wire_parts) > 2:
            ground_match = re.search(r'(\d+)-#([^,]+)', wire_parts[2])
            if ground_match:
                parsed_values['CKT_Wire Ground Quantity_CED'] = int(ground_match.group(1))
                parsed_values['CKT_Wire Ground Size_CEDT'] = ground_match.group(2)

    # Prepend the `#` sign only if it's missing
    if parsed_values['CKT_Wire Hot Size_CEDT'] and not parsed_values['CKT_Wire Hot Size_CEDT'].startswith('#'):
        parsed_values['CKT_Wire Hot Size_CEDT'] = "#" + parsed_values['CKT_Wire Hot Size_CEDT']
    if parsed_values['CKT_Wire Neutral Size_CEDT'] != 'n/a' and not parsed_values['CKT_Wire Neutral Size_CEDT'].startswith('#'):
        parsed_values['CKT_Wire Neutral Size_CEDT'] = "#" + parsed_values['CKT_Wire Neutral Size_CEDT']
    if parsed_values['CKT_Wire Ground Size_CEDT'] and not parsed_values['CKT_Wire Ground Size_CEDT'].startswith('#'):
        parsed_values['CKT_Wire Ground Size_CEDT'] = "#" + parsed_values['CKT_Wire Ground Size_CEDT']

    parsed_values['CKT_Number of Wires_CED'] = (
        parsed_values['CKT_Wire Hot Quantity_CED'] +
        parsed_values['CKT_Wire Neutral Quantity_CED']
    )

    return parsed_values


def assign_wire_parameters(circuit_info, doc):
    """Assign parsed wire size values to circuit shared parameters."""
    wire_size_string = circuit_info['wire_size']
    if not wire_size_string or wire_size_string == "N/A":
        return

    circuit = circuit_info['circuit']

    # Check if the circuit is overridden by the user
    user_override_param = circuit.LookupParameter("CKT_User Override_CED")
    if user_override_param and user_override_param.AsInteger() == 1:  # 1 means checked
        print("circuit: {}:{} is overridden by user. Values not changed".format(
            circuit_info['panel'], circuit_info['circuit_number']
        ))
        return  # Stop further processing for this circuit

    # Use the existing parsing function to get the parameter values
    parsed_values = parse_wire_size_string(wire_size_string)

    # Start a transaction to set parameters
    with Transaction(doc, "Sync Wire Info") as trans:
        trans.Start()
        try:
            # Use LookupParameter to set shared parameters
            def set_parameter(circuit, param_name, value):
                param = circuit.LookupParameter(param_name)
                if param and value is not None:
                    param.Set(value)

            set_parameter(circuit, "CKT_Wire Size_CEDT", parsed_values['CKT_Wire Size_CEDT'])
            set_parameter(circuit, "CKT_Number of Sets_CED", parsed_values['CKT_Number of Sets_CED'])
            set_parameter(circuit, "CKT_Number of Wires_CED", parsed_values['CKT_Number of Wires_CED'])
            set_parameter(circuit, "CKT_Wire Hot Quantity_CED", parsed_values['CKT_Wire Hot Quantity_CED'])
            set_parameter(circuit, "CKT_Wire Hot Size_CEDT", parsed_values['CKT_Wire Hot Size_CEDT'])
            set_parameter(circuit, "CKT_Wire Neutral Quantity_CED", parsed_values['CKT_Wire Neutral Quantity_CED'])
            set_parameter(circuit, "CKT_Wire Neutral Size_CEDT", parsed_values['CKT_Wire Neutral Size_CEDT'])
            set_parameter(circuit, "CKT_Wire Ground Quantity_CED", parsed_values['CKT_Wire Ground Quantity_CED'])
            set_parameter(circuit, "CKT_Wire Ground Size_CEDT", parsed_values['CKT_Wire Ground Size_CEDT'])

        except Exception as e:
            output.print_md("**Error updating circuit parameters: {}**".format(str(e)))
        finally:
            trans.Commit()

# Function to create a list with Distribution System, Panel Name, and Element ID
def format_panel_display(panel, doc):
    """Returns a string with the distribution system name, panel name, and element ID for display."""
    panel_data = get_panel_dist_system(panel, doc)
    dist_system_name = panel_data['dist_system_name'] if panel_data['dist_system_name'] else "Unknown Dist. System"
    return "{} - {} (ID: {})".format(panel.Name,dist_system_name, panel.Id)

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


def main():
    # Get all panels in the project
    all_panels = get_all_panels(doc)
    sorted_panels = get_sorted_filtered_panels(all_panels, doc)
    panel_options = [format_panel_display(panel, doc) for panel in sorted_panels]

    # Prompt user to select multiple panels
    selected_panels_display = forms.SelectFromList.show(
        panel_options,
        title="Select Panels",
        prompt="Choose the panels that contain the circuits to move:",
        multiselect=True
    )
    if not selected_panels_display:
        script.exit()

    # Extract Element IDs for the selected panels
    selected_panel_ids = [
        ElementId(int(panel_display.split("(ID: ")[-1].rstrip(")")))
        for panel_display in selected_panels_display
    ]

    selected_panels = [doc.GetElement(panel_id) for panel_id in selected_panel_ids if panel_id]

    if not selected_panels:
        script.exit()

    # Fetch and collect circuits from all selected panels
    all_circuits = {}
    for panel in selected_panels:
        panel_circuits = get_circuits_from_panel(panel, doc)
        all_circuits.update(panel_circuits)

    if not all_circuits:
        script.exit()

    # Sort circuits by panel name and then by start slot
    sorted_circuits = sorted(
        all_circuits.items(),
        key=lambda item: (item[1]['panel'], item[1]['start_slot'])
    )

    # Prepare circuit options for user selection
    circuit_options = [
        "{}:{} - {}".format(info['panel'], info['circuit_number'], info['load_name'])
        for _, info in sorted_circuits
    ]

    selected_circuits = forms.SelectFromList.show(
        circuit_options,
        title="Select Circuits to Sync",
        prompt="Choose circuits to sync wire info:",
        multiselect=True
    )
    if selected_circuits is None:
        script.exit()

    # Filter the circuits based on user selection
    selected_circuit_objects = [
        info for key, info in all_circuits.items()
        if "{}:{} - {}".format(info['panel'], info['circuit_number'], info['load_name']) in selected_circuits
    ]

    # Sync wire info for the selected circuits
    for circuit_info in selected_circuit_objects:
        assign_wire_parameters(circuit_info, doc)

    # Output success message
    output.print_md("**Wire information synced successfully for selected circuits.**")

# Execute the main function
if __name__ == '__main__':
    main()
