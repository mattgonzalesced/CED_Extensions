# -*- coding: utf-8 -*-
__title__ = "Batch Swap Circuits"
__doc__ = """Version = 1.1"""
from Autodesk.Revit.DB import Electrical, BuiltInCategory, ElementId
# TODO FIX SORTING ISSUE
from pyrevit import script, forms

from Snippets._elecutils import get_panel_dist_system, get_compatible_panels, move_circuits_to_panel, \
    get_circuits_from_panel, get_all_panels

# Get the current document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Create output linkable table
output = script.get_output()


# Function to auto-detect selected panel or circuit's panel
def auto_detect_starting_panel():
    """Auto-detects the selected panel or the panel associated with a selected circuit."""
    selection = uidoc.Selection.GetElementIds()

    if selection:
        for element_id in selection:
            element = doc.GetElement(element_id)
            # If it's an electrical system (circuit)
            if isinstance(element, Electrical.ElectricalSystem):
                return element.BaseEquipment
            # If it's a panel (FamilyInstance)
            if element.Category.Id == ElementId(BuiltInCategory.OST_ElectricalEquipment):
                return element
    return None


def format_panel_display(panel, doc):
    """Returns a string with the distribution system name, panel name, and element ID for display."""
    panel_data = get_panel_dist_system(panel, doc)
    dist_system_name = panel_data['dist_system_name'] if panel_data['dist_system_name'] else "Unknown Dist. System"
    return "{} - {} (ID: {})".format(panel.Name,dist_system_name, panel.Id)


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


# Function to sort and filter the panels by distribution system name and panel name
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


# Main script logic
def main():
    # Auto-detect starting panel
    starting_panel = auto_detect_starting_panel()

    # Get all panels in the project
    panels = get_all_panels(doc)

    if not starting_panel:
        starting_panel = prompt_for_panel(
            doc,
            panels,
            title="Select Starting Panel",
            prompt_msg="Choose the panel that contains the circuits to move:"
        )

    while True:  # Keep asking for a panel until circuits are found or user cancels
        # Step 2: Get the circuits from the starting panel
        circuits = get_circuits_from_panel(starting_panel, doc, 0)

        if circuits:
            break  # Exit the loop if circuits are found

        # Show alert with Retry and Cancel options
        retry = forms.alert(
            "No circuits found in the selected panel. Would you like to retry or cancel?",
            title="No Circuits Found",
            retry=True,
            cancel=True,
            ok=False,
            exitscript=False
        )

        if not retry:
            script.exit()  # Exit the script if user cancels

        # If Retry is selected, prompt for a new starting panel
        starting_panel = prompt_for_panel(
            doc,
            panels,
            title="Select Starting Panel",
            prompt_msg="Choose the panel that contains the circuits to move:"
        )

    # Step 3: Sort circuits by odd and even start slots
    circuit_options = sorted(
        circuits,
        key=lambda circuit: ((circuit['start_slot'] or 0) % 2 == 0, circuit['start_slot'])
        # Odd first, then sort by start slot
    )

    # Format sorted circuits for display
    circuit_options = [
        "{} - {}".format(circuit['circuit_number'], circuit['load_name']) for circuit in circuit_options
    ]

    # Display circuits in checkboxes and get user selection
    selected_circuits = forms.SelectFromList.show(
        circuit_options,
        title="Select Circuits to Move (Starting Panel: {})".format(starting_panel.Name),
        multiselect=True
    )

    if selected_circuits is None:
        script.exit()  # User closed the window

    # Step 4: Map selected descriptions back to circuit objects using the circuit data
    selected_circuit_objects = [
        circuit['circuit'] for circuit in circuits
        if "{} - {}".format(circuit['circuit_number'], circuit['load_name']) in selected_circuits
    ]

    # Step 4: Get compatible panels based on the selected circuits
    compatible_panels = []
    for circuit in selected_circuit_objects:
        compatible_panels.extend(get_compatible_panels(circuit, panels, doc))

    # Use a set to remove duplicate panels
    compatible_panels = list(set(compatible_panels))

    if not compatible_panels:
        script.exit()  # No compatible panels found

    # Sort and filter compatible panels
    # Prompt the user to select a target panel
    target_panel = prompt_for_panel(
        doc,
        compatible_panels,
        title="Select Target Panel (Starting Panel: {})".format(starting_panel.Name),
        prompt_msg="Choose the target panel to move the circuits to:"
    )

    if not target_panel:
        script.exit()

    # Step 5: Move the selected circuits to the target panel and store data for final output
    try:
        circuit_data = move_circuits_to_panel(selected_circuit_objects, target_panel, doc, output)
    except Exception as e:
        output.print_md("**Error occurred while transferring circuits: {}**".format(str(e)))
        return

    # Step 6: Output success message and table
    output.print_md("**Circuits transferred successfully.**")
    output.print_table(circuit_data, ["Circuit ID", "Previous Circuit", "New Circuit"])


# Execute the main function
if __name__ == '__main__':
    main()
