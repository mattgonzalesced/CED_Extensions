# -*- coding: utf-8 -*-
__title__ = "Batch Swap Circuits"
__doc__ = """Version = 1.1"""

from pyrevit import script, forms
from Snippets._elecutils import get_panel_dist_system, get_compatible_panels, move_circuits_to_panel, \
    get_circuits_from_panel, get_all_panels
from Autodesk.Revit.DB import Electrical, BuiltInCategory, Transaction, ElementId, FilteredElementCollector, \
    BuiltInParameter

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
    all_panels = get_all_panels(doc)

    # If a starting panel wasn't auto-detected, prompt the user to select one
    if starting_panel:
        starting_panel_name = starting_panel.Name
    else:
        # Sort and filter panels
        sorted_panels = get_sorted_filtered_panels(all_panels, doc)
        panel_options = [format_panel_display(panel, doc) for panel in sorted_panels]

        # Prompt user to select the starting panel
        selected_panel_display = forms.SelectFromList.show(
            panel_options,
            title="Select Starting Panel",
            prompt="Choose the panel that contains the circuits to move:",
            multiselect=False
        )

        if not selected_panel_display:
            script.exit()

        # Extract the Element ID from the selected display option
        selected_panel_id_str = selected_panel_display.split("(ID: ")[-1].rstrip(")")
        selected_panel_id = ElementId(int(selected_panel_id_str))

        # Find the selected starting panel by its ElementId
        starting_panel = doc.GetElement(selected_panel_id)

    if not starting_panel:
        script.exit()

    # Step 2: Get the circuits from the starting panel
    circuits = get_circuits_from_panel(starting_panel, doc)
    if not circuits:
        script.exit()

    # Step 3: Display circuits in checkboxes and get user selection
    circuit_options = ["{} - {}".format(info['circuit_number'], info['load_name']) for _, info in circuits.items()]

    selected_circuits = forms.SelectFromList.show(
        circuit_options,
        title="Select Circuits to Move (Starting Panel: {})".format(starting_panel.Name),
        multiselect=True
    )

    if selected_circuits is None:
        script.exit()  # User closed the window

    # Step 4: Map selected descriptions back to circuit objects using the dictionary
    selected_circuit_objects = [
        info['circuit'] for key, info in circuits.items()
        if "{} - {}".format(info['circuit_number'], info['load_name']) in selected_circuits
    ]

    # Step 4: Get compatible panels based on the selected circuits
    compatible_panels = []
    for circuit in selected_circuit_objects:
        compatible_panels.extend(get_compatible_panels(circuit, all_panels, doc))

    # Use a set to remove duplicate panels
    compatible_panels = list(set(compatible_panels))

    if not compatible_panels:
        script.exit()  # No compatible panels found

    # Sort and filter compatible panels
    sorted_compatible_panels = get_sorted_filtered_panels(compatible_panels, doc)
    compatible_panel_options = [format_panel_display(panel, doc) for panel in sorted_compatible_panels]

    # Prompt the user to select a target panel (with distribution system, panel name, and ID)
    target_panel_display = forms.SelectFromList.show(
        compatible_panel_options,
        title="Select Target Panel (Starting Panel: {})".format(starting_panel.Name),
        prompt="Choose the target panel to move the circuits to:",
        multiselect=False
    )

    if not target_panel_display:
        script.exit()

    # Extract the Element ID from the selected display option
    target_panel_id_str = target_panel_display.split("(ID: ")[-1].rstrip(")")
    target_panel_id = ElementId(int(target_panel_id_str))

    # Find the selected target panel by its ElementId
    target_panel = doc.GetElement(target_panel_id)

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
