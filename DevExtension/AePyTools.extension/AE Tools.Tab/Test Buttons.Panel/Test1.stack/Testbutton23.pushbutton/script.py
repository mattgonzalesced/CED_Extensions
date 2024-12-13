# -*- coding: utf-8 -*-
__title__ = "List Circuits by Panel"
from pyrevit import script, revit, DB, forms
from Snippets._elecutils import get_panel_dist_system

# Initialize output
output = script.get_output()

# --- Second table: Circuits filtered by selected Panels ---

# Create a second output section within the same output window
output.set_title("Circuit Table")

# Get the document
doc = revit.doc

# Function to display panels with Distribution System info
def format_panel_display(panel, doc):
    """Returns a string with the distribution system name, panel name, and element ID for display."""
    panel_data = get_panel_dist_system(panel, doc)
    dist_system_name = panel_data['dist_system_name'] if panel_data and panel_data['dist_system_name'] else "Unknown Dist. System"
    return "{} - {} (ID: {})".format( panel.Name,dist_system_name, panel.Id)


# Function to get circuits from a panel (based on your provided approach)
def get_circuits_from_panel(panel, doc, include_spares=True):
    """Get circuits associated with a selected panel, with an option to include or exclude spare/space circuits."""
    circuits = []
    # Collect all electrical circuits (ElectricalSystem objects)
    panel_circuits = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem).ToElements()

    for circuit in panel_circuits:
        # Check if the circuit's BaseEquipment matches the selected panel
        if circuit.BaseEquipment and circuit.BaseEquipment.Id == panel.Id:
            if not include_spares and circuit.CircuitType in [DB.Electrical.CircuitType.Spare, DB.Electrical.CircuitType.Space]:
                continue

            # Get circuit number and load name
            circuit_number = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
            load_name = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).AsString()

            circuits.append((circuit.Id, load_name, panel.Name, circuit_number))

    # Sort circuits by circuit number
    circuits_sorted = sorted(circuits, key=lambda x: (x[2], x[3]))  # Sorting by panel name, then by circuit number

    return circuits_sorted


# Main script logic
def main():
    # Collect all panels
    panel_collector = DB.FilteredElementCollector(doc)\
                        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)\
                        .WhereElementIsNotElementType().ToElements()  # Convert to list of elements

    # Get sorted panels
    sorted_panels = panel_collector  # We no longer need to sort here since the panel selection list will be human-readable

    # Create a list of formatted panel display names for the user selection
    panel_options = [format_panel_display(panel, doc) for panel in sorted_panels]

    # Show a dropdown list for the user to select panels
    selected_panel_options = forms.SelectFromList.show(
        panel_options,
        title="Select Panels",
        prompt="Select Panels",
        multiselect=True
    )

    if not selected_panel_options:
        output.print_md("No panels selected. Exiting...")
        return

    # Convert the selected panel options back to panel elements
    selected_panels = [panel for panel in sorted_panels if format_panel_display(panel, doc) in selected_panel_options]

    # List to store all matching circuits
    all_circuits = []

    # Iterate over selected panels and get circuits for each panel
    for panel in selected_panels:
        circuits = get_circuits_from_panel(panel, doc)
        all_circuits.extend(circuits)

    # Sort the circuits by panel name, then by circuit number
    sorted_circuits = sorted(all_circuits, key=lambda x: (x[2], x[3]))

    # Print the second table with circuits
    if sorted_circuits:
        output.print_table(
            table_data=[["Circuit ID", "Load Name", "Panel", "Circuit Number"]] +
                       [[output.linkify(c_id), load_name, panel, circuit_number] for c_id, load_name, panel, circuit_number in sorted_circuits],
            title="Filtered Circuits on Selected Panels",
            columns=["Circuit ID", "Load Name", "Panel", "Circuit Number"]
        )
    else:
        output.print_md("No circuits found for the selected panels.")


# Execute the main function
if __name__ == '__main__':
    main()
