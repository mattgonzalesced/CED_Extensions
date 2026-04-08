# -*- coding: utf-8 -*-
__title__ = "Create Dedicated Circuits"
__doc__ = '''Working ok. 
- Need to NOT create circuit if the user accidentally makes a bad voltage/ph selection i.e. 120V/2P, 480V/1P
    check out test button 3 for the dist system logic
- The panels list should not include panels (or switches) that DONT have a dist. assigned.
    '''

from pyrevit import forms, script, revit, DB

import Snippets._elecutils as eu

# Start the transaction for modifying the Revit model
doc = revit.doc
uidoc = revit.uidoc

# Initialize output window
output = script.get_output()


# Function to get the distribution system name for a panel
def get_distribution_system_name(panel):
    # Get the parameter containing the distribution system type (RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
    dist_system_param = panel.get_Parameter(DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)

    # Check if the parameter has a value and retrieve the AsValueString
    if dist_system_param and dist_system_param.HasValue:
        dist_system_value = dist_system_param.AsValueString()

        # If AsValueString is populated, return it
        if dist_system_value:
            return dist_system_value


# Step 1: Get selected elements (fixtures) from user selection
selection = revit.get_selection()  # Directly use pyRevit's get_selection() to retrieve selected elements

# Filter selected elements to only those that are valid for electrical systems
fixtures_with_mep_model = [elem for elem in selection if hasattr(elem, 'MEPModel') and elem.MEPModel]

# If no valid fixtures are selected, alert the user
if not fixtures_with_mep_model:
    forms.alert("No valid electrical fixtures selected. Please select fixtures with MEP models.", exitscript=True)



# Step 2: Prompt the user to select a panel using pyRevit forms
# Use pyRevit's query to filter for electrical equipment

panels = eu.get_all_panels(doc)
# If auto-detection fails, prompt the user


# Create a dictionary with Distribution System and Panel Name
panel_dict = {}

for panel in panels:
    # Check if the panel has an MEPModel
    mep_model = panel.MEPModel
    if mep_model:

        try:
            # Try to get the Distribution System Type from the ElectricalEquipment instance
            dist_sys_name = get_distribution_system_name(panel)

            # Create a unique key combining the Distribution System and Panel Name
            panel_key = "[{}] {}".format(dist_sys_name, panel.Name)

            # Ensure the key is unique by adding the Element Id if necessary
            panel_dict[panel_key + " (ID: {})".format(panel.Id)] = panel

        except Exception as e:
            # Output any issues with retrieving the distribution system
            output.print_md("Error accessing Distribution System for panel '{}': {}".format(panel.Name, str(e)))

# Check if any valid panels were found
if not panel_dict:
    output.print_md("No valid panels found.")
    forms.alert("No panels found in the project.", exitscript=True)
else:
    # Display a selection form for the panels
    selected_panel_key = forms.SelectFromList.show(sorted(panel_dict.keys()), title="Select a Panel to Assign Circuits",
                                                   multiselect=False)

    # If no panel was selected, alert the user
    if not selected_panel_key:
        forms.alert("No panel selected. Please select a panel to continue.", exitscript=True)

    # Get the selected panel
    selected_panel = panel_dict[selected_panel_key]

# Step 3: Create circuits for each fixture and assign them to the selected panel
circuit_data = []  # This will hold the data for output (Fixture, Circuit, Panel, Circuit Number, Result)


def _short_error(ex):
    text = str(ex or "").strip()
    if not text:
        return "Unknown error"
    return " ".join(text.splitlines())

with revit.Transaction("Create Circuit and Assign to Panel"):
    for fixture in fixtures_with_mep_model:
        try:
            system = DB.Electrical.ElectricalSystem.Create(
                doc,
                [fixture.Id],
                DB.Electrical.ElectricalSystemType.PowerCircuit
            )
        except Exception as create_ex:
            circuit_data.append([
                output.linkify(fixture.Id),
                "-",
                selected_panel.Name,
                "-",
                "Create failed: {}".format(_short_error(create_ex)),
            ])
            continue

        try:
            # Assign the created circuit to the selected panel
            system.SelectPanel(selected_panel)
            # Get the circuit number and other relevant data
            circuit_number = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
            circuit_data.append([
                output.linkify(fixture.Id),
                output.linkify(system.Id),  # Linkified element ID of the circuit
                selected_panel.Name,         # Panel name
                circuit_number,              # Circuit number
                "Created"
            ])
        except Exception as assign_ex:
            circuit_data.append([
                output.linkify(fixture.Id),
                output.linkify(system.Id),  # Linkified element ID of the circuit
                selected_panel.Name,
                "-",
                "Assign failed: {}".format(_short_error(assign_ex))
            ])

# Step 4: Output success message and table
if circuit_data:
    output.print_md("**Create Dedicated Circuits results:**")
    output.print_table(
        circuit_data,
        ["Fixture ID", "Circuit ID", "Panel", "Circuit Number", "Result"]
    )
else:
    output.print_md("No circuits were created.")


