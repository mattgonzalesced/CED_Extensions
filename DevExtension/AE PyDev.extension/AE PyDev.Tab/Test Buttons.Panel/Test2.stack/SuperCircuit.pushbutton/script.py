# -*- coding: utf-8 -*-
__title__ = "SUPER CIRCUIT"

from pyrevit import DB, revit
from pyrevit.revit.db import transaction, query
from pyrevit import script
import re
from Snippets._elecutils import get_all_light_devices, get_all_panels, get_all_elec_fixtures
from System.Collections.Generic import List  # Import .NET List for IList compatibility


doc = revit.doc
uidoc = revit.uidoc

# Helper function to extract the first number from a circuit string
def get_first_number_from_circuit(circuit_str):
    if not circuit_str:
        return float('inf')  # Return a high value to sort None or empty circuit numbers last
    first_part = circuit_str.split(",")[0]
    match = re.search(r'\d+', first_part)
    return int(match.group()) if match else float('inf')

# Function to create an electrical system and select a panel
def create_electrical_system(doc, element_ids, system_type, panel_element):
    if not element_ids or len(element_ids) == 0:
        return None

    element_id_list = List[DB.ElementId](element_ids)
    new_system = DB.Electrical.ElectricalSystem.Create(doc, element_id_list, system_type)
    if new_system and panel_element:
        new_system.SelectPanel(panel_element)
        doc.Regenerate()
    return new_system

# Function to organize elements by Panel and Circuit Number
def group_elements_by_circuit(elements, panel_elements):
    grouped_dict = {}
    unnamed_group = {}
    unnamed_counter = 1

    # Create a lookup dictionary for quick panel name to element mapping
    panel_lookup = {query.get_param_value(query.get_param(panel, "Panel Name")): panel for panel in panel_elements}

    for element in elements:
        # Retrieve parameters using query functions
        panel_param = query.get_param(element, "ckt-Panel")
        circuit_param = query.get_param(element, "ckt-Circuit Number")
        rating_param = query.get_param(element, "ckt-Rating")
        load_name_param = query.get_param(element, "ckt-Load Name")
        ckt_notes_param = query.get_param(element, "ckt-Schedule Notes")

        # Get actual parameter values
        panel_name = query.get_param_value(panel_param)
        circuit_number = query.get_param_value(circuit_param)
        rating = query.get_param_value(rating_param)
        load_name = query.get_param_value(load_name_param)
        ckt_notes = query.get_param_value(ckt_notes_param)

        # Skip elements without valid `ckt-Panel` and `ckt-Circuit Number` entirely
        if not panel_name and not circuit_number:
            continue

        # Handle elements with "<unnamed>" circuit numbers individually
        if circuit_number == "<unnamed>" and not panel_name:
            key = "<unnamed>{}".format(unnamed_counter)
            unnamed_group[key] = {
                "elements": [element],
                "element_ids": [element.Id],
                "panel_name": panel_name,
                "panel_element": None,
                "circuit_number": circuit_number,
                "rating": rating,
                "load_name": load_name,
                "ckt_notes": ckt_notes
            }
            unnamed_counter += 1
            continue

        # Find the corresponding panel element using the lookup dictionary
        panel_element = panel_lookup.get(panel_name)

        # Create a unique key using panel and circuit number
        key = (panel_name, circuit_number)

        # Group elements by key
        if key not in grouped_dict:
            grouped_dict[key] = {
                "elements": [],
                "element_ids": [],
                "panel_name": panel_name,
                "panel_element": panel_element,
                "circuit_number": circuit_number,
                "rating": rating,
                "load_name": load_name,
                "ckt_notes": ckt_notes
            }

        # Ensure ElementId is collected
        grouped_dict[key]["elements"].append(element)
        grouped_dict[key]["element_ids"].append(element.Id)

    # Sort the keys by panel name, then by the first number in the circuit number
    sorted_keys = sorted(
        grouped_dict.keys(),
        key=lambda k: (
            k[0],  # Panel Name
            get_first_number_from_circuit(k[1]) % 2 == 0,  # True for even, False for odd (prioritizes odd)
            get_first_number_from_circuit(k[1])  # Sort numerically after odd/even is prioritized
        )
    )

    # Return sorted list of groups and unnamed group separately for easier processing
    return [(key, grouped_dict[key]) for key in sorted_keys], unnamed_group

def main():

    # Collectors for the elements
    ee_collector = list(get_all_panels(doc))  # Panels
    ef_collector = list(get_all_elec_fixtures(doc))  # Electrical Fixtures
    ld_collector = list(get_all_light_devices(doc))  # Lighting Devices

    selection = revit.get_selection()
    # Combine all elements that need circuiting
    if not selection:
        elements_to_circuit = ef_collector + ld_collector + ee_collector
    else:
        elements_to_circuit = selection

    # Group elements by panel and circuit
    grouped_elements, unnamed_elements = group_elements_by_circuit(elements_to_circuit, ee_collector)

    # Define the electrical system type
    system_type = DB.Electrical.ElectricalSystemType.PowerCircuit

    # Use a Transaction Group to handle multiple transactions together
    tg = DB.TransactionGroup(doc, "Create and Update Circuits")
    tg.Start()

    try:
        # First Transaction: Create circuits and assign to panels
        created_systems = {}  # To store created systems and associate them with original groupings

        with revit.Transaction("Create Circuits and Assign Panels"):
            for key, data in grouped_elements:
                if data["panel_element"] and data["element_ids"]:
                    created_system = create_electrical_system(doc, data["element_ids"], system_type, data["panel_element"])
                    if created_system:
                        created_systems[created_system.Id] = key  # Link to the group key
                    else:
                        print("Skipped creating system for: {}".format(key))

            for key, data in unnamed_elements.items():
                if data["element_ids"]:
                    created_system = create_electrical_system(doc, data["element_ids"], system_type, data["panel_element"])
                    if created_system:
                        created_systems[created_system.Id] = key
                    else:
                        print("Skipped creating system for unnamed group: {}".format(key))

        # Collect all electrical systems created in the project
        all_systems = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem).ToElements()

        # Second Transaction: Update circuit parameters based on original group data
        with revit.Transaction("Update Circuit Parameters"):
            for system in all_systems:
                if system.Id in created_systems:
                    key = created_systems[system.Id]
                    # Check if it's a grouped element or unnamed
                    if key in dict(grouped_elements):
                        data = dict(grouped_elements)[key]
                    else:
                        data = unnamed_elements.get(key)

                    # Update parameters
                    if data:
                        # Example of updating circuit parameters: "ckt-Rating" and "ckt-Load Name"
                        rating_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
                        load_name_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
                        ckt_notes_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)

                        if rating_param and data["rating"]:
                            rating_param.Set(data["rating"])

                        if load_name_param and data["load_name"]:
                            load_name_param.Set(data["load_name"])

                        if ckt_notes_param and data["ckt_notes"]:
                            ckt_notes_param.Set(data["ckt_notes"])

        # Commit the transaction group to save all changes
        tg.Assimilate()

    except Exception as e:
        print("Error occurred: {}. Rolling back all changes.".format(e))
        tg.RollBack()



    # Commented Output Section (for reference and future use)
    """
    output = script.get_output()
    output.print_md("## Grouped Elements by Panel and Circuit:")
    for key, data in grouped_elements:
        output.print_md("**Panel: {} | Circuit: {}**".format(data["panel_name"], data["circuit_number"]))
        if data["panel_element"]:
            output.print_md("Panel Element ID: {}".format(data["panel_element"].Id))
            output.print_md("Rating: {}, Load Name: {}".format(data["rating"], data["load_name"]))
            output.print_md("Elements: {}".format(len(data["elements"])))
            output.print_md("Element IDs for Create method: {}".format([str(eid) for eid in data["element_ids"]]))
    
    output.print_md("## Unnamed Circuit Elements:")
    for key, data in unnamed_elements.items():
        output.print_md("**Key: {} | Element ID: {}**".format(key, data["element_ids"][0]))
        output.print_md("Rating: {}, Load Name: {}".format(data["rating"], data["load_name"]))
    
    """

if __name__ == "__main__":
        main()