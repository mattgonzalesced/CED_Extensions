# -*- coding: utf-8 -*-
__title__ = "List Circuits by Panel"

from pyrevit import script, revit, DB, forms
from collections import OrderedDict

output = script.get_output()
doc = revit.doc

# Dictionary to map parameter names to BuiltInParameter enums

PARAMETER_MAP = OrderedDict([
    ("SLOT NUMBER", DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT),
    ("PANEL NAME", DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM),
    ("CIRCUIT NUMBER", DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER),
    ("LOAD NAME", DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME),
    ("VOLTAGE", DB.BuiltInParameter.RBS_ELEC_VOLTAGE),
    ("POLES", DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES),
    ("APPARENT CURRENT", DB.BuiltInParameter.RBS_ELEC_APPARENT_CURRENT_PARAM),
    ("RATING", DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM),
    ("FRAME", DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM)
])


def format_panel_display(panel):
    """Returns a string with panel name and element ID for display."""
    return "{} (ID: {})".format(panel.Name, panel.Id)

def get_circuits_from_panel(panel):
    """Get circuits associated with the selected panel."""
    panel_circuits = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem).ToElements()
    return [circuit for circuit in panel_circuits if circuit.BaseEquipment and circuit.BaseEquipment.Id == panel.Id]

def retrieve_parameter_value(element, param_enum):
    """Retrieve the value of a parameter, handling different cases for 'n/a' and empty strings."""
    param = element.get_Parameter(param_enum)

    if not param:
        # Parameter does not exist on this element
        return "n/a"

    if param.HasValue:
        return format_parameter_value(param)

    # If parameter exists but has no value, return an empty string
    return ""

def format_parameter_value(param):
    """Format the parameter value based on its storage type."""
    try:
        # Use AsValueString if possible
        return param.AsValueString()
    except:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == DB.StorageType.Integer:
            return str(param.AsInteger())
        elif param.StorageType == DB.StorageType.Double:
            return str(DB.UnitUtils.ConvertFromInternalUnits(param.AsDouble(), param.DisplayUnitType))
        elif param.StorageType == DB.StorageType.ElementId:
            # Retrieve and display the name of the linked element
            linked_element = revit.doc.GetElement(param.AsElementId())
            return linked_element.Name if linked_element else str(param.AsElementId().IntegerValue)
        return ""

def collect_circuit_data(circuits):
    """Collect all parameter values for circuits."""
    circuit_data = []
    for circuit in circuits:
        row = []
        for param_name in PARAMETER_MAP.keys():  # Iterates in defined order
            param_enum = PARAMETER_MAP[param_name]
            value = retrieve_parameter_value(circuit, param_enum)
            row.append(value)
        circuit_data.append(row)
    return circuit_data


def sort_circuit_data(data):
    """Sort circuit data by panel name and circuit number."""
    panel_index = list(PARAMETER_MAP.keys()).index("PANEL NAME")
    circuit_number_index = list(PARAMETER_MAP.keys()).index("CIRCUIT NUMBER")
    return sorted(data, key=lambda x: (x[panel_index], x[circuit_number_index]))

def print_report(circuit_data, columns):
    """Print the report in a table with the specified columns."""
    output.print_md("## Circuit Parameter Report")
    output.print_md("**Legend**\n- `n/a`: Parameter not available in this project.")
    output.insert_divider()
    output.print_table(
        table_data=circuit_data,
        columns=columns
    )

# Main execution
with forms.ProgressBar(title="Executing Script...", max_value=4) as pb:
    # Step 1: Select Panels
    pb.title = "Selecting Panels (Step 1 of 4)"
    panel_collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType()
    panel_options = [format_panel_display(panel) for panel in panel_collector]

    selected_panel_options = forms.SelectFromList.show(
        panel_options,
        title="Select Panels",
        multiselect=True
    )
    if not selected_panel_options:
        output.print_md("No panels selected. Exiting...")
        script.exit()

    # Map selected panel display names back to panel objects
    selected_panels = [panel for panel in panel_collector if format_panel_display(panel) in selected_panel_options]
    pb.update_progress(1, 4)

    columns = list(PARAMETER_MAP.keys())
    all_circuits = []

    # Step 2: Get Circuits from Selected Panels
    pb.title = "Collecting Circuit Data (Step 2 of 4)"
    for panel in selected_panels:
        circuits = get_circuits_from_panel(panel)
        all_circuits.extend(circuits)
    pb.update_progress(2, 4)

    # Step 3: Retrieve and Sort Circuit Data
    pb.title = "Retrieving and Sorting Circuit Data (Step 3 of 4)"
    circuit_data = collect_circuit_data(all_circuits)
    sorted_circuit_data = sort_circuit_data(circuit_data)
    pb.update_progress(3, 4)

    # Step 4: Print Report
    pb.title = "Printing Report (Step 4 of 4)"
    if not sorted_circuit_data:
        output.print_md("No circuits found for the selected panels.")
        script.exit()
    else:
        print_report(sorted_circuit_data, columns)
    pb.update_progress(4, 4)
