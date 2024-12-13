# -*- coding: utf-8 -*-
__title__ = "Panel Breaker Report"

import clr
from pyrevit import script
from pyrevit import revit, DB
from pyrevit.revit import query
from collections import defaultdict

# Access the current Revit document
doc = revit.doc

# Set up the output window
output = script.get_output()
output.close_others()
output.set_width(800)

# Collect all electrical circuits
circuits = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Collect all internal circuits (spares and spaces)
internal_circuits = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalInternalCircuits) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Collect all Electrical Panels
elec_equip = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Combine both sets of circuits
all_circuits = list(circuits)
all_circuits.extend(internal_circuits)

# Dictionary to store panel information
panel_dict = defaultdict(lambda: {'circuits': defaultdict(int), 'spares': defaultdict(int), 'spaces': defaultdict(int)})

# Process each electrical system
for system in all_circuits:
    if isinstance(system, DB.Electrical.ElectricalSystem):
        panel = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM).AsString()
        if panel:
            poles = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES).AsInteger()
            amps = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM).AsDouble()
            notes = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM).AsString() or ""

            # Accessing the CircuitType enum
            circuit_type = system.CircuitType

            # Categorize the circuit based on its type and increment the count
            key = (poles, int(amps), notes)
            if circuit_type == DB.Electrical.CircuitType.Spare:
                panel_dict[panel]['spares'][key] += 1
            elif circuit_type == DB.Electrical.CircuitType.Space:
                panel_dict[panel]['spaces'][key] += 1
            else:
                panel_dict[panel]['circuits'][key] += 1


# Function to get parameter information by name
def get_parameter_by_name(element, param_name):
    """Return the value of the parameter with the given name."""
    param = next((p for p in element.Parameters if p.Definition.Name == param_name), None)
    if param and param.HasValue:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == DB.StorageType.Double:
            return param.AsDouble()  # Convert if necessary
        elif param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == DB.StorageType.ElementId:
            return param.AsElementId().IntegerValue
    return None


# Prepare to link equipment to panels
equipment_info = {}

# Extract information from each electrical equipment instance
for equip in elec_equip:
    panel_name = get_parameter_by_name(equip, "Panel Name")
    if not panel_name:
        continue

    # Collect both instance and type parameters
    equip_params = {
        "Wiring Configuration_CEDT": get_parameter_by_name(equip, "Wiring Configuration_CEDT"),
        "Mains Rating_CED": get_parameter_by_name(equip, "Mains Rating_CED"),
        "Mains Type_CEDT": get_parameter_by_name(equip, "Mains Type_CEDT"),
        "Main Breaker Rating_CED": get_parameter_by_name(equip, "Main Breaker Rating_CED"),
        "Short Circuit Rating_CEDT": get_parameter_by_name(equip, "Short Circuit Rating_CEDT"),
        "Panel Modifications_CEDT": get_parameter_by_name(equip, "Panel Modifications CEDT"),
        "Mounting_CEDT": get_parameter_by_name(equip, "Mounting_CEDT"),
        "Panel Feed_CEDT": get_parameter_by_name(equip, "Panel Feed_CEDT"),
        "Max Number of Circuits_CED": get_parameter_by_name(equip, "Max Number of Circuits_CED"),
        "Max Number of Single Pole Breakers_CED": get_parameter_by_name(equip,
                                                                        "Max Number of Single Pole Breakers_CED"),
    }

    # Include type parameters
    equip_type = doc.GetElement(equip.GetTypeId())
    for param_name in equip_params.keys():
        if equip_params[param_name] is None:
            equip_params[param_name] = get_parameter_by_name(equip_type, param_name)

    # Store information using the panel name as a key
    equipment_info[panel_name] = equip_params


# Function to display equipment parameters
def print_equipment_parameters(panel_name):
    if panel_name in equipment_info:
        params = equipment_info[panel_name]
        # Prepare content to display in two columns
        param_lines = []
        for param, value in params.items():
            if value is not None:
                param_lines.append("<p><strong>{}</strong>: {}</p>".format(param, value))

        # Split lines into two columns
        half = len(param_lines) // 2 + len(param_lines) % 2
        left_col = param_lines[:half]
        right_col = param_lines[half:]

        # Create HTML with two columns
        output_html = '<div style="display: flex; justify-content: space-between;">'
        output_html += '<div style="width: 45%;">' + "".join(left_col) + '</div>'
        output_html += '<div style="width: 45%;">' + "".join(right_col) + '</div>'
        output_html += '</div>'

        # Print the formatted header and parameters
        output.print_html("<h1 style='font-weight: bold;'>{}</h1>".format(panel_name))
        output.print_html(output_html)


# Function to process and format rows safely
def process_row(count, amps, poles, notes, label=""):
    if amps:
        amps_poles = "{}A / {}P".format(amps, poles)
    else:
        amps_poles = "{}P".format(poles)

    if label:
        note_label = "{} - {}".format(notes, label) if notes else label
    else:
        note_label = notes

    return count, amps_poles, note_label, label


# Update the table generation function to include equipment parameters
# Update the table generation function to ensure consistent spacing
def generate_report_tables():
    for panel_name, data in sorted(panel_dict.items()):
        # Optionally print equipment parameters first
        print_equipment_parameters(panel_name)

        # Use a wrapper to control spacing and positioning
        output_html = '<div style="margin-bottom: 10px; padding: 0;">'

        # Start HTML table with CSS for column widths and reduced spacing
        output_html += '<table style="width: 100%; border-collapse: collapse; margin: 0; padding: 0;">'
        output_html += ('<tr>'
                        '<th style="width: 10%; text-align: left; margin: 0; padding: 5px 0;">Count</th>'
                        '<th style="width: 20%; text-align: left; margin: 0; padding: 5px 0;">Amps / Poles</th>'
                        '<th style="width: 50%; text-align: left; margin: 0; padding: 5px 0;">Notes</th>'
                        '<th style="width: 20%; text-align: left; margin: 0; padding: 5px 0;">Label</th>'
                        '</tr>')

        # Helper function to generate a row safely
        def generate_table_row(count, amps, poles, notes, label=""):
            # Ensure empty placeholders where data may be missing
            if amps and poles:
                amps_poles = "{}A / {}P".format(amps, poles)
            elif poles:
                amps_poles = "{}P".format(poles)
            else:
                amps_poles = ""

            # Build the row string with all placeholders and reduced spacing
            row = (
                '<tr>'
                '<td style="text-align: left; margin: 0; padding: 5px 0;">({})</td>'
                '<td style="text-align: left; margin: 0; padding: 5px 0;">{}</td>'
                '<td style="text-align: left; margin: 0; padding: 5px 0;">{}</td>'
                '<td style="text-align: left; margin: 0; padding: 5px 0;">{}</td>'
                '</tr>'
            ).format(count, amps_poles, notes or "", label)
            return row

        # Add circuits first
        for key, count in sorted(data['circuits'].items()):
            poles, amps, notes = key
            row = generate_table_row(count, amps, poles, notes)
            output_html += row

        # Add spares
        for key, count in sorted(data['spares'].items()):
            poles, amps, notes = key
            row = generate_table_row(count, amps, poles, notes, "SPARE")
            output_html += row

        # Add spaces
        for key, count in sorted(data['spaces'].items()):
            poles, amps, notes = key
            row = generate_table_row(count, amps if amps else "", poles, notes, "SPACE")
            output_html += row

        # End HTML table
        output_html += '</table>'
        output_html += '</div>'  # Close wrapper div for consistent spacing

        # Output the HTML with controlled spacing
        output.print_html(output_html)
        output.insert_divider()
        output.print_md("<br><br><br>")


# Run the report generation
generate_report_tables()

