import re

from pyrevit import DB, script, forms

logger = script.get_logger()
output = script.get_output()

# Get the active Revit document
doc = __revit__.ActiveUIDocument.Document

# Markdown and table data containers
device_critical = []
circuit_critical = []
equipment_critical = []



# Regular expression to extract the VA value from the RBS_ELECTRICAL_DATA parameter
va_regex = re.compile(r'-(\d+)\s*VA')

# Define categories for electrical devices and equipment
category_enum = {
    "Electrical Equipment": DB.BuiltInCategory.OST_ElectricalEquipment,
    "Electrical Fixtures": DB.BuiltInCategory.OST_ElectricalFixtures,
    "Lighting Fixtures": DB.BuiltInCategory.OST_LightingFixtures,
    "Lighting Devices": DB.BuiltInCategory.OST_LightingDevices,
    "Data Devices": DB.BuiltInCategory.OST_DataDevices,
    "Fire Alarm Devices": DB.BuiltInCategory.OST_FireAlarmDevices,
    "Security Devices": DB.BuiltInCategory.OST_SecurityDevices
}


# --- Function to Check Thresholds ---
def check_thresholds(device_count, circuit_count, equipment_count):
    # Thresholds for alerts
    TOTAL_THRESHOLD = 300
    INDIVIDUAL_THRESHOLD = 100

    total_issues = device_count + circuit_count + equipment_count
    if total_issues > TOTAL_THRESHOLD or \
       device_count > INDIVIDUAL_THRESHOLD or \
       circuit_count > INDIVIDUAL_THRESHOLD or \
       equipment_count > INDIVIDUAL_THRESHOLD:
        message = (
            "Critical Issue Counts:\n"
            "Devices: {}\n"
            "Circuits: {}\n"
            "Equipment: {}\n\n"
            "The form may take a while to print. Do you wish to continue?"
        ).format(device_count, circuit_count, equipment_count)
        return forms.alert(message, title="Large Report", yes=True, no=True, exitscript=True)
    return True


# Step 1: Collect Electrical Equipment
elec_equipment_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
    .WhereElementIsNotElementType() \
    .ToElements()

logger.info("Found %d electrical equipment instances" % len(elec_equipment_collector))

# Step 2: Collect Electrical Devices
elec_devices = []
for name, cat in category_enum.items():
    if name != "Electrical Equipment":
        devices_collector = DB.FilteredElementCollector(doc) \
            .OfCategory(cat) \
            .WhereElementIsNotElementType() \
            .ToElements()
        elec_devices.extend(devices_collector)

logger.info("Found %d electrical device instances" % len(elec_devices))

# Step 3: Collect Electrical Circuits
circuits_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit) \
    .WhereElementIsNotElementType() \
    .ToElements()

logger.info("Found %d electrical circuit instances" % len(circuits_collector))

# --- Electrical Equipment Checks ---
for equipment in elec_equipment_collector:
    try:
        mains_param = equipment.get_Parameter(DB.BuiltInParameter.RBS_ELEC_MAINS)
        demand_param = equipment.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM)
        mcb_rating_param = equipment.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM)
        panel_name_param = equipment.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)

        mains = mains_param.AsDouble() if mains_param else None
        demand = demand_param.AsDouble() if demand_param else None
        mcb_rating = mcb_rating_param.AsDouble() if mcb_rating_param else None
        panel_name = panel_name_param.AsString() if panel_name_param else "N/A"

        if mains is not None and demand is not None and mains < demand:
            equipment_critical.append([
                output.linkify(equipment.Id),
                panel_name,
                "Mains (%s) < Total Demand (%s)" % (mains, demand)
            ])

        if mcb_rating is not None and demand is not None and mcb_rating < demand:
            equipment_critical.append([
                output.linkify(equipment.Id),
                panel_name,
                "MCB Rating (%s) < Total Demand (%s)" % (mcb_rating, demand)
            ])

    except Exception as e:
        logger.error("Error checking equipment %s: %s" % (equipment.Id, str(e)))

# --- Device Checks ---
for device in elec_devices:
    try:
        circuit_number_param = device.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        panel_param = device.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        electrical_data_param = device.get_Parameter(DB.BuiltInParameter.RBS_ELECTRICAL_DATA)
        circuit_number = circuit_number_param.AsString() if circuit_number_param else ""
        panel_name = panel_param.AsString() if panel_param else ""
        electrical_data = electrical_data_param.AsString() if electrical_data_param else ""

        va_value = 0
        if electrical_data:
            va_match = va_regex.search(electrical_data)
            if va_match:
                va_value = int(va_match.group(1))

        if va_value > 0:
            if not circuit_number and not panel_name:
                device_critical.append([
                    output.linkify(device.Id),
                    device.Symbol.FamilyName,
                    device.Name,
                    "Non-zero load (%s VA) not connected to any electrical system" % va_value
                ])
            elif circuit_number == "<unnamed>" and not panel_name:
                device_critical.append([
                    output.linkify(device.Id),
                    device.Symbol.FamilyName,
                    device.Name,
                    "Non-zero load (%s VA) not assigned to panel" % va_value
                ])

    except Exception as e:
        logger.error("Error checking device %s: %s" % (device.Id, str(e)))

# --- Circuit Checks ---
for circuit in circuits_collector:
    try:
        circuit_rating_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
        apparent_current_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_APPARENT_CURRENT_PARAM)
        panel_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        circuit_number_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        load_name_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)

        circuit_rating = circuit_rating_param.AsDouble() if circuit_rating_param else None
        apparent_current = apparent_current_param.AsDouble() if apparent_current_param else None
        panel_name = panel_param.AsString() if panel_param else ""
        circuit_number = circuit_number_param.AsString() if circuit_number_param else ""
        load_name = load_name_param.AsString() if load_name_param else ""
        panel_circuit = "%s:%s" % (panel_name, circuit_number) if panel_name and circuit_number else "N/A"

        if circuit_rating is not None and apparent_current is not None and circuit_rating < apparent_current:
            circuit_critical.append([
                output.linkify(circuit.Id),
                panel_circuit,
                load_name,
                "Circuit Rating (%s A) < Apparent Current (%s A)" % (circuit_rating, apparent_current)
            ])

        if apparent_current == 0 and panel_name is not None:
            circuit_critical.append([
                output.linkify(circuit.Id),
                panel_circuit,
                load_name,
                "Load on Circuit %s is Zero" % panel_circuit
            ])

    except Exception as e:
        logger.error("Error checking circuit %s: %s" % (circuit.Id, str(e)))


critical_count = check_thresholds(len(device_critical),
                                  len(circuit_critical),
                                  len(equipment_critical))


# --- Print Critical Issues ---
output.print_md("# Electrical System Check")
if equipment_critical or device_critical or circuit_critical:

    if equipment_critical:
        output.print_md("## Electrical Equipment Issues")
        output.print_table(equipment_critical, columns=["Element ID", "Panel Name", "Error"])

    if device_critical:
        output.print_md("## Electrical Device Issues")
        output.print_table(device_critical, columns=["Element ID","Family", "Type", "Error"])

    if circuit_critical:
        output.print_md("## Electrical Circuit Issues")
        output.print_table(circuit_critical, columns=["Element ID", "Circuit", "Load Name", "Error"])
else:
    output.print_md("### No issues found!")

logger.info("QC Check Completed")
