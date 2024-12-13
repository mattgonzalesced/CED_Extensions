from Autodesk.Revit.DB import FilteredElementCollector, Electrical, Transaction, BuiltInCategory, BuiltInParameter, \
    ElementId
from pyrevit import script, forms, output


def get_all_panels(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_panel_types(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_circuits(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_elec_fixtures(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_data_devices(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DataDevices).WhereElementIsNotElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_light_devices(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_LightingDevices).WhereElementIsNotElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_light_fixtures(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_LightingFixtures).WhereElementIsNotElementType()
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


# Helper function to get panel's distribution system and voltage capacity
def get_panel_dist_system(panel, doc, debug=False):
    """Returns a dictionary with the panel's distribution system name, voltage, and phase."""
    panel_data = {
        'dist_system_name': None,
        'phase': None,
        'lg_voltage': None,
        'll_voltage': None
    }

    # Try to get the secondary distribution system (for transformers)
    secondary_dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS)
    dist_system_id = None  # Initialize dist_system_id

    if secondary_dist_system_param and secondary_dist_system_param.HasValue:
        dist_system_id = secondary_dist_system_param.AsElementId()
        if debug:
            print("Secondary distribution system found for panel: {}".format(panel.Name))
    else:
        # Fallback to primary distribution system (for panels or switchboards)
        dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
        if dist_system_param and dist_system_param.HasValue:
            dist_system_id = dist_system_param.AsElementId()
            if debug:
                print("Primary distribution system found for panel: {}".format(panel.Name))
        else:
            if debug:
                print("Warning: No distribution system found for panel: {}".format(panel.Name))
            return panel_data  # Return early if no distribution system is found

    # Retrieve the DistributionSysType element using the ID
    dist_system_type = doc.GetElement(dist_system_id)

    if dist_system_type is None:
        if debug:
            print("Warning: Distribution system element not found for panel: {}".format(panel.Name))
        return panel_data  # Return early if the distribution system element is not found

    # Retrieve the Name using the SYMBOL_NAME_PARAM built-in parameter
    name_param = dist_system_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if name_param and name_param.HasValue:
        panel_data['dist_system_name'] = name_param.AsString()
        if debug:
            print("Distribution system name for panel {}: {}".format(panel.Name, panel_data['dist_system_name']))
    else:
        if debug:
            print("Warning: No name found for the distribution system of panel: {}".format(panel.Name))
        panel_data['dist_system_name'] = "Unnamed Distribution System"

    # Get phase (check if ElectricalPhase exists)
    if hasattr(dist_system_type, "ElectricalPhase"):
        panel_data['phase'] = dist_system_type.ElectricalPhase
    else:
        if debug:
            print("Warning: No phase information found for distribution system: {}".format(panel.Name))

    # Retrieve Line-to-Ground and Line-to-Line voltages
    lg_voltage = getattr(dist_system_type, "VoltageLineToGround", None)
    ll_voltage = getattr(dist_system_type, "VoltageLineToLine", None)

    # Fetch voltage values
    if lg_voltage:
        lg_voltage_param = lg_voltage.get_Parameter(BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
        panel_data['lg_voltage'] = lg_voltage_param.AsDouble() if lg_voltage_param else None
    else:
        if debug:
            print("Warning: No L-G voltage found for panel: {}".format(panel.Name))

    if ll_voltage:
        ll_voltage_param = ll_voltage.get_Parameter(BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
        panel_data['ll_voltage'] = ll_voltage_param.AsDouble() if ll_voltage_param else None
    else:
        if debug:
            print("Warning: No L-L voltage found for panel: {}".format(panel.Name))

    return panel_data


def get_compatible_panels(selected_circuit, all_panels, doc):
    """Returns a list of compatible panels based on the selected circuit's poles and voltage."""
    circuit_poles = selected_circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES).AsInteger()
    circuit_voltage = selected_circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE).AsDouble()

    compatible_panels = []

    for panel in all_panels:
        panel_data = get_panel_dist_system(panel, doc)
        panel_lg_voltage = panel_data['lg_voltage']
        panel_ll_voltage = panel_data['ll_voltage']

        if circuit_poles == 1 and panel_lg_voltage and abs(panel_lg_voltage - circuit_voltage) < 1.0:
            compatible_panels.append(panel)
        elif circuit_poles >= 2 and panel_ll_voltage and abs(panel_ll_voltage - circuit_voltage) < 1.0:
            compatible_panels.append(panel)

    return compatible_panels


def get_circuit_data(circuit):
    """Returns a dictionary containing the number of poles and voltage for the circuit."""
    circuit_data = {'poles': None, 'voltage': None}

    poles_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
    if poles_param and poles_param.HasValue:
        circuit_data['poles'] = poles_param.AsInteger()

    voltage_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
    if voltage_param and voltage_param.HasValue:
        circuit_data['voltage'] = voltage_param.AsDouble()

    return circuit_data


def move_circuits_to_panel(circuits, target_panel, doc, output):
    """Moves the selected circuits to the target panel and stores old and new info."""
    data = []
    with Transaction(doc, "Move Circuits to New Panel") as trans:
        trans.Start()
        for circuit in circuits:
            old_panel = circuit.BaseEquipment.Name
            old_circuit_number = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()

            circuit.SelectPanel(target_panel)
            doc.Regenerate()

            new_circuit_number = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
            prev_circuit = "{} / {}".format(old_panel, old_circuit_number)
            new_circuit = "{} / {}".format(target_panel.Name, new_circuit_number)
            data.append([output.linkify(circuit.Id), prev_circuit, new_circuit])

        trans.Commit()

    return data


def find_open_slots(target_panel):
    """Find available slots in the target panel, prioritizing odd-numbered slots."""
    available_slots = list(range(1, 43))
    odd_slots = [slot for slot in available_slots if slot % 2 == 1]
    even_slots = [slot for slot in available_slots if slot % 2 == 0]
    return odd_slots + even_slots


def get_circuits_from_panel(panel, doc, sort_method=0, include_spares=True):
    """Get circuits from a selected panel with sorting and inclusion of spare/space circuits."""
    circuits = {}
    panel_circuits = FilteredElementCollector(doc).OfClass(Electrical.ElectricalSystem).ToElements()

    for circuit in panel_circuits:
        if circuit.BaseEquipment and circuit.BaseEquipment.Id == panel.Id:
            if not include_spares and circuit.CircuitType in [Electrical.CircuitType.Spare, Electrical.CircuitType.Space]:
                continue

            # Get circuit parameters
            circuit_number = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
            load_name = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).AsString()
            start_slot_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT)
            wire_size_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM)

            # Retrieve wire size as string if available
            wire_size = wire_size_param.AsString() if wire_size_param and wire_size_param.HasValue else "N/A"

            # Retrieve the start slot if available
            start_slot = start_slot_param.AsInteger() if start_slot_param and start_slot_param.HasValue else None

            # Retrieve the panel name
            panel_name = circuit.BaseEquipment.Name if circuit.BaseEquipment else "N/A"

            # Store data in a dictionary with element ID as the key
            circuits[circuit.Id] = {
                'circuit_number': circuit_number,
                'load_name': load_name,
                'start_slot': start_slot,
                'wire_size': wire_size,
                'panel': panel_name,
                'circuit': circuit
            }

    # Sort circuits based on the selected method
    if sort_method == 1:
        circuits_sorted = dict(sorted(circuits.items(), key=lambda item: item[1]['start_slot']))
    else:
        circuits_sorted = dict(sorted(circuits.items(), key=lambda item: (item[1]['start_slot'] % 2 == 0, item[1]['start_slot'])))

    return circuits_sorted

