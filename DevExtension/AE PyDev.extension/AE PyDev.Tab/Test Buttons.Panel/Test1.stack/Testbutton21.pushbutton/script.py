# -*- coding: utf-8 -*-

from pyrevit import DB, script, forms, revit, output
from pyrevit.revit import query
import clr
from Snippets import _elecutils as eu
from Autodesk.Revit.DB.Electrical import *

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

console = script.get_output()
logger = script.get_logger()

WIRE_AMPACITY_TABLE = {
    'Copper': {
        60: [('14', 15), ('12', 20), ('10', 30), ('8', 40), ('6', 55),
             ('4', 70), ('3', 85), ('2', 95), ('1', 110), ('1/0', 125),
             ('2/0', 145), ('3/0', 165), ('4/0', 195)],
        75: [('14', 20), ('12', 25), ('10', 35), ('8', 50), ('6', 65),
             ('4', 85), ('3', 100), ('2', 115), ('1', 130), ('1/0', 150),
             ('2/0', 175), ('3/0', 200), ('4/0', 230)],
        90: [('14', 25), ('12', 30), ('10', 40), ('8', 55), ('6', 75),
             ('4', 95), ('3', 110), ('2', 125), ('1', 145), ('1/0', 170),
             ('2/0', 195), ('3/0', 225), ('4/0', 260)]
    },
    'Aluminum': {
        60: [('14', 15), ('12', 15), ('10', 25), ('8', 35), ('6', 40),
             ('4', 55), ('2', 75), ('1', 90), ('1/0', 100), ('2/0', 120),
             ('3/0', 135), ('4/0', 155)],
        75: [('14', 20), ('12', 20), ('10', 30), ('8', 40), ('6', 50),
             ('4', 65), ('2', 85), ('1', 100), ('1/0', 110), ('2/0', 130),
             ('3/0', 150), ('4/0', 175)],
        90: [('14', 25), ('12', 25), ('10', 35), ('8', 50), ('6', 60),
             ('4', 75), ('2', 100), ('1', 115), ('1/0', 130), ('2/0', 150),
             ('3/0', 175), ('4/0', 205)]
    }
}


def pick_circuits_from_list():
    ckts = DB.FilteredElementCollector(doc) \
        .OfClass(ElectricalSystem) \
        .WhereElementIsNotElementType()

    print("Total Circuits in Doc: {}".format(ckts.GetElementCount()))

    ckt_options = {" All": []}

    for ckt in ckts:
        ckt_supply = ckt.BaseEquipment.Name
        ckt_number = ckt.CircuitNumber
        ckt_load_name = ckt.LoadName
        ckt_rating = ckt.Rating
        ckt_wireType = ckt.WireType
        print("{}/{} ({}) - {}".format(ckt_supply, ckt_number, ckt_rating, ckt_load_name))
        ckt_options[" All"].append(ckt)

        if ckt_supply not in ckt_options:
            ckt_options[ckt_supply] = []
        ckt_options[ckt_supply].append(ckt)

    ckt_lookup = {}
    grouped_options = {}
    for group, circuits in ckt_options.items():
        option_strings = []
        for ckt in circuits:
            ckt_string = "{} | {} - {}".format(ckt.BaseEquipment.Name, ckt.CircuitNumber, ckt.LoadName)
            option_strings.append(ckt_string)
            ckt_lookup[ckt_string] = ckt  # Map string to circuit
        option_strings.sort()
        grouped_options[group] = option_strings

    selected_option = forms.SelectFromList.show(
        grouped_options,
        title="Select a CKT",
        group_selector_title="Panel:",
        multiselect=False
    )

    if not selected_option:
        logger.info("No circuit selected. Exiting script.")
        script.exit()

    selected_ckt = ckt_lookup[selected_option]
    print("Selected Circuit Element ID: {}".format(selected_ckt.Id))
    return selected_ckt


def get_wire_type(ckt):
    wire_type = ckt.WireType
    wire_material = DB.Element.Name.__get__(wire_type.WireMaterial)
    wire_temp = DB.Element.Name.__get__(wire_type.TemperatureRating)
    wire_cond = WireConduitType.Name.__get__(wire_type.Conduit)
    wire_insulation = DB.Element.Name.__get__(wire_type.Insulation)

    wire_type_info = {
        DB.Element.Id.__get__(ckt):
            {'wire_material': wire_material,
             'wire_temp': wire_temp,
             'wire_cond': wire_cond,
             'wire_insulation': wire_insulation
             }
    }

    return wire_type_info


def get_circuit_info(ckt):
    ckt_supply = ckt.BaseEquipment.Name
    ckt_number = ckt.CircuitNumber
    ckt_load_name = ckt.LoadName
    ckt_rating = ckt.Rating
    ckt_apparent_power = ElectricalSystem.ApparentLoad.__get__(ckt)
    ckt_apparent_current = ElectricalSystem.ApparentCurrent.__get__(ckt)
    ckt_voltage = ElectricalSystem.Voltage.__get__(ckt)
    ckt_poles_number = ElectricalSystem.PolesNumber.__get__(ckt)
    ckt_power_factor = ElectricalSystem.PowerFactor.__get__(ckt)

    circuit_info = {
        DB.Element.Id.__get__(ckt):
            {'ckt_supply': ckt_supply,
             'ckt_number': ckt_number,
             'ckt_load_name': ckt_load_name,
             'ckt_rating': ckt_rating,
             'ckt_apparent_power': ckt_apparent_power,
             'ckt_apparent_current': ckt_apparent_current,
             'ckt_voltage': ckt_voltage,
             'ckt_poles_number': ckt_poles_number,
             'ckt_power_factor': ckt_power_factor
             }
    }

    return circuit_info


def print_nested_dict(nested_dict, title="Data"):
    print("\n=== {} ===\n".format(title))
    for circuit_id, data in nested_dict.items():
        print("Circuit ID: {}".format(circuit_id))
        for key, value in data.items():
            print("    {}: {}".format(key, value))
        print("")  # blank line between circuits


def get_voltage_drop(voltage,current,wire_qty,wire_impedance):
    return


test_condition = 0

if test_condition == 0:
    test_circuit = revit.get_selection()

else:
    test_circuit = pick_circuits_from_list()

for circuit in test_circuit:
    print("\n\nGetting wire type info...\n")
    wire_type_info = get_wire_type(circuit)
    print_nested_dict(wire_type_info,"Wire Type Stuff")

    print("\n\nGetting circuit info...\n")
    circuit_info = get_circuit_info(circuit)

    print_nested_dict(circuit_info,"Circuit Type Stuff")

