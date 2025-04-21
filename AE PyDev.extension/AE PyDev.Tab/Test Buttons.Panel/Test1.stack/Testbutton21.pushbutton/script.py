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
        60: [
            ('14', 15), ('12', 20), ('10', 30), ('8', 40), ('6', 55),
            ('4', 70), ('3', 85), ('2', 95), ('1', 110), ('1/0', 125),
            ('2/0', 145), ('3/0', 165), ('4/0', 195), ('250', 215), ('300', 240),
            ('350', 260), ('400', 280), ('500', 320), ('600', 355), ('700', 385),
            ('750', 400), ('800', 410), ('900', 435), ('1000', 455), ('1250', 495),
            ('1500', 520), ('1750', 545), ('2000', 560)
        ],
        75: [
            ('14', 20), ('12', 25), ('10', 35), ('8', 50), ('6', 65),
            ('4', 85), ('3', 100), ('2', 115), ('1', 130), ('1/0', 150),
            ('2/0', 175), ('3/0', 200), ('4/0', 230), ('250', 255), ('300', 285),
            ('350', 310), ('400', 335), ('500', 380), ('600', 420), ('700', 460),
            ('750', 475), ('800', 490), ('900', 520), ('1000', 545), ('1250', 590),
            ('1500', 625), ('1750', 650), ('2000', 665)
        ],
        90: [
            ('14', 25), ('12', 30), ('10', 40), ('8', 55), ('6', 75),
            ('4', 95), ('3', 110), ('2', 130), ('1', 150), ('1/0', 170),
            ('2/0', 195), ('3/0', 225), ('4/0', 260), ('250', 290), ('300', 320),
            ('350', 350), ('400', 380), ('500', 430), ('600', 475), ('700', 520),
            ('750', 535), ('800', 555), ('900', 585), ('1000', 615), ('1250', 665),
            ('1500', 705), ('1750', 735), ('2000', 750)
        ]
    },
    'Aluminium': {
        60: [
            ('12', 15), ('10', 25), ('8', 35), ('6', 40), ('4', 55),
            ('3', 65), ('2', 75), ('1', 90), ('1/0', 100), ('2/0', 120),
            ('3/0', 135), ('4/0', 155), ('250', 170), ('300', 190),
            ('350', 210), ('400', 225), ('500', 260), ('600', 285), ('700', 310),
            ('750', 320), ('800', 330), ('900', 355), ('1000', 375), ('1250', 405),
            ('1500', 435), ('1750', 455), ('2000', 470)
        ],
        75: [
            ('12', 20), ('10', 30), ('8', 40), ('6', 50), ('4', 65),
            ('3', 75), ('2', 90), ('1', 100), ('1/0', 120), ('2/0', 135),
            ('3/0', 155), ('4/0', 180), ('250', 205), ('300', 230),
            ('350', 250), ('400', 270), ('500', 310), ('600', 340), ('700', 375),
            ('750', 385), ('800', 395), ('900', 425), ('1000', 445), ('1250', 485),
            ('1500', 520), ('1750', 545), ('2000', 560)
        ],
        90: [
            ('12', 25), ('10', 35), ('8', 45), ('6', 60), ('4', 75),
            ('3', 85), ('2', 100), ('1', 115), ('1/0', 135), ('2/0', 150),
            ('3/0', 175), ('4/0', 205), ('250', 230), ('300', 255),
            ('350', 280), ('400', 305), ('500', 350), ('600', 385), ('700', 420),
            ('750', 435), ('800', 450), ('900', 480), ('1000', 500), ('1250', 545),
            ('1500', 585), ('1750', 615), ('2000', 630)
        ]
    }
}

EGC_TABLE = {
    'Copper': [
        (15, '14'), (20, '12'), (30, '10'), (40, '10'), (60, '10'),
        (100, '8'), (200, '6'), (300, '4'), (400, '3'), (500, '2'),
        (600, '1'), (800, '1/0'), (1000, '2/0'), (1200, '3/0'),
        (1600, '4/0'), (2000, '250'), (2500, '350'), (3000, '400'),
        (4000, '500'), (5000, '700'), (6000, '800')
    ],
    'Aluminium': [
        (15, '12'), (20, '10'), (30, '8'), (40, '8'), (60, '8'),
        (100, '6'), (200, '4'), (300, '2'), (400, '1'), (500, '1/0'),
        (600, '2/0'), (800, '3/0'), (1000, '4/0'), (1200, '250'),
        (1600, '350'), (2000, '400'), (2500, '600'), (3000, '600'),
        (4000, '800'), (5000, '1200'), (6000, '1200')
    ]
}

BREAKER_FRAME_SWITCH_TABLE = {
    15: {'frame': 30, 'switch': 30},
    20: {'frame': 30, 'switch': 30},
    25: {'frame': 30, 'switch': 30},
    30: {'frame': 30, 'switch': 30},
    35: {'frame': 60, 'switch': 60},
    40: {'frame': 60, 'switch': 60},
    45: {'frame': 60, 'switch': 60},
    50: {'frame': 60, 'switch': 60},
    60: {'frame': 60, 'switch': 60},
    70: {'frame': 100, 'switch': 100},
    80: {'frame': 100, 'switch': 100},
    90: {'frame': 100, 'switch': 100},
    100: {'frame': 100, 'switch': 100},
    125: {'frame': 200, 'switch': 200},
    150: {'frame': 200, 'switch': 200},
    175: {'frame': 200, 'switch': 200},
    200: {'frame': 200, 'switch': 200},
    225: {'frame': 225, 'switch': 400},
    250: {'frame': 250, 'switch': 400},
    300: {'frame': 400, 'switch': 400},
    350: {'frame': 400, 'switch': 400},
    400: {'frame': 400, 'switch': 400},
    450: {'frame': 600, 'switch': 600},
    500: {'frame': 600, 'switch': 600},
    600: {'frame': 600, 'switch': 600},
    700: {'frame': 800, 'switch': 800},
    800: {'frame': 800, 'switch': 800},
    1000: {'frame': 1000, 'switch': 1000},
    1200: {'frame': 1200, 'switch': 1200},
    1600: {'frame': 1600, 'switch': 1600},
    2000: {'frame': 2000, 'switch': 2000},
    2500: {'frame': 2500, 'switch': 2500},
    3000: {'frame': 3000, 'switch': 3000},
    4000: {'frame': 4000, 'switch': 4000},
    5000: {'frame': 5000, 'switch': 5000},
    6000: {'frame': 6000, 'switch': 6000}
}



def pick_circuits_from_list():
    ckts = DB.FilteredElementCollector(doc) \
        .OfClass(ElectricalSystem) \
        .WhereElementIsNotElementType()

    print("Total Circuits in Doc: {}".format(ckts.GetElementCount()))

    ckt_options = {" All": []}

    for ckt in ckts:
        ckt_supply = DB.Element.Name.__get__(ckt.BaseEquipment)
        ckt_number = ckt.CircuitNumber
        ckt_load_name = ckt.LoadName
        if ckt.SystemType == ElectricalSystemType.PowerCircuit:
            ckt_rating = ckt.Rating
            ckt_wireType = ckt.WireType
        # print("{}/{} ({}) - {}".format(ckt_supply, ckt_number, ckt_rating, ckt_load_name))

        ckt_options[" All"].append(ckt)

        if ckt_supply not in ckt_options:
            ckt_options[ckt_supply] = []
        ckt_options[ckt_supply].append(ckt)

    ckt_lookup = {}
    grouped_options = {}
    for group, circuits in ckt_options.items():
        option_strings = []
        for ckt in circuits:
            ckt_string = "{} | {} - {}".format(DB.Element.Name.__get__(ckt.BaseEquipment), ckt.CircuitNumber,
                                               ckt.LoadName)
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


class CircuitSettings(object):
    def __init__(self):
        # User-adjustable settings (can be loaded from file or UI later)
        self.min_wire_size = '12'
        self.max_wire_size = '600'
        self.min_breaker_size = 20
        self.max_parallel_size = '500'  # largest wire allowed before parallel
        self.excluded_wire_sizes = []

    def to_dict(self):
        return {
            'min_wire_size': self.min_wire_size,
            'max_wire_size': self.max_wire_size,
            'min_breaker_size': self.min_breaker_size,
            'max_parallel_size': self.max_parallel_size
        }

    def load_from_dict(self, data):
        self.min_wire_size = data.get('min_wire_size', self.min_wire_size)
        self.max_wire_size = data.get('max_wire_size', self.max_wire_size)
        self.min_breaker_size = data.get('min_breaker_size', self.min_breaker_size)
        self.max_parallel_size = data.get('max_parallel_size', self.max_parallel_size)


class CircuitBranch(object):
    def __init__(self, circuit, settings=None):
        self.circuit = circuit
        self.settings = settings if settings else CircuitSettings()
        self.circuit_id = circuit.Id.IntegerValue
        self.name = "{}-{}".format(circuit.BaseEquipment.Name, circuit.CircuitNumber)

        self._wire_info = None  # Lazy-loaded wire info dictionary

        # User overrides (None = no override)
        self._breaker_override = None
        self._hot_wire_override = None
        self._ground_wire_override = None
        self._max_single_wire_size = '500'  # Max size before parallel sets

        # Calculated values (set by calculation methods)
        self._calculated_breaker = None
        self._calculated_hot_wire = None
        self._calculated_hot_sets = None
        self._calculated_hot_ampacity = None
        self._calculated_ground_wire = None

    # ----------- Classification -----------

    @property
    def is_power_circuit(self):
        return self.circuit.SystemType == ElectricalSystemType.PowerCircuit

    @property
    def is_spare(self):
        return self.circuit.CircuitType == CircuitType.Spare

    @property
    def is_space(self):
        return self.circuit.CircuitType == CircuitType.Space

    # ----------- Wire Info Dictionary -----------

    @property
    def wire_info(self):
        if not self.is_power_circuit:
            return {}
        if self._wire_info is None:
            try:
                wt = self.circuit.WireType
                self._wire_info = {
                    'material': DB.Element.Name.__get__(wt.WireMaterial),
                    'temperature': DB.Element.Name.__get__(wt.TemperatureRating),
                    'conduit': WireConduitType.Name.__get__(wt.Conduit),
                    'insulation': DB.Element.Name.__get__(wt.Insulation)
                }
            except Exception:
                self._wire_info = {}
        return self._wire_info

    # ----------- Circuit Properties -----------

    @property
    def rating(self):
        try:
            if self.is_power_circuit and not self.is_space:
                return self.circuit.Rating
        except:
            return None

    @property
    def length(self):
        try:
            if self.is_power_circuit and not self.is_spare and not self.is_space:
                return self.circuit.Length
        except:
            return None

    @property
    def voltage(self):
        try:
            return ElectricalSystem.Voltage.__get__(self.circuit)
        except:
            return None

    @property
    def apparent_power(self):
        try:
            return ElectricalSystem.ApparentLoad.__get__(self.circuit)
        except:
            return None

    @property
    def apparent_current(self):
        try:
            return ElectricalSystem.ApparentCurrent.__get__(self.circuit)
        except:
            return None

    @property
    def poles(self):
        try:
            return ElectricalSystem.PolesNumber.__get__(self.circuit)
        except:
            return None

    @property
    def power_factor(self):
        try:
            return ElectricalSystem.PowerFactor.__get__(self.circuit)
        except:
            return None

    # ----------- Override Setters -----------

    def set_breaker_override(self, value):
        self._breaker_override = value

    def set_hot_wire_override(self, wire_size):
        self._hot_wire_override = wire_size

    def set_ground_wire_override(self, wire_size):
        self._ground_wire_override = wire_size

    def set_max_single_wire_size(self, wire_size):
        self._max_single_wire_size = wire_size

    # ----------- Public Access Properties -----------

    @property
    def breaker_rating(self):
        return self._breaker_override if self._breaker_override is not None else self._calculated_breaker

    @property
    def hot_wire_size(self):
        return self._hot_wire_override if self._hot_wire_override is not None else self._calculated_hot_wire

    @property
    def ground_wire_size(self):
        return self._ground_wire_override if self._ground_wire_override is not None else self._calculated_ground_wire

    @property
    def number_of_sets(self):
        return self._calculated_hot_sets

    @property
    def circuit_base_ampacity(self):
        return self._calculated_hot_ampacity

    # ----------- Calculations -----------

    def calculate_breaker_size(self):
        try:
            amps = self.apparent_current
            if amps:
                amps *= 1.25
                if amps < self.settings.min_breaker_size:
                    amps = self.settings.min_breaker_size

                for b in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
                    if b >= amps:
                        self._calculated_breaker = b
                        break
        except:
            self._calculated_breaker = None

    def calculate_hot_wire_size(self):
        try:
            breaker_amps = self.breaker_rating
            if breaker_amps is None:
                return

            temp = int(self.wire_info.get('temperature', '75').replace('C', '').replace('°', ''))
            material = self.wire_info.get('material', 'Copper')
            wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp, [])

            min_size = self.settings.min_wire_size
            max_size = self.settings.max_parallel_size
            sets = 1

            # Find index of minimum allowed wire size
            start_index = 0
            for i, (wire_size, _) in enumerate(wire_set):
                if wire_size == min_size:
                    start_index = i
                    break

            while sets < 10:
                for wire, ampacity in wire_set[start_index:]:
                    if ampacity * sets >= breaker_amps:
                        self._calculated_hot_wire = wire
                        self._calculated_hot_sets = sets
                        self._calculated_hot_ampacity = ampacity * sets
                        return
                    if wire == max_size:
                        sets += 1  # we’ve reached max wire size, need to parallel
                        break  # restart loop with new set count
                else:
                    # No wire found in this set range — break to avoid looping forever
                    break

            # Fallback if no suitable wire was found
            self._calculated_hot_wire = None
            self._calculated_hot_sets = None
            self._calculated_hot_ampacity = None

        except Exception as e:
            self._calculated_hot_wire = None
            self._calculated_hot_sets = None
            self._calculated_hot_ampacity = None

    def calculate_ground_wire_size(self):
        try:
            amps = self.breaker_rating
            if amps is None:
                return

            material = self.wire_info.get('material', 'Copper')
            min_size = self.settings.min_wire_size
            egc_list = EGC_TABLE.get(material, [])

            # Find index of first entry that matches or exceeds min wire size
            start_index = 0
            for i, (_, wire_size) in enumerate(egc_list):
                if wire_size == min_size:
                    start_index = i
                    break

            # Filter from the min wire size forward
            for amp_limit, wire_size in egc_list[start_index:]:
                if amps <= amp_limit:
                    self._calculated_ground_wire = wire_size
                    return

            # If nothing matches
            self._calculated_ground_wire = None

        except Exception as e:
            self._calculated_ground_wire = None

    def calculate_hot_wire_quantity(self):
        return self.poles or 0

    def calculate_neutral_quantity(self):
        # Placeholder for future logic
        return 0

    def calculate_ground_wire_quantity(self):
        return 1

    def calculate_isolated_ground_quantity(self):
        # Placeholder for future logic
        return 0

    def calculate_voltage_drop(self):
        # Placeholder for voltage drop calculation (3-phase)
        return None

    def print_info(self, include_wire_info=True, include_all_properties=False):
        print("\n=== CircuitBranch: {} (ID: {}) ===".format(self.name, self.circuit_id))

        # Wire Info
        if include_wire_info:
            print("\nWire Info:")
            if self.wire_info:
                for key, val in self.wire_info.items():
                    print("    {}: {}".format(key, val if val else "N/A"))
            else:
                print("    No wire info available.")

        # Circuit Info
        print("\nCircuit Info:")
        info_fields = [
            'rating', 'voltage', 'length',
            'apparent_power', 'apparent_current',
            'poles', 'power_factor'
        ]
        for attr in info_fields:
            try:
                value = getattr(self, attr)
                print("    {}: {}".format(attr, value if value is not None else "N/A"))
            except:
                print("    {}: [Error]".format(attr))

        # Calculation Results
        print("\nCalculated/Resolved Values:")
        print("    Breaker Rating: {}".format(self.breaker_rating or "N/A"))
        print("    Hot Wire Size: {}".format(self.hot_wire_size or "N/A"))
        print("    Number of Sets: {}".format(self.number_of_sets or "N/A"))
        print("    Circuit Base Ampacity: {}".format(self.circuit_base_ampacity or "N/A"))
        print("    Ground Wire Size: {}".format(self.ground_wire_size or "N/A"))

        # Quantities
        print("\nWire Quantities:")
        print("    Hot Conductors: {}".format(self.calculate_hot_wire_quantity()))
        print("    Ground Conductors: {}".format(self.calculate_ground_wire_quantity()))
        print("    Neutral Conductors: {}".format(self.calculate_neutral_quantity()))
        print("    Isolated Ground Conductors: {}".format(self.calculate_isolated_ground_quantity()))


def main():
    test_condition = 1

    if test_condition == 0:
        test_circuit = revit.get_selection()
    else:
        test_circuit = pick_circuits_from_list()

    for circuit in test_circuit:
        branch = CircuitBranch(circuit)
        branch.calculate_breaker_size()
        branch.calculate_hot_wire_size()
        branch.calculate_ground_wire_size()

        branch.print_info()
        # branch.debug("Post Calculation")


if __name__ == "__main__":
    main()
