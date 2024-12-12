from Autodesk.Revit.DB import FilteredElementCollector, Electrical, Transaction, BuiltInCategory, BuiltInParameter, \
    ElementId
from pyrevit import script, forms, output


class Circuit:
    """
    Manages wire sizing operations.
    """
    def __init__(self, circuit_type, rating, wire_type, wire_set, ):
        self.size = size  # e.g., "12", "500 kcmil"
        self.unit = unit  # Either "AWG" or "kcmil"

    def __repr__(self):
        return "<WireSize(size=" + str(self.size) + ", unit=" + str(self.unit) + ")>"


class Conductor:
    """
    Represents an electrical conductor.
    """
    def __init__(self, material, temp_rating, insulation, wire_size, base_ampacity):
        """
        Initialize a conductor with its NEC properties.
        """
        self.material = material  # e.g., "Copper", "Aluminum"
        self.temp_rating = temp_rating  # e.g., "90C"
        self.insulation = insulation  # e.g., "THHN"
        self.wire_size = wire_size  # Instance of WireSize
        self.base_ampacity = base_ampacity  # Ampacity at the rated temperature
        self.reactance_per_1000ft = None  # Impedance properties (set externally)
        self.resistance_per_1000ft = None
        self.area = None  # Conductor area (calculated based on insulation type)

    def set_impedance(self, reactance, resistance):
        """
        Sets impedance values.
        """
        self.reactance_per_1000ft = reactance
        self.resistance_per_1000ft = resistance

    def set_area(self, area):
        """
        Sets the conductor's cross-sectional area.
        """
        self.area = area

    def __repr__(self):
        return "<Conductor(material={material}, temp_rating={temp_rating}, insulation={insulation}, wire_size={wire_size}, base_ampacity={base_ampacity})>".format(
            material=self.material,
            temp_rating=self.temp_rating,
            insulation=self.insulation,
            wire_size=self.wire_size,
            base_ampacity=self.base_ampacity
        )


class Conduit:
    """
    Represents a conduit with its properties.
    """
    def __init__(self, type, diameter):
        """
        Initialize a conduit type.
        """
        self.type = type  # e.g., "PVC", "Steel", "EMT"
        self.diameter = diameter  # Diameter in inches

    def __repr__(self):
        return "<Conduit(type=" + str(self.type) + ", diameter=" + str(self.diameter) + ")>"
