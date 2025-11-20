# -*- coding: utf-8 -*-
"""Core data models for circuit calculations.

These classes intentionally avoid Revit business logic so that
calculation engines can operate on plain Python objects.
"""


class CircuitSettings(object):
    """User-configurable calculation settings."""

    def __init__(self):
        self.min_wire_size = '12'
        self.max_wire_size = '600'
        self.min_breaker_size = 20
        self.auto_calculate_breaker = False
        self.min_conduit_size = '3/4"'
        self.max_conduit_fill = 0.36
        self.max_branch_voltage_drop = 0.03
        self.max_feeder_voltage_drop = 0.02
        self.wire_size_prefix = '#'
        self.conduit_size_suffix = 'C'


class CircuitBranchModel(object):
    """Plain data extracted from a Revit electrical system."""

    def __init__(
        self,
        circuit_id=None,
        panel='',
        circuit_number=None,
        branch_type='Unknown',
        load_name=None,
        rating=None,
        frame=None,
        length=None,
        circuit_notes=None,
        voltage=None,
        apparent_current=None,
        apparent_power=None,
        power_factor=None,
        poles=0,
        include_neutral=None,
        include_isolated_ground=None,
        auto_calculate_override=False,
        base_wire_info=None,
        overrides=None,
        is_power_circuit=False,
        is_feeder=False,
        name=None,
        settings=None,
    ):
        self.circuit_id = circuit_id
        self.panel = panel or ''
        self.circuit_number = circuit_number
        self.branch_type = branch_type or 'Unknown'
        self.load_name = load_name
        self.rating = rating
        self.frame = frame
        self.length = length
        self.circuit_notes = circuit_notes
        self.voltage = voltage
        self.apparent_current = apparent_current
        self.apparent_power = apparent_power
        self.power_factor = power_factor
        self.poles = poles or 0

        self.include_neutral = include_neutral
        self.include_isolated_ground = include_isolated_ground
        self.auto_calculate_override = auto_calculate_override

        self.base_wire_info = base_wire_info or {}
        self.overrides = overrides or {}
        self.is_power_circuit = bool(is_power_circuit)
        self.is_feeder = bool(is_feeder)
        self.settings = settings or CircuitSettings()
        self.name = name or "%s-%s" % (self.panel, self.circuit_number)

    @property
    def max_voltage_drop(self):
        if self.is_feeder:
            return self.settings.max_feeder_voltage_drop
        return self.settings.max_branch_voltage_drop

    @property
    def circuit_load_current(self):
        if not self.is_power_circuit:
            return None
        if self.is_feeder:
            return self.apparent_current
        return self.apparent_current


class CircuitCalculationResult(object):
    """Calculated outputs for a circuit branch."""

    def __init__(self, model):
        self.model = model
        self.breaker_rating = None
        self.frame = model.frame
        self.length = model.length
        self.circuit_notes = model.circuit_notes
        self.voltage_drop_percentage = None

        self.hot_wire_size = None
        self.neutral_wire_size = None
        self.ground_wire_size = None
        self.isolated_ground_wire_size = None
        self.number_of_sets = None
        self.hot_wire_quantity = None
        self.neutral_wire_quantity = None
        self.ground_wire_quantity = None
        self.isolated_ground_wire_quantity = None
        self.number_of_wires = None

        self.wire_material = None
        self.wire_temp_rating = None
        self.wire_insulation = None

        self.conduit_type = None
        self.conduit_size = None
        self.conduit_fill_percentage = None

        self.circuit_load_current = model.circuit_load_current
        self.circuit_base_ampacity = None

    @property
    def wire_size_callout(self):
        parts = []
        if self.hot_wire_size and self.hot_wire_quantity:
            parts.append("%s%s" % (self.hot_wire_quantity, self.hot_wire_size))
        if self.neutral_wire_size and self.neutral_wire_quantity:
            parts.append("%sN%s" % (self.neutral_wire_quantity, self.neutral_wire_size))
        if self.ground_wire_size and self.ground_wire_quantity:
            parts.append("%sG%s" % (self.ground_wire_quantity, self.ground_wire_size))
        if self.isolated_ground_wire_size and self.isolated_ground_wire_quantity:
            parts.append("%sIG%s" % (self.isolated_ground_wire_quantity, self.isolated_ground_wire_size))
        return ', '.join(parts)

    @property
    def conduit_and_wire_size(self):
        if self.conduit_size and self.conduit_type:
            if self.wire_size_callout:
                return "%s %s - %s" % (self.conduit_size, self.conduit_type, self.wire_size_callout)
            return "%s %s" % (self.conduit_size, self.conduit_type)
        return self.wire_size_callout
