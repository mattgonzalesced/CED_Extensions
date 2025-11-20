# -*- coding: utf-8 -*-
"""Core data models for circuit calculations.

These classes intentionally avoid Revit business logic so that
calculation engines can operate on plain Python objects.
"""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.refdata.ocp_cable_defaults import OCP_CABLE_DEFAULTS


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

    def __init__(self, circuit, settings, overrides=None):
        self.circuit = circuit
        self.settings = settings or CircuitSettings()
        self.overrides = overrides or {}

        self.circuit_id = circuit.Id.Value if hasattr(circuit, 'Id') else None
        self.panel = self._get_panel_name(circuit)
        self.circuit_number = getattr(circuit, 'CircuitNumber', None)
        self.name = "%s-%s" % (self.panel, self.circuit_number)
        self.branch_type = self._get_branch_type(circuit)
        self.load_name = self._safe_get(circuit, 'LoadName')
        self.rating = self._safe_get(circuit, 'Rating')
        self.frame = self._safe_get(circuit, 'Frame')
        self.length = self._get_length(circuit)
        self.circuit_notes = self._get_parameter_value(circuit, 'CKT_Schedule Notes_CEDT')
        self.voltage = self._get_voltage(circuit)
        self.apparent_current = self._safe_get(DBE.ElectricalSystem, 'ApparentCurrent', circuit)
        self.apparent_power = self._safe_get(DBE.ElectricalSystem, 'ApparentLoad', circuit)
        self.power_factor = self._safe_get(DBE.ElectricalSystem, 'PowerFactor', circuit)
        self.poles = self._safe_get(DBE.ElectricalSystem, 'PolesNumber', circuit) or 0

        self.include_neutral = overrides.get('include_neutral') if overrides else None
        self.include_isolated_ground = overrides.get('include_isolated_ground') if overrides else None
        self.auto_calculate_override = overrides.get('auto_calculate_override') if overrides else False

        self.base_wire_info = self._get_base_wire_info()

    def _safe_get(self, source, attr_name, target=None):
        try:
            if target is None:
                return getattr(source, attr_name)
            return getattr(source, attr_name).__get__(target)
        except Exception:
            try:
                return getattr(target, attr_name)
            except Exception:
                return None

    def _get_parameter_value(self, element, param_name):
        try:
            param = element.LookupParameter(param_name)
            if param:
                if param.StorageType == DB.StorageType.String:
                    return param.AsString()
                if param.StorageType == DB.StorageType.Integer:
                    return param.AsInteger()
                if param.StorageType == DB.StorageType.Double:
                    return param.AsDouble()
        except Exception:
            pass
        return None

    def _get_voltage(self, circuit):
        try:
            return self._safe_get(DBE.ElectricalSystem, 'Voltage', circuit)
        except Exception:
            return None

    def _get_length(self, circuit):
        try:
            if hasattr(circuit, 'Length'):
                return circuit.Length
        except Exception:
            pass
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM)
            if param:
                return param.AsDouble()
        except Exception:
            pass
        return None

    def _get_panel_name(self, circuit):
        try:
            base = getattr(circuit, 'BaseEquipment', None)
            if base:
                return getattr(base, 'Name', '')
        except Exception:
            pass
        return ''

    def _get_branch_type(self, circuit):
        try:
            system_type = getattr(circuit, 'SystemType', None)
            if system_type:
                return str(system_type)
        except Exception:
            pass
        return 'Unknown'

    def _get_base_wire_info(self):
        rating = self.rating
        if rating is None:
            return {}
        try:
            rating_key = int(rating)
        except Exception:
            return {}

        table = OCP_CABLE_DEFAULTS
        if rating_key in table:
            return table[rating_key]

        sorted_keys = sorted(table.keys())
        for key in sorted_keys:
            if key >= rating_key:
                return table[key]
        if sorted_keys:
            return table[sorted_keys[-1]]
        return {}

    @property
    def is_power_circuit(self):
        try:
            return self.circuit.CircuitType == DBE.CircuitType.Circuit
        except Exception:
            return False

    @property
    def is_feeder(self):
        try:
            return self.circuit.IsFeedToPanel
        except Exception:
            return False

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
