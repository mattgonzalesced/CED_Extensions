# -*- coding: utf-8 -*-
import json

class FeederVDMethod(object):
    DEMAND = "demand"
    CONNECTED = "connected"
    EIGHTY_PERCENT = "80_percent"

    @classmethod
    def all(cls):
        return [cls.DEMAND, cls.CONNECTED, cls.EIGHTY_PERCENT]


class NeutralBehavior(object):
    MATCH_HOT = "match_hot"
    MANUAL = "manual"

    @classmethod
    def all(cls):
        return [cls.MATCH_HOT, cls.MANUAL]


class CircuitSettings(object):
    DEFAULTS = {
        # ORIGINAL settings you still need internally:
        "min_wire_size": "12",
        "max_wire_size": "600",
        "min_breaker_size": 20,
        "auto_calculate_breaker": False,
        "wire_size_prefix": "#",
        "conduit_size_suffix": "C",

        # USER-EXPOSED settings:
        "min_conduit_size": '3/4"',
        "max_conduit_fill": 0.36,
        "neutral_behavior": NeutralBehavior.MATCH_HOT,
        "max_branch_voltage_drop": 0.03,
        "max_feeder_voltage_drop": 0.02,
        "feeder_vd_method": FeederVDMethod.DEMAND,
        "write_equipment_results": True,
        "write_fixture_results": False,
    }

    def __init__(self, values=None):
        values = values or {}
        self._values = {}

        for key in self.DEFAULTS:
            if key in values:
                self._values[key] = values[key]
            else:
                self._values[key] = self.DEFAULTS[key]

    def get(self, key):
        return self._values.get(key)

    def set(self, key, value):
        # Validation
        if key == "feeder_vd_method":
            if value not in FeederVDMethod.all():
                # backward compatibility for old persisted values
                legacy_map = {
                    "80_percent": FeederVDMethod.EIGHTY_PERCENT,
                }
                value = legacy_map.get(value, value)
            if value not in FeederVDMethod.all():
                raise ValueError("Invalid feeder_vd_method: {}".format(value))

        if key == "neutral_behavior":
            if value not in NeutralBehavior.all():
                raise ValueError("Invalid neutral_behavior: {}".format(value))

        if key in ("max_conduit_fill",
                   "max_branch_voltage_drop",
                   "max_feeder_voltage_drop"):
            value = round(float(value), 3)  # ensures it is numeric and rounded

        if key in ("write_equipment_results", "write_fixture_results"):
            value = bool(value)

        self._values[key] = value

    def to_json(self):
        payload = dict(self._values)
        for key in ("max_conduit_fill", "max_branch_voltage_drop", "max_feeder_voltage_drop"):
            try:
                payload[key] = round(float(payload[key]), 3)
            except Exception:
                pass
        return json.dumps(payload)

    @classmethod
    def from_json(cls, text):
        if not text:
            return cls()
        try:
            data = json.loads(text)
        except:
            data = {}
        return cls(data)

# Attribute accessors ----------------------------------------

    @property
    def min_wire_size(self):
        return self._values["min_wire_size"]

    @property
    def max_wire_size(self):
        return self._values["max_wire_size"]

    @property
    def min_breaker_size(self):
        return self._values["min_breaker_size"]

    @property
    def auto_calculate_breaker(self):
        return self._values["auto_calculate_breaker"]

    @property
    def wire_size_prefix(self):
        return self._values["wire_size_prefix"]

    @property
    def conduit_size_suffix(self):
        return self._values["conduit_size_suffix"]


    # User-exposed settings
    @property
    def min_conduit_size(self):
        return self._values["min_conduit_size"]

    @property
    def max_conduit_fill(self):
        return float(self._values["max_conduit_fill"])

    @property
    def neutral_behavior(self):
        return self._values["neutral_behavior"]

    @property
    def max_branch_voltage_drop(self):
        return float(self._values["max_branch_voltage_drop"])

    @property
    def max_feeder_voltage_drop(self):
        return float(self._values["max_feeder_voltage_drop"])

    @property
    def feeder_vd_method(self):
        return self._values["feeder_vd_method"]

    @property
    def write_equipment_results(self):
        return bool(self._values["write_equipment_results"])

    @property
    def write_fixture_results(self):
        return bool(self._values["write_fixture_results"])
