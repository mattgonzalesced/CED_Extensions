# -*- coding: utf-8 -*-
"""Voltage drop calculations."""

from math import sqrt
from pyrevit import DB

from CEDElectrical.refdata.impedance_table import WIRE_IMPEDANCE_TABLE


class VoltageDropCalculator(object):
    def __init__(self):
        pass

    def _length_in_feet(self, length_internal):
        try:
            return DB.UnitUtils.ConvertFromInternalUnits(length_internal, DB.DisplayUnitType.DUT_FEET)
        except Exception:
            return length_internal

    def calculate_percentage(self, model, hot_size, wire_sets, material, conduit_material_type):
        if not hot_size or model.length is None or model.voltage is None or model.circuit_load_current is None:
            return None

        data = WIRE_IMPEDANCE_TABLE.get(str(hot_size))
        if not data:
            return None

        x_table = data.get('X', {})
        r_table = data.get('R', {}).get(material, {})
        if conduit_material_type not in x_table or conduit_material_type not in r_table:
            return None

        x_val = x_table.get(conduit_material_type)
        r_val = r_table.get(conduit_material_type)
        if x_val is None or r_val is None:
            return None

        impedance = sqrt(r_val ** 2 + x_val ** 2)
        length_ft = self._length_in_feet(model.length) or 0
        if length_ft <= 0:
            return None

        vd_volts = (2 * length_ft * impedance * (model.circuit_load_current or 0)) / (1000.0 * (wire_sets or 1))
        try:
            return vd_volts / float(model.voltage)
        except Exception:
            return None
