# -*- coding: utf-8 -*-
"""Wire sizing utilities."""

from CEDElectrical.refdata.ampacity_table import WIRE_AMPACITY_TABLE


class WireSizer(object):
    def __init__(self, settings):
        self.settings = settings

    def _get_ampacity_table(self, material, temp):
        material_table = WIRE_AMPACITY_TABLE.get(material, {})
        return material_table.get(temp, [])

    def size_hot_conductor(self, model, overrides, breaker_rating):
        if breaker_rating is None and model.circuit_load_current is None:
            return None

        base_info = model.base_wire_info or {}
        material = overrides.get('wire_material_override') or base_info.get('wire_material') or 'CU'
        temp_str = overrides.get('wire_temp_rating_override') or base_info.get('wire_temperature_rating') or '75'
        insulation = overrides.get('wire_insulation_override') or base_info.get('wire_insulation', 'THHN')

        try:
            temp = int(str(temp_str).replace('C', '').strip())
        except Exception:
            temp = 75

        ampacity_table = self._get_ampacity_table(material, temp)
        if not ampacity_table:
            return None

        sets = overrides.get('wire_sets_override') or base_info.get('number_of_parallel_sets') or 1
        try:
            sets = int(sets)
        except Exception:
            sets = 1
        if sets < 1:
            sets = 1

        target = breaker_rating
        if target is None:
            target = 0
        load_current = model.circuit_load_current or 0
        if load_current:
            target = max(target, load_current * 1.25)

        hot_override = overrides.get('wire_hot_size_override')
        if hot_override:
            for size, ampacity in ampacity_table:
                if size == hot_override and ampacity * sets >= target:
                    return {
                        'hot_size': size,
                        'wire_sets': sets,
                        'ampacity': ampacity * sets,
                        'material': material,
                        'temperature': temp,
                        'insulation': insulation
                    }

        for size, ampacity in ampacity_table:
            total = ampacity * sets
            if total >= target:
                return {
                    'hot_size': size,
                    'wire_sets': sets,
                    'ampacity': total,
                    'material': material,
                    'temperature': temp,
                    'insulation': insulation
                }
        return None

    def size_neutral(self, hot_selection, overrides):
        if not hot_selection:
            return None
        neutral_override = overrides.get('wire_neutral_size_override')
        if neutral_override:
            return neutral_override
        return hot_selection.get('hot_size')
