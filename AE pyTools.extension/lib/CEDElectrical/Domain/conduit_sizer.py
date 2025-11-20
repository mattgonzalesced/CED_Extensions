# -*- coding: utf-8 -*-
"""Conduit sizing utilities."""

from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE
from CEDElectrical.refdata.conduit_area_table import CONDUIT_AREA_TABLE, CONDUIT_SIZE_INDEX


class ConduitSizer(object):
    def __init__(self, settings):
        self.settings = settings

    def _conductor_area(self, wire_size):
        if wire_size is None:
            return 0
        return CONDUCTOR_AREA_TABLE.get(str(wire_size), 0)

    def _conduit_area(self, conduit_type, conduit_size):
        if conduit_type is None or conduit_size is None:
            return None
        try:
            material_table = CONDUIT_AREA_TABLE.get(conduit_type, {})
            if not material_table:
                return None
            return material_table.get(conduit_size)
        except Exception:
            return None

    def pick_conduit_size(self, conduit_type, total_area):
        material_table = CONDUIT_AREA_TABLE.get(conduit_type, {})
        if not material_table or total_area is None:
            return None

        for size in CONDUIT_SIZE_INDEX:
            if size in material_table:
                area = material_table[size]
                if area * self.settings.max_conduit_fill >= total_area:
                    return size
        return None

    def size_conduit(self, conduit_type, hot_size, neutral_size, ground_size, isolated_ground_size, quantities):
        total_area = 0
        total_area += self._conductor_area(hot_size) * quantities.get('hot', 0)
        total_area += self._conductor_area(neutral_size) * quantities.get('neutral', 0)
        total_area += self._conductor_area(ground_size) * quantities.get('ground', 0)
        total_area += self._conductor_area(isolated_ground_size) * quantities.get('isolated_ground', 0)

        conduit_size = self.pick_conduit_size(conduit_type, total_area)
        conduit_area = self._conduit_area(conduit_type, conduit_size)
        fill_pct = None
        if conduit_area:
            fill_pct = total_area / float(conduit_area)
        return conduit_size, fill_pct
