# -*- coding: utf-8 -*-
"""Panel slot data model used by panel schedule orchestration."""


class PanelSlot(object):
    """Represents one schedule slot and its resolved slot metadata."""

    def __init__(
        self,
        slot,
        cells=None,
        is_locked=False,
        is_spare=False,
        is_space=False,
        is_circuit=False,
        poles=1,
        group_number=0,
    ):
        self.slot = int(slot or 0)
        self.cells = list(cells or [])
        self.is_locked = bool(is_locked)
        self.is_spare = bool(is_spare)
        self.is_space = bool(is_space)
        self.is_circuit = bool(is_circuit)
        self.poles = int(max(1, poles or 1))
        self.group_number = int(group_number or 0)

    @property
    def is_grouped(self):
        """Return True when this slot belongs to a grouped schedule row."""
        return bool(int(self.group_number or 0) > 0)

    def get_slot_range(self):
        """Return resolved row/column coordinates for this slot."""
        return list(self.cells or [])

