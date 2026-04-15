# -*- coding: utf-8 -*-
"""Circuit read/lock query contract."""


class ICircuitRepository(object):
    """Provides Revit-backed circuit collection and lock information."""

    def get_target_circuits(self, doc, circuit_ids=None):
        """Return target circuits from explicit ids or model scope."""
        raise NotImplementedError

    def partition_locked_elements(self, doc, circuits, settings):
        """Split circuits into editable and locked subsets."""
        raise NotImplementedError

    def summarize_locked(self, doc, locked_ids):
        """Return summary counts for locked elements."""
        raise NotImplementedError
