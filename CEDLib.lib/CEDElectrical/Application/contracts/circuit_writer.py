# -*- coding: utf-8 -*-
"""Circuit write contract."""


class ICircuitWriter(object):
    """Writes calculated results onto circuits and downstream elements."""

    def write_circuit_parameters(self, circuit, param_values):
        """Write calculated parameter map to circuit."""
        raise NotImplementedError

    def write_connected_elements(self, branch, param_values, settings, locked_ids=None):
        """Write calculated parameter map to connected elements."""
        raise NotImplementedError
