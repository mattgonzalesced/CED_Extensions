# -*- coding: utf-8 -*-
"""DTO for operation execution requests."""


class OperationRequest(object):
    """Carries operation key, targets, and options."""

    def __init__(self, operation_key, circuit_ids=None, source='ribbon', options=None):
        self.operation_key = operation_key
        self.circuit_ids = list(circuit_ids or [])
        self.source = source
        self.options = dict(options or {})
