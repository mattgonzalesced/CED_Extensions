# -*- coding: utf-8 -*-
"""Operation registration map."""


class OperationRegistry(object):
    """Maps operation keys to operation instances."""

    def __init__(self):
        self._operations = {}

    def register(self, operation):
        """Register an operation by its ``key`` field."""
        key = getattr(operation, 'key', None)
        if not key:
            raise ValueError('Operation key is required.')
        self._operations[key] = operation

    def get(self, key):
        """Return operation instance for key or ``None``."""
        return self._operations.get(key)
