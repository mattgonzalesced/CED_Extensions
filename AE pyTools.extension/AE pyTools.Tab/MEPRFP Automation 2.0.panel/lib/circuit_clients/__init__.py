# -*- coding: utf-8 -*-
"""
Client registry for SuperCircuit.

Each client is a module under ``circuit_clients/`` that subclasses
``CircuitClient`` (base.py). The registry below maps a stable key
(``"heb"``, ``"pf"``, ...) to the client instance. Adding a new client
is two steps: drop a new module in this package and add it to
``_CLIENTS``.

This package is import-safe outside Revit — the base class and clients
are pure Python (no Revit-API imports). The Revit-side adapters live in
``circuit_workflow.py`` / ``circuit_apply.py``.
"""

from circuit_clients.base import CircuitClient
from circuit_clients.heb import HebClient
from circuit_clients.pf import PfClient


_CLIENTS = (
    HebClient(),
    PfClient(),
)


def all_clients():
    """Return the list of registered client instances, in display order."""
    return list(_CLIENTS)


def by_key(key):
    """Look up a client by its stable key. Returns None if not found."""
    if not key:
        return None
    target = str(key).strip().lower()
    for c in _CLIENTS:
        if c.key == target:
            return c
    return None


def display_names():
    """``[(display_name, key), ...]`` for picker dialogs."""
    return [(c.display_name, c.key) for c in _CLIENTS]


__all__ = (
    "CircuitClient",
    "HebClient",
    "PfClient",
    "all_clients",
    "by_key",
    "display_names",
)
