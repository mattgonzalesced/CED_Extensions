# -*- coding: utf-8 -*-
"""Reusable selection helpers for circuit-centric UI tools."""

from System.Collections.Generic import List
from pyrevit import DB, revit
from pyrevit.compat import get_elementid_value_func

_get_elid_value = get_elementid_value_func()


def _idval(item):
    try:
        return int(_get_elid_value(item))
    except Exception:
        return int(getattr(item, "IntegerValue", 0))


def set_revit_selection(elements, uidoc=None):
    uidoc = uidoc or getattr(revit, "uidoc", None)
    if uidoc is None:
        return False
    ids = List[DB.ElementId]()
    seen = set()
    for element in list(elements or []):
        try:
            element_id = getattr(element, "Id", None)
        except Exception:
            element_id = None
        if element_id is None:
            continue
        key = _idval(element_id)
        if key <= 0 or key in seen:
            continue
        seen.add(key)
        ids.Add(element_id)
    try:
        uidoc.Selection.SetElementIds(ids)
        return True
    except Exception:
        return False


def clear_revit_selection(uidoc=None):
    uidoc = uidoc or getattr(revit, "uidoc", None)
    if uidoc is None:
        return False
    try:
        uidoc.Selection.SetElementIds(List[DB.ElementId]())
        return True
    except Exception:
        return False


def collect_circuit_targets(circuit, mode):
    mode_key = str(mode or "").strip().lower()
    if circuit is None:
        return []
    if mode_key == "circuit":
        return [circuit]
    if mode_key == "panel":
        try:
            base_equipment = circuit.BaseEquipment
        except Exception:
            base_equipment = None
        return [base_equipment] if base_equipment is not None else []
    if mode_key == "device":
        try:
            return [el for el in list(circuit.Elements or []) if el is not None]
        except Exception:
            return []
    return []


def format_writeback_lock_reason(row):
    if not isinstance(row, dict):
        return "Blocked by ownership"
    circuit_locked = bool(row.get("circuit_locked", False))
    circuit_owner = str(row.get("circuit_owner") or "").strip()
    device_owner = str(row.get("device_owner") or "").strip()
    sync_writeback = bool(row.get("sync_writeback", False))
    if circuit_locked and circuit_owner and device_owner:
        return "Locked by {}; downstream owned by {}".format(circuit_owner, device_owner)
    if circuit_locked and circuit_owner:
        return "Locked by {}".format(circuit_owner)
    if circuit_locked:
        return "Locked by another user"
    if device_owner:
        if sync_writeback:
            return "Blocked by writeback ownership ({})".format(device_owner)
        return "Blocked by downstream ownership ({})".format(device_owner)
    return "Blocked by ownership"
