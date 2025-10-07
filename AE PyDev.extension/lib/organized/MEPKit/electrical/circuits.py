# -*- coding: utf-8 -*-
from __future__ import absolute_import
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType

from organized.MEPKit.revit.transactions import RunInTransaction
from organized.MEPKit.electrical.devices import is_circuited

def _get_existing_power_system(elem):
    mep = getattr(elem, "MEPModel", None)
    if not mep: return None
    try:
        # Prefer enumerating and filtering by type
        systems = getattr(mep, "ElectricalSystems", None)
        if systems:
            for s in systems:
                try:
                    if s.SystemType == ElectricalSystemType.PowerCircuit:
                        return s
                except: pass
        if hasattr(mep, "GetElectricalSystems"):
            ss = list(mep.GetElectricalSystems())  # may return IList
            for s in ss:
                if s.SystemType == ElectricalSystemType.PowerCircuit:
                    return s
    except:
        pass
    return None

def get_or_create_power_circuit(doc, device):
    """Return a power ElectricalSystem for device; create if missing."""
    sys = _get_existing_power_system(device)
    if sys: return sys
    ids = List[ElementId](); ids.Add(device.Id)
    return ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)

def add_devices(system, devices):
    """Add multiple devices to an existing system (ignores already-added)."""
    for d in devices:
        try:
            system.Add(d)
        except:
            pass
    return system

def assign_to_panel(system, panel):
    """Select a panel for the circuit (no transaction inside)."""
    if not system or not panel: return False
    try:
        system.SelectPanel(panel)
        return True
    except:
        return False

@RunInTransaction("Electrical::AutoCircuitToNearestPanel")
def auto_circuit_devices_to_panel(doc, devices, panel):
    """Make (or reuse) a single circuit and add the given devices, then assign to panel."""
    if not devices: return None
    sys = get_or_create_power_circuit(doc, devices[0])
    add_devices(sys, devices[1:])
    assign_to_panel(sys, panel)
    return sys