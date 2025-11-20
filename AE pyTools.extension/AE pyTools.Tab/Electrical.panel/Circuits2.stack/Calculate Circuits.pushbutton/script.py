# -*- coding: utf-8 -*-
"""Modaless circuit browser and calculator for electrical systems."""
from pyrevit import DB, forms, revit

from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitSettings
from CEDElectrical.ui.circuit_browser import CircuitSelectionWindow
from Snippets import _elecutils as eu

doc = revit.doc


def collect_circuits():
    """Collect electrical circuits from selection or active document."""
    selection = revit.get_selection()
    circuits = []
    if selection:
        circuits = eu.get_circuits_from_selection(selection)
    if not circuits:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem)
        circuits = [c for c in collector if c.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit]
    return circuits


def main():
    settings = CircuitSettings()
    circuits = collect_circuits()
    if not circuits:
        forms.alert("No electrical circuits found for display.", title="Circuit Browser")
        return
    window = CircuitSelectionWindow(doc, circuits, settings)
    window.show(modal=False)


def collect_circuits():
    selection = revit.get_selection()
    circuits = []
    if selection:
        circuits = eu.get_circuits_from_selection(selection)
    if not circuits:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem)
        circuits = [c for c in collector if c.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit]
    return circuits


def main():
    settings = CircuitSettings()
    circuits = collect_circuits()
    if not circuits:
        forms.alert("No electrical circuits found for display.", title="Circuit Browser")
        return
    window = CircuitSelectionWindow(doc, circuits, settings)
    window.show(modal=False)


if __name__ == "__main__":
    main()
