# -*- coding: utf-8 -*-
"""Circuit sizing calculator with searchable circuit picker."""

from pyrevit import DB, forms, revit

from CEDElectrical.circuit_sizing.services.revit_reader import CircuitListProvider
from CEDElectrical.circuit_sizing.services.sizing_runner import CircuitSizingRunner
from CEDElectrical.circuit_sizing.ui.circuit_selector import CircuitSelectorWindow
from Snippets import _elecutils as eu

USE_SELECTION_WINDOW = True  # Toggle to False to quickly test sizing logic without the picker

doc = revit.doc
logger = script.get_logger()


def _launch_selection_window():
    provider = CircuitListProvider(doc)
    window = CircuitSelectorWindow(provider)
    return window.show_dialog()


def main():
    selection = revit.get_selection()
    selected_circuits = []

    if selection:
        selected_circuits = [el for el in selection if isinstance(el, DB.Electrical.ElectricalSystem)]

    if not selected_circuits:
        if USE_SELECTION_WINDOW:
            selected_circuits = _launch_selection_window() or []
        else:
            selected_circuits = eu.pick_circuits_from_list(doc, select_multiple=True)

    if not selected_circuits:
        logger.info("No circuits selected. Exiting.")
        return

    runner = CircuitSizingRunner(doc, logger)
    results = runner.calculate_and_update(selected_circuits)
    if not results:
        return

    output = script.get_output()
    output.close_others()
    output.print_md("## ✅ Shared Parameters Updated")
    output.print_md("* Circuits updated: **{}**".format(results['circuits']))
    output.print_md("* Electrical Fixtures updated: **{}**".format(results['fixtures']))
    output.print_md("* Electrical Equipment updated: **{}**".format(results['equipment']))


if __name__ == "__main__":
    main()
