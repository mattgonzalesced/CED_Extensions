# -*- coding: utf-8 -*-
"""Calculate circuits using the layered CEDElectrical architecture."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB, forms, revit, script

from CEDElectrical.Model.CircuitBranch import CircuitSettings
from CEDElectrical.Domain.circuit_evaluator import CircuitEvaluator
from CEDElectrical.Services.revit_reader import RevitCircuitReader
from CEDElectrical.Services.revit_writer import RevitCircuitWriter
from Snippets import _elecutils as eu

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
logger = script.get_logger()


def pick_circuits_from_ui():
    return eu.pick_circuits_from_list(doc, select_multiple=True)


def alert_large_selection(count):
    proceed = forms.alert(
        "%s circuits selected.\n\nThis may take a while.\n\n" % count,
        title="⚠️ Large Selection Warning",
        options=["Continue", "Cancel"]
    )
    return proceed == "Continue"


def calculate_and_write(circuits, settings):
    reader = RevitCircuitReader(doc, settings)
    writer = RevitCircuitWriter(doc)
    evaluator = CircuitEvaluator(settings)

    results = []
    for circuit in circuits:
        model = reader.read(circuit)
        if not model.is_power_circuit:
            continue
        results.append(evaluator.evaluate(model))

    fixture_total = 0
    equipment_total = 0

    tg = DB.TransactionGroup(doc, "Calculate Circuits")
    tg.Start()
    t = DB.Transaction(doc, "Write Shared Parameters")
    try:
        t.Start()
        for result in results:
            writer.write_circuit(result.model.circuit, result)
            f, e = writer.write_connected(result.model.circuit, result)
            fixture_total += f
            equipment_total += e
        t.Commit()
        tg.Assimilate()
    except Exception as ex:
        t.RollBack()
        tg.RollBack()
        logger.error("Transaction failed: %s" % ex)
        raise

    return results, fixture_total, equipment_total


def main():
    settings = CircuitSettings()
    reader = RevitCircuitReader(doc, settings)

    selection = revit.get_selection()
    circuits = reader.get_selected_circuits(selection, pick_circuits_from_ui)

    count = len(circuits)
    if count > 1000:
        if not alert_large_selection(count):
            script.exit()

    results, fixtures, equipment = calculate_and_write(circuits, settings)

    output = script.get_output()
    output.close_others()
    output.print_md("## ✅ Shared Parameters Updated")
    output.print_md("* Circuits updated: **%s**" % len(results))
    output.print_md("* Electrical Fixtures updated: **%s**" % fixtures)
    output.print_md("* Electrical Equipment updated: **%s**" % equipment)


if __name__ == "__main__":
    main()
