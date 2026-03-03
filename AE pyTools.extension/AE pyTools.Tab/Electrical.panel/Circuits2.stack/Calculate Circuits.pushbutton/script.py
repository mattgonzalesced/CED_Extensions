# -*- coding: utf-8 -*-

from pyrevit import forms, revit, DB, script

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_runner import build_default_runner
from Snippets import _elecutils as eu

doc = revit.doc
logger = script.get_logger()


def _collect_target_circuit_ids(doc):
    selection = list(revit.get_selection() or [])
    if selection:
        selected = []
        for el in selection:
            if isinstance(el, DB.Electrical.ElectricalSystem):
                selected.append(el)
        if not selected:
            selected = eu.get_circuits_from_selection(selection)
    else:
        selected = eu.pick_circuits_from_list(doc, select_multiple=True)

    return [c.Id.IntegerValue for c in selected if isinstance(c, DB.Electrical.ElectricalSystem)]


def main():
    circuit_ids = _collect_target_circuit_ids(doc)
    if not circuit_ids:
        forms.alert('No circuits selected.', exitscript=True)

    request = OperationRequest(
        operation_key='calculate_circuits',
        circuit_ids=circuit_ids,
        source='ribbon',
        options={'show_output': True},
    )

    runner = build_default_runner(alert_parameter_name='Circuit Data_CED')
    result = runner.run(request, doc)
    if not result:
        return
    if result.get('status') != 'ok':
        logger.info('Calculate circuits request ended: %s', result)


if __name__ == '__main__':
    main()
