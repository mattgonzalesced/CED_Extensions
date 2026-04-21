# -*- coding: utf-8 -*-
"""Operation runner and default service wiring."""

from CEDElectrical.Application.operations.autosize_breaker_and_recalculate_operation import (
    AutosizeBreakerAndRecalculateOperation,
)
from CEDElectrical.Application.operations.calculate_circuits_operation import CalculateCircuitsOperation
from CEDElectrical.Application.operations.calculate_circuits_preview_operation import (
    CalculateCircuitsPreviewOperation,
)
from CEDElectrical.Application.operations.edit_circuit_properties_and_recalculate_operation import (
    EditCircuitPropertiesAndRecalculateOperation,
)
from CEDElectrical.Application.operations.mark_existing_and_recalculate_operation import (
    MarkExistingAndRecalculateOperation,
)
from CEDElectrical.Application.operations.move_selected_circuits_operation import MoveSelectedCircuitsOperation
from CEDElectrical.Application.operations.set_hidden_alert_types_operation import SetHiddenAlertTypesOperation
from CEDElectrical.Application.operations.set_include_and_recalculate_operation import SetIncludeAndRecalculateOperation
from CEDElectrical.Application.services.operation_registry import OperationRegistry
from CEDElectrical.Infrastructure.Revit.repositories.revit_circuit_repository import RevitCircuitRepository
from CEDElectrical.Infrastructure.Revit.stores.parameter_alert_store import ParameterAlertStore
from CEDElectrical.Infrastructure.Revit.writers.revit_circuit_writer import RevitCircuitWriter


class OperationRunner(object):
    """Executes operation requests through the registry."""

    def __init__(self, registry):
        self.registry = registry

    def run(self, request, doc):
        """Run a request against the active Revit document."""
        operation = self.registry.get(request.operation_key)
        if operation is None:
            raise ValueError('Unknown operation: {}'.format(request.operation_key))
        return operation.execute(request, doc)


def build_default_runner(alert_parameter_name='Circuit Data_CED'):
    """Build default runner with Revit adapters and parameter alert store."""
    registry = OperationRegistry()

    repository = RevitCircuitRepository()
    writer = RevitCircuitWriter()
    alert_store = ParameterAlertStore(parameter_name=alert_parameter_name)

    calc_operation = CalculateCircuitsOperation(repository, writer, alert_store)
    registry.register(calc_operation)
    registry.register(CalculateCircuitsPreviewOperation(repository))
    registry.register(SetHiddenAlertTypesOperation(repository, alert_store))
    registry.register(EditCircuitPropertiesAndRecalculateOperation(calculate_operation=calc_operation))
    registry.register(
        SetIncludeAndRecalculateOperation(
            key='set_neutral_and_recalculate',
            include_param_name='CKT_Include Neutral_CED',
            mode_name='Neutral',
            allowed_branch_types=['BRANCH'],
            calculate_operation=calc_operation,
        )
    )
    registry.register(
        SetIncludeAndRecalculateOperation(
            key='set_ig_and_recalculate',
            include_param_name='CKT_Include Isolated Ground_CED',
            mode_name='Isolated Ground',
            allowed_branch_types=['BRANCH', 'FEEDER', 'XFMR PRI', 'XFMR SEC'],
            calculate_operation=calc_operation,
        )
    )
    registry.register(AutosizeBreakerAndRecalculateOperation(calculate_operation=calc_operation))
    registry.register(MarkExistingAndRecalculateOperation(calculate_operation=calc_operation))
    registry.register(MoveSelectedCircuitsOperation(calculate_operation=calc_operation))
    return OperationRunner(registry)
