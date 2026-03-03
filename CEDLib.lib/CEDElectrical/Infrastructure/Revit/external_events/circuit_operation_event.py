# -*- coding: utf-8 -*-
"""ExternalEvent gateway for circuit operations triggered by modeless UI."""

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_runner import build_default_runner


class CircuitOperationExternalEventGateway(object):
    """Queues operation requests and executes them through a shared ExternalEvent."""

    def __init__(self, logger=None, alert_parameter_name='Circuit Data_CED'):
        self.logger = logger
        self.alert_parameter_name = alert_parameter_name
        self._pending = None
        self._handler = _CircuitOperationHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def is_busy(self):
        return self._pending is not None

    def raise_operation(self, operation_key, circuit_ids, source='pane', options=None, callback=None):
        if self._pending is not None:
            return False

        request = OperationRequest(
            operation_key=operation_key,
            circuit_ids=list(circuit_ids or []),
            source=source,
            options=dict(options or {}),
        )
        self._pending = {
            'request': request,
            'callback': callback,
        }
        self._event.Raise()
        return True

    def _consume_pending(self):
        pending = self._pending
        self._pending = None
        return pending


class _CircuitOperationHandler(IExternalEventHandler):
    """Executes queued operation requests in valid Revit API context."""

    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, application):
        pending = self._gateway._consume_pending()
        if not pending:
            return

        callback = pending.get('callback')
        request = pending.get('request')

        status = 'ok'
        result = None
        error = None
        try:
            uidoc = application.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if doc is None:
                raise Exception('No active Revit document available.')

            runner = build_default_runner(alert_parameter_name=self._gateway.alert_parameter_name)
            result = runner.run(request, doc)
        except Exception as ex:
            status = 'error'
            error = ex
            if self._gateway.logger:
                self._gateway.logger.exception('External operation failed: %s', ex)

        if callback:
            try:
                callback(status, request, result, error)
            except Exception as cb_ex:
                if self._gateway.logger:
                    self._gateway.logger.exception('External operation callback failed: %s', cb_ex)

    def GetName(self):
        return 'CED Circuit Operation External Event'

