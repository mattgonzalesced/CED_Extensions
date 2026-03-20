# -*- coding: utf-8 -*-
"""Parameter-backed alert persistence adapter."""

import json

from CEDElectrical.Application.contracts.alert_store import IAlertStore


class ParameterAlertStore(IAlertStore):
    """Stores serialized alert payloads in a multiline text parameter."""

    def __init__(self, parameter_name='Circuit Data_CED'):
        self.parameter_name = parameter_name

    def _resolve_param(self, circuit):
        """Resolve writable parameter on circuit."""
        param = None
        try:
            param = circuit.LookupParameter(self.parameter_name)
        except Exception:
            param = None
        return param

    def read_alert_payload(self, circuit):
        """Read and deserialize payload from circuit parameter."""
        param = self._resolve_param(circuit)
        if not param:
            return None
        try:
            text = param.AsString()
            if not text:
                return None
            return json.loads(text)
        except Exception:
            return None

    def write_alert_payload(self, circuit, payload):
        """Serialize and write payload to circuit parameter."""
        param = self._resolve_param(circuit)
        if not param:
            return False
        try:
            text = json.dumps(payload, indent=2, sort_keys=True)
            param.Set(text)
            return True
        except Exception:
            return False

    def clear_alert_payload(self, circuit):
        """Clear persisted payload from circuit parameter."""
        param = self._resolve_param(circuit)
        if not param:
            return False
        try:
            param.Set('')
            return True
        except Exception:
            return False
