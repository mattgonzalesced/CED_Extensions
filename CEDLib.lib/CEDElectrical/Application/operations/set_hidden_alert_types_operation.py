# -*- coding: utf-8 -*-
"""Operation for persisting per-circuit hidden alert type selections."""

from pyrevit import DB


class SetHiddenAlertTypesOperation(object):
    """Updates hidden alert definition ids for target circuits."""
    key = 'set_hidden_alert_types'
    _HIDABLE_ALERT_IDS = set([
        'Design.NonStandardOCPRating',
        'Design.BreakerLugSizeLimitOverride',
        'Design.BreakerLugQuantityLimitOverride',
        'Calculations.BreakerLugSizeLimit',
        'Calculations.BreakerLugQuantityLimit',
    ])

    def __init__(self, repository, alert_store):
        self.repository = repository
        self.alert_store = alert_store

    def execute(self, request, doc):
        circuits = self.repository.get_target_circuits(doc, request.circuit_ids)
        if not circuits:
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        requested_hidden = set(request.options.get('hidden_definition_ids') or []).intersection(self._HIDABLE_ALERT_IDS)

        tx = DB.Transaction(doc, 'Update Hidden Circuit Alerts')
        tx.Start()
        updated = 0
        try:
            for circuit in circuits:
                payload = self.alert_store.read_alert_payload(circuit) or {}
                alerts = payload.get('alerts') if isinstance(payload, dict) else None
                if not isinstance(alerts, list):
                    alerts = []

                present_ids = set()
                for item in alerts:
                    if not isinstance(item, dict):
                        continue
                    definition_id = item.get('definition_id') or item.get('id')
                    if definition_id:
                        present_ids.add(definition_id)

                hidden_ids = sorted(list(requested_hidden.intersection(present_ids)))
                payload['hidden_definition_ids'] = hidden_ids
                payload['version'] = payload.get('version') or 1

                if self.alert_store.write_alert_payload(circuit, payload):
                    updated += 1
            tx.Commit()
        except Exception:
            tx.RollBack()
            raise

        return {'status': 'ok', 'updated_circuits': updated}
