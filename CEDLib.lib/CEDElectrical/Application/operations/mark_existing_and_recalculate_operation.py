# -*- coding: utf-8 -*-
"""Mark circuits as existing, then run calculate operation."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB
from pyrevit.compat import get_elementid_value_func, get_elementid_from_value_func

from CEDElectrical.Application.dto.operation_request import OperationRequest

_get_elid_value = get_elementid_value_func()
_get_elid_from_value = get_elementid_from_value_func()


def _elid_value(item):
    try:
        return int(_get_elid_value(item))
    except Exception:
        return int(getattr(item, 'IntegerValue', 0))


def _elid_from_value(value):
    return _get_elid_from_value(int(value))


class MarkExistingAndRecalculateOperation(object):
    """Sets override/notes/clear flags then recalculates affected circuits."""

    key = 'mark_existing_and_recalculate'

    def __init__(self, calculate_operation):
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        circuits = self._get_target_circuits(doc, request.circuit_ids)
        if not circuits:
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        mode = str(request.options.get('mode', 'existing') or 'existing').strip().lower()
        if mode not in ('existing', 'new'):
            mode = 'existing'
        set_notes = bool(request.options.get('set_notes', True))
        notes_text = '' if mode == 'new' else 'EX'
        clear_wire = bool(request.options.get('clear_wire', False))
        clear_conduit = bool(request.options.get('clear_conduit', False))
        user_override_value = 0 if mode == 'new' else 1

        changed_ids = []
        locked_rows = []
        tg = DB.TransactionGroup(doc, 'Mark New/Existing + Calculate Circuits')
        tg.Start()
        tx = DB.Transaction(doc, 'Mark New/Existing Circuit Data')
        tx.Start()
        try:
            for circuit in circuits:
                if self._is_locked(doc, circuit.Id):
                    locked_rows.append(self._locked_row(circuit, doc))
                    continue

                did_change = False
                did_change = self._set_int_param(circuit, 'CKT_User Override_CED', user_override_value) or did_change

                if set_notes:
                    did_change = self._set_schedule_notes(circuit, notes_text) or did_change
                    did_change = self._set_str_param(circuit, 'CKT_Schedule Notes_CEDT', notes_text) or did_change

                if mode == 'existing' and clear_wire:
                    did_change = self._set_str_param(circuit, 'CKT_Wire Hot Size_CEDT', '-') or did_change
                    did_change = self._set_str_param(circuit, 'Wire Hot Size_CEDT', '-') or did_change

                if mode == 'existing' and clear_conduit:
                    did_change = self._set_str_param(circuit, 'Conduit Size_CEDT', '-') or did_change

                if did_change:
                    changed_ids.append(_elid_value(circuit.Id))

            tx.Commit()
        except Exception:
            tx.RollBack()
            try:
                tg.RollBack()
            except Exception:
                pass
            raise

        if not changed_ids:
            try:
                tg.RollBack()
            except Exception:
                pass
            return {
                'status': 'cancelled',
                'reason': 'no_changes',
                'locked_rows': locked_rows,
                'runtime_alert_rows': [],
            }

        calc_request = OperationRequest(
            operation_key='calculate_circuits',
            circuit_ids=changed_ids,
            source=request.source,
            options={
                'show_output': bool(request.options.get('show_output', False)),
                'use_existing_transaction_group': True,
            },
        )
        try:
            calc_result = self._calculate_operation.execute(calc_request, doc) or {}
            tg.Assimilate()
        except Exception:
            try:
                tg.RollBack()
            except Exception:
                pass
            raise

        if locked_rows:
            existing = list(calc_result.get('locked_rows') or [])
            calc_result['locked_rows'] = existing + locked_rows
        return calc_result

    def _get_target_circuits(self, doc, circuit_ids):
        circuits = []
        for raw_id in list(circuit_ids or []):
            try:
                el = doc.GetElement(_elid_from_value(raw_id))
            except Exception:
                el = None
            if isinstance(el, DBE.ElectricalSystem):
                circuits.append(el)
        return circuits

    def _set_schedule_notes(self, circuit, value):
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
        except Exception:
            param = None
        if not param or param.StorageType != DB.StorageType.String:
            return False
        try:
            current = param.AsString() or ''
        except Exception:
            current = ''
        if current == value:
            return False
        try:
            return bool(param.Set(value))
        except Exception:
            return False

    def _set_str_param(self, circuit, param_name, value):
        try:
            param = circuit.LookupParameter(param_name)
        except Exception:
            param = None
        if not param or param.StorageType != DB.StorageType.String:
            return False
        try:
            current = param.AsString() or ''
        except Exception:
            current = ''
        if current == value:
            return False
        try:
            return bool(param.Set(value))
        except Exception:
            return False

    def _set_int_param(self, circuit, param_name, value):
        try:
            param = circuit.LookupParameter(param_name)
        except Exception:
            param = None
        if not param or param.StorageType != DB.StorageType.Integer:
            return False
        try:
            current = param.AsInteger()
        except Exception:
            current = None
        if current == int(value):
            return False
        try:
            return bool(param.Set(int(value)))
        except Exception:
            return False

    def _is_locked(self, doc, eid):
        if not getattr(doc, 'IsWorkshared', False):
            return False
        try:
            return DB.WorksharingUtils.GetCheckoutStatus(doc, eid) == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _locked_row(self, circuit, doc):
        panel = ''
        try:
            panel = circuit.BaseEquipment.Name if circuit.BaseEquipment else ''
        except Exception:
            panel = ''
        number = ''
        try:
            number = circuit.CircuitNumber or ''
        except Exception:
            number = ''
        owner = ''
        try:
            owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, circuit.Id).Owner or ''
        except Exception:
            owner = ''
        return {
            'circuit_id': _elid_value(circuit.Id),
            'circuit': '{}-{}'.format(panel, number),
            'load_name': getattr(circuit, 'LoadName', '') or '',
            'circuit_owner': owner,
            'device_owner': '',
            'sync_writeback': False,
        }

