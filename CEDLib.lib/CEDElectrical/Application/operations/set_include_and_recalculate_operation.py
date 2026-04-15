# -*- coding: utf-8 -*-
"""Set include flags on circuits, then run calculate operation."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from Snippets import revit_helpers


def _elid_value(item):
    return int(revit_helpers.get_elementid_value(item))


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


class SetIncludeAndRecalculateOperation(object):
    """Applies include-neutral/include-IG flags and recalculates affected circuits."""

    def __init__(self, key, include_param_name, mode_name, allowed_branch_types, calculate_operation):
        self.key = key
        self._include_param_name = include_param_name
        self._mode_name = mode_name
        self._allowed_branch_types = set([x.upper() for x in (allowed_branch_types or [])])
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        circuits = self._get_target_circuits(doc, request.circuit_ids)
        if not circuits:
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        updates_by_id = self._build_updates_map(request.options.get('updates'))
        mode = str(request.options.get('mode') or '').lower()
        include_value = None
        if mode in ('add', 'remove'):
            include_value = 1 if mode == 'add' else 0
        if include_value is None and not updates_by_id:
            return {'status': 'cancelled', 'reason': 'invalid_mode'}

        target_ids = []
        target_id_set = set()
        locked_rows = []
        tg = DB.TransactionGroup(doc, '{} + Calculate Circuits'.format(self._mode_name))
        tg.Start()
        tx = DB.Transaction(doc, 'Update Circuit Include Flags ({})'.format(self._mode_name))
        tx.Start()
        try:
            for circuit in circuits:
                blocked_reason = self._block_reason(circuit)
                if blocked_reason:
                    continue

                if self._is_locked(doc, circuit.Id):
                    locked_rows.append(self._locked_row(circuit, doc))
                    continue

                param = None
                try:
                    param = circuit.LookupParameter(self._include_param_name)
                except Exception:
                    param = None
                if not param or param.StorageType != DB.StorageType.Integer:
                    continue

                circuit_id = _elid_value(circuit.Id)
                requested_include = updates_by_id.get(circuit_id, include_value)
                if requested_include not in (0, 1):
                    continue

                current = None
                try:
                    current = param.AsInteger()
                except Exception:
                    current = None
                include_matches = current == requested_include
                explicit_update = circuit_id in updates_by_id
                applied_or_confirmed = False

                if include_matches:
                    applied_or_confirmed = True
                else:
                    try:
                        param.Set(requested_include)
                        applied_or_confirmed = True
                    except Exception:
                        continue

                if not applied_or_confirmed:
                    continue

                if explicit_update or (not include_matches):
                    if circuit_id not in target_id_set:
                        target_id_set.add(circuit_id)
                        target_ids.append(circuit_id)

            tx.Commit()
        except Exception:
            tx.RollBack()
            try:
                tg.RollBack()
            except Exception:
                pass
            raise

        if not target_ids:
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
            circuit_ids=target_ids,
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

    def _build_updates_map(self, updates):
        mapping = {}
        for row in list(updates or []):
            item = dict(row or {})
            try:
                cid = int(item.get('circuit_id') or 0)
            except Exception:
                cid = 0
            if cid <= 0:
                continue
            try:
                include_value = int(item.get('include'))
            except Exception:
                include_value = None
            if include_value not in (0, 1):
                continue
            mapping[cid] = include_value
        return mapping

    def _block_reason(self, circuit):
        try:
            branch = CircuitBranch(circuit, settings=settings_manager.load_circuit_settings(circuit.Document))
        except Exception:
            return 'invalid_circuit'

        branch_type = (branch.branch_type or '').upper()
        if branch_type not in self._allowed_branch_types:
            return 'unsupported_type'
        if self._include_param_name == 'CKT_Include Neutral_CED':
            poles = branch.poles or 0
            if int(poles) <= 1:
                return 'neutral_required_1p'
        if self._include_param_name == 'CKT_Include Isolated Ground_CED':
            override_flag = 0
            try:
                p = circuit.LookupParameter('CKT_User Override_CED')
                override_flag = p.AsInteger() if p else 0
            except Exception:
                override_flag = 0
            ground_size = ''
            try:
                p = circuit.LookupParameter('CKT_Wire Ground Size_CEDT')
                ground_size = (p.AsString() if p else '') or ''
            except Exception:
                ground_size = ''
            if int(override_flag or 0) == 1 and str(ground_size).strip() == '-':
                return 'ig_blocked_ground_clear'
        return None

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


