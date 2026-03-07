# -*- coding: utf-8 -*-
"""Apply breaker/frame updates and run calculate operation."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import CircuitBranch


class AutosizeBreakerAndRecalculateOperation(object):
    """Writes selected breaker/frame values then recalculates circuits."""

    key = 'autosize_breaker_and_recalculate'

    _ALLOWED_TYPES = set(['BRANCH', 'FEEDER', 'XFMR PRI', 'XFMR SEC'])

    def __init__(self, calculate_operation):
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        updates = list(request.options.get('updates') or [])
        if not updates:
            return {'status': 'cancelled', 'reason': 'no_updates'}

        by_id = {}
        for row in updates:
            try:
                cid = int(row.get('circuit_id'))
            except Exception:
                continue
            by_id[cid] = row

        circuits = []
        for cid in by_id.keys():
            try:
                el = doc.GetElement(DB.ElementId(cid))
            except Exception:
                el = None
            if isinstance(el, DBE.ElectricalSystem):
                circuits.append(el)
        if not circuits:
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        changed_ids = []
        locked_rows = []
        tg = DB.TransactionGroup(doc, 'Auto Size Breaker/Frame + Calculate Circuits')
        tg.Start()
        tx = DB.Transaction(doc, 'Auto Size Breaker/Frame')
        tx.Start()
        try:
            for circuit in circuits:
                branch_type = self._branch_type(circuit)
                if branch_type not in self._ALLOWED_TYPES:
                    continue
                if self._is_locked(doc, circuit.Id):
                    locked_rows.append(self._locked_row(circuit, doc))
                    continue

                spec = by_id.get(circuit.Id.IntegerValue) or {}
                set_rating = bool(spec.get('set_rating', True))
                set_frame = bool(spec.get('set_frame', True))
                if not (set_rating or set_frame):
                    continue

                did_change = False
                if set_rating:
                    did_change = self._set_numeric(circuit, 'Rating', 'RBS_ELEC_CIRCUIT_RATING_PARAM', spec.get('rating')) or did_change
                if set_frame:
                    did_change = self._set_numeric(circuit, 'Frame', 'RBS_ELEC_CIRCUIT_FRAME_PARAM', spec.get('frame')) or did_change
                if did_change:
                    changed_ids.append(circuit.Id.IntegerValue)
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

    def _branch_type(self, circuit):
        try:
            settings = settings_manager.load_circuit_settings(circuit.Document)
            branch = CircuitBranch(circuit, settings=settings)
            return (branch.branch_type or '').upper()
        except Exception:
            return ''

    def _is_locked(self, doc, eid):
        if not getattr(doc, 'IsWorkshared', False):
            return False
        try:
            return DB.WorksharingUtils.GetCheckoutStatus(doc, eid) == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _set_numeric(self, circuit, prop_name, bip_name, value):
        try:
            numeric = float(value)
        except Exception:
            return False

        changed = False
        try:
            current = getattr(circuit, prop_name)
            if current is None or abs(float(current) - numeric) > 0.0001:
                setattr(circuit, prop_name, numeric)
                changed = True
        except Exception:
            pass

        try:
            bip = getattr(DB.BuiltInParameter, bip_name)
        except Exception:
            bip = None
        if bip is not None:
            try:
                param = circuit.get_Parameter(bip)
            except Exception:
                param = None
            if param:
                try:
                    if param.StorageType == DB.StorageType.Double:
                        cur = param.AsDouble()
                        if cur is None or abs(float(cur) - numeric) > 0.0001:
                            param.Set(numeric)
                            changed = True
                    elif param.StorageType == DB.StorageType.Integer:
                        iv = int(round(numeric))
                        cur = param.AsInteger()
                        if cur != iv:
                            param.Set(iv)
                            changed = True
                except Exception:
                    pass
        return changed

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
            'circuit': '{}-{}'.format(panel, number),
            'circuit_owner': owner,
            'device_owner': '',
        }
