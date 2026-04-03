# -*- coding: utf-8 -*-
"""Apply staged circuit property edits, then run calculate operation."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Application.dto.operation_request import OperationRequest
from Snippets import revit_helpers


def _elid_value(item):
    return int(revit_helpers.get_elementid_value(item))


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


class EditCircuitPropertiesAndRecalculateOperation(object):
    """Writes staged user edits on circuits, then recalculates affected circuits."""

    key = "edit_circuit_properties_and_recalculate"

    def __init__(self, calculate_operation):
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        circuits = self._get_target_circuits(doc, request.circuit_ids)
        if not circuits:
            return {"status": "cancelled", "reason": "no_circuits"}

        updates_by_id = self._normalize_updates(request.options.get("updates"))
        if not updates_by_id:
            return {"status": "cancelled", "reason": "no_updates"}

        changed_ids = []
        locked_rows = []

        tg = DB.TransactionGroup(doc, "Edit Circuit Properties + Calculate Circuits")
        tg.Start()
        tx = DB.Transaction(doc, "Edit Circuit Properties")
        tx.Start()
        try:
            for circuit in circuits:
                cid = _elid_value(circuit.Id)
                param_values = updates_by_id.get(cid)
                if not param_values:
                    continue

                if self._is_locked(doc, circuit.Id):
                    locked_rows.append(self._locked_row(circuit, doc))
                    continue

                did_change = False
                for param_name, value in list(param_values.items()):
                    did_change = self._set_param_value(circuit, param_name, value) or did_change
                if did_change:
                    changed_ids.append(cid)

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
                "status": "cancelled",
                "reason": "no_changes",
                "locked_rows": locked_rows,
                "runtime_alert_rows": [],
            }

        calc_request = OperationRequest(
            operation_key="calculate_circuits",
            circuit_ids=changed_ids,
            source=request.source,
            options={
                "show_output": bool(request.options.get("show_output", False)),
                "use_existing_transaction_group": True,
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
            existing = list(calc_result.get("locked_rows") or [])
            calc_result["locked_rows"] = existing + locked_rows
        calc_result["edited_circuits"] = len(changed_ids)
        return calc_result

    def _normalize_updates(self, updates):
        by_id = {}
        for row in list(updates or []):
            if not isinstance(row, dict):
                continue
            try:
                cid = int(row.get("circuit_id") or 0)
            except Exception:
                cid = 0
            if cid <= 0:
                continue
            param_values = dict(row.get("param_values") or {})
            if not param_values:
                continue
            by_id[cid] = param_values
        return by_id

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

    def _set_param_value(self, circuit, param_name, value):
        try:
            param = circuit.LookupParameter(str(param_name or ""))
        except Exception:
            param = None
        if not param:
            return False

        try:
            storage_type = param.StorageType
            if storage_type == DB.StorageType.Integer:
                current = param.AsInteger()
                target = int(round(float(value or 0)))
            elif storage_type == DB.StorageType.Double:
                current = param.AsDouble()
                target = float(value or 0.0)
                if abs(float(current) - float(target)) < 0.000001:
                    return False
            elif storage_type == DB.StorageType.String:
                current = param.AsString() or ""
                target = str(value or "")
            else:
                return False
        except Exception:
            return False

        try:
            if storage_type == DB.StorageType.Integer:
                if int(current) == int(target):
                    return False
                return bool(param.Set(int(target)))
            if storage_type == DB.StorageType.Double:
                return bool(param.Set(float(target)))
            if storage_type == DB.StorageType.String:
                if str(current or "") == str(target or ""):
                    return False
                return bool(param.Set(str(target)))
        except Exception:
            return False
        return False

    def _is_locked(self, doc, eid):
        if not getattr(doc, "IsWorkshared", False):
            return False
        try:
            return DB.WorksharingUtils.GetCheckoutStatus(doc, eid) == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _locked_row(self, circuit, doc):
        panel = ""
        try:
            panel = circuit.BaseEquipment.Name if circuit.BaseEquipment else ""
        except Exception:
            panel = ""
        number = ""
        try:
            number = circuit.CircuitNumber or ""
        except Exception:
            number = ""
        owner = ""
        try:
            owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, circuit.Id).Owner or ""
        except Exception:
            owner = ""
        return {
            "circuit_id": _elid_value(circuit.Id),
            "circuit": "{}-{}".format(panel, number),
            "load_name": getattr(circuit, "LoadName", "") or "",
            "circuit_owner": owner,
            "device_owner": "",
            "sync_writeback": False,
        }

