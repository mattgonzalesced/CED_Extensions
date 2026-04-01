# -*- coding: utf-8 -*-
"""Service-layer functions for Alerts Browser."""

import json

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_runner import build_default_runner
from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.alerts import get_alert_definition
from Snippets.circuit_ui_actions import format_writeback_lock_reason
from alerts_browser_view_models import AlertCircuitItem
from alerts_browser_view_models import AlertRow


def _lookup_param_text(element, name):
    param = element.LookupParameter(name)
    if not param:
        return None
    value = param.AsString()
    if value is None:
        value = param.AsValueString()
    return value


def _read_alert_payload(circuit, alert_data_param):
    raw = _lookup_param_text(circuit, alert_data_param)
    if not raw:
        return None
    try:
        return json.loads(raw)
    except Exception:
        return None


def _payload_alert_records(payload):
    if not isinstance(payload, dict):
        return []
    alerts = payload.get("alerts")
    return alerts if isinstance(alerts, list) else []


def _payload_hidden_ids(payload):
    if not isinstance(payload, dict):
        return set()
    hidden = payload.get("hidden_definition_ids")
    if not isinstance(hidden, list):
        return set()
    return set([x for x in hidden if x])


def _alert_rows_from_payload(payload):
    hidden_ids = _payload_hidden_ids(payload)
    rows = []
    for item in _payload_alert_records(payload):
        if not isinstance(item, dict):
            continue
        definition_id = item.get("definition_id") or item.get("id")
        is_hidden = definition_id in hidden_ids
        severity = str(item.get("severity") or "NONE").upper()
        group = str(item.get("group") or "Other")
        definition = get_alert_definition(definition_id) if definition_id else None
        message_value = item.get("message")
        if message_value:
            if isinstance(message_value, (dict, list)):
                try:
                    text = json.dumps(message_value, ensure_ascii=False)
                except Exception:
                    text = str(message_value)
            else:
                text = str(message_value)
        elif definition:
            text = definition.GetDescriptionText()
        elif definition_id:
            text = definition_id
        else:
            text = "Unmapped alert"
        row = AlertRow(severity, group, definition_id or "-", text)
        row.is_hidden = bool(is_hidden)
        rows.append(row)
    return rows


def _build_writeback_lock_map(doc, circuits, idval_fn, lock_repository):
    if doc is None or not getattr(doc, "IsWorkshared", False):
        return {}
    circuit_list = [c for c in list(circuits or []) if c is not None]
    if not circuit_list:
        return {}
    settings = settings_manager.load_circuit_settings(doc)
    if settings is None:
        return {}
    _, _, locked_rows = lock_repository.partition_locked_elements(
        doc,
        circuit_list,
        settings,
        collect_all_device_owners=False,
    )
    lock_map = {}
    for row in list(locked_rows or []):
        try:
            cid = int((row or {}).get("circuit_id") or 0)
        except Exception:
            cid = 0
        if cid <= 0:
            continue
        lock_map[cid] = row
    return lock_map


def build_snapshot(doc, alert_data_param, idval_fn, lock_repository):
    if doc is None:
        return {"doc_title": "-", "items": []}
    circuits = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DBE.ElectricalSystem)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    circuits.sort(
        key=lambda c: (
            (getattr(getattr(c, "BaseEquipment", None), "Name", "") or ""),
            (getattr(c, "StartSlot", 0) or 0),
            (getattr(c, "LoadName", "") or ""),
        )
    )
    lock_map = _build_writeback_lock_map(doc, circuits, idval_fn, lock_repository)
    items = []
    for circuit in circuits:
        rows = _alert_rows_from_payload(_read_alert_payload(circuit, alert_data_param))
        if not rows:
            continue
        circuit_id = idval_fn(circuit.Id)
        lock_row = lock_map.get(circuit_id)
        blocked = lock_row is not None
        reason = format_writeback_lock_reason(lock_row) if blocked else ""
        items.append(AlertCircuitItem(circuit, circuit_id, rows, blocked=blocked, block_reason=reason))
    return {
        "doc_title": getattr(doc, "Title", "-") or "-",
        "items": items,
    }


def recalculate_and_snapshot(doc, circuit_id, alert_data_param, idval_fn, lock_repository):
    request = OperationRequest(
        operation_key="calculate_circuits",
        circuit_ids=[int(circuit_id)],
        source="alerts_browser",
        options={"show_output": False},
    )
    runner = build_default_runner(alert_parameter_name=alert_data_param)
    operation_result = runner.run(request, doc) or {}
    return {
        "operation_result": operation_result,
        "snapshot": build_snapshot(doc, alert_data_param, idval_fn, lock_repository),
        "circuit_id": int(circuit_id),
    }
