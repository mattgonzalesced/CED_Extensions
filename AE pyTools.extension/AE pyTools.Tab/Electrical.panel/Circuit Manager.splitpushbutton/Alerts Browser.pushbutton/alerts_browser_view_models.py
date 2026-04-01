# -*- coding: utf-8 -*-
"""View-models for Alerts Browser."""


class AlertRow(object):
    def __init__(self, severity, group, definition_id, message):
        self.severity = str(severity or "NONE")
        self.group = str(group or "Other")
        self.definition_id = str(definition_id or "-")
        self.message = str(message or "")
        self.is_hidden = False


class AlertCircuitItem(object):
    def __init__(self, circuit, circuit_id, rows, blocked=False, block_reason=""):
        self.circuit = circuit
        self.circuit_id = int(circuit_id or 0)
        self.panel = "-"
        if getattr(circuit, "BaseEquipment", None):
            self.panel = getattr(circuit.BaseEquipment, "Name", self.panel) or self.panel
        self.circuit_number = getattr(circuit, "CircuitNumber", "") or ""
        self.load_name = getattr(circuit, "LoadName", "") or ""
        self.panel_ckt_text = "{} / {}".format(self.panel or "-", self.circuit_number or "-")
        self.rows = list(rows or [])
        self.active_rows = [x for x in self.rows if not bool(getattr(x, "is_hidden", False))]
        self.hidden_rows = [x for x in self.rows if bool(getattr(x, "is_hidden", False))]
        self.total_count = len(self.rows)
        self.active_count = len(self.active_rows)
        self.hidden_count = len(self.hidden_rows)
        self.counts_text = "Alerts: {} | Active: {} | Hidden: {}".format(
            self.total_count,
            self.active_count,
            self.hidden_count,
        )
        self.recalc_blocked = bool(blocked)
        self.recalc_block_reason = str(block_reason or "")
