# -*- coding: utf-8 -*-
"""In-memory circuit calculation preview (no parameter writeback)."""

from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Model.circuit_settings import IsolatedGroundBehavior, NeutralBehavior
from Snippets import revit_helpers


def _elid_value(item):
    return int(revit_helpers.get_elementid_value(item))


class CalculateCircuitsPreviewOperation(object):
    """Computes calculated circuit values for UI preview only."""

    key = "calculate_circuits_preview"

    def __init__(self, repository):
        self.repository = repository

    def execute(self, request, doc):
        settings = settings_manager.load_circuit_settings(doc)
        circuits = self.repository.get_target_circuits(doc, request.circuit_ids)
        overrides_by_circuit = self._normalize_overrides(request.options.get("preview_values_by_circuit"))
        previews = []

        for circuit in list(circuits or []):
            if circuit is None:
                continue
            cid = int(_elid_value(getattr(circuit, "Id", None)))
            preview_values = dict(overrides_by_circuit.get(cid, {}))
            branch = CircuitBranch(circuit, settings=settings, preview_values=preview_values)

            can_calculate = bool(branch.is_power_circuit and not branch.is_space and not branch.is_spare)
            if can_calculate:
                branch.calculate_hot_wire_size()
                branch.calculate_neutral_wire_size()
                branch.calculate_ground_wire_size()
                branch.calculate_isolated_ground_wire_size()
                branch.calculate_conduit_size()

            values = self._collect_shared_param_values(branch)
            previews.append(self._build_preview_row(branch, values, settings, can_calculate))

        return {
            "status": "ok",
            "previews": previews,
            "settings": {
                "multi_pole_branch_neutral_behavior": settings.multi_pole_branch_neutral_behavior,
                "neutral_behavior": settings.neutral_behavior,
                "isolated_ground_behavior": settings.isolated_ground_behavior,
            },
        }

    def _normalize_overrides(self, raw):
        normalized = {}
        source = dict(raw or {})
        for key, value in list(source.items()):
            try:
                cid = int(key)
            except Exception:
                try:
                    cid = int(str(key or "").strip())
                except Exception:
                    cid = 0
            if cid <= 0:
                continue
            normalized[cid] = dict(value or {})
        return normalized

    def _build_preview_row(self, branch, values, settings, can_calculate):
        notices = []
        collector = getattr(branch, "notices", None)
        if collector and collector.has_items():
            for definition, severity, group, message in collector.items:
                definition_id = ""
                try:
                    definition_id = definition.GetId() if definition else ""
                except Exception:
                    definition_id = ""
                notices.append(
                    {
                        "definition_id": definition_id or "",
                        "severity": severity or "",
                        "group": group or "",
                        "message": message or "",
                    }
                )

        user_override = bool(getattr(branch, "_auto_calculate_override", False))
        hot_cleared = bool(getattr(branch, "_user_clear_hot", False))
        ground_cleared = bool(getattr(branch, "_user_clear_ground", False))
        conduit_cleared = bool(getattr(branch, "_user_clear_conduit", False))
        wire_manual_enabled = bool(user_override and not hot_cleared)
        ground_manual_enabled = bool(wire_manual_enabled and not ground_cleared)
        conduit_manual_enabled = bool(user_override and not conduit_cleared)
        neutral_manual_enabled = bool(
            wire_manual_enabled and settings.neutral_behavior == NeutralBehavior.MANUAL
        )
        isolated_ground_manual_enabled = bool(
            ground_manual_enabled and settings.isolated_ground_behavior == IsolatedGroundBehavior.MANUAL
        )

        circuit_id = int(_elid_value(getattr(branch.circuit, "Id", None)))
        return {
            "circuit_id": circuit_id,
            "panel": branch.panel or "",
            "circuit_number": branch.circuit_number or "",
            "load_name": branch.load_name or "",
            "branch_type": branch.branch_type or "",
            "can_calculate": bool(can_calculate),
            "values": dict(values or {}),
            "editability": {
                "user_override": bool(user_override),
                "wire_manual_enabled": wire_manual_enabled,
                "ground_manual_enabled": ground_manual_enabled,
                "conduit_manual_enabled": conduit_manual_enabled,
                "neutral_size_manual_enabled": neutral_manual_enabled,
                "isolated_ground_size_manual_enabled": isolated_ground_manual_enabled,
                "hot_cleared": bool(hot_cleared),
                "ground_cleared": bool(ground_cleared),
                "conduit_cleared": bool(conduit_cleared),
            },
            "notices": notices,
        }

    def _collect_shared_param_values(self, branch):
        neutral_qty = branch.neutral_wire_quantity or 0
        ig_qty = branch.isolated_ground_wire_quantity or 0
        include_neutral = 1 if neutral_qty > 0 else 0
        include_ig = 1 if ig_qty > 0 else 0

        return {
            "CKT_Circuit Type_CEDT": branch.branch_type,
            "CKT_Panel_CEDT": branch.panel,
            "CKT_Circuit Number_CEDT": branch.circuit_number,
            "CKT_Load Name_CEDT": branch.load_name,
            "CKT_Rating_CED": branch.rating,
            "CKT_Frame_CED": branch.frame,
            "CKT_Length_CED": branch.length,
            "CKT_Schedule Notes_CEDT": branch.circuit_notes,
            "Voltage Drop Percentage_CED": branch.voltage_drop_percentage,
            "CKT_Wire Hot Size_CEDT": branch.hot_wire_size,
            "CKT_Number of Wires_CED": branch.number_of_wires,
            "CKT_Number of Sets_CED": branch.number_of_sets,
            "CKT_Wire Hot Quantity_CED": branch.hot_wire_quantity,
            "CKT_Wire Ground Size_CEDT": branch.ground_wire_size,
            "CKT_Wire Ground Quantity_CED": branch.ground_wire_quantity,
            "CKT_Wire Neutral Size_CEDT": branch.neutral_wire_size,
            "CKT_Wire Neutral Quantity_CED": neutral_qty,
            "CKT_Wire Isolated Ground Size_CEDT": branch.isolated_ground_wire_size,
            "CKT_Wire Isolated Ground Quantity_CED": ig_qty,
            "CKT_Include Neutral_CED": include_neutral,
            "CKT_Include Isolated Ground_CED": include_ig,
            "Wire Material_CEDT": branch.wire_material,
            "Wire Temparature Rating_CEDT": branch.wire_temp_rating,
            "Wire Insulation_CEDT": branch.wire_insulation,
            "Conduit Size_CEDT": branch.conduit_size,
            "Conduit Type_CEDT": branch.conduit_type,
            "Conduit Fill Percentage_CED": branch.conduit_fill_percentage,
            "Wire Size_CEDT": branch.get_wire_size_callout(),
            "Conduit and Wire Size_CEDT": branch.get_conduit_and_wire_size(),
            "Circuit Load Current_CED": branch.circuit_load_current,
            "Circuit Ampacity_CED": branch.circuit_base_ampacity,
            "CKT_Length Makeup_CED": branch.wire_length_makeup,
        }
