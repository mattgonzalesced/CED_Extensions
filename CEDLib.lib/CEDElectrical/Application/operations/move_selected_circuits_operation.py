# -*- coding: utf-8 -*-
"""Move selected circuits to a target panel with optional post-move recalculation."""

import Autodesk.Revit.DB.Electrical as DBE

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.move_circuits_to_panel_service import move_circuits_to_panel
from Snippets import revit_helpers
from Snippets._elecutils import panel_has_schedule_view


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


class BufferedMoveOutput(object):
    """Captures move output so callers can decide when to show it."""

    def __init__(self):
        self._events = []

    def linkify(self, element_id):
        return _elid_value(element_id)

    def print_md(self, text):
        self._events.append(("md", str(text or "")))

    def print_table(self, table_data, columns):
        self._events.append(
            (
                "table",
                list(table_data or []),
                list(columns or []),
            )
        )

    def flush_to(self, output):
        if output is None:
            return
        for event in list(self._events or []):
            kind = event[0]
            if kind == "md":
                output.print_md(event[1])
                continue
            if kind == "table":
                output.print_table(event[1], event[2])


class MoveSelectedCircuitsOperation(object):
    """Executes Move Selected Circuits behavior through the operation runner."""

    key = "move_selected_circuits"

    def __init__(self, calculate_operation=None):
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        options = dict(getattr(request, "options", None) or {})
        target_panel_id = int(options.get("target_panel_id", 0) or 0)
        if target_panel_id <= 0:
            raise Exception("Target panel selection is invalid.")

        target_panel = doc.GetElement(_elid_from_value(target_panel_id))
        if target_panel is None:
            raise Exception("Target panel was not found in the active document.")
        if not panel_has_schedule_view(doc, target_panel):
            raise Exception(
                "Target panel has no panel schedule view.\n"
                "Create the panel schedule first, then retry the move."
            )

        circuit_ids = [int(x) for x in list(getattr(request, "circuit_ids", None) or []) if int(x or 0) > 0]
        circuits = []
        pre_on_target_ids = set()
        for cid in list(circuit_ids or []):
            circuit = doc.GetElement(_elid_from_value(int(cid)))
            if not isinstance(circuit, DBE.ElectricalSystem):
                continue
            circuits.append(circuit)
            base_equipment = getattr(circuit, "BaseEquipment", None)
            if base_equipment is None:
                continue
            if _elid_value(getattr(base_equipment, "Id", None)) == target_panel_id:
                pre_on_target_ids.add(_elid_value(circuit.Id))

        if not circuits:
            raise Exception("No valid circuits were found to move.")

        buffered_output = BufferedMoveOutput()
        move_result = move_circuits_to_panel(circuits, target_panel, doc, buffered_output)

        moved_ids = []
        for circuit in list(circuits or []):
            cid = _elid_value(circuit.Id)
            if cid in pre_on_target_ids:
                continue
            base_equipment = getattr(circuit, "BaseEquipment", None)
            if base_equipment is None:
                continue
            if _elid_value(getattr(base_equipment, "Id", None)) == target_panel_id:
                moved_ids.append(cid)
        moved_ids = sorted(list(set([int(x) for x in list(moved_ids or []) if int(x) > 0])))

        recalc_result = None
        recalc_error = None
        if bool(options.get("recalculate", False)) and moved_ids:
            if self._calculate_operation is None:
                recalc_error = Exception("Calculate operation is not configured for move execution.")
            else:
                try:
                    calc_request = OperationRequest(
                        operation_key="calculate_circuits",
                        circuit_ids=moved_ids,
                        source=getattr(request, "source", "ribbon"),
                        options={
                            "show_output": bool(options.get("show_recalc_output", False)),
                        },
                    )
                    recalc_result = self._calculate_operation.execute(calc_request, doc)
                except Exception as ex:
                    recalc_error = ex

        return {
            "move_result": move_result,
            "buffered_output": buffered_output,
            "moved_ids": moved_ids,
            "recalc_result": recalc_result,
            "recalc_error": recalc_error,
        }
