# -*- coding: utf-8 -*-
"""Move Selected Circuits legacy implementation extracted from Snippets._elecutils."""

import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.DB import Transaction, BuiltInParameter
from pyrevit import script, forms, DB

from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Model.panel_schedule_manager import PanelScheduleManager
from Snippets import revit_helpers

logger = script.get_logger()


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from(value):
    return revit_helpers.elementid_from_value(value)

def move_circuits_to_panel(circuits, target_panel, doc, output):
    """Move selected circuits and optionally replace default spares/spaces when target is full."""

    def _safe_text(value, fallback=""):
        try:
            if value is None:
                return fallback
            return str(value)
        except Exception:
            return fallback

    def _panel_id(element):
        try:
            return int(_elid_value(getattr(element, "Id", None)))
        except Exception:
            return 0

    def _safe_circuit_number(circuit):
        try:
            param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
            if param and param.HasValue:
                return _safe_text(param.AsString(), "-")
        except Exception:
            pass
        return "-"

    def _safe_panel_name(circuit):
        try:
            base = getattr(circuit, "BaseEquipment", None)
            if base is not None:
                return _safe_text(getattr(base, "Name", None), "N/A")
        except Exception:
            pass
        return "N/A"

    def _run_select_panel_moves(allow_partial=False, sort_by_poles=False, phase="primary"):
        def _resolve_start_slot_on_target(circuit, hint_slot=0):
            slot_hint = int(hint_slot or 0)
            if slot_hint > 0 and target_option is not None and bool(ps_repo.is_slot_valid_for_option(target_option, slot_hint)):
                return slot_hint
            if schedule_view is None:
                return slot_hint
            cid_target = int(_elid_value(getattr(circuit, "Id", None)))
            if cid_target <= 0:
                return slot_hint
            slot_scan = []
            if target_option is not None:
                slot_scan = list(ps_repo.get_option_slot_order(target_option, include_excess=True) or [])
            elif slot_hint > 0:
                slot_scan = [slot_hint]
            if slot_hint > 0 and int(slot_hint) not in [int(x) for x in list(slot_scan or [])]:
                slot_scan.insert(0, int(slot_hint))
            for slot in list(slot_scan or []):
                slot_value = int(slot or 0)
                if slot_value <= 0:
                    continue
                cells = list(ps_repo.get_cells_by_slot_number(schedule_view, slot_value) or [])
                if not cells:
                    continue
                for row, col in list(cells or []):
                    try:
                        cell_id = schedule_view.GetCircuitIdByCell(int(row), int(col))
                    except Exception:
                        cell_id = DB.ElementId.InvalidElementId
                    if cell_id is None or cell_id == DB.ElementId.InvalidElementId:
                        continue
                    if int(_elid_value(cell_id)) == int(cid_target):
                        return int(slot_value)
            return slot_hint

        def _validate_target_assignment(circuit):
            if target_option is None:
                return True, 0, []
            start_slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
            start_slot = int(_resolve_start_slot_on_target(circuit, hint_slot=start_slot))
            poles = int(max(1, _get_circuit_poles(circuit)))
            covered_valid = list(
                ps_repo.get_slot_span_slots_for_option(
                    target_option,
                    start_slot=int(start_slot),
                    pole_count=int(poles),
                    require_valid=True,
                )
                or []
            )
            if covered_valid:
                return True, int(start_slot), [int(x) for x in covered_valid if int(x) > 0]
            covered_all = list(
                ps_repo.get_slot_span_slots_for_option(
                    target_option,
                    start_slot=int(start_slot),
                    pole_count=int(poles),
                    require_valid=False,
                )
                or []
            )
            return False, int(start_slot), [int(x) for x in covered_all if int(x) > 0]

        def _try_revert_to_original_panel(snap):
            try:
                old_panel = snap.get("old_panel_element")
                if old_panel is None:
                    return False
                result = snap["circuit"].SelectPanel(old_panel)
                if isinstance(result, bool) and not result:
                    return False
                base_after = getattr(snap["circuit"], "BaseEquipment", None)
                return bool(int(_panel_id(base_after)) == int(_panel_id(old_panel)))
            except Exception:
                return False

        planner_available = None
        planner_slot_order = []
        if target_option is not None:
            planner_slot_order = [int(x) for x in list(ps_repo.get_option_slot_order(target_option, include_excess=False) or []) if int(x) > 0]
            planner_available = set([int(x) for x in list(_collect_empty_slots(target_option) or []) if int(x) > 0])

        def _plan_next_span(circuit):
            if planner_available is None or not planner_slot_order:
                return 0, []
            pole_count = int(max(1, _get_circuit_poles(circuit)))
            for start in list(planner_slot_order or []):
                covered = list(
                    ps_repo.get_slot_span_slots_for_option(
                        target_option,
                        start_slot=int(start),
                        pole_count=int(pole_count),
                        require_valid=True,
                    )
                    or []
                )
                covered = [int(x) for x in list(covered or []) if int(x) > 0]
                if not covered:
                    continue
                if all(int(slot) in planner_available for slot in list(covered or [])):
                    return int(start), list(covered)
            return 0, []

        move_list = list(circuits or [])
        if bool(sort_by_poles):
            move_list = sorted(move_list, key=lambda c: int(_get_circuit_poles(c)), reverse=True)
        phase_name = _safe_text(phase, "primary")
        target_id = _panel_id(target_panel)
        try:
            logger.info(
                "[MoveSelectedCircuits] SelectPanel phase=%s start requested=%s allow_partial=%s sort_by_poles=%s regenerate=%s",
                phase_name,
                int(len(move_list)),
                bool(allow_partial),
                bool(sort_by_poles),
                False,
            )
        except Exception:
            pass
        snapshots = []
        for ckt in move_list:
            old_panel_element = getattr(ckt, "BaseEquipment", None)
            snapshots.append({
                "circuit": ckt,
                "old_panel": _safe_panel_name(ckt),
                "old_circuit_number": _safe_circuit_number(ckt),
                "old_panel_element": old_panel_element,
            })
        data_rows = []
        failed_rows = []
        for snap in snapshots:
            try:
                _, planned_span = _plan_next_span(snap["circuit"])
                if planner_available is not None and not planned_span:
                    raise Exception(
                        "Insufficient valid slot capacity on target panel for this circuit."
                    )
                result = snap["circuit"].SelectPanel(target_panel)
                if isinstance(result, bool) and not result:
                    raise Exception("SelectPanel returned False.")
                base_after = getattr(snap["circuit"], "BaseEquipment", None)
                if int(_panel_id(base_after)) != int(target_id):
                    raise Exception("SelectPanel did not place circuit on target panel.")
                is_valid, start_slot, covered_slots = _validate_target_assignment(snap["circuit"])
                if not bool(is_valid) and (int(start_slot) <= 0 or not list(covered_slots or [])):
                    prior_start = int(start_slot or 0)
                    prior_slots = [int(x) for x in list(covered_slots or []) if int(x) > 0]
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                    is_valid, start_slot, covered_slots = _validate_target_assignment(snap["circuit"])
                    try:
                        logger.info(
                            "[MoveSelectedCircuits] SelectPanel validation retry regenerate prior_start=%s prior_slots=%s new_start=%s new_slots=%s",
                            int(prior_start),
                            ",".join([str(int(x)) for x in list(prior_slots or [])]) or "-",
                            int(start_slot or 0),
                            ",".join([str(int(x)) for x in list(covered_slots or []) if int(x) > 0]) or "-",
                        )
                    except Exception:
                        pass
                if not bool(is_valid):
                    raise Exception(
                        "SelectPanel placed circuit outside usable target capacity. start_slot={0} covered_slots={1}".format(
                            int(start_slot),
                            ",".join([str(int(x)) for x in list(covered_slots or [])]) or "-",
                        )
                    )
                if planner_available is not None:
                    for slot in list(planned_span or []):
                        planner_available.discard(int(slot))
            except Exception as ex:
                if not bool(allow_partial):
                    raise
                reverted = _try_revert_to_original_panel(snap)
                if not bool(reverted):
                    raise Exception(
                        "{0} (revert to original panel failed)".format(
                            _safe_text(ex, "Move failed.")
                        )
                    )
                prev_circuit = "{} / {}".format(snap["old_panel"], snap["old_circuit_number"])
                err_text = _safe_text(ex, "Move failed.")
                failed_rows.append([
                    output.linkify(snap["circuit"].Id),
                    prev_circuit,
                    err_text,
                ])
                continue
            new_circuit_number = _safe_circuit_number(snap["circuit"])
            prev_circuit = "{} / {}".format(snap["old_panel"], snap["old_circuit_number"])
            new_circuit = "{} / {}".format(_safe_text(getattr(target_panel, "Name", None), "N/A"), new_circuit_number)
            data_rows.append([output.linkify(snap["circuit"].Id), prev_circuit, new_circuit])
        try:
            logger.info(
                "[MoveSelectedCircuits] SelectPanel phase=%s done requested=%s moved=%s failed=%s regenerate=%s",
                phase_name,
                int(len(snapshots)),
                int(len(data_rows)),
                int(len(failed_rows)),
                False,
            )
        except Exception:
            pass
        return data_rows, failed_rows

    def _get_panel_schedule_view(panel):
        panel_id = _elid_value(getattr(panel, "Id", None))
        if panel_id <= 0:
            return None
        mapped = ps_repo.map_panel_schedule_views(doc, panels=[panel])
        return mapped.get(panel_id)

    def _get_target_option(panel, schedule_view):
        panel_id = _panel_id(panel)
        if panel_id <= 0:
            return None
        options = list(ps_repo.collect_panel_equipment_options(doc, panels=[panel], include_without_schedule=True) or [])
        for option in options:
            if int(option.get("panel_id", 0) or 0) != panel_id:
                continue
            if schedule_view is not None:
                try:
                    ps_repo.attach_schedule_to_option(doc, option, schedule_view)
                except Exception:
                    pass
            return option
        return None

    def _get_circuit_poles(circuit):
        for attr in ("PolesNumber", "NumberOfPoles"):
            try:
                value = getattr(circuit, attr, None)
                if value is not None:
                    return int(max(1, value))
            except Exception:
                continue
        try:
            param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
            if param and param.HasValue:
                return int(max(1, param.AsInteger()))
        except Exception:
            pass
        return 1

    def _set_circuit_poles(circuit, poles):
        try:
            target = int(max(1, poles or 1))
        except Exception:
            target = 1
        for attr in ("NumberOfPoles", "PolesNumber"):
            try:
                setattr(circuit, attr, int(target))
                return True
            except Exception:
                continue
        try:
            param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
            if param and not bool(getattr(param, "IsReadOnly", False)):
                param.Set(int(target))
                return True
        except Exception:
            pass
        return False

    def _covered_slots(schedule_view, circuit, start_slot):
        slot_value = int(start_slot or 0)
        if slot_value <= 0:
            return []
        try:
            layout = ps_repo.get_schedule_layout_info(schedule_view)
            max_slot = int(layout.get("max_slot", 0) or 0)
            sort_mode = layout.get("sort_mode", ps_repo.SORT_MODE_PANELBOARD_ACROSS)
        except Exception:
            max_slot = 0
            sort_mode = ps_repo.SORT_MODE_PANELBOARD_ACROSS
        poles = _get_circuit_poles(circuit)
        slots = ps_repo.get_slot_span_slots(
            start_slot=slot_value,
            pole_count=poles,
            max_slot=max_slot,
            sort_mode=sort_mode,
        )
        if not slots:
            return [slot_value]
        return [int(x) for x in list(slots or []) if int(x) > 0]

    def _set_slot_locked(schedule_view, slot, is_locked):
        cells = list(ps_repo.get_cells_by_slot_number(schedule_view, int(slot or 0)) or [])
        methods = ("SetSlotLocked", "SetLockSlot", "SetCellLocked")
        for row, col in cells:
            for method_name in methods:
                method = getattr(schedule_view, method_name, None)
                if method is None:
                    continue
                for args in (
                    (int(row), int(col), bool(is_locked)),
                    (int(row), int(col)),
                    (int(slot), bool(is_locked)),
                    (int(slot),),
                ):
                    try:
                        result = method(*args)
                        if isinstance(result, bool) and not result:
                            continue
                        return True
                    except Exception:
                        continue
        return False

    def _slot_is_occupied(schedule_view, slot):
        for row, col in list(ps_repo.get_cells_by_slot_number(schedule_view, int(slot or 0)) or []):
            try:
                cid = schedule_view.GetCircuitIdByCell(int(row), int(col))
            except Exception:
                cid = DB.ElementId.InvalidElementId
            if cid is not None and cid != DB.ElementId.InvalidElementId:
                return True
        return False

    def _get_slot_circuit(schedule_view, slot):
        for row, col in list(ps_repo.get_cells_by_slot_number(schedule_view, int(slot or 0)) or []):
            try:
                cid = schedule_view.GetCircuitIdByCell(int(row), int(col))
            except Exception:
                cid = DB.ElementId.InvalidElementId
            if cid is None or cid == DB.ElementId.InvalidElementId:
                continue
            element = doc.GetElement(cid)
            if isinstance(element, DBE.ElectricalSystem):
                return element
        return None

    def _add_special_to_slot(schedule_view, slot, kind, poles=1):
        action = str(kind or "").strip().lower()
        if action not in ("spare", "space"):
            raise Exception("Unsupported special row type: {0}".format(kind))
        method_names = ("AddSpare",) if action == "spare" else ("AddSpace",)
        slot_value = int(slot or 0)
        cells = list(ps_repo.get_cells_by_slot_number(schedule_view, slot_value) or [])
        if not cells:
            try:
                empties = dict(ps_repo.gather_empty_slot_cells(schedule_view) or {})
                cells = list(empties.get(slot_value) or [])
            except Exception:
                cells = []
        try:
            table = schedule_view.GetTableData()
            body = table.GetSectionData(DB.SectionType.Body)
            body_rows = int(getattr(body, "NumberOfRows", 0) or 0)
            body_cols = int(getattr(body, "NumberOfColumns", 0) or 0)
        except Exception:
            body_rows = 0
            body_cols = 0
        is_circuit_cell = getattr(schedule_view, "IsCellInCircuitTable", None)
        unique_cells = []
        seen_cells = set()
        for pair in list(cells or []):
            if not pair or len(pair) < 2:
                continue
            row = int(pair[0])
            col = int(pair[1])
            if body_rows > 0 and (row < 0 or row >= body_rows):
                continue
            if body_cols > 0 and (col < 0 or col >= body_cols):
                continue
            try:
                if is_circuit_cell is not None and not bool(is_circuit_cell(int(row), int(col))):
                    continue
            except Exception:
                continue
            try:
                if int(schedule_view.GetSlotNumberByCell(int(row), int(col)) or 0) != int(slot_value):
                    continue
            except Exception:
                continue
            key = (row, col)
            if key in seen_cells:
                continue
            seen_cells.add(key)
            unique_cells.append(key)

        errors = []
        for method_name in method_names:
            method = getattr(schedule_view, method_name, None)
            if method is None:
                continue
            attempts = []
            for row, col in list(unique_cells or []):
                attempts.append((int(row), int(col)))
            for args in attempts:
                try:
                    result = method(*args)
                    if isinstance(result, bool) and not result:
                        continue
                    occupant = _get_slot_circuit(schedule_view, slot_value)
                    if isinstance(occupant, DBE.ElectricalSystem):
                        _set_circuit_poles(occupant, int(max(1, poles or 1)))
                    return True
                except Exception as ex:
                    errors.append("{0}{1} -> {2}".format(method_name, tuple(args), ex))
                    continue
        if errors:
            try:
                logger.warning(
                    "Restore {0} failed at slot {1}; attempts: {2}".format(
                        action.upper(),
                        slot_value,
                        " | ".join(list(errors or [])[:8]),
                    )
                )
            except Exception:
                pass
        return False

    def _collect_default_special_rows(panel, schedule_view):
        panel_id = _elid_value(getattr(panel, "Id", None))
        if panel_id <= 0 or schedule_view is None:
            return []
        settings = ps_repo.get_electrical_settings(doc)
        default_rating = ps_repo.get_default_circuit_rating(settings)
        entries = []
        circuits_in_panel = (
            DB.FilteredElementCollector(doc)
            .OfClass(DBE.ElectricalSystem)
            .WhereElementIsNotElementType()
            .WherePasses(option_filter)
            .ToElements()
        )
        for circuit in list(circuits_in_panel or []):
            base = getattr(circuit, "BaseEquipment", None)
            if base is None or _elid_value(getattr(base, "Id", None)) != panel_id:
                continue
            kind = _safe_text(ps_repo._kind_from_circuit(circuit), "").strip().lower()
            if kind not in ("spare", "space"):
                continue
            start_slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
            if start_slot <= 0:
                continue
            removable = False
            cells = list(ps_repo.get_cells_by_slot_number(schedule_view, start_slot) or [])
            for row, col in cells:
                try:
                    cid = schedule_view.GetCircuitIdByCell(int(row), int(col))
                    if cid is None or cid == DB.ElementId.InvalidElementId:
                        continue
                    if _elid_value(cid) != _elid_value(circuit.Id):
                        continue
                except Exception:
                    pass
                if kind == "spare":
                    removable = bool(
                        ps_repo.is_removable_spare(
                            schedule_view,
                            int(row),
                            int(col),
                            circuit,
                            electrical_settings=settings,
                            default_rating=default_rating,
                        )
                    )
                else:
                    removable = bool(ps_repo.is_removable_space(schedule_view, int(row), int(col), circuit))
                if removable:
                    break
            if not removable:
                continue
            entries.append({
                "circuit_id": _elid_value(circuit.Id),
                "kind": kind,
                "start_slot": int(start_slot),
                "poles": int(max(1, _get_circuit_poles(circuit))),
                "cells": [(int(r), int(c)) for r, c in list(ps_repo.get_cells_by_slot_number(schedule_view, start_slot) or [])],
                "slots": _covered_slots(schedule_view, circuit, start_slot),
            })
        entries.sort(key=lambda x: (int(x.get("start_slot", 0) or 0), -int(x.get("poles", 1) or 1)))
        return entries

    def _circuits_requiring_new_slots():
        target_id = _panel_id(target_panel)
        needing = []
        for circuit in list(circuits or []):
            base = getattr(circuit, "BaseEquipment", None)
            base_id = _panel_id(base)
            if base_id == target_id:
                continue
            needing.append(circuit)
        return needing

    def _fit_count(option, free_slots, pole_counts):
        if option is None:
            return 0
        slot_order = list(ps_repo.get_option_slot_order(option, include_excess=False) or [])
        if not slot_order:
            return 0
        available = set([int(x) for x in list(free_slots or []) if int(x) > 0])
        moved = 0
        for poles in list(pole_counts or []):
            pole_count = int(max(1, poles or 1))
            placed = False
            for start in slot_order:
                covered = ps_repo.get_slot_span_slots_for_option(
                    option,
                    start_slot=int(start),
                    pole_count=int(pole_count),
                    require_valid=True,
                )
                if not covered:
                    continue
                if all(int(slot) in available for slot in list(covered or [])):
                    for slot in covered:
                        try:
                            available.remove(int(slot))
                        except Exception:
                            pass
                    moved += 1
                    placed = True
                    break
            if not placed:
                break
        return int(moved)

    def _accept_partial_prompt(moved_count, total_count):
        return forms.alert(
            (
                "Only some circuits could be moved.\n\n"
                "Moved: {0} of {1}\n\nAccept partial result?".format(
                    int(moved_count),
                    int(total_count),
                )
            ),
            title="Move Selected Circuits",
            ok=False,
            yes=True,
            no=True,
        )

    def _collect_removable_slots(default_entries):
        removable_slots = set()
        for entry in list(default_entries or []):
            for slot in list(entry.get("slots") or []):
                sval = int(slot or 0)
                if sval > 0:
                    removable_slots.add(sval)
        return removable_slots

    def _capacity_plan(schedule_view, default_entries):
        moving = _circuits_requiring_new_slots()
        pole_counts = sorted([int(max(1, _get_circuit_poles(c))) for c in list(moving or [])], reverse=True)
        required = int(len(pole_counts))
        option = target_option
        if option is None and schedule_view is not None:
            option = _get_target_option(target_panel, schedule_view)
        empty_slots = set()
        if option is not None:
            empty_slots = set([int(x) for x in list(_collect_empty_slots(option) or []) if int(x) > 0])
        removable_slots = _collect_removable_slots(default_entries)
        fit_without = int(_fit_count(option, empty_slots, pole_counts)) if required > 0 else 0
        fit_with = int(_fit_count(option, empty_slots.union(removable_slots), pole_counts)) if required > 0 else 0
        is_switchboard = bool(
            option is not None
            and option.get("schedule_type") == ps_repo.PSTYPE_SWITCHBOARD
        )
        return {
            "option": option,
            "required": int(required),
            "fit_without": int(fit_without),
            "fit_with": int(fit_with),
            "has_removable_defaults": bool(len(list(removable_slots or [])) > 0),
            "replace_improves_fit": bool(int(fit_with) > int(fit_without)),
            "is_switchboard": bool(is_switchboard),
        }

    def _should_offer_default_replace(schedule_view, default_entries):
        if schedule_view is None:
            return False
        moving = _circuits_requiring_new_slots()
        if not moving:
            return False
        option = _get_target_option(target_panel, schedule_view)
        if option is None:
            return False
        rows = list(ps_repo.build_panel_rows(doc, option) or [])
        empty_slots = set()
        for row in rows:
            kind = _safe_text(row.get("kind", ""), "").strip().lower()
            if kind != "empty":
                continue
            for slot in list(ps_repo.get_row_covered_slots(row, option=option) or []):
                sval = int(slot or 0)
                if sval > 0:
                    empty_slots.add(sval)
        removable_slots = set()
        for entry in list(default_entries or []):
            for slot in list(entry.get("slots") or []):
                sval = int(slot or 0)
                if sval > 0:
                    removable_slots.add(sval)
        if not removable_slots:
            return False
        pole_counts = sorted([_get_circuit_poles(c) for c in moving], reverse=True)
        required = int(len(pole_counts))
        if required <= 0:
            return False
        fit_without = _fit_count(option, empty_slots, pole_counts)
        fit_with = _fit_count(option, empty_slots.union(removable_slots), pole_counts)
        if fit_without >= required:
            return False
        return bool(fit_with > fit_without)

    def _partition_requested_circuits():
        target_id = _panel_id(target_panel)
        to_move = []
        skipped = []
        for circuit in list(circuits or []):
            old_ref = "{} / {}".format(_safe_panel_name(circuit), _safe_circuit_number(circuit))
            base = getattr(circuit, "BaseEquipment", None)
            if _panel_id(base) == target_id:
                skipped.append([output.linkify(circuit.Id), old_ref, "Already on target panel"])
                continue
            to_move.append(circuit)
        return to_move, skipped

    def _collect_empty_slots(option):
        if option is None:
            return set()
        rows = list(ps_repo.build_panel_rows(doc, option) or [])
        empty_slots = set()
        for row in rows:
            kind = _safe_text(row.get("kind", ""), "").strip().lower()
            if kind != "empty":
                continue
            if not bool(row.get("is_editable", True)):
                continue
            for slot in list(ps_repo.get_row_covered_slots(row, option=option) or []):
                sval = int(slot or 0)
                if sval > 0:
                    empty_slots.add(sval)
        return empty_slots

    def _log_default_snapshot(entries):
        if not entries:
            return
        try:
            rows = []
            for entry in list(entries or []):
                cell_text = ", ".join(["({0},{1})".format(int(rc[0]), int(rc[1])) for rc in list(entry.get("cells") or [])])
                rows.append([
                    str(entry.get("kind", "")).upper(),
                    int(entry.get("start_slot", 0) or 0),
                    int(entry.get("poles", 1) or 1),
                    ",".join([str(int(x)) for x in list(entry.get("slots") or []) if int(x) > 0]),
                    cell_text,
                ])
            output.print_md("**Default SPARE/SPACE snapshot on target panel (pre-move):**")
            output.print_table(rows, ["Type", "Start Slot", "Poles", "Covered Slots", "Cells (row,col)"])
        except Exception:
            pass

    def _remove_default_special_rows(entries):
        deleted = set()
        for entry in list(entries or []):
            cid = int(entry.get("circuit_id", 0) or 0)
            if cid <= 0 or cid in deleted:
                continue
            element = doc.GetElement(_elid_from(cid))
            if element is None:
                continue
            doc.Delete(element.Id)
            deleted.add(cid)

    def _restore_default_special_rows(schedule_view, entries):
        restore_order = sorted(
            list(entries or []),
            key=lambda x: (-int(x.get("poles", 1) or 1), int(x.get("start_slot", 0) or 0)),
        )
        for entry in restore_order:
            kind = _safe_text(entry.get("kind", ""), "").strip().lower()
            start_slot = int(entry.get("start_slot", 0) or 0)
            poles = int(max(1, entry.get("poles", 1) or 1))
            if kind not in ("spare", "space") or start_slot <= 0:
                continue
            intended_slots = [int(x) for x in list(entry.get("slots") or []) if int(x) > 0]
            if not intended_slots:
                intended_slots = [int(start_slot)]
            if any(_slot_is_occupied(schedule_view, slot) for slot in intended_slots):
                continue
            for slot in intended_slots:
                _set_slot_locked(schedule_view, slot, False)
            try:
                if kind == "spare":
                    psm.add_spare(
                        panel_id=target_panel_id,
                        panel_slot=int(start_slot),
                        poles=int(poles),
                        rating=0,
                        frame=0,
                        unlock=True,
                        load_name=None,
                        schedule_notes=None,
                    )
                else:
                    psm.add_space(
                        panel_id=target_panel_id,
                        panel_slot=int(start_slot),
                        poles=int(poles),
                        unlock=True,
                        load_name=None,
                        schedule_notes=None,
                    )
            except Exception:
                continue

    def _backfill_new_empty_with_default_spaces(schedule_view, option, baseline_empty_slots):
        if option is None:
            return 0
        baseline = set([int(x) for x in list(baseline_empty_slots or []) if int(x) > 0])
        current_empty = _collect_empty_slots(option)
        newly_open = set([int(x) for x in list(current_empty or []) if int(x) > 0 and int(x) not in baseline])
        if not newly_open:
            return 0
        slot_order = list(ps_repo.get_option_slot_order(option, include_excess=False) or [])
        if not slot_order:
            slot_order = sorted(list(newly_open))
        added = 0
        for slot in slot_order:
            sval = int(slot or 0)
            if sval <= 0 or sval not in newly_open:
                continue
            if _slot_is_occupied(schedule_view, sval):
                continue
            try:
                psm.add_space_default(panel_id=target_panel_id, panel_slot=int(sval), unlock=True)
                added += 1
            except Exception:
                continue
        return int(added)

    schedule_view = _get_panel_schedule_view(target_panel)
    target_option = _get_target_option(target_panel, schedule_view)
    target_panel_id = int(_panel_id(target_panel))
    panel_option_lookup = {}
    if target_option is not None and target_panel_id > 0:
        panel_option_lookup[int(target_panel_id)] = target_option
    psm = PanelScheduleManager(doc, panel_option_lookup=panel_option_lookup, logger=logger)
    default_entries = _collect_default_special_rows(target_panel, schedule_view)
    requested_moves, skipped_rows = _partition_requested_circuits()
    circuits = list(requested_moves)
    if not circuits:
        return {
            "moved": [],
            "failed": [],
            "skipped": list(skipped_rows),
            "partial": False,
            "fallback_used": False,
        }
    tx_group = DB.TransactionGroup(doc, "Move Selected Circuits")
    tx_group.Start()
    try:
        first_error = ""

        initial_tx = Transaction(doc, "Move Circuits to New Panel")
        initial_tx.Start()
        try:
            data, failed = _run_select_panel_moves(allow_partial=False, sort_by_poles=False, phase="primary")
            initial_tx.Commit()
            tx_group.Assimilate()
            return {
                "moved": data,
                "failed": failed,
                "skipped": list(skipped_rows),
                "partial": False,
                "fallback_used": False,
            }
        except Exception as ex:
            first_error = _safe_text(ex, "Move failed.")
            try:
                logger.info(
                    "[MoveSelectedCircuits] Primary move failed; evaluating fallback. error=%s",
                    _safe_text(first_error, "Move failed."),
                )
            except Exception:
                pass
            try:
                initial_tx.RollBack()
            except Exception:
                pass

        def _run_partial_move_batch(phase_name):
            partial_tx = Transaction(doc, "Move Circuits to New Panel")
            partial_tx.Start()
            try:
                try:
                    logger.info(
                        "[MoveSelectedCircuits] Partial move batch start phase=%s regenerate=%s",
                        _safe_text(phase_name, "partial"),
                        False,
                    )
                except Exception:
                    pass
                data_rows, failed_rows = _run_select_panel_moves(
                    allow_partial=True,
                    sort_by_poles=True,
                    phase=phase_name,
                )
                partial_tx.Commit()
                try:
                    logger.info(
                        "[MoveSelectedCircuits] Partial move batch committed phase=%s moved=%s failed=%s.",
                        _safe_text(phase_name, "partial"),
                        int(len(data_rows or [])),
                        int(len(failed_rows or [])),
                    )
                except Exception:
                    pass
                return list(data_rows or []), list(failed_rows or [])
            except Exception:
                try:
                    partial_tx.RollBack()
                except Exception:
                    pass
                raise

        capacity = _capacity_plan(schedule_view, default_entries)
        required = int(capacity.get("required", 0) or 0)
        fit_without = int(capacity.get("fit_without", 0) or 0)
        fit_with = int(capacity.get("fit_with", 0) or 0)
        has_removable_defaults = bool(capacity.get("has_removable_defaults", False))
        replace_improves_fit = bool(capacity.get("replace_improves_fit", False))
        is_switchboard_target = bool(capacity.get("is_switchboard", False))
        cap_unit = "positions" if bool(is_switchboard_target) else "slots"
        try:
            logger.info(
                "[MoveSelectedCircuits] Capacity plan required=%s fit_without=%s fit_with_defaults=%s removable_defaults=%s switchboard=%s",
                int(required),
                int(fit_without),
                int(fit_with),
                bool(has_removable_defaults),
                bool(is_switchboard_target),
            )
        except Exception:
            pass

        if required <= 0:
            raise Exception(first_error)

        # Capacity is sufficient without fallback; primary failure is unrelated.
        if fit_without >= required:
            raise Exception(first_error)

        # Fast-fail impossible scenarios (e.g., 3P request with only two usable slots total).
        if fit_with <= 0:
            raise Exception(
                "{0}\n\nTarget panel cannot fit any selected circuits with available/removable {1}.".format(
                    _safe_text(first_error, "Move failed."),
                    _safe_text(cap_unit, "slots"),
                )
            )

        # Allow partial move even when no defaults are available (or replacement offers no benefit).
        if fit_without > 0 and not bool(replace_improves_fit):
            proceed_partial = forms.alert(
                (
                    "Only {0} of {1} selected circuit(s) can fit on the target panel with current {2}.\n\n"
                    "Continue with partial move?"
                ).format(int(fit_without), int(required), _safe_text(cap_unit, "slots")),
                title="Move Selected Circuits",
                ok=False,
                yes=True,
                no=True,
            )
            if not proceed_partial:
                raise Exception(first_error)
            data, failed = _run_partial_move_batch("partial-no-replace")
            if failed:
                moved_count = int(len(data or []))
                total_count = int(moved_count + len(failed or []))
                if moved_count <= 0:
                    raise Exception("No circuits could be moved with available target {0}.".format(_safe_text(cap_unit, "slots")))
                if not bool(_accept_partial_prompt(moved_count, total_count)):
                    raise Exception("User chose rollback after partial move.")
            tx_group.Assimilate()
            return {
                "moved": data,
                "failed": failed,
                "skipped": list(skipped_rows),
                "partial": bool(failed),
                "fallback_used": False,
            }

        offer_replace = bool(has_removable_defaults and replace_improves_fit)
        if not offer_replace:
            raise Exception(first_error)

        proceed = forms.alert(
            (
                "Target capacity can improve from {0} to {1} of {2} circuits by removing default SPARE/SPACE rows.\n\n"
                "Attempt default SPARE/SPACE replacement workflow?"
            ).format(int(fit_without), int(fit_with), int(required)),
            title="Move Selected Circuits",
            ok=False,
            yes=True,
            no=True,
        )
        if not proceed:
            if fit_without > 0:
                proceed_partial = forms.alert(
                    (
                        "Without removing defaults, {0} of {1} circuits can still be moved.\n\n"
                        "Continue with partial move?"
                    ).format(int(fit_without), int(required)),
                    title="Move Selected Circuits",
                    ok=False,
                    yes=True,
                    no=True,
                )
                if proceed_partial:
                    data, failed = _run_partial_move_batch("partial-no-replace")
                    if failed:
                        moved_count = int(len(data or []))
                        total_count = int(moved_count + len(failed or []))
                        if moved_count <= 0:
                            raise Exception("No circuits could be moved with available target {0}.".format(_safe_text(cap_unit, "slots")))
                        if not bool(_accept_partial_prompt(moved_count, total_count)):
                            raise Exception("User chose rollback after partial move.")
                    tx_group.Assimilate()
                    return {
                        "moved": data,
                        "failed": failed,
                        "skipped": list(skipped_rows),
                        "partial": bool(failed),
                        "fallback_used": False,
                    }
            raise Exception(first_error)

        _log_default_snapshot(default_entries)
        baseline_empty_slots = _collect_empty_slots(target_option)

        remove_tx = Transaction(doc, "Remove Default SPARE/SPACE Rows")
        remove_tx.Start()
        try:
            try:
                logger.info(
                    "[MoveSelectedCircuits] Fallback remove defaults start entries=%s regenerate=%s",
                    int(len(list(default_entries or []))),
                    False,
                )
            except Exception:
                pass
            _remove_default_special_rows(default_entries)
            remove_tx.Commit()
            try:
                logger.info("[MoveSelectedCircuits] Fallback remove defaults committed.")
            except Exception:
                pass
        except Exception:
            try:
                remove_tx.RollBack()
            except Exception:
                pass
            raise

        move_tx = Transaction(doc, "Move Circuits to New Panel")
        move_tx.Start()
        try:
            try:
                logger.info("[MoveSelectedCircuits] Fallback move batch start regenerate=%s", False)
            except Exception:
                pass
            data, failed = _run_select_panel_moves(allow_partial=True, sort_by_poles=True, phase="fallback")
            move_tx.Commit()
            try:
                logger.info(
                    "[MoveSelectedCircuits] Fallback move batch committed moved=%s failed=%s.",
                    int(len(data or [])),
                    int(len(failed or [])),
                )
            except Exception:
                pass
        except Exception:
            try:
                move_tx.RollBack()
            except Exception:
                pass
            raise

        restore_tx = Transaction(doc, "Restore Default SPARE/SPACE Rows")
        restore_tx.Start()
        try:
            try:
                logger.info("[MoveSelectedCircuits] Fallback restore defaults start regenerate=%s", False)
            except Exception:
                pass
            _restore_default_special_rows(schedule_view, default_entries)
            _backfill_new_empty_with_default_spaces(schedule_view, target_option, baseline_empty_slots)
            restore_tx.Commit()
            try:
                logger.info("[MoveSelectedCircuits] Fallback restore defaults committed.")
            except Exception:
                pass
        except Exception:
            try:
                restore_tx.RollBack()
            except Exception:
                pass
            raise

        if failed:
            moved_count = int(len(data or []))
            total_count = int(moved_count + len(failed or []))
            try:
                logger.warning(
                    "[MoveSelectedCircuits] Fallback move partial result moved=%s failed=%s.",
                    int(moved_count),
                    int(len(failed or [])),
                )
            except Exception:
                pass
            if moved_count <= 0:
                raise Exception("No circuits could be moved after removing default SPARE/SPACE rows.")
            if not bool(_accept_partial_prompt(moved_count, total_count)):
                raise Exception("User chose rollback after partial move.")

        tx_group.Assimilate()
        return {
            "moved": data,
            "failed": failed,
            "skipped": list(skipped_rows),
            "partial": bool(failed),
            "fallback_used": True,
        }
    except Exception:
        try:
            tx_group.RollBack()
        except Exception:
            pass
        raise
