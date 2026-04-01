# -*- coding: utf-8 -*-
"""Service-layer execution for Add/Remove Spares and Spaces."""

from pyrevit import DB, revit

from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Model.panel_schedule_enums import PanelSpecialKind as SpecialKind
from CEDElectrical.Model.panel_schedule_enums import PanelUiActionType as UiActionType
from CEDElectrical.Model.panel_schedule_enums import PanelUiMode as UiMode
from CEDElectrical.Model.panel_schedule_manager import PanelScheduleManager
from Snippets import revit_helpers

_elid_from_value = revit_helpers.elementid_from_value


def _collect_all_circuits(doc):
    return list(
        DB.FilteredElementCollector(doc)
        .OfClass(ps_repo.DBE.ElectricalSystem)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def collect_panel_circuit_index(doc, circuits=None):
    by_panel = {}
    items = list(circuits or _collect_all_circuits(doc) or [])
    for circuit in list(items or []):
        base = getattr(circuit, "BaseEquipment", None)
        if base is None:
            continue
        panel_id = int(ps_repo._idval(getattr(base, "Id", None)))
        if panel_id <= 0:
            continue
        by_panel.setdefault(int(panel_id), []).append(circuit)
    return by_panel


def _safe_int(value, default=0):
    try:
        return int(value)
    except Exception:
        return int(default or 0)


def _circuit_poles(circuit):
    for attr in ("PolesNumber", "NumberOfPoles"):
        value = getattr(circuit, attr, None)
        if value is None:
            continue
        poles = int(value)
        if poles > 0:
            return poles
    param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
    if param and param.HasValue:
        poles = int(param.AsInteger() or 0)
        if poles > 0:
            return poles
    return 1


def _panel_capacity(option):
    if option is None:
        return 0
    usable = int(ps_repo.get_option_usable_slot_count(option) or 0)
    if usable > 0:
        return int(usable)
    return int(max(0, _safe_int(option.get("max_slot", 0), 0)))


def collect_panel_assignment_usage(doc):
    usage = {}
    if doc is None:
        return usage

    circuits = (
        DB.FilteredElementCollector(doc)
        .OfClass(ps_repo.DBE.ElectricalSystem)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for circuit in list(circuits or []):
        base = getattr(circuit, "BaseEquipment", None)
        if base is None:
            continue
        panel_id = int(ps_repo._idval(getattr(base, "Id", None)))
        if panel_id <= 0:
            continue
        start_slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
        if start_slot <= 0:
            continue
        entry = usage.get(panel_id)
        if entry is None:
            entry = {"circuits": 0, "poles": 0}
            usage[panel_id] = entry
        entry["circuits"] = int(entry.get("circuits", 0) or 0) + 1
        entry["poles"] = int(entry.get("poles", 0) or 0) + int(max(1, _circuit_poles(circuit)))
    return usage


def _get_panel_circuits(option, panel_circuit_index):
    if option is None:
        return []
    panel_id = int(option.get("panel_id", 0) or 0)
    if panel_id <= 0:
        return []
    if isinstance(panel_circuit_index, dict):
        return list(panel_circuit_index.get(panel_id) or [])
    return []


def ordered_add_slots(option, panel_circuit_index=None):
    if option is None:
        return []
    slot_order = list(ps_repo.get_option_slot_order(option, include_excess=False) or [])
    if not slot_order:
        return []
    slot_set = set([int(x) for x in list(slot_order or []) if int(x) > 0])
    occupied = set()
    for circuit in list(_get_panel_circuits(option, panel_circuit_index) or []):
        start_slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
        if start_slot <= 0:
            continue
        poles = int(max(1, _circuit_poles(circuit)))
        covered = list(ps_repo.get_slot_span_slots_for_option(option, int(start_slot), int(poles), require_valid=True) or [])
        if not covered:
            continue
        for slot in list(covered or []):
            sval = int(slot or 0)
            if sval > 0 and sval in slot_set:
                occupied.add(sval)
    slots = [int(slot) for slot in list(slot_order or []) if int(slot) > 0 and int(slot) not in occupied]
    if not slots:
        return []
    schedule_type = option.get("schedule_type")
    if schedule_type == ps_repo.PSTYPE_SWITCHBOARD:
        return list(slots)
    odds = [x for x in slots if int(x % 2) == 1]
    evens = [x for x in slots if int(x % 2) == 0]
    return odds + evens


def count_open_slots_fast(option, usage_by_panel):
    panel_id = int(option.get("panel_id", 0) or 0)
    if panel_id <= 0:
        return 0
    capacity = int(max(0, _panel_capacity(option)))
    consumed = usage_by_panel.get(panel_id, {"circuits": 0, "poles": 0})
    if option.get("schedule_type") == ps_repo.PSTYPE_SWITCHBOARD:
        used = int(max(0, _safe_int(consumed.get("circuits", 0), 0)))
    else:
        used = int(max(0, _safe_int(consumed.get("poles", 0), 0)))
    open_slots = int(capacity - used)
    return max(0, open_slots)


def _removable_targets_for_option(doc, option, mode, panel_circuit_index):
    mode_key = UiMode.normalize_for_remove(mode, default=UiMode.BOTH)
    remove_spares = mode_key in (UiMode.SPARE, UiMode.BOTH)
    remove_spaces = mode_key in (UiMode.SPACE, UiMode.BOTH)

    schedule_view = option.get("schedule_view") if option else None
    if schedule_view is None:
        return []
    settings = ps_repo.get_electrical_settings(doc)
    default_rating = ps_repo.get_default_circuit_rating(settings)

    targets = []
    seen_ids = set()
    for circuit in list(_get_panel_circuits(option, panel_circuit_index) or []):
        if not isinstance(circuit, ps_repo.DBE.ElectricalSystem):
            continue
        circuit_id = int(ps_repo._idval(getattr(circuit, "Id", None)))
        if circuit_id <= 0 or circuit_id in seen_ids:
            continue
        kind = str(ps_repo._kind_from_circuit(circuit) or "").strip().lower()
        if kind not in (SpecialKind.SPARE, SpecialKind.SPACE):
            continue
        slot_value = int(ps_repo.get_circuit_start_slot(circuit) or 0)
        if slot_value <= 0 or circuit_id <= 0:
            continue
        cells = list(ps_repo.get_cells_by_slot_number(schedule_view, int(slot_value)) or [])
        if not cells:
            continue
        row, col = cells[0]

        if kind == SpecialKind.SPARE and remove_spares:
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
            if removable:
                seen_ids.add(circuit_id)
                targets.append((int(slot_value), SpecialKind.SPARE, int(circuit_id)))
        elif kind == SpecialKind.SPACE and remove_spaces:
            removable = bool(ps_repo.is_removable_space(schedule_view, int(row), int(col), circuit))
            if removable:
                seen_ids.add(circuit_id)
                targets.append((int(slot_value), SpecialKind.SPACE, int(circuit_id)))
    targets.sort(key=lambda x: int(x[0]), reverse=True)
    return list(targets)


def _execute_add_for_option(manager, doc, option, mode, panel_circuit_index=None):
    panel_id = int(option.get("panel_id", 0) or 0)
    slots = list(ordered_add_slots(option, panel_circuit_index=panel_circuit_index) or [])
    if panel_id <= 0 or not slots:
        return {"added_spares": 0, "added_spaces": 0, "added_slots": [], "switchboard_added_circuit_ids": []}

    mode_key = UiMode.normalize_for_add(mode, default=UiMode.SPACE)
    is_switchboard = bool(option.get("schedule_type") == ps_repo.PSTYPE_SWITCHBOARD)
    add_spares = 0
    add_spaces = 0
    added_slots = []
    added_circuit_ids = []
    switchboard_added_circuit_ids = []

    if mode_key == UiMode.SPARE:
        add_plan = [SpecialKind.SPARE] * len(slots)
    elif mode_key == UiMode.SPACE:
        add_plan = [SpecialKind.SPACE] * len(slots)
    else:
        spare_count = int((len(slots) + 1) / 2)
        add_plan = ([SpecialKind.SPARE] * spare_count) + ([SpecialKind.SPACE] * (len(slots) - spare_count))

    for slot, kind in zip(slots, add_plan):
        if kind == SpecialKind.SPARE:
            result = manager.add_spare_default(
                panel_id=panel_id,
                panel_slot=int(slot),
                unlock=False,
                apply_switchboard_default_poles=False,
            )
            add_spares += 1
        else:
            result = manager.add_space_default(
                panel_id=panel_id,
                panel_slot=int(slot),
                unlock=False,
                apply_switchboard_default_poles=False,
            )
            add_spaces += 1
        added_slots.append(int(slot))
        if is_switchboard:
            circuit_id = int((result or {}).get("circuit_id", 0) or 0)
            if circuit_id > 0:
                switchboard_added_circuit_ids.append(int(circuit_id))
        circuit_id = int((result or {}).get("circuit_id", 0) or 0)
        if circuit_id > 0:
            added_circuit_ids.append(int(circuit_id))

    return {
        "added_spares": int(add_spares),
        "added_spaces": int(add_spaces),
        "added_slots": [int(x) for x in list(added_slots or []) if int(x) > 0],
        "added_circuit_ids": [int(x) for x in list(added_circuit_ids or []) if int(x) > 0],
        "switchboard_added_circuit_ids": [int(x) for x in list(switchboard_added_circuit_ids or []) if int(x) > 0],
    }


def _unlock_added_slots(manager, doc, added_circuit_ids, fallback_unlock_requests):
    requests = []
    seen = set()
    for circuit_id in list(added_circuit_ids or []):
        cid = int(circuit_id or 0)
        if cid <= 0:
            continue
        circuit = doc.GetElement(_elid_from_value(cid))
        if not isinstance(circuit, ps_repo.DBE.ElectricalSystem):
            continue
        base = getattr(circuit, "BaseEquipment", None)
        panel_id = int(ps_repo._idval(getattr(base, "Id", None))) if base is not None else 0
        slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
        key = (panel_id, slot)
        if panel_id <= 0 or slot <= 0 or key in seen:
            continue
        seen.add(key)
        requests.append(key)
    for panel_id, slot in list(fallback_unlock_requests or []):
        pid = int(panel_id or 0)
        s = int(slot or 0)
        key = (pid, s)
        if pid <= 0 or s <= 0 or key in seen:
            continue
        seen.add(key)
        requests.append(key)
    if not requests:
        return {"attempted": 0, "failed": 0}

    def _unlock_slot_cells(schedule_view, slot_value):
        slot = int(slot_value or 0)
        if schedule_view is None or slot <= 0:
            return False
        cells = list(ps_repo.get_cells_by_slot_number(schedule_view, slot) or [])
        if not cells:
            return False

        per_cell_ok = False
        lock_cell = getattr(schedule_view, "SetLockSlot", None)
        set_cell_locked = getattr(schedule_view, "SetCellLocked", None)
        for row, col in list(cells or []):
            unlocked = False
            if lock_cell is not None:
                try:
                    result = lock_cell(int(row), int(col), 0)
                    if not (isinstance(result, bool) and not result):
                        unlocked = True
                except Exception:
                    pass
            if (not unlocked) and set_cell_locked is not None:
                try:
                    result = set_cell_locked(int(row), int(col), False)
                    if not (isinstance(result, bool) and not result):
                        unlocked = True
                except Exception:
                    pass
            if unlocked:
                per_cell_ok = True

        if per_cell_ok:
            return True

        set_slot_locked = getattr(schedule_view, "SetSlotLocked", None)
        if set_slot_locked is not None:
            try:
                result = set_slot_locked(int(slot), False)
                if not (isinstance(result, bool) and not result):
                    return True
            except Exception:
                pass
        return False

    failed = 0
    for panel_id, slot in list(requests or []):
        schedule_view = manager.get_panel_schedule_view(int(panel_id))
        if not bool(_unlock_slot_cells(schedule_view, int(slot))):
            failed += 1
    return {"attempted": int(len(requests)), "failed": int(failed)}


def _finalize_added_defaults(manager, doc, added_circuit_ids, fallback_unlock_requests, switchboard_added_circuit_ids):
    unlock_summary = _unlock_added_slots(manager, doc, added_circuit_ids, fallback_unlock_requests)
    pole_attempted = 0
    pole_failed = 0
    seen = set()
    for circuit_id in list(switchboard_added_circuit_ids or []):
        cid = int(circuit_id or 0)
        if cid <= 0 or cid in seen:
            continue
        seen.add(cid)
        pole_attempted += 1
        if not bool(manager.set_circuit_poles(cid, 3)):
            pole_failed += 1
    return {
        "unlock_attempted": int(unlock_summary.get("attempted", 0) or 0),
        "unlock_failed": int(unlock_summary.get("failed", 0) or 0),
        "pole_attempted": int(pole_attempted),
        "pole_failed": int(pole_failed),
    }


def _execute_remove_for_option(manager, doc, option, mode, panel_circuit_index):
    panel_id = int(option.get("panel_id", 0) or 0)
    if panel_id <= 0:
        return {"removed_spares": 0, "removed_spaces": 0}
    removed_spares = 0
    removed_spaces = 0
    targets = _removable_targets_for_option(doc, option, mode, panel_circuit_index)
    for _, kind, circuit_id in list(targets or []):
        circuit = doc.GetElement(_elid_from_value(int(circuit_id)))
        if circuit is None:
            continue
        doc.Delete(circuit.Id)
        if kind == SpecialKind.SPARE:
            removed_spares += 1
        else:
            removed_spaces += 1
    return {"removed_spares": int(removed_spares), "removed_spaces": int(removed_spaces)}


def format_panel_info(option, usage_by_panel):
    if not option:
        return "Unknown Panel"
    return "{0} ({1}) | {2} | Open Slots: {3}".format(
        str(option.get("panel_name", "") or "Unnamed Panel"),
        str(option.get("part_type_name", "") or option.get("board_type", "Unknown")),
        str(option.get("dist_system_name", "") or "Unknown Dist. System"),
        str(count_open_slots_fast(option, usage_by_panel)),
    )


def execute_quick_action(doc, panel_options, action_type, mode):
    options = [x for x in list(panel_options or []) if isinstance(x, dict) and int(x.get("panel_id", 0) or 0) > 0]
    if not options:
        raise Exception("Could not resolve panels for quick action.")

    action_kind = UiActionType.normalize(action_type, default=UiActionType.ADD)
    if action_kind == UiActionType.ADD:
        mode_key = UiMode.normalize_for_add(mode, default=UiMode.SPACE)
    else:
        mode_key = UiMode.normalize_for_remove(mode, default=UiMode.BOTH)

    manager = PanelScheduleManager(
        doc,
        panel_option_lookup={int(x.get("panel_id", 0) or 0): x for x in list(options or [])},
    )
    panel_circuit_index = collect_panel_circuit_index(doc)
    added_spares = 0
    added_spaces = 0
    removed_spares = 0
    removed_spaces = 0
    touched = 0
    unlock_requests = []
    added_circuit_ids = []
    switchboard_added_circuit_ids = []
    finalize_summary = {"unlock_attempted": 0, "unlock_failed": 0, "pole_attempted": 0, "pole_failed": 0}

    tx_group = DB.TransactionGroup(doc, "Quick Spare/Space")
    tx_group.Start()
    try:
        with revit.Transaction("Quick Spare/Space - Apply", doc):
            for option in list(options or []):
                panel_id = int(option.get("panel_id", 0) or 0)
                if panel_id <= 0:
                    continue
                touched += 1
                if action_kind == UiActionType.ADD:
                    result = _execute_add_for_option(
                        manager,
                        doc,
                        option,
                        mode_key,
                        panel_circuit_index=panel_circuit_index,
                    )
                    added_spares += int(result.get("added_spares", 0) or 0)
                    added_spaces += int(result.get("added_spaces", 0) or 0)
                    unlock_requests.extend([(int(panel_id), int(x)) for x in list(result.get("added_slots", []) or [])])
                    added_circuit_ids.extend(
                        [int(x) for x in list(result.get("added_circuit_ids", []) or []) if int(x) > 0]
                    )
                    switchboard_added_circuit_ids.extend(
                        [int(x) for x in list(result.get("switchboard_added_circuit_ids", []) or []) if int(x) > 0]
                    )
                else:
                    result = _execute_remove_for_option(
                        manager,
                        doc,
                        option,
                        mode_key,
                        panel_circuit_index=panel_circuit_index,
                    )
                    removed_spares += int(result.get("removed_spares", 0) or 0)
                    removed_spaces += int(result.get("removed_spaces", 0) or 0)
        if action_kind == UiActionType.ADD and (unlock_requests or added_circuit_ids or switchboard_added_circuit_ids):
            with revit.Transaction("Quick Spare/Space - Finalize Added", doc):
                finalize_summary = _finalize_added_defaults(
                    manager,
                    doc,
                    added_circuit_ids,
                    unlock_requests,
                    switchboard_added_circuit_ids,
                )
        tx_group.Assimilate()
    except Exception:
        tx_group.RollBack()
        raise

    return {
        "action_kind": action_kind,
        "touched": int(touched),
        "added_spares": int(added_spares),
        "added_spaces": int(added_spaces),
        "removed_spares": int(removed_spares),
        "removed_spaces": int(removed_spaces),
        "finalize_summary": dict(finalize_summary or {}),
    }


def execute_staged_actions(doc, staged_actions, option_lookup):
    manager = PanelScheduleManager(doc, panel_option_lookup=option_lookup)
    need_panel_scan = any(
        UiActionType.normalize(x.get("action_type", ""), default="") in (UiActionType.ADD, UiActionType.REMOVE)
        for x in list(staged_actions or [])
    )
    panel_circuit_index = collect_panel_circuit_index(doc) if need_panel_scan else {}
    unlock_requests = []
    added_circuit_ids = []
    switchboard_added_circuit_ids = []
    finalize_summary = {"unlock_attempted": 0, "unlock_failed": 0, "pole_attempted": 0, "pole_failed": 0}

    tx_group = DB.TransactionGroup(doc, "Add/Remove Spares and Spaces")
    tx_group.Start()
    try:
        with revit.Transaction("Add/Remove Spares and Spaces - Apply", doc):
            for action in list(staged_actions or []):
                panel_id = int(action.get("panel_id", 0) or 0)
                option = option_lookup.get(int(panel_id))
                if option is None:
                    continue
                kind = UiActionType.normalize(action.get("action_type", ""), default="")
                mode = str(action.get("mode", "") or "")
                if kind == UiActionType.ADD:
                    result = _execute_add_for_option(
                        manager,
                        doc,
                        option,
                        mode,
                        panel_circuit_index=panel_circuit_index,
                    )
                    unlock_requests.extend(
                        [(int(panel_id), int(x)) for x in list(result.get("added_slots", []) or [])]
                    )
                    added_circuit_ids.extend(
                        [int(x) for x in list(result.get("added_circuit_ids", []) or []) if int(x) > 0]
                    )
                    switchboard_added_circuit_ids.extend(
                        [int(x) for x in list(result.get("switchboard_added_circuit_ids", []) or []) if int(x) > 0]
                    )
                elif kind == UiActionType.REMOVE:
                    _execute_remove_for_option(
                        manager,
                        doc,
                        option,
                        mode,
                        panel_circuit_index=panel_circuit_index,
                    )
        if unlock_requests or added_circuit_ids or switchboard_added_circuit_ids:
            with revit.Transaction("Add/Remove Spares and Spaces - Finalize Added", doc):
                finalize_summary = _finalize_added_defaults(
                    manager,
                    doc,
                    added_circuit_ids,
                    unlock_requests,
                    switchboard_added_circuit_ids,
                )
        tx_group.Assimilate()
    except Exception:
        tx_group.RollBack()
        raise

    return {
        "finalize_summary": dict(finalize_summary or {}),
    }

