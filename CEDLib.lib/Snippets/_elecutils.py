# -*- coding: utf-8 -*-
import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.DB import FilteredElementCollector, Electrical, Transaction, BuiltInCategory, BuiltInParameter
from pyrevit import script, forms, DB

from CEDElectrical.Infrastructure.Revit.repositories import distribution_equipment_repository as de_repo
from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Model.panel_schedule_manager import PanelScheduleManager
from Snippets import revit_helpers

logger = script.get_logger()
def _elid_value(item):
    return int(revit_helpers.get_elementid_value(item))


def _elid_from(value):
    return revit_helpers.elementid_from_value(value)


#design option filter
option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
def get_all_panels(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()

    return collector


def get_all_panel_types(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_circuits(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_elec_fixtures(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_data_devices(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DataDevices).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_light_devices(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_LightingDevices).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_light_fixtures(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_LightingFixtures).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


def get_all_mech_control_devices(doc, el_id=False):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_MechanicalControlDevices).WhereElementIsNotElementType().WherePasses(option_filter)
    if el_id:
        collector = collector.ToElementIds()
    else:
        collector = collector.ToElements()
    return collector


# Helper function to get panel's distribution system and voltage capacity
def get_panel_dist_system(panel, doc, debug=False):
    """Returns a dictionary with the panel's distribution system name, voltage, and phase."""
    panel_data = {
        'dist_system_name': None,
        'phase': None,
        'lg_voltage': None,
        'll_voltage': None
    }

    # Try to get the secondary distribution system (for transformers)
    secondary_dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS)
    dist_system_id = None  # Initialize dist_system_id

    if secondary_dist_system_param and secondary_dist_system_param.HasValue:
        dist_system_id = secondary_dist_system_param.AsElementId()
        if debug:
            print("Secondary distribution system found for panel: {}".format(panel.Name))
    else:
        # Fallback to primary distribution system (for panels or switchboards)
        dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
        if dist_system_param and dist_system_param.HasValue:
            dist_system_id = dist_system_param.AsElementId()
            if debug:
                print("Primary distribution system found for panel: {}".format(panel.Name))
        else:
            if debug:
                print("Warning: No distribution system found for panel: {}".format(panel.Name))
            return panel_data  # Return early if no distribution system is found

    # Retrieve the DistributionSysType element using the ID
    dist_system_type = doc.GetElement(dist_system_id)

    if dist_system_type is None:
        if debug:
            print("Warning: Distribution system element not found for panel: {}".format(panel.Name))
        return panel_data  # Return early if the distribution system element is not found

    # Retrieve the Name using the SYMBOL_NAME_PARAM built-in parameter
    name_param = dist_system_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if name_param and name_param.HasValue:
        panel_data['dist_system_name'] = name_param.AsString()
        if debug:
            print("Distribution system name for panel {}: {}".format(panel.Name, panel_data['dist_system_name']))
    else:
        if debug:
            print("Warning: No name found for the distribution system of panel: {}".format(panel.Name))
        panel_data['dist_system_name'] = "Unnamed Distribution System"

    # Get phase (check if ElectricalPhase exists)
    if hasattr(dist_system_type, "ElectricalPhase"):
        panel_data['phase'] = dist_system_type.ElectricalPhase
    else:
        if debug:
            print("Warning: No phase information found for distribution system: {}".format(panel.Name))

    # Retrieve Line-to-Ground and Line-to-Line voltages
    lg_voltage = getattr(dist_system_type, "VoltageLineToGround", None)
    ll_voltage = getattr(dist_system_type, "VoltageLineToLine", None)

    # Fetch voltage values
    if lg_voltage:
        lg_voltage_param = lg_voltage.get_Parameter(BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
        panel_data['lg_voltage'] = lg_voltage_param.AsDouble() if lg_voltage_param else None
    else:
        if debug:
            print("Warning: No L-G voltage found for panel: {}".format(panel.Name))

    if ll_voltage:
        ll_voltage_param = ll_voltage.get_Parameter(BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
        panel_data['ll_voltage'] = ll_voltage_param.AsDouble() if ll_voltage_param else None
    else:
        if debug:
            print("Warning: No L-L voltage found for panel: {}".format(panel.Name))

    return panel_data


def get_compatible_panels(selected_circuit, all_panels, doc):
    """Return panels that can accept the circuit voltage/pole configuration."""
    circuit_voltage, circuit_poles = ps_repo.get_circuit_voltage_poles(selected_circuit)
    if circuit_poles is None:
        try:
            poles_param = selected_circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
            if poles_param and poles_param.HasValue:
                circuit_poles = int(poles_param.AsInteger())
        except Exception:
            circuit_poles = None
    if circuit_voltage is None:
        try:
            voltage_param = selected_circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
            if voltage_param and voltage_param.HasValue:
                circuit_voltage = float(voltage_param.AsDouble())
        except Exception:
            circuit_voltage = None
    if circuit_poles is None:
        return []

    target_poles = int(max(1, circuit_poles))
    tolerance = 1.0
    compatible_panels = []

    for panel in list(all_panels or []):
        model = de_repo.build_distribution_equipment(doc, panel, schedule_view=None)
        options = list(getattr(model, "branch_circuit_options", []) or [])
        matched = False
        for option in options:
            try:
                opt_poles = int(option.get("poles", 0) or 0)
            except Exception:
                opt_poles = 0
            opt_voltage = option.get("voltage")
            if opt_poles != target_poles:
                continue
            if circuit_voltage is None or opt_voltage is None:
                matched = True
                break
            try:
                if abs(float(opt_voltage) - float(circuit_voltage)) <= float(tolerance):
                    matched = True
                    break
            except Exception:
                continue
        if matched:
            compatible_panels.append(panel)

    return compatible_panels


def get_circuit_data(circuit):
    """Returns a dictionary containing the number of poles and voltage for the circuit."""
    circuit_data = {'poles': None, 'voltage': None}

    poles_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
    if poles_param and poles_param.HasValue:
        circuit_data['poles'] = poles_param.AsInteger()

    voltage_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
    if voltage_param and voltage_param.HasValue:
        circuit_data['voltage'] = voltage_param.AsDouble()

    return circuit_data


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
            snapshots.append({
                "circuit": ckt,
                "old_panel": _safe_panel_name(ckt),
                "old_circuit_number": _safe_circuit_number(ckt),
            })
        data_rows = []
        failed_rows = []
        for snap in snapshots:
            try:
                result = snap["circuit"].SelectPanel(target_panel)
                if isinstance(result, bool) and not result:
                    raise Exception("SelectPanel returned False.")
                base_after = getattr(snap["circuit"], "BaseEquipment", None)
                if int(_panel_id(base_after)) != int(target_id):
                    raise Exception("SelectPanel did not place circuit on target panel.")
            except Exception as ex:
                if not bool(allow_partial):
                    raise
                prev_circuit = "{} / {}".format(snap["old_panel"], snap["old_circuit_number"])
                failed_rows.append([
                    output.linkify(snap["circuit"].Id),
                    prev_circuit,
                    _safe_text(ex, "Move failed."),
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

        offer_replace = _should_offer_default_replace(schedule_view, default_entries)
        if not offer_replace:
            # Keep fallback deterministic and user-driven: if the initial move failed and we can
            # actually manipulate default SPARE/SPACE rows on this target, still offer retry.
            offer_replace = bool(
                schedule_view is not None
                and target_option is not None
                and len(list(default_entries or [])) > 0
                and len(_circuits_requiring_new_slots()) > 0
            )
        if not offer_replace:
            raise Exception(first_error)

        proceed = forms.alert(
            (
                "Insufficient slots. Attempt to replace default SPARE/SPACE rows to accommodate?\n\n"
                "This removes default SPARE/SPACE rows, attempts to move selected circuits, then restores remaining rows."
            ),
            title="Move Selected Circuits",
            ok=False,
            yes=True,
            no=True,
        )
        if not proceed:
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
            keep_partial = forms.alert(
                (
                    "Only some circuits could be moved.\n\n"
                    "Moved: {0} of {1}\n\nAccept partial result?".format(moved_count, total_count)
                ),
                title="Move Selected Circuits",
                ok=False,
                yes=True,
                no=True,
            )
            if not keep_partial:
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


def find_open_slots(target_panel):
    """Find available slots in the target panel, prioritizing odd-numbered slots."""
    available_slots = list(range(1, 43))
    odd_slots = [slot for slot in available_slots if slot % 2 == 1]
    even_slots = [slot for slot in available_slots if slot % 2 == 0]
    return odd_slots + even_slots


def get_circuits_from_panel(panel, doc, sort_method=0, include_spares=True):
    """Get circuits from a selected panel with sorting and inclusion of spare/space circuits."""
    circuits = []
    panel_circuits = FilteredElementCollector(doc).OfClass(DBE.ElectricalSystem).WherePasses(option_filter).ToElements()

    for circuit in panel_circuits:
        if circuit.BaseEquipment and circuit.BaseEquipment.Id == panel.Id:
            if not include_spares and circuit.CircuitType in [Electrical.CircuitType.Spare, Electrical.CircuitType.Space]:
                continue

            # Get circuit parameters
            circuit_number = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
            load_name = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).AsString()
            start_slot_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT)
            wire_size_param = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM)

            # Retrieve wire size as string if available
            wire_size = wire_size_param.AsString() if wire_size_param and wire_size_param.HasValue else "N/A"

            # Retrieve the start slot value
            start_slot = start_slot_param.AsInteger() if start_slot_param and start_slot_param.HasValue else 0

            # Retrieve the panel name
            panel_name = circuit.BaseEquipment.Name if circuit.BaseEquipment else "N/A"

            # Store data in a list of dictionaries
            circuits.append({
                'element_id': _elid_value(circuit.Id),
                'circuit_number': circuit_number,
                'load_name': load_name,
                'start_slot': start_slot,
                'wire_size': wire_size,
                'panel': panel_name,
                'circuit': circuit
            })

    # Sort circuits based on the selected method
    if sort_method == 1:
        circuits_sorted = sorted(circuits, key=lambda item: item['start_slot'])
    else:
        circuits_sorted = sorted(circuits, key=lambda item: (item['start_slot'] % 2 == 0, item['start_slot']))

    return circuits_sorted


def pick_circuits_from_list(doc, select_multiple=False, include_spares_and_spaces=False):
    ckts = DB.FilteredElementCollector(doc) \
        .OfClass(DBE.ElectricalSystem) \
        .WhereElementIsNotElementType().WherePasses(option_filter)

    grouped_options = {" All": []}
    ckt_lookup = {}
    panel_groups = {}  # key: panel name, value: list of (sort_key, label)
    all_labels = []  # list of (sort_key, label)

    for ckt in ckts:
        # Skip spares/spaces if not included
        if not include_spares_and_spaces and ckt.CircuitType in [DBE.CircuitType.Spare, DBE.CircuitType.Space]:
            continue

        # Safely get rating and poles if circuit is a PowerCircuit
        if ckt.SystemType == DBE.ElectricalSystemType.PowerCircuit:
            try:
                rating = int(round(ckt.Rating,0))
            except:
                rating = "N/A"

            try:
                pole = ckt.PolesNumber
            except:
                pole = "?"
        else:
            rating = "N/A"
            pole = "?"

        ckt_id = _elid_value(ckt.Id)
        base_equipment = ckt.BaseEquipment
        panel_name = getattr(base_equipment, 'Name', None) if base_equipment else None
        panel_name = panel_name or " No Panel"
        load_name = ckt.LoadName or ""
        circuit_number = ckt.CircuitNumber
        start_slot = ckt.StartSlot if hasattr(ckt, 'StartSlot') else 0
        sort_key = (panel_name, start_slot, load_name.strip())

        if ckt.CircuitType == DBE.CircuitType.Space:
            # Space: no rating/poles, just panel and label
            label = "[{}]  {}/{} - {}({}P)".format(ckt_id, panel_name, circuit_number, load_name.strip(),pole)

        elif ckt.CircuitType == DBE.CircuitType.Spare:
            # Spare: show circuit number and panel, label as [SPARE]
            label = "[{}]  {}/{} - {}  ({} A/{}P)".format(ckt_id, panel_name, circuit_number, load_name.strip(), rating, pole)

        else:
            # Normal circuit
            label = "[{}]  {}/{} - {}  ({} A/{}P)".format(ckt_id, panel_name, circuit_number, load_name.strip(), rating,
                                                       pole)

        all_labels.append((sort_key, label))

        if panel_name not in panel_groups:
            panel_groups[panel_name] = []
        panel_groups[panel_name].append((sort_key, label))

        ckt_lookup[label] = ckt

    # Build grouped options sorted by panel/circuit number
    grouped_options[" All"] = [label for _, label in sorted(all_labels)]

    for panel_name, label_list in panel_groups.items():
        grouped_options[panel_name] = [label for _, label in sorted(label_list)]

    selected_option = forms.SelectFromList.show(
        grouped_options,
        title="Select a Circuit",
        group_selector_title="Panel:",
        multiselect=select_multiple
    )

    if not selected_option:
        logger.info("No circuit selected. Exiting script.")
        script.exit()

    if not isinstance(selected_option, list):
        selected_option = [selected_option]

    selected_ckts = [ckt_lookup[label] for label in selected_option]
    logger.info("Selected {} Circuit(s).".format(len(selected_ckts)))
    return selected_ckts



def pick_panel_from_list(doc, select_multiple=False):
    panels = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().WherePasses(option_filter)

    panel_lookup = {}
    grouped_options = {" All": []}

    for panel in panels:

        panel_name = DB.Element.Name.__get__(panel)
        panel_data = get_panel_dist_system(panel, doc)
        dist_system = panel_data.get('dist_system_name', 'Unspecified')
        grouped_options[' All'].append(panel_name)
        if dist_system not in grouped_options:
            grouped_options[dist_system] = []

        grouped_options[dist_system].append(panel_name)
        panel_lookup[panel_name] = panel

    # Sort each group
    for group in grouped_options:
        grouped_options[group].sort()

    selected_names = forms.SelectFromList.show(
        grouped_options,
        title="Select Panel(s)",
        group_selector_title="Distribution System:",
        multiselect=select_multiple
    )

    if not selected_names:
        logger.info("No panel selected. Exiting script.")
        script.exit()

    selected_panels = [panel_lookup[name] for name in selected_names] if select_multiple else panel_lookup[selected_names]
    return selected_panels


def get_circuits_from_selection(selection):
    circuits = []

    if not isinstance(selection, (list, tuple, set)):
        selection = [selection]

    for item in selection:
        if isinstance(item, DBE.ElectricalSystem):
            logger.debug("item {} is electrical circuit".format(item.Id.Value))
            circuits.append(item)
            continue

        try:
            mep = item.MEPModel
        except Exception as e:
            logger.debug("{}".format(e))
            continue

        if item.Category == DB.BuiltInCategory.OST_ElectricalEquipment:
            all_systems = mep.GetElectricalSystems() or []
            assigned_systems = mep.GetAssignedElectricalSystems() or []
            assigned_ids = set([sys.Id for sys in assigned_systems])

            supply_systems = [sys for sys in all_systems if sys.Id not in assigned_ids]

            circuits.extend(supply_systems)
        else:
            all_systems = mep.GetElectricalSystems() or []
            circuits.extend(all_systems)

    return circuits


