# -*- coding: utf-8 -*-
"""Panel schedule orchestration helpers for panel/switch/data equipment."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB, script

from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from Snippets import revit_helpers
from .panel_schedule_enums import PanelSpecialKind as SpecialKind
from .panel_schedule_enums import PanelStagedAction as StagedAction
from .panel_slot import PanelSlot


class PanelScheduleManager(object):
    """Model-layer orchestrator for panel schedule operations."""

    def __init__(self, doc, distribution_bus=None, panel_option_lookup=None, logger=None):
        self.doc = doc
        self.distribution_bus = distribution_bus
        self.logger = logger or script.get_logger()
        self._panel_option_lookup = dict(panel_option_lookup or {})
        self._slot_cells_cache = {}
        self._layout_context_cache = {}

    # -------------------------------------------------------------------------
    # Public metadata surface
    # -------------------------------------------------------------------------
    def set_panel_option_lookup(self, panel_option_lookup):
        """Set panel-option lookup map keyed by panel id."""
        self._panel_option_lookup = dict(panel_option_lookup or {})

    def has_panel_schedule(self, panel_id):
        """Return True when panel has an instance schedule view in project."""
        option = self._option_for_panel_id(panel_id)
        return bool(option and isinstance(option.get("schedule_view"), DBE.PanelScheduleView))

    def get_panel_configuration(self, panel_id):
        """Return DBE.PanelConfiguration for panel option when available."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return None
        return option.get("panel_configuration")

    def get_schedule_type(self, panel_id):
        """Return DBE.PanelScheduleType for panel option when available."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return None
        return option.get("schedule_type")

    def get_panel_schedule_view(self, panel_id):
        """Return PanelScheduleView for panel id, or None when missing."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return None
        return option.get("schedule_view")

    def get_sheet_instances(self, panel_id):
        """Return panel schedule sheet instances for the panel schedule view."""
        view = self.get_panel_schedule_view(panel_id)
        if not isinstance(view, DB.View):
            return []
        collector = (
            DB.FilteredElementCollector(self.doc)
            .OfClass(DBE.PanelScheduleSheetInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        view_id_val = int(self._idval(view.Id))
        matches = []
        for item in list(collector or []):
            try:
                sid = int(self._idval(getattr(item, "ScheduleId", None)))
                if sid == view_id_val:
                    matches.append(item)
            except Exception:
                continue
        return matches

    def is_switchboard(self, panel_id):
        """Return True when panel board type is switchboard."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return False
        text = str(option.get("board_type", "") or "").strip().lower()
        return text == "switchboard"

    def number_of_slots(self, panel_id):
        """Return slot count from mapped panel option."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return 0
        return int(option.get("max_slot", 0) or 0)

    def circuit_options(self, panel_id):
        """Return branch-circuit options from DistributionBus model."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return []
        model = option.get("equipment_model")
        if model is None:
            return []
        return list(getattr(model, "branch_circuit_options", []) or [])

    def get_valid_templates(self, panel_id, probe_assignability=True):
        """Return compatible templates for one distribution bus."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return []
        panel = option.get("panel")
        if panel is None:
            return []
        return list(
            ps_repo.get_compatible_panel_schedule_templates(
                self.doc,
                panel,
                probe_assignability=bool(probe_assignability),
            )
            or []
        )

    def get_row_col_from_slot(self, panel_id, slot):
        """Return schedule row/column pairs that represent one slot."""
        view = self.get_panel_schedule_view(panel_id)
        if view is None:
            return []
        return list(self._slot_cells(view, int(slot or 0)) or [])

    def get_slot_from_circuit(self, circuit):
        """Return circuit start slot using repository-compatible lookup."""
        return int(ps_repo.get_circuit_start_slot(circuit) or 0)

    def get_available_slots(self, panel_id):
        """Return empty slot numbers for a panel schedule."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return []
        rows = list(ps_repo.build_panel_rows(self.doc, option) or [])
        slots = []
        for row in rows:
            if str(row.get("kind", "") or "").strip().lower() != "empty":
                continue
            if not bool(row.get("is_editable", True)):
                continue
            slot = int(row.get("slot", 0) or 0)
            if slot > 0:
                slots.append(slot)
        return sorted(set(slots))

    def build_panel_slot(self, panel_id, slot):
        """Build a PanelSlot model from current schedule state."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return PanelSlot(slot=slot)
        schedule_view = option.get("schedule_view")
        slot_value = int(slot or 0)
        cells = list(self._slot_cells(schedule_view, slot_value) or [])
        occupant = self._get_circuit_at_slot(schedule_view, slot_value)
        kind = str(ps_repo._kind_from_circuit(occupant) or "").lower() if occupant is not None else ""
        poles = self._get_circuit_poles(occupant, fallback=1) if occupant is not None else 1
        group_no = int(ps_repo.get_slot_group_number(schedule_view, slot_value) or 0) if schedule_view is not None else 0
        return PanelSlot(
            slot=slot_value,
            cells=cells,
            is_locked=bool(self._slot_is_locked(schedule_view, slot_value)),
            is_spare=bool(kind == SpecialKind.SPARE),
            is_space=bool(kind == SpecialKind.SPACE),
            is_circuit=bool(kind == "circuit"),
            poles=int(max(1, poles or 1)),
            group_number=group_no,
        )

    # -------------------------------------------------------------------------
    # Primitive operations
    # -------------------------------------------------------------------------
    def unlock_slot(self, panel_id, slot):
        """Unlock one slot for mutation."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return False
        return bool(self._set_slot_locked(option.get("schedule_view"), int(slot), False))

    def set_slot_locked(self, panel_id, slot, is_locked):
        """Set one slot lock state."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            return False
        return bool(self._set_slot_locked(option.get("schedule_view"), int(slot), bool(is_locked)))

    def add_spare(self, panel_id, panel_slot, poles=1, rating=0, frame=0, unlock=True, load_name=None, schedule_notes=None):
        """Add SPARE with custom poles/rating/frame behavior (Batch Swap path)."""
        return self._add_special(
            panel_id=panel_id,
            panel_slot=panel_slot,
            kind=SpecialKind.SPARE,
            poles=poles,
            rating=rating,
            frame=frame,
            unlock=unlock,
            load_name=load_name,
            schedule_notes=schedule_notes,
            apply_custom_properties=True,
            force_switchboard_three_p=False,
            restrict_data_to_space=True,
        )

    def add_space(self, panel_id, panel_slot, poles=1, unlock=True, load_name=None, schedule_notes=None):
        """Add SPACE with custom pole behavior (Batch Swap path)."""
        return self._add_special(
            panel_id=panel_id,
            panel_slot=panel_slot,
            kind=SpecialKind.SPACE,
            poles=poles,
            rating=0,
            unlock=unlock,
            load_name=load_name,
            schedule_notes=schedule_notes,
            apply_custom_properties=True,
            force_switchboard_three_p=False,
            restrict_data_to_space=True,
        )

    def add_spare_default(self, panel_id, panel_slot, unlock=True, apply_switchboard_default_poles=True):
        """Add SPARE with add/remove-tool defaults and no post-add parameter edits."""
        option = self._option_for_panel_id(panel_id)
        poles = 3 if option and option.get("schedule_type") == ps_repo.PSTYPE_SWITCHBOARD else 1
        return self._add_special(
            panel_id=panel_id,
            panel_slot=panel_slot,
            kind=SpecialKind.SPARE,
            poles=int(poles),
            rating=0,
            frame=0,
            unlock=unlock,
            load_name=None,
            schedule_notes=None,
            apply_custom_properties=False,
            force_switchboard_three_p=bool(apply_switchboard_default_poles),
            restrict_data_to_space=False,
        )

    def add_space_default(self, panel_id, panel_slot, unlock=True, apply_switchboard_default_poles=True):
        """Add SPACE with add/remove-tool defaults and no post-add parameter edits."""
        option = self._option_for_panel_id(panel_id)
        poles = 3 if option and option.get("schedule_type") == ps_repo.PSTYPE_SWITCHBOARD else 1
        return self._add_special(
            panel_id=panel_id,
            panel_slot=panel_slot,
            kind=SpecialKind.SPACE,
            poles=int(poles),
            rating=0,
            frame=0,
            unlock=unlock,
            load_name=None,
            schedule_notes=None,
            apply_custom_properties=False,
            force_switchboard_three_p=bool(apply_switchboard_default_poles),
            restrict_data_to_space=False,
        )

    def set_circuit_poles(self, circuit_id, poles):
        """Set poles on one circuit by id when writable."""
        circuit = self._element_by_id_value(circuit_id)
        if not isinstance(circuit, DBE.ElectricalSystem):
            return False
        self._set_circuit_poles(circuit, int(max(1, poles or 1)))
        return True

    def remove_spare(self, panel_id, panel_slot):
        """Remove spare row/circuit from one slot."""
        return self._remove_special(panel_id=panel_id, panel_slot=panel_slot, kind_hint=SpecialKind.SPARE)

    def remove_space(self, panel_id, panel_slot):
        """Remove space row/circuit from one slot."""
        return self._remove_special(panel_id=panel_id, panel_slot=panel_slot, kind_hint=SpecialKind.SPACE)

    def move_circuit_to_panel(self, circuit_id, target_panel_id):
        """Move circuit to new panel using ElectricalSystem.SelectPanel."""
        circuit = self._element_by_id_value(circuit_id)
        if not isinstance(circuit, DBE.ElectricalSystem):
            raise Exception("Circuit {0} could not be resolved.".format(int(circuit_id or 0)))
        target_option = self._option_for_panel_id(target_panel_id)
        if not target_option:
            raise Exception("Target panel option not found: {0}".format(int(target_panel_id or 0)))
        panel = target_option.get("panel")
        if panel is None:
            raise Exception("Target panel element is unavailable.")
        self._select_panel_for_circuit(circuit, panel)
        self.doc.Regenerate()
        self._invalidate_schedule_cache(target_option.get("schedule_view"))
        return int(ps_repo.get_circuit_start_slot(circuit) or 0)

    def move_circuit_in_panel(self, panel_id, circuit_id, target_slot):
        """Move circuit to target slot within one panel schedule."""
        option = self._option_for_panel_id(panel_id)
        if not option:
            raise Exception("Panel option not found: {0}".format(int(panel_id or 0)))
        target_slot_value = int(target_slot or 0)
        if not bool(ps_repo.is_slot_valid_for_option(option, int(target_slot_value))):
            raise Exception(
                "Target slot {0} exceeds equipment-supported slot capacity for this panel.".format(
                    int(target_slot_value)
                )
            )
        schedule_view = option.get("schedule_view")
        if schedule_view is None:
            raise Exception("Panel schedule view is unavailable.")
        circuit = self._element_by_id_value(circuit_id)
        if not isinstance(circuit, DBE.ElectricalSystem):
            raise Exception("Circuit {0} could not be resolved.".format(int(circuit_id or 0)))
        source_slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
        if source_slot <= 0:
            raise Exception("Could not resolve current slot for circuit {0}.".format(int(circuit_id)))
        self._move_slot_to(schedule_view, source_slot, int(target_slot_value), circuit_id=int(circuit_id))
        self.doc.Regenerate()
        self._invalidate_schedule_cache(schedule_view)
        return int(ps_repo.get_circuit_start_slot(circuit) or 0)

    # -------------------------------------------------------------------------
    # Composite actions used by Batch Swap
    # -------------------------------------------------------------------------
    def apply_add_action(self, placement):
        """Apply staged add-spare/add-space action."""
        panel_id = int(placement.get("to_panel_id", 0) or 0)
        slot_value = int(placement.get("new_slot", 0) or 0)
        option = self._option_for_panel_id(panel_id)
        is_data_panel = bool(option and option.get("schedule_type") == ps_repo.PSTYPE_DATA)
        poles = int(max(1, placement.get("poles", 1) or 1))
        spare_rating = int(placement.get("spare_rating", 0) or 0)
        action = placement.get("action", "")
        if StagedAction.is_add_spare(action):
            kind = SpecialKind.SPARE
        elif StagedAction.is_add_space(action):
            kind = SpecialKind.SPACE
        else:
            raise Exception("Unsupported add action: {0}".format(str(action or "")))
        if bool(is_data_panel):
            if kind != SpecialKind.SPACE:
                raise Exception("Data panels only allow adding 1P SPACE.")
            poles = 1
            spare_rating = 0
        if kind == SpecialKind.SPARE:
            return self.add_spare(
                panel_id=panel_id,
                panel_slot=slot_value,
                poles=poles,
                rating=spare_rating,
                frame=(0 if is_data_panel else int(placement.get("spare_frame", 0) or 0)),
                unlock=True,
                load_name=placement.get("load_name"),
                schedule_notes=placement.get("schedule_notes"),
            )
        return self.add_space(
            panel_id=panel_id,
            panel_slot=slot_value,
            poles=poles,
            unlock=True,
            load_name=placement.get("load_name"),
            schedule_notes=placement.get("schedule_notes"),
        )

    def apply_remove_action(self, placement):
        """Apply staged remove-spare/remove-space action."""
        panel_id = int(placement.get("from_panel_id", 0) or 0)
        slot_value = int(placement.get("old_slot", 0) or 0)
        action = placement.get("action", "")
        restore_lock_state = not bool(placement.get("leave_unlocked", False))
        if StagedAction.is_remove_spare(action):
            return self._remove_special(
                panel_id=panel_id,
                panel_slot=slot_value,
                kind_hint=SpecialKind.SPARE,
                restore_lock_state=restore_lock_state,
            )
        if StagedAction.is_remove_space(action):
            return self._remove_special(
                panel_id=panel_id,
                panel_slot=slot_value,
                kind_hint=SpecialKind.SPACE,
                restore_lock_state=restore_lock_state,
            )
        return self._remove_special(
            panel_id=panel_id,
            panel_slot=slot_value,
            kind_hint=None,
            restore_lock_state=restore_lock_state,
        )

    def apply_move_action(self, placement):
        """Apply staged move action including replacement of target specials."""
        circuit_id = int(placement.get("circuit_id", 0) or 0)
        if circuit_id <= 0:
            raise Exception("Invalid circuit id in move action.")
        target_panel_id = int(placement.get("to_panel_id", 0) or 0)
        target_slot = int(placement.get("new_slot", 0) or 0)
        if target_slot <= 0:
            raise Exception("Move action has invalid target slot.")

        target_option = self._option_for_panel_id(target_panel_id)
        if target_option is None:
            raise Exception("Target panel option not found: {0}".format(int(target_panel_id)))
        if not bool(ps_repo.is_slot_valid_for_option(target_option, int(target_slot))):
            raise Exception(
                "Target slot {0} exceeds equipment-supported slot capacity for panel {1}.".format(
                    int(target_slot),
                    int(target_panel_id),
                )
            )
        target_schedule = target_option.get("schedule_view")
        if target_schedule is None:
            raise Exception("Target panel has no schedule view.")

        circuit = self._element_by_id_value(circuit_id)
        if not isinstance(circuit, DBE.ElectricalSystem):
            raise Exception("Circuit {0} could not be resolved.".format(int(circuit_id)))

        current_panel = getattr(circuit, "BaseEquipment", None)
        current_panel_id = int(self._idval(getattr(current_panel, "Id", None)))
        current_option = self._option_for_panel_id(current_panel_id)
        if current_option is None:
            raise Exception("Current panel option not found for circuit {0}.".format(int(circuit_id)))
        current_schedule = current_option.get("schedule_view")
        if current_schedule is None:
            raise Exception("Current panel schedule view is unavailable for circuit {0}.".format(int(circuit_id)))

        target_slots = [int(x) for x in list(placement.get("new_covered_slots") or []) if int(x) > 0]
        if not target_slots:
            poles_hint = int(max(1, placement.get("poles", 1) or 1))
            target_slots = ps_repo.get_slot_span_slots_for_option(
                target_option,
                start_slot=int(target_slot),
                pole_count=int(poles_hint),
                require_valid=True,
            )
        if not target_slots:
            raise Exception(
                "Target slot {0} / pole span exceeds equipment-supported slot capacity.".format(int(target_slot))
            )
        for slot in list(target_slots or []):
            if not bool(ps_repo.is_slot_valid_for_option(target_option, int(slot))):
                raise Exception(
                    "Target slot {0} exceeds equipment-supported slot capacity.".format(int(slot))
                )

        current_slots = self._covered_slots_for_circuit(
            current_option,
            circuit,
            fallback_slot=int(placement.get("old_slot", 0) or 0),
            fallback_poles=int(max(1, len(list(placement.get("old_covered_slots") or [])) or 1)),
        )

        source_lock_snapshot = self._unlock_slots_with_snapshot(current_schedule, current_slots)
        target_lock_snapshot = self._unlock_slots_with_snapshot(target_schedule, target_slots)
        source_was_locked = any(bool(x) for x in source_lock_snapshot.values())

        try:
            if bool(placement.get("is_regular_circuit", True)):
                deleted_snapshot = self._remove_specials_in_slots(target_option, target_slots, protected_circuit_id=circuit_id)
                for slot, was_locked in deleted_snapshot.items():
                    if int(slot) not in target_lock_snapshot:
                        target_lock_snapshot[int(slot)] = bool(was_locked)
                    else:
                        target_lock_snapshot[int(slot)] = bool(target_lock_snapshot[int(slot)] or bool(was_locked))

            same_panel = bool(current_panel_id == target_panel_id)
            if same_panel:
                current_start = self._resolve_circuit_start_slot_in_schedule(
                    current_schedule,
                    circuit,
                    fallback_slot=int(placement.get("old_slot", 0) or 0),
                )
                if current_start <= 0:
                    raise Exception("Could not resolve current slot for circuit {0}.".format(int(circuit_id)))
                self._move_slot_to(target_schedule, current_start, int(target_slot), circuit_id=int(circuit_id))
            else:
                placed_start = self.move_circuit_to_panel(circuit_id, target_panel_id)
                if placed_start <= 0:
                    raise Exception("Circuit {0} has invalid placement after SelectPanel.".format(int(circuit_id)))
                placed_start = self._resolve_circuit_start_slot_in_schedule(
                    target_schedule,
                    circuit,
                    fallback_slot=int(placed_start),
                )
                if int(placed_start) != int(target_slot):
                    self._move_slot_to(target_schedule, int(placed_start), int(target_slot), circuit_id=int(circuit_id))

            self.doc.Regenerate()
            self._invalidate_schedule_cache(target_schedule)
            if not same_panel:
                self._invalidate_schedule_cache(current_schedule)
            final_slots = self._covered_slots_for_circuit(
                target_option,
                circuit,
                fallback_slot=int(target_slot),
                fallback_poles=int(max(1, len(target_slots))),
            )
            if source_was_locked:
                for slot in list(final_slots or []):
                    self._set_slot_locked(target_schedule, int(slot), True)
            return {"final_slots": [int(x) for x in list(final_slots or [])]}
        finally:
            self._restore_slot_locks(current_schedule, source_lock_snapshot)
            self._restore_slot_locks(target_schedule, target_lock_snapshot)

    # -------------------------------------------------------------------------
    # Internal helpers
    # -------------------------------------------------------------------------
    def _idval(self, item):
        return revit_helpers.get_elementid_value(item)

    def _option_for_panel_id(self, panel_id):
        panel_key = int(panel_id or 0)
        if panel_key <= 0:
            return None
        option = self._panel_option_lookup.get(panel_key)
        if option is not None:
            return option
        options = list(ps_repo.collect_panel_equipment_options(self.doc, include_without_schedule=True) or [])
        for item in options:
            pid = int(item.get("panel_id", 0) or 0)
            if pid <= 0:
                continue
            if pid not in self._panel_option_lookup:
                self._panel_option_lookup[pid] = item
        return self._panel_option_lookup.get(panel_key)

    def _element_by_id_value(self, id_value):
        value = int(id_value or 0)
        if value <= 0:
            return None
        try:
            return self.doc.GetElement(revit_helpers.elementid_from_value(value))
        except Exception:
            return None

    def _schedule_id_value(self, schedule_view):
        if schedule_view is None:
            return 0
        try:
            return int(self._idval(getattr(schedule_view, "Id", None)))
        except Exception:
            return 0

    def _invalidate_schedule_cache(self, schedule_view=None, slots=None):
        """Clear cached slot-cell/layout mapping after schedule mutations."""
        if schedule_view is None:
            self._slot_cells_cache = {}
            self._layout_context_cache = {}
            return
        sid = self._schedule_id_value(schedule_view)
        if sid <= 0:
            self._slot_cells_cache = {}
            self._layout_context_cache = {}
            return
        if slots is None:
            self._layout_context_cache.pop(int(sid), None)
            stale_keys = [key for key in list(self._slot_cells_cache.keys()) if int(key[0]) == int(sid)]
            for key in stale_keys:
                self._slot_cells_cache.pop(key, None)
            return
        for slot in list(slots or []):
            slot_value = int(slot or 0)
            if slot_value <= 0:
                continue
            self._slot_cells_cache.pop((int(sid), int(slot_value)), None)

    def _slot_cells(self, schedule_view, slot, refresh=False):
        cache_key = self._schedule_slot_cache_key(schedule_view, slot)
        if (not bool(refresh)) and cache_key is not None and cache_key in self._slot_cells_cache:
            return list(self._slot_cells_cache.get(cache_key) or [])
        slot_value = int(slot or 0)
        if slot_value <= 0:
            return []
        if bool(refresh):
            self._invalidate_schedule_cache(schedule_view, slots=[int(slot_value)])
        try:
            cells = list(ps_repo.get_cells_by_slot_number(schedule_view, slot_value) or [])
        except Exception:
            cells = []
        ordered = self._order_cells_for_slot(schedule_view, slot_value, cells)
        if cache_key is not None:
            self._slot_cells_cache[cache_key] = list(ordered)
        return ordered

    def _schedule_slot_cache_key(self, schedule_view, slot):
        if schedule_view is None:
            return None
        sid = int(self._idval(getattr(schedule_view, "Id", None)))
        slot_value = int(slot or 0)
        if sid <= 0 or slot_value <= 0:
            return None
        return (sid, slot_value)

    def _layout_context(self, schedule_view):
        if schedule_view is None:
            return {"max_slot": 0, "sort_mode": "panelboard", "preferred_cols": {}}
        sid = int(self._idval(getattr(schedule_view, "Id", None)))
        if sid > 0 and sid in self._layout_context_cache:
            return dict(self._layout_context_cache.get(sid) or {})
        max_slot = 0
        try:
            table = schedule_view.GetTableData()
            max_slot = int(getattr(table, "NumberOfSlots", 0) or 0)
        except Exception:
            max_slot = 0
        sort_mode = ps_repo.classify_schedule_layout(schedule_view)
        col_counts = {1: {}, 2: {}}
        if max_slot > 0:
            for slot in range(1, int(max_slot) + 1):
                display_col = int(ps_repo.get_slot_display_column(slot, max_slot, sort_mode) or 1)
                for row, col in list(ps_repo.get_cells_by_slot_number(schedule_view, slot) or []):
                    cid = int(self._cell_circuit_id(schedule_view, row, col))
                    if cid <= 0:
                        continue
                    bucket = col_counts.setdefault(display_col, {})
                    bucket[int(col)] = int(bucket.get(int(col), 0) or 0) + 1
        preferred_cols = {}
        for display_col, bucket in col_counts.items():
            ordered = [pair[0] for pair in sorted(bucket.items(), key=lambda x: (-int(x[1]), int(x[0])))]
            preferred_cols[int(display_col)] = [int(x) for x in ordered]
        context = {"max_slot": int(max_slot), "sort_mode": sort_mode, "preferred_cols": preferred_cols}
        if sid > 0:
            self._layout_context_cache[sid] = dict(context)
        return context

    def _order_cells_for_slot(self, schedule_view, slot, cells):
        seen = set()
        ordered = []
        for pair in list(cells or []):
            if not pair or len(pair) < 2:
                continue
            key = (int(pair[0]), int(pair[1]))
            if key in seen:
                continue
            seen.add(key)
            ordered.append(key)
        if not ordered:
            return []
        context = self._layout_context(schedule_view)
        max_slot = int(context.get("max_slot", 0) or 0)
        sort_mode = context.get("sort_mode", "panelboard")
        display_col = int(ps_repo.get_slot_display_column(int(slot or 0), max_slot, sort_mode) or 1)
        preferred_cols = list((context.get("preferred_cols") or {}).get(display_col) or [])
        priority = {}
        for idx, col in enumerate(preferred_cols):
            priority[int(col)] = int(idx)
        fallback_rank = 999999
        ordered.sort(key=lambda pair: (int(priority.get(int(pair[1]), fallback_rank)), int(pair[1]), int(pair[0])))
        return ordered

    def _slot_cells_for_add(self, schedule_view, slot):
        """Return native slot cells for AddSpare/AddSpace calls."""
        slot_value = int(slot or 0)
        if slot_value <= 0 or schedule_view is None:
            return []
        cells = list(ps_repo.get_cells_by_slot_number(schedule_view, slot_value) or [])
        deduped = []
        seen = set()
        for row, col in list(cells or []):
            key = (int(row), int(col))
            if key in seen:
                continue
            seen.add(key)
            deduped.append(key)
        deduped.sort(key=lambda x: (int(x[0]), int(x[1])))
        return deduped

    def _cell_circuit_id(self, schedule_view, row, col):
        getter_id = getattr(schedule_view, "GetCircuitIdByCell", None)
        if getter_id is None:
            return 0
        try:
            cid = getter_id(int(row), int(col))
            if cid is None or cid == DB.ElementId.InvalidElementId:
                return 0
            return int(self._idval(cid))
        except Exception:
            return 0

    def _scan_cells_for_slot(self, schedule_view, slot, circuit_id=None):
        slot_value = int(slot or 0)
        if schedule_view is None or slot_value <= 0:
            return []
        try:
            table = schedule_view.GetTableData()
            body = table.GetSectionData(DB.SectionType.Body)
        except Exception:
            return []
        rows = int(getattr(body, "NumberOfRows", 0) or 0)
        cols = int(getattr(body, "NumberOfColumns", 0) or 0)
        if rows <= 0 or cols <= 0:
            return []
        target_cid = None
        if circuit_id is not None:
            target_cid = int(circuit_id or 0)

        matches = []
        for row in range(rows):
            for col in range(cols):
                if not self._is_cell_in_circuit_table(schedule_view, row, col):
                    continue
                try:
                    slot_no = int(schedule_view.GetSlotNumberByCell(int(row), int(col)) or 0)
                except Exception:
                    slot_no = 0
                if int(slot_no) != int(slot_value):
                    continue
                if target_cid is not None and target_cid > 0:
                    cid = int(self._cell_circuit_id(schedule_view, int(row), int(col)))
                    if int(cid) != int(target_cid):
                        continue
                matches.append((int(row), int(col)))
        return sorted(set(matches), key=lambda x: (int(x[1]), int(x[0])))

    def _is_cell_in_circuit_table(self, schedule_view, row, col):
        checker = getattr(schedule_view, "IsCellInCircuitTable", None)
        if checker is None:
            return True
        try:
            return bool(checker(int(row), int(col)))
        except Exception:
            return True

    def _slot_is_locked(self, schedule_view, slot):
        # Some schedules expose multiple cells per slot; treat slot as locked
        # when any mapped cell reports locked.
        for row, col in self._slot_cells(schedule_view, slot):
            try:
                if bool(ps_repo._slot_is_locked(schedule_view, row, col)):
                    return True
            except Exception:
                continue
        return False

    def _set_slot_locked(self, schedule_view, slot, is_locked):
        slot_value = int(slot or 0)
        if schedule_view is None or slot_value <= 0:
            return False
        lock_value = bool(is_locked)
        cells = [c for c in list(self._slot_cells(schedule_view, slot_value) or []) if len(c) >= 2]
        if not cells:
            cells = list(self._scan_cells_for_slot(schedule_view, slot_value, circuit_id=None) or [])
        setter = getattr(schedule_view, "SetLockSlot", None)
        if setter is None:
            return False
        for row, col in list(cells or []):
            try:
                setter(int(row), int(col), bool(lock_value))
                return True
            except Exception:
                continue
        return False

    def _unlock_slots_with_snapshot(self, schedule_view, slots):
        snapshot = {}
        for slot in sorted(set([int(x) for x in list(slots or []) if int(x) > 0])):
            locked = bool(self._slot_is_locked(schedule_view, slot))
            snapshot[int(slot)] = locked
            if not bool(self._set_slot_locked(schedule_view, slot, False)):
                self.logger.warning("Unlock attempt returned no-op for slot %s.", int(slot))
            elif bool(self._slot_is_locked(schedule_view, slot)):
                self.logger.warning("Slot %s still reports locked after unlock attempt.", int(slot))
        return snapshot

    def _restore_slot_locks(self, schedule_view, snapshot):
        for slot, was_locked in dict(snapshot or {}).items():
            if bool(was_locked):
                self._set_slot_locked(schedule_view, int(slot), True)

    def _get_circuit_at_slot(self, schedule_view, slot):
        def _resolve_from_cells(cells):
            for row, col in list(cells or []):
                getter = getattr(schedule_view, "GetCircuitByCell", None)
                if getter is not None:
                    try:
                        circuit = getter(int(row), int(col))
                        if isinstance(circuit, DBE.ElectricalSystem):
                            return circuit
                    except Exception:
                        pass
                getter_id = getattr(schedule_view, "GetCircuitIdByCell", None)
                if getter_id is not None:
                    try:
                        cid = getter_id(int(row), int(col))
                        if cid is None or cid == DB.ElementId.InvalidElementId:
                            continue
                        circuit = self.doc.GetElement(cid)
                        if isinstance(circuit, DBE.ElectricalSystem):
                            return circuit
                    except Exception:
                        pass
            return None

        cached_cells = self._slot_cells(schedule_view, slot, refresh=False)
        resolved = _resolve_from_cells(cached_cells)
        if isinstance(resolved, DBE.ElectricalSystem):
            return resolved
        # Retry once from fresh API cells in case cached row/col map is stale.
        fresh_cells = self._slot_cells(schedule_view, slot, refresh=True)
        return _resolve_from_cells(fresh_cells)

    def _find_circuit_slots_in_schedule(self, schedule_view, circuit_id):
        target_id = int(circuit_id or 0)
        if schedule_view is None or target_id <= 0:
            return []
        try:
            table = schedule_view.GetTableData()
            body = table.GetSectionData(DB.SectionType.Body)
        except Exception:
            return []
        rows = int(getattr(body, "NumberOfRows", 0) or 0)
        cols = int(getattr(body, "NumberOfColumns", 0) or 0)
        if rows <= 0 or cols <= 0:
            return []
        slots = set()
        for row in range(rows):
            for col in range(cols):
                if not self._is_cell_in_circuit_table(schedule_view, row, col):
                    continue
                cid = int(self._cell_circuit_id(schedule_view, int(row), int(col)))
                if int(cid) != int(target_id):
                    continue
                try:
                    slot_no = int(schedule_view.GetSlotNumberByCell(int(row), int(col)) or 0)
                except Exception:
                    slot_no = 0
                if slot_no > 0:
                    slots.add(int(slot_no))
        return sorted(slots)

    def _resolve_circuit_start_slot_in_schedule(self, schedule_view, circuit, fallback_slot=0):
        circuit_id = int(self._idval(getattr(circuit, "Id", None)))
        slots = self._find_circuit_slots_in_schedule(schedule_view, circuit_id)
        if slots:
            return int(slots[0])
        try:
            start_slot = int(ps_repo.get_circuit_start_slot(circuit) or 0)
        except Exception:
            start_slot = 0
        if start_slot > 0:
            return int(start_slot)
        return int(fallback_slot or 0)

    def _get_circuit_poles(self, circuit, fallback=1):
        poles = None
        for attr in ("PolesNumber", "NumberOfPoles"):
            try:
                value = getattr(circuit, attr, None)
                if value is not None:
                    poles = int(value)
                    break
            except Exception:
                continue
        if poles is None:
            try:
                poles = int(ps_repo.get_circuit_voltage_poles(circuit)[1] or 0)
            except Exception:
                poles = 0
        return int(max(1, poles or fallback or 1))

    def _covered_slots_for_circuit(self, option, circuit, fallback_slot=0, fallback_poles=1):
        if option is None or circuit is None:
            slot_value = int(fallback_slot or 0)
            return [slot_value] if slot_value > 0 else []
        slot_value = int(ps_repo.get_circuit_start_slot(circuit) or fallback_slot or 0)
        if slot_value <= 0:
            return []
        poles = self._get_circuit_poles(circuit, fallback=fallback_poles)
        covered = ps_repo.get_slot_span_slots(
            start_slot=slot_value,
            pole_count=poles,
            max_slot=option.get("max_slot", 0),
            sort_mode=option.get("sort_mode", "panelboard"),
        )
        return covered or [int(slot_value)]

    def _move_slot_to(self, schedule_view, from_slot, to_slot, circuit_id=0):
        source_slot = int(from_slot or 0)
        target_slot = int(to_slot or 0)
        if source_slot <= 0 or target_slot <= 0 or source_slot == target_slot:
            return
        mover = getattr(schedule_view, "MoveSlotTo", None)
        if mover is None:
            raise Exception("PanelScheduleView.MoveSlotTo is unavailable.")

        src_cells = list(self._slot_cells(schedule_view, source_slot, refresh=True))
        dst_cells = list(self._slot_cells(schedule_view, target_slot, refresh=True))
        src_cells = [c for c in list(src_cells or []) if self._is_cell_in_circuit_table(schedule_view, c[0], c[1])]
        dst_cells = [c for c in list(dst_cells or []) if self._is_cell_in_circuit_table(schedule_view, c[0], c[1])]
        src_ordered = list(src_cells)
        if int(circuit_id or 0) > 0 and src_cells:
            owned = [cell for cell in src_cells if int(self._cell_circuit_id(schedule_view, cell[0], cell[1])) == int(circuit_id)]
            if not owned:
                owned = list(self._scan_cells_for_slot(schedule_view, source_slot, circuit_id=int(circuit_id)))
            if owned:
                src_ordered = owned
        if not src_ordered:
            src_ordered = list(self._scan_cells_for_slot(schedule_view, source_slot, circuit_id=None))
        dst_ordered = list(dst_cells)
        scanned_dst = list(self._scan_cells_for_slot(schedule_view, target_slot, circuit_id=None))
        if scanned_dst:
            dst_seen = set([tuple([int(x[0]), int(x[1])]) for x in list(dst_ordered or [])])
            for row, col in list(scanned_dst or []):
                key = (int(row), int(col))
                if key in dst_seen:
                    continue
                dst_ordered.append(key)
                dst_seen.add(key)

        attempts = []
        seen_attempts = set()

        # Prefer same-column mapping between source and destination.
        for s_row, s_col in src_ordered:
            for d_row, d_col in dst_ordered:
                if int(s_col) != int(d_col):
                    continue
                key = (int(s_row), int(s_col), int(d_row), int(d_col))
                if key in seen_attempts:
                    continue
                seen_attempts.add(key)
                attempts.append(key)

        # Add low-cost directional fallbacks (avoid full cartesian noise first).
        if src_ordered and dst_ordered:
            s0_row, s0_col = src_ordered[0]
            d0_row, d0_col = dst_ordered[0]
            key = (int(s0_row), int(s0_col), int(d0_row), int(d0_col))
            if key not in seen_attempts:
                seen_attempts.add(key)
                attempts.append(key)
            for d_row, d_col in dst_ordered:
                key = (int(s0_row), int(s0_col), int(d_row), int(d_col))
                if key in seen_attempts:
                    continue
                seen_attempts.add(key)
                attempts.append(key)
            for s_row, s_col in src_ordered:
                key = (int(s_row), int(s_col), int(d0_row), int(d0_col))
                if key in seen_attempts:
                    continue
                seen_attempts.add(key)
                attempts.append(key)

        # Final fallback: full cartesian attempts.
        for s_row, s_col in src_ordered:
            for d_row, d_col in dst_ordered:
                key = (int(s_row), int(s_col), int(d_row), int(d_col))
                if key in seen_attempts:
                    continue
                seen_attempts.add(key)
                attempts.append(key)
        if not attempts:
            raise Exception("No valid source/target body cells resolved for MoveSlotTo.")

        self.logger.info(
            "MoveSlotTo try from_slot=%s to_slot=%s ckt=%s src_cells=%s dst_cells=%s attempts=%s",
            int(source_slot),
            int(target_slot),
            int(circuit_id or 0),
            str(src_ordered),
            str(dst_ordered),
            int(len(attempts)),
        )
        errors = []
        for idx, args in enumerate(attempts, 1):
            try:
                result = mover(*args)
                if isinstance(result, bool) and not result:
                    errors.append("args={0} -> False".format(args))
                    continue
                self.logger.info(
                    "MoveSlotTo success from_slot=%s to_slot=%s ckt=%s attempt=%s args=%s",
                    int(source_slot),
                    int(target_slot),
                    int(circuit_id or 0),
                    int(idx),
                    str(args),
                )
                self._invalidate_schedule_cache(schedule_view)
                return
            except Exception as ex:
                errors.append("args={0} -> {1}".format(args, str(ex)))
                continue
        raise Exception(
            "MoveSlotTo failed from slot {0} to {1}. attempts={2} details={3}".format(
                int(source_slot),
                int(target_slot),
                int(len(attempts)),
                "; ".join(errors[:6]) if errors else "no-details",
            )
        )

    def _select_panel_for_circuit(self, circuit, panel):
        selector = getattr(circuit, "SelectPanel", None)
        if selector is None:
            raise Exception("ElectricalSystem.SelectPanel is unavailable.")
        result = selector(panel)
        if isinstance(result, bool) and not result:
            raise Exception("SelectPanel returned False.")

    def _set_circuit_poles(self, circuit, poles):
        try:
            target = int(max(1, poles or 1))
        except Exception:
            return
        for attr in ("NumberOfPoles", "PolesNumber"):
            try:
                setattr(circuit, attr, int(target))
                return
            except Exception:
                continue
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
        except Exception:
            param = None
        if not param:
            return
        try:
            if bool(getattr(param, "IsReadOnly", False)):
                return
        except Exception:
            pass
        try:
            param.Set(int(target))
        except Exception:
            pass

    def _set_circuit_rating(self, circuit, amps):
        try:
            target_value = float(int(amps))
        except Exception:
            return
        if target_value <= 0:
            return
        try:
            setattr(circuit, "Rating", float(target_value))
            return
        except Exception:
            pass
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
        except Exception:
            param = None
        if not param:
            return
        try:
            if bool(getattr(param, "IsReadOnly", False)):
                return
        except Exception:
            pass
        try:
            param.Set(float(target_value))
        except Exception:
            pass

    def _set_circuit_frame(self, circuit, amps):
        """Set circuit frame/ampere frame when a writable parameter is available."""
        try:
            target_value = int(amps)
        except Exception:
            return
        if target_value <= 0:
            return

        for attr in ("Frame", "FrameRating"):
            try:
                setattr(circuit, attr, int(target_value))
                return
            except Exception:
                continue

        bip_names = (
            "RBS_ELEC_CIRCUIT_FRAME_PARAM",
            "RBS_ELEC_FRAME",
        )
        for bip_name in bip_names:
            try:
                bip = getattr(DB.BuiltInParameter, bip_name)
            except Exception:
                bip = None
            if bip is None:
                continue
            try:
                param = circuit.get_Parameter(bip)
            except Exception:
                param = None
            if not param:
                continue
            try:
                if bool(getattr(param, "IsReadOnly", False)):
                    continue
            except Exception:
                pass
            try:
                if bool(param.Set(int(target_value))):
                    return
            except Exception:
                continue

        try:
            frame_param = revit_helpers.get_parameter(
                circuit,
                "Frame",
                include_type=False,
                case_insensitive=True,
            )
        except Exception:
            frame_param = None
        if not frame_param:
            return
        try:
            if bool(getattr(frame_param, "IsReadOnly", False)):
                return
        except Exception:
            pass
        try:
            frame_param.Set(int(target_value))
        except Exception:
            pass

    def _set_circuit_notes(self, circuit, notes_text):
        text = str(notes_text or "").strip()
        if text == "":
            return
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
        except Exception:
            param = None
        if not param:
            return
        try:
            if bool(getattr(param, "IsReadOnly", False)):
                return
        except Exception:
            pass
        try:
            param.Set(text)
        except Exception:
            pass

    def _set_circuit_load_name(self, circuit, load_name):
        text = str(load_name or "").strip()
        if text == "":
            return
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
        except Exception:
            param = None
        if not param:
            return
        try:
            if bool(getattr(param, "IsReadOnly", False)):
                return
        except Exception:
            pass
        try:
            param.Set(text)
        except Exception:
            pass

    def _add_special_to_slot(self, schedule_view, slot, kind):
        action = SpecialKind.normalize(kind, None)
        if action is None:
            raise Exception("Unsupported special kind: {0}".format(kind))
        method_name = "AddSpare" if action == SpecialKind.SPARE else "AddSpace"
        method = getattr(schedule_view, method_name, None)
        if method is None:
            raise Exception("{0} is unavailable.".format(method_name))

        cells = list(self._slot_cells_for_add(schedule_view, slot))
        errors = []
        for row, col in list(cells or []):
            try:
                result = method(int(row), int(col))
                if isinstance(result, bool) and not result:
                    errors.append("({0},{1}) -> False".format(int(row), int(col)))
                    continue
                return
            except Exception as ex:
                errors.append("({0},{1}) -> {2}".format(int(row), int(col), str(ex)))
                continue

        self.logger.warning(
            "Add %s failed at slot %s attempts=%s details=%s",
            str(action).upper(),
            int(slot),
            int(len(errors)),
            " | ".join(list(errors or [])[:8]),
        )
        raise Exception("Could not add {0} at slot {1}.".format(action.upper(), int(slot)))

    def _add_special(
        self,
        panel_id,
        panel_slot,
        kind,
        poles=1,
        rating=0,
        frame=0,
        unlock=True,
        load_name=None,
        schedule_notes=None,
        apply_custom_properties=True,
        force_switchboard_three_p=False,
        restrict_data_to_space=True,
    ):
        panel_id_value = int(panel_id or 0)
        option = self._option_for_panel_id(panel_id_value)
        if option is None:
            raise Exception("Missing panel option for add operation.")
        schedule_view = option.get("schedule_view")
        if schedule_view is None:
            raise Exception("Panel has no schedule view.")
        slot_value = int(panel_slot or 0)
        if slot_value <= 0:
            raise Exception("Invalid slot for add operation.")
        if not bool(ps_repo.is_slot_valid_for_option(option, int(slot_value))):
            raise Exception(
                "Target slot {0} exceeds equipment-supported slot capacity for this panel.".format(
                    int(slot_value)
                )
            )

        requested_poles = int(max(1, poles or 1))
        requested_rating = int(rating or 0)
        requested_frame = int(frame or 0)
        slot_span_hint = ps_repo.get_slot_span_slots_for_option(
            option,
            start_slot=int(slot_value),
            pole_count=int(requested_poles),
            require_valid=True,
        )
        if not slot_span_hint:
            raise Exception(
                "Target slot {0} / pole span exceeds equipment-supported slot capacity.".format(int(slot_value))
            )

        if bool(unlock):
            self._unlock_slots_with_snapshot(schedule_view, [int(x) for x in slot_span_hint if int(x) > 0])

        self._add_special_to_slot(schedule_view, int(slot_value), kind)
        self._invalidate_schedule_cache(schedule_view, slots=[int(slot_value)])

        is_switchboard = False
        try:
            is_switchboard = bool(option.get("schedule_type") == ps_repo.PSTYPE_SWITCHBOARD)
        except Exception:
            is_switchboard = False
        is_data_panel = False
        try:
            is_data_panel = bool(option.get("schedule_type") == ps_repo.PSTYPE_DATA)
        except Exception:
            is_data_panel = False
        kind_key = SpecialKind.normalize(kind, None)
        if kind_key is None:
            raise Exception("Unsupported special kind: {0}".format(kind))
        if bool(restrict_data_to_space and is_data_panel and kind_key != SpecialKind.SPACE):
            raise Exception("Data panels only allow adding SPACE.")

        occupant = self._get_circuit_at_slot(schedule_view, int(slot_value))
        if not isinstance(occupant, DBE.ElectricalSystem):
            # Fallback only when immediate lookup fails after add.
            self.doc.Regenerate()
            self._invalidate_schedule_cache(schedule_view, slots=[int(slot_value)])
            occupant = self._get_circuit_at_slot(schedule_view, int(slot_value))
        if not isinstance(occupant, DBE.ElectricalSystem):
            raise Exception("Added {0} could not be resolved at slot {1}.".format(str(kind).upper(), int(slot_value)))

        target_poles = int(requested_poles)
        if bool(force_switchboard_three_p and is_switchboard):
            target_poles = 3
        if bool(is_data_panel):
            target_poles = 1

        if not bool(is_data_panel):
            if bool(apply_custom_properties):
                self._set_circuit_poles(occupant, int(target_poles))
            elif bool(force_switchboard_three_p and is_switchboard):
                self._set_circuit_poles(occupant, 3)
            if bool(apply_custom_properties) and kind_key == SpecialKind.SPARE:
                if requested_rating > 0:
                    self._set_circuit_rating(occupant, requested_rating)
                if requested_frame > 0:
                    self._set_circuit_frame(occupant, requested_frame)
        if bool(apply_custom_properties):
            self._set_circuit_load_name(occupant, load_name)
            self._set_circuit_notes(occupant, schedule_notes)

        if bool(unlock):
            covered_slots = self._covered_slots_for_circuit(
                option,
                occupant,
                fallback_slot=int(slot_value),
                fallback_poles=int(target_poles),
            )
            for slot in [int(x) for x in list(covered_slots or []) if int(x) > 0]:
                self._set_slot_locked(schedule_view, int(slot), False)

        return {
            "panel_id": panel_id_value,
            "slot": int(slot_value),
            "circuit_id": int(self._idval(occupant.Id)),
        }

    def _remove_special(self, panel_id, panel_slot, kind_hint=None, restore_lock_state=True):
        panel_id_value = int(panel_id or 0)
        option = self._option_for_panel_id(panel_id_value)
        if option is None:
            raise Exception("Missing panel option for remove operation.")
        schedule_view = option.get("schedule_view")
        if schedule_view is None:
            raise Exception("Panel has no schedule view.")
        slot_value = int(panel_slot or 0)
        if slot_value <= 0:
            raise Exception("Invalid slot for remove operation.")
        self._invalidate_schedule_cache(schedule_view, slots=[int(slot_value)])
        snapshot = self._unlock_slots_with_snapshot(schedule_view, [int(slot_value)])
        try:
            target = self._get_circuit_at_slot(schedule_view, int(slot_value))
            if not isinstance(target, DBE.ElectricalSystem):
                raise Exception("No removable spare/space found at target slot.")
            kind = str(ps_repo._kind_from_circuit(target) or "").lower()
            if kind_hint and kind != str(kind_hint).lower():
                raise Exception("Target slot does not contain {0}.".format(str(kind_hint).upper()))
            if kind not in SpecialKind.all():
                raise Exception("Target slot is occupied by non spare/space.")
            removed_id = int(self._idval(target.Id))
            self.doc.Delete(target.Id)
            self._invalidate_schedule_cache(schedule_view)
            return {"panel_id": panel_id_value, "slot": int(slot_value), "removed_circuit_id": int(removed_id)}
        finally:
            if bool(restore_lock_state):
                self._restore_slot_locks(schedule_view, snapshot)

    def _remove_specials_in_slots(self, target_option, slots, protected_circuit_id=0):
        schedule_view = target_option.get("schedule_view")
        to_delete = {}
        covered_slots = set([int(x) for x in list(slots or []) if int(x) > 0])
        self._invalidate_schedule_cache(schedule_view, slots=list(covered_slots))
        for slot in list(covered_slots):
            occupant = self._get_circuit_at_slot(schedule_view, slot)
            if not isinstance(occupant, DBE.ElectricalSystem):
                continue
            occ_id = int(self._idval(occupant.Id))
            if occ_id <= 0 or occ_id == int(protected_circuit_id or 0):
                continue
            kind = str(ps_repo._kind_from_circuit(occupant) or "").lower()
            if kind not in SpecialKind.all():
                raise Exception("Target slot {0} is occupied by a non-spare/space circuit.".format(int(slot)))
            occ_slots = self._covered_slots_for_circuit(target_option, occupant, fallback_slot=slot, fallback_poles=1)
            to_delete[occ_id] = {"circuit": occupant, "slots": occ_slots}
            for occ_slot in occ_slots:
                covered_slots.add(int(occ_slot))

        snapshot = self._unlock_slots_with_snapshot(schedule_view, covered_slots)
        for occ in to_delete.values():
            self.doc.Delete(occ["circuit"].Id)
        self._invalidate_schedule_cache(schedule_view)
        return snapshot
