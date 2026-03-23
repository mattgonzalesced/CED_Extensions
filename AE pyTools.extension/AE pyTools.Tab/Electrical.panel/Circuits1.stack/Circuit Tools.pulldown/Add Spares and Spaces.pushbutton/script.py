# -*- coding: utf-8 -*-
"""Stage add/remove spare-space actions across panel schedules."""

import os
import sys

from System.Windows.Controls import Button, DataGridRow
from System.Windows.Media import VisualTreeHelper
from pyrevit import DB, forms, revit, script

TITLE = "Add / Remove Spares and Spaces"
THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_ACCENT_KEY = "accent_mode"
VALID_ACCENT_MODES = ("blue", "red", "green", "neutral")

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
_FALLBACK_LIB_ROOT = os.path.abspath(os.path.join(THIS_DIR, "..", "..", "..", "..", "..", "CEDLib.lib"))
if os.path.isdir(_FALLBACK_LIB_ROOT) and _FALLBACK_LIB_ROOT not in sys.path:
    sys.path.append(_FALLBACK_LIB_ROOT)

from UIClasses import pathing as ui_pathing

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate workspace root for Add/Remove Spares and Spaces.", title=TITLE)
    raise SystemExit

from UIClasses import resource_loader
from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Model.panel_schedule_manager import PanelScheduleManager

LOGGER = script.get_logger()
UI_RESOURCES_ROOT = ui_pathing.resolve_ui_resources_root(LIB_ROOT)


def _find_visual_ancestor(node, target_type):
    """Return nearest ancestor of target_type for a WPF visual node."""
    current = node
    while current is not None:
        if isinstance(current, target_type):
            return current
        try:
            current = VisualTreeHelper.GetParent(current)
        except Exception:
            return None
    return None


def _is_descendant_of_control(node, control):
    """Return True when node is inside the given WPF control."""
    current = node
    while current is not None:
        if current is control:
            return True
        try:
            current = VisualTreeHelper.GetParent(current)
        except Exception:
            return False
    return False


def _normalize_accent_mode(value, fallback="blue"):
    mode = str(value or fallback).strip().lower()
    return mode if mode in VALID_ACCENT_MODES else fallback


def _load_accent_mode(default_accent="blue"):
    accent_mode = _normalize_accent_mode(default_accent, "blue")
    try:
        cfg = script.get_config(THEME_CONFIG_SECTION)
        if cfg is None:
            return accent_mode
        accent_mode = _normalize_accent_mode(cfg.get_option(THEME_CONFIG_ACCENT_KEY, accent_mode), accent_mode)
    except Exception:
        pass
    return accent_mode


def _collect_all_circuits(doc):
    """Return all electrical systems once for batch removable evaluation."""
    try:
        return list(
            DB.FilteredElementCollector(doc)
            .OfClass(ps_repo.DBE.ElectricalSystem)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


def _open_slot_numbers(option):
    """Return open slots for a panel option in schedule display order."""
    empties = dict(ps_repo.gather_empty_slot_cells(option.get("schedule_view")) or {})
    if not empties:
        return []
    max_slot = int(option.get("max_slot", 0) or 0)
    sort_mode = option.get("sort_mode", "panelboard")
    order = list(ps_repo.get_slot_order(max_slot, sort_mode) or [])
    if order:
        return [int(slot) for slot in order if int(slot) in empties]
    return sorted([int(x) for x in empties.keys() if int(x) > 0])


def _count_open_slots_fast(option):
    """Return open slot count using empty-cell slot map (fast path)."""
    return int(len(_open_slot_numbers(option)))


def _removable_slots_for_option(doc, option, all_circuits):
    """Return removable spare/space slots for one panel option."""
    rows = list(ps_repo.build_panel_rows(doc, option, all_circuits=all_circuits) or [])
    removable_spare_slots = []
    removable_space_slots = []
    for row in rows:
        kind = str(row.get("kind", "") or "").lower()
        slot_value = int(row.get("slot", 0) or 0)
        if slot_value <= 0:
            continue
        if kind == "spare" and bool(row.get("is_spare_removable", False)):
            removable_spare_slots.append(slot_value)
        elif kind == "space" and bool(row.get("is_space_removable", False)):
            removable_space_slots.append(slot_value)
    return sorted(set(removable_spare_slots)), sorted(set(removable_space_slots))


def _execute_add_for_option(manager, option, mode):
    """Execute add special action for one panel option."""
    panel_id = int(option.get("panel_id", 0) or 0)
    slots = list(_open_slot_numbers(option) or [])
    if panel_id <= 0 or not slots:
        return {"added_spares": 0, "added_spaces": 0}

    mode_key = str(mode or "").lower()
    if mode_key not in ("spare", "space", "mixed"):
        mode_key = "space"

    add_spares = 0
    add_spaces = 0
    for index, slot in enumerate(slots):
        if mode_key == "spare":
            kind = "spare"
        elif mode_key == "space":
            kind = "space"
        else:
            kind = "spare" if int(index % 2) == 0 else "space"

        if kind == "spare":
            manager.add_spare(
                panel_id=panel_id,
                panel_slot=int(slot),
                poles=1,
                rating=20,
                frame=20,
                unlock=True,
            )
            add_spares += 1
        else:
            manager.add_space(
                panel_id=panel_id,
                panel_slot=int(slot),
                poles=1,
                unlock=True,
            )
            add_spaces += 1

    return {"added_spares": int(add_spares), "added_spaces": int(add_spaces)}


def _execute_remove_for_option(manager, doc, option, mode, all_circuits):
    """Execute removable-only remove special action for one panel option."""
    panel_id = int(option.get("panel_id", 0) or 0)
    if panel_id <= 0:
        return {"removed_spares": 0, "removed_spaces": 0}

    mode_key = str(mode or "").lower()
    remove_spares = mode_key in ("spare", "both")
    remove_spaces = mode_key in ("space", "both")
    removable_spare_slots, removable_space_slots = _removable_slots_for_option(doc, option, all_circuits)

    removed_spares = 0
    removed_spaces = 0
    if remove_spares:
        for slot in list(removable_spare_slots or []):
            manager.remove_spare(panel_id=panel_id, panel_slot=int(slot))
            removed_spares += 1
    if remove_spaces:
        for slot in list(removable_space_slots or []):
            manager.remove_space(panel_id=panel_id, panel_slot=int(slot))
            removed_spaces += 1

    return {"removed_spares": int(removed_spares), "removed_spaces": int(removed_spaces)}


def _format_panel_info(option):
    """Return one-line panel info for quick mode header."""
    if not option:
        return "Unknown Panel"
    return "{0} ({1}) | {2} | Open Slots: {3}".format(
        str(option.get("panel_name", "") or "Unnamed Panel"),
        str(option.get("part_type_name", "") or option.get("board_type", "Unknown")),
        str(option.get("dist_system_name", "") or "Unknown Dist. System"),
        str(_count_open_slots_fast(option)),
    )


def _get_active_schedule_option(doc):
    """Return panel option for active PanelScheduleView, or None."""
    try:
        active_view = revit.uidoc.ActiveView
    except Exception:
        active_view = None
    if not isinstance(active_view, ps_repo.DBE.PanelScheduleView):
        return None

    panel = ps_repo.resolve_schedule_panel(doc, active_view)
    if panel is None:
        return None
    options = list(
        ps_repo.collect_panel_equipment_options(
            doc,
            panels=[panel],
            include_without_schedule=True,
        )
        or []
    )
    if not options:
        return None
    option = options[0]
    ps_repo.attach_schedule_to_option(doc, option, active_view)
    return option


def _run_quick_action(doc, option, action_type, mode):
    """Run single-panel quick action in one transaction."""
    panel_id = int((option or {}).get("panel_id", 0) or 0)
    if panel_id <= 0:
        forms.alert("Could not resolve active panel for quick action.", title=TITLE)
        return

    mode_key = str(mode or "").lower()
    if str(action_type or "").lower() == "add":
        if mode_key == "both":
            mode_key = "mixed"
        elif mode_key not in ("spare", "space"):
            mode_key = "space"
    else:
        if mode_key not in ("spare", "space", "both"):
            mode_key = "both"

    manager = PanelScheduleManager(doc, panel_option_lookup={int(panel_id): option})
    all_circuits = _collect_all_circuits(doc) if str(action_type or "").lower() == "remove" else None
    tx = DB.Transaction(doc, "Panel Quick Action - Spares/Spaces")
    tx.Start()
    try:
        if str(action_type or "").lower() == "add":
            result = _execute_add_for_option(manager, option, mode_key)
            message = "Added Spare: {0}, Space: {1}".format(
                int(result.get("added_spares", 0) or 0),
                int(result.get("added_spaces", 0) or 0),
            )
        else:
            result = _execute_remove_for_option(manager, doc, option, mode_key, all_circuits)
            message = "Removed Spare: {0}, Space: {1}".format(
                int(result.get("removed_spares", 0) or 0),
                int(result.get("removed_spaces", 0) or 0),
            )
        status = tx.Commit()
        if status != DB.TransactionStatus.Committed:
            raise Exception("Transaction did not commit.")
    except Exception as ex:
        try:
            if tx.GetStatus() == DB.TransactionStatus.Started:
                tx.RollBack()
        except Exception:
            pass
        forms.alert("Quick action failed and was rolled back.\n\n{0}".format(str(ex)), title=TITLE)
        return

    forms.alert(message, title=TITLE)


class PanelListItem(object):
    """List row view-model for one panel schedule option."""

    def __init__(self, option, open_slots):
        self.option = option
        self.panel_id = int(option.get("panel_id", 0) or 0)
        self.panel_name = str(option.get("panel_name", "") or "Unnamed Panel")
        self.part_type = str(option.get("part_type_name", "") or option.get("board_type", "Unknown"))
        self.dist_system_name = str(option.get("dist_system_name", "") or "Unknown Dist. System")
        self.open_slots = int(max(0, open_slots or 0))
        self.open_slots_text = str(self.open_slots)
        self.action_text = ""
        self.is_checked = False


class QuickPanelActionWindow(forms.WPFWindow):
    """Super-lightweight active-panel action window."""

    def __init__(self, option, accent_mode):
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "AddSparesSpacesQuickWindow.xaml"))
        self._accent_mode = _normalize_accent_mode(accent_mode, "blue")
        self.option = option
        self.result = None
        forms.WPFWindow.__init__(self, xaml_path)
        self._apply_theme()
        self._init_controls()

    def _apply_theme(self):
        try:
            resource_loader.apply_theme(
                self,
                resources_root=UI_RESOURCES_ROOT,
                theme_mode="light",
                accent_mode=self._accent_mode,
            )
        except Exception as ex:
            LOGGER.warning("Quick panel action theme apply failed: %s", ex)

    def _init_controls(self):
        self.PanelInfoText = self.FindName("PanelInfoText")
        self.ModeFillRadio = self.FindName("ModeFillRadio")
        self.ModeRemoveRadio = self.FindName("ModeRemoveRadio")
        if self.PanelInfoText is not None:
            self.PanelInfoText.Text = _format_panel_info(self.option)

    def _selected_action_type(self):
        return "remove" if bool(getattr(self.ModeRemoveRadio, "IsChecked", False)) else "add"

    def _finish(self, mode):
        self.result = {"action_type": self._selected_action_type(), "mode": str(mode or "")}
        self.Close()

    def quick_spare_clicked(self, sender, args):
        self._finish("spare")

    def quick_space_clicked(self, sender, args):
        self._finish("space")

    def quick_both_clicked(self, sender, args):
        self._finish("both")


class AddRemoveSparesSpacesWindow(forms.WPFWindow):
    """Staged planner for adding/removing panel schedule specials."""

    def __init__(self, accent_mode):
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "AddSparesSpacesWindow.xaml"))
        self._accent_mode = _normalize_accent_mode(accent_mode, "blue")
        self._items = []
        self._item_by_panel_id = {}
        self._staged_actions_by_panel = {}
        self._is_syncing_checks = False
        self._suppress_check_events = False
        self._is_ready = False
        self._last_selected_rows = []
        forms.WPFWindow.__init__(self, xaml_path)
        self._apply_theme()
        self._init_controls()
        self._load_panels()
        self.Loaded += self.window_loaded

    def _active_doc(self):
        """Return current active document."""
        try:
            return revit.doc
        except Exception:
            return None

    def _apply_theme(self):
        """Apply forced light UI theme with configured accent."""
        try:
            resource_loader.apply_theme(
                self,
                resources_root=UI_RESOURCES_ROOT,
                theme_mode="light",
                accent_mode=self._accent_mode,
            )
        except Exception as ex:
            LOGGER.warning("Add/Remove Spares and Spaces theme apply failed: %s", ex)

    def _init_controls(self):
        """Resolve XAML controls."""
        self.StatusText = self.FindName("StatusText")
        self.PanelsList = self.FindName("PanelsList")
        self.ModeSpareRadio = self.FindName("ModeSpareRadio")
        self.ModeSpaceRadio = self.FindName("ModeSpaceRadio")
        self.ModeBothRadio = self.FindName("ModeBothRadio")
        self.CheckedStatusText = self.FindName("CheckedStatusText")
        self.StagedStatusText = self.FindName("StagedStatusText")
        self.StageAddButton = self.FindName("StageAddButton")
        self.StageRemoveButton = self.FindName("StageRemoveButton")
        self.ResetSelectedButton = self.FindName("ResetSelectedButton")
        self.ApplyButton = self.FindName("ApplyButton")

    def _set_status(self, text):
        """Set status line text."""
        if self.StatusText is not None:
            self.StatusText.Text = str(text or "")

    def _load_panels(self):
        """Load all non-template panel schedule options into list."""
        doc = self._active_doc()
        if doc is None:
            forms.alert("No active Revit document.", title=TITLE, exitscript=True)

        options = list(
            ps_repo.collect_panel_equipment_options(
                doc,
                include_without_schedule=False,
            )
            or []
        )
        options = [x for x in options if x is not None and x.get("schedule_view") is not None]
        options.sort(key=lambda x: (str(x.get("panel_name", "") or ""), str(x.get("dist_system_name", "") or "")))
        if not options:
            forms.alert("No panel schedule views found in this model.", title=TITLE, exitscript=True)

        self._items = []
        self._item_by_panel_id = {}
        for option in options:
            open_slots = _count_open_slots_fast(option)
            item = PanelListItem(option, open_slots)
            self._items.append(item)
            self._item_by_panel_id[int(item.panel_id)] = item

        self.PanelsList.ItemsSource = self._items
        self._set_status("Loaded {0} panel schedules.".format(len(self._items)))
        self._refresh_status()

    def _refresh_panel_list(self):
        """Refresh list control after row updates."""
        if self.PanelsList is None:
            return
        try:
            self.PanelsList.Items.Refresh()
        except Exception:
            self.PanelsList.ItemsSource = None
            self.PanelsList.ItemsSource = self._items

    def _selected_items(self):
        """Return checked panel rows."""
        selected = []
        for item in list(self._items or []):
            if bool(getattr(item, "is_checked", False)):
                selected.append(item)
        return selected

    def _selected_grid_items(self):
        """Return currently selected rows in grid."""
        if self.PanelsList is None:
            return []
        try:
            return list(self.PanelsList.SelectedItems or [])
        except Exception:
            return []

    def _selected_mode(self):
        """Return selected mode key from shared mode radio group."""
        if bool(getattr(self.ModeBothRadio, "IsChecked", False)):
            return "both"
        if bool(getattr(self.ModeSpaceRadio, "IsChecked", False)):
            return "space"
        return "spare"

    def _stage_actions(self, action_type, mode, selected_items):
        """Stage one action per selected panel, overwriting previous row stage."""
        count = 0
        for item in list(selected_items or []):
            self._staged_actions_by_panel[int(item.panel_id)] = {
                "panel_id": int(item.panel_id),
                "action_type": str(action_type),
                "mode": str(mode),
            }
            count += 1
        self._rebuild_action_column()
        return int(count)

    def _rebuild_action_column(self):
        """Update action-column summary text from staged panel-action map."""
        actions_by_panel = dict(self._staged_actions_by_panel or {})
        for item in list(self._items or []):
            action = actions_by_panel.get(int(item.panel_id))
            if not action:
                item.action_text = ""
                continue
            kind = str(action.get("action_type", "") or "").lower()
            mode = str(action.get("mode", "") or "").lower()
            if kind == "add":
                if mode == "spare":
                    item.action_text = "Add Spare"
                elif mode == "space":
                    item.action_text = "Add Space"
                else:
                    item.action_text = "Add 50/50"
            elif kind == "remove":
                if mode == "both":
                    item.action_text = "Remove Both"
                elif mode == "space":
                    item.action_text = "Remove Space"
                else:
                    item.action_text = "Remove Spare"
            else:
                item.action_text = ""
        self._refresh_panel_list()
        self._refresh_status()

    def _panel_option_lookup(self):
        """Return panel option lookup map keyed by panel id."""
        lookup = {}
        for item in list(self._items or []):
            lookup[int(item.panel_id)] = item.option
        return lookup

    def _refresh_open_slot_counts(self):
        """Refresh open-slot counts after apply."""
        for item in list(self._items or []):
            item.open_slots = int(_count_open_slots_fast(item.option))
            item.open_slots_text = str(int(item.open_slots))

    def _refresh_status(self):
        """Refresh checked/staged counters and action button states."""
        checked = len([x for x in self._items if bool(getattr(x, "is_checked", False))])
        total = len(list(self._items or []))
        staged = len(list((self._staged_actions_by_panel or {}).keys()))
        if self.CheckedStatusText is not None:
            self.CheckedStatusText.Text = "{0} of {1} checked".format(int(checked), int(total))
        if self.StagedStatusText is not None:
            self.StagedStatusText.Text = "{0} panel(s) staged".format(int(staged))

        has_checked = checked > 0
        has_selected = len(self._selected_grid_items()) > 0
        if self.StageAddButton is not None:
            self.StageAddButton.IsEnabled = bool(has_checked)
        if self.StageRemoveButton is not None:
            self.StageRemoveButton.IsEnabled = bool(has_checked)
        if self.ResetSelectedButton is not None:
            self.ResetSelectedButton.IsEnabled = bool(has_selected)
        if self.ApplyButton is not None:
            self.ApplyButton.IsEnabled = staged > 0

    def window_loaded(self, sender, args):
        """Mark window ready after load so checkbox sync can run safely."""
        self._is_ready = True
        self._refresh_status()

    def _apply_checkbox_to_selected(self, sender, state):
        """Apply checkbox click to row or entire selected set when applicable."""
        if not self._is_ready or self._is_syncing_checks or self._suppress_check_events:
            return
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        targets = [row]
        selected = self._selected_grid_items()
        if len(selected) > 1 and row in selected:
            targets = selected

        self._is_syncing_checks = True
        try:
            for item in list(targets or []):
                item.is_checked = bool(state)
        finally:
            self._is_syncing_checks = False

        self._suppress_check_events = True
        try:
            self._refresh_panel_list()
            self._refresh_status()
        finally:
            self._suppress_check_events = False

    # ------------------------------------------------------------------
    # UI handlers
    # ------------------------------------------------------------------
    def item_checkbox_clicked(self, sender, args):
        """Toggle checkbox for row; if multi-selected, apply to all selected rows."""
        self._apply_checkbox_to_selected(sender, bool(getattr(sender, "IsChecked", False)))

    def grid_selection_changed(self, sender, args):
        """Track selected rows and sync reset button state."""
        selected = self._selected_grid_items()
        if selected:
            self._last_selected_rows = list(selected)
        self._refresh_status()

    def _clear_grid_selection(self):
        """Clear grid row selection."""
        if self.PanelsList is None:
            return
        try:
            self.PanelsList.UnselectAll()
        except Exception:
            pass

    def window_preview_mouse_down(self, sender, args):
        """Clear list selection when clicking outside the data grid."""
        source = getattr(args, "OriginalSource", None)
        if self.PanelsList is None or source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if not _is_descendant_of_control(source, self.PanelsList):
            self._clear_grid_selection()
            self._refresh_status()

    def grid_preview_mouse_down(self, sender, args):
        """Clear selection if click is not on a row."""
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, DataGridRow) is None:
            self._clear_grid_selection()
            self._refresh_status()

    def check_all_clicked(self, sender, args):
        """Check all panel rows."""
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for item in list(self._items or []):
                item.is_checked = True
        finally:
            self._is_syncing_checks = False
        self._refresh_panel_list()
        self._suppress_check_events = False
        self._refresh_status()
        self._set_status("Checked all panels.")

    def uncheck_all_clicked(self, sender, args):
        """Uncheck all panel rows."""
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for item in list(self._items or []):
                item.is_checked = False
        finally:
            self._is_syncing_checks = False
        self._refresh_panel_list()
        self._suppress_check_events = False
        self._refresh_status()
        self._set_status("Unchecked all panels.")

    def stage_add_clicked(self, sender, args):
        """Stage add action for checked panels."""
        selected = self._selected_items()
        if not selected:
            self._set_status("Select one or more panel rows.")
            return
        mode = self._selected_mode()
        add_mode = "mixed" if mode == "both" else mode
        count = self._stage_actions("add", add_mode, selected)
        self._set_status("Staged ADD action for {0} panel(s).".format(int(count)))

    def stage_remove_clicked(self, sender, args):
        """Stage remove action for checked panels."""
        selected = self._selected_items()
        if not selected:
            self._set_status("Select one or more panel rows.")
            return
        mode = self._selected_mode()
        remove_mode = "both" if mode == "both" else mode
        count = self._stage_actions("remove", remove_mode, selected)
        self._set_status("Staged REMOVE action for {0} panel(s).".format(int(count)))

    def reset_selected_clicked(self, sender, args):
        """Clear staged actions for currently selected rows."""
        selected = self._selected_grid_items()
        if not selected:
            self._set_status("Select one or more rows to reset.")
            return
        count = 0
        for item in list(selected or []):
            panel_id = int(getattr(item, "panel_id", 0) or 0)
            if panel_id > 0 and panel_id in self._staged_actions_by_panel:
                del self._staged_actions_by_panel[panel_id]
                count += 1
        self._rebuild_action_column()
        self._set_status("Reset staged action for {0} row(s).".format(int(count)))

    def apply_clicked(self, sender, args):
        """Apply staged actions in sequence order."""
        if not self._staged_actions_by_panel:
            self._set_status("No staged actions to apply.")
            return

        doc = self._active_doc()
        if doc is None:
            self._set_status("No active Revit document.")
            return

        option_lookup = self._panel_option_lookup()
        manager = PanelScheduleManager(doc, panel_option_lookup=option_lookup)
        staged = []
        for item in list(self._items or []):
            action = (self._staged_actions_by_panel or {}).get(int(item.panel_id))
            if action:
                staged.append(action)
        need_remove_scan = any(str(x.get("action_type", "")).lower() == "remove" for x in staged)
        all_circuits = _collect_all_circuits(doc) if need_remove_scan else []

        tx = DB.Transaction(doc, "Add/Remove Spares and Spaces")
        tx.Start()
        try:
            for action in staged:
                panel_id = int(action.get("panel_id", 0) or 0)
                option = option_lookup.get(int(panel_id))
                if option is None:
                    continue
                kind = str(action.get("action_type", "") or "")
                mode = str(action.get("mode", "") or "")
                if kind == "add":
                    _execute_add_for_option(manager, option, mode)
                elif kind == "remove":
                    _execute_remove_for_option(manager, doc, option, mode, all_circuits)
            status = tx.Commit()
            if status != DB.TransactionStatus.Committed:
                raise Exception("Transaction did not commit.")
        except Exception as ex:
            try:
                if tx.GetStatus() == DB.TransactionStatus.Started:
                    tx.RollBack()
            except Exception:
                pass
            forms.alert("Apply failed. Changes were rolled back.\n\n{0}".format(str(ex)), title=TITLE)
            self._set_status("Apply failed and rolled back.")
            return

        self._staged_actions_by_panel = {}
        self._refresh_open_slot_counts()
        self._rebuild_action_column()
        self._set_status("Apply completed.")

    def cancel_clicked(self, sender, args):
        """Close window without applying staged actions."""
        self.Close()


def _show_modal():
    """Show staged planner window."""
    window = AddRemoveSparesSpacesWindow(accent_mode=_load_accent_mode("blue"))
    try:
        window.ShowDialog()
    except Exception:
        window.Show()
    try:
        window.Activate()
    except Exception:
        pass


def _show_quick_modal(option):
    """Show quick active-panel action window and return selected action dict."""
    window = QuickPanelActionWindow(option=option, accent_mode=_load_accent_mode("blue"))
    try:
        window.ShowDialog()
    except Exception:
        window.Show()
    return getattr(window, "result", None)


if __name__ == "__main__":
    active_doc = revit.doc
    quick_option = _get_active_schedule_option(active_doc)
    if quick_option is not None:
        quick_action = _show_quick_modal(quick_option)
        if quick_action:
            _run_quick_action(
                active_doc,
                quick_option,
                quick_action.get("action_type"),
                quick_action.get("mode"),
            )
    else:
        _show_modal()
