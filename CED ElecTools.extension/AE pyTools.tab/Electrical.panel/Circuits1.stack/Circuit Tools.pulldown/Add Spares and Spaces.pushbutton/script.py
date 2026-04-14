# -*- coding: utf-8 -*-
"""Stage add/remove spare-space actions across panel schedules."""

import os

from System.Windows.Controls import Button, DataGridRow
from System.Windows.Media import VisualTreeHelper
from pyrevit import forms, revit, script

TITLE = "Add / Remove Spares and Spaces"

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
from UIClasses import pathing as ui_pathing

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate workspace root for Add/Remove Spares and Spaces.", title=TITLE)
    raise SystemExit

from UIClasses import resource_loader
from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Model.panel_schedule_enums import PanelUiActionType as UiActionType
from add_remove_execution import collect_panel_assignment_usage
from add_remove_execution import count_open_slots_fast
from add_remove_execution import execute_quick_action
from add_remove_execution import execute_staged_actions
from add_remove_execution import format_panel_info
from add_spares_spaces_view_models import PanelListItem
from add_spares_spaces_view_models import action_label

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


def _load_theme_state(default_theme="light", default_accent="blue"):
    from UIClasses import load_theme_state_from_config

    return load_theme_state_from_config(
        default_theme=default_theme,
        default_accent=default_accent,
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


def _get_selected_sheet_instance_options(doc):
    """Return panel options from selected PanelScheduleSheetInstance elements."""
    try:
        uidoc = revit.uidoc
        selected_ids = list(uidoc.Selection.GetElementIds() or [])
    except Exception:
        selected_ids = []
    if not selected_ids:
        return []
    selected_panels = []
    seen_panel_ids = set()
    for selected_id in list(selected_ids or []):
        try:
            element = doc.GetElement(selected_id)
        except Exception:
            element = None
        if not isinstance(element, ps_repo.DBE.PanelScheduleSheetInstance):
            continue
        try:
            schedule_view = doc.GetElement(element.ScheduleId)
        except Exception:
            schedule_view = None
        if not isinstance(schedule_view, ps_repo.DBE.PanelScheduleView):
            continue
        panel = ps_repo.resolve_schedule_panel(doc, schedule_view)
        if panel is None:
            continue
        panel_id = int(ps_repo._idval(getattr(panel, "Id", None)))
        if panel_id <= 0 or panel_id in seen_panel_ids:
            continue
        seen_panel_ids.add(panel_id)
        selected_panels.append(panel)
    if not selected_panels:
        return []
    options = list(
        ps_repo.collect_panel_equipment_options(
            doc,
            panels=selected_panels,
            include_without_schedule=True,
        )
        or []
    )
    options.sort(key=lambda x: (str(x.get("panel_name", "") or ""), str(x.get("dist_system_name", "") or "")))
    return options


def _get_quick_options(doc):
    """Return quick-action panel options from active schedule or selected sheet instances."""
    active_option = _get_active_schedule_option(doc)
    if active_option is not None:
        return [active_option]
    return list(_get_selected_sheet_instance_options(doc) or [])


def _run_quick_action(doc, options, action_type, mode):
    """Run quick action for one or more panel options and report summary."""
    try:
        result = execute_quick_action(doc, options, action_type, mode)
        action_kind = UiActionType.normalize(result.get("action_kind", ""), default=UiActionType.ADD)
        if action_kind == UiActionType.ADD:
            message = "Panels: {0}\nAdded Spare: {1}, Space: {2}".format(
                int(result.get("touched", 0) or 0),
                int(result.get("added_spares", 0) or 0),
                int(result.get("added_spaces", 0) or 0),
            )
        else:
            message = "Panels: {0}\nRemoved Spare: {1}, Space: {2}".format(
                int(result.get("touched", 0) or 0),
                int(result.get("removed_spares", 0) or 0),
                int(result.get("removed_spaces", 0) or 0),
            )
        finalize_summary = dict(result.get("finalize_summary") or {})
    except Exception as ex:
        forms.alert("Quick action failed and was rolled back.\n\n{0}".format(str(ex)), title=TITLE)
        return

    if action_kind == UiActionType.ADD and int(finalize_summary.get("unlock_failed", 0) or 0) > 0:
        message = "{0}\nUnlock warnings: {1}/{2} slot(s) remained locked.".format(
            str(message or ""),
            int(finalize_summary.get("unlock_failed", 0) or 0),
            int(finalize_summary.get("unlock_attempted", 0) or 0),
        )
    if action_kind == UiActionType.ADD and int(finalize_summary.get("pole_failed", 0) or 0) > 0:
        message = "{0}\nPole warnings: {1}/{2} switchboard rows could not be set to 3P.".format(
            str(message or ""),
            int(finalize_summary.get("pole_failed", 0) or 0),
            int(finalize_summary.get("pole_attempted", 0) or 0),
        )
    forms.alert(message, title=TITLE)


class QuickPanelActionWindow(forms.WPFWindow):
    """Super-lightweight active-panel action window."""

    def __init__(self, options, theme_mode, accent_mode):
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "AddSparesSpacesQuickWindow.xaml"))
        self._theme_mode = resource_loader.normalize_theme_mode(theme_mode, "light")
        self._accent_mode = resource_loader.normalize_accent_mode(accent_mode, "blue")
        self.options = [x for x in list(options or []) if x is not None]
        self.result = None
        forms.WPFWindow.__init__(self, xaml_path)
        self._apply_theme()
        self._init_controls()

    def _apply_theme(self):
        try:
            resource_loader.apply_theme(
                self,
                resources_root=UI_RESOURCES_ROOT,
                theme_mode=self._theme_mode,
                accent_mode=self._accent_mode,
            )
        except Exception as ex:
            LOGGER.warning("Quick panel action theme apply failed: %s", ex)

    def _init_controls(self):
        self.PanelInfoText = self.FindName("PanelInfoText")
        self.ModeFillRadio = self.FindName("ModeFillRadio")
        self.ModeRemoveRadio = self.FindName("ModeRemoveRadio")
        self.QuickGuidanceText = self.FindName("QuickGuidanceText")
        self.QuickDefaultDefinitionText = self.FindName("QuickDefaultDefinitionText")
        usage = collect_panel_assignment_usage(revit.doc)
        if self.PanelInfoText is not None:
            count = int(len(list(self.options or [])))
            if count <= 1:
                single = self.options[0] if count == 1 else None
                self.PanelInfoText.Text = format_panel_info(single, usage)
            else:
                self.PanelInfoText.Text = "({0}) Panels selected".format(int(count))
        self._update_quick_guidance()

    def _selected_action_type(self):
        return UiActionType.REMOVE if bool(getattr(self.ModeRemoveRadio, "IsChecked", False)) else UiActionType.ADD

    def _update_quick_guidance(self):
        """Update quick-mode guidance copy based on selected action mode."""
        action_type = self._selected_action_type()
        if self.QuickGuidanceText is not None:
            if action_type == UiActionType.REMOVE:
                self.QuickGuidanceText.Text = "All default spares/spaces will be removed from selected panels."
            else:
                self.QuickGuidanceText.Text = "All available slots on selected panels will be filled with default spares/spaces."
        if self.QuickDefaultDefinitionText is not None:
            self.QuickDefaultDefinitionText.Text = (
                "Default spare/space definition:\n"
                "- Unlocked\n"
                "- No schedule notes\n"
                "- Default load name\n"
                "- Default rating"
            )

    def quick_mode_changed(self, sender, args):
        """Handle Fill/Remove radio mode toggle for guidance text."""
        self._update_quick_guidance()

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

    def __init__(self, theme_mode, accent_mode):
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "AddSparesSpacesWindow.xaml"))
        self._theme_mode = resource_loader.normalize_theme_mode(theme_mode, "light")
        self._accent_mode = resource_loader.normalize_accent_mode(accent_mode, "blue")
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
        """Apply configured UI theme + accent."""

        resource_loader.apply_theme(
            self,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )

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

        usage_by_panel = collect_panel_assignment_usage(doc)
        self._items = []
        self._item_by_panel_id = {}
        for option in options:
            open_slots = count_open_slots_fast(option, usage_by_panel)
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
        """Return union of checked rows and grid-selected rows."""
        selected_by_id = {}
        for item in list(self._items or []):
            if bool(getattr(item, "is_checked", False)):
                selected_by_id[int(getattr(item, "panel_id", 0) or 0)] = item
        for item in list(self._selected_grid_items() or []):
            panel_id = int(getattr(item, "panel_id", 0) or 0)
            if panel_id > 0:
                selected_by_id[panel_id] = item
        return list(selected_by_id.values())

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
            item.action_text = action_label(action.get("action_type", ""), action.get("mode", ""))
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
        doc = self._active_doc()
        if doc is None:
            return
        usage_by_panel = collect_panel_assignment_usage(doc)
        for item in list(self._items or []):
            item.open_slots = int(count_open_slots_fast(item.option, usage_by_panel))
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
        has_targets = bool(has_checked or has_selected)
        if self.StageAddButton is not None:
            self.StageAddButton.IsEnabled = bool(has_targets)
        if self.StageRemoveButton is not None:
            self.StageRemoveButton.IsEnabled = bool(has_targets)
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
        staged = []
        for item in list(self._items or []):
            action = (self._staged_actions_by_panel or {}).get(int(item.panel_id))
            if action:
                staged.append(action)
        try:
            result = execute_staged_actions(doc, staged, option_lookup)
            finalize_summary = dict(result.get("finalize_summary") or {})
        except Exception as ex:
            forms.alert("Apply failed. Changes were rolled back.\n\n{0}".format(str(ex)), title=TITLE)
            self._set_status("Apply failed and rolled back.")
            return

        self._staged_actions_by_panel = {}
        self._refresh_open_slot_counts()
        self._rebuild_action_column()
        if int(finalize_summary.get("unlock_failed", 0) or 0) > 0:
            self._set_status(
                "Apply completed with unlock warnings ({0}/{1} failed).".format(
                    int(finalize_summary.get("unlock_failed", 0) or 0),
                    int(finalize_summary.get("unlock_attempted", 0) or 0),
                )
            )
        elif int(finalize_summary.get("pole_failed", 0) or 0) > 0:
            self._set_status(
                "Apply completed with pole warnings ({0}/{1} failed).".format(
                    int(finalize_summary.get("pole_failed", 0) or 0),
                    int(finalize_summary.get("pole_attempted", 0) or 0),
                )
            )
        else:
            self._set_status("Apply completed.")

    def cancel_clicked(self, sender, args):
        """Close window without applying staged actions."""
        self.Close()


def _show_modal():
    """Show staged planner window."""
    theme_mode, accent_mode = _load_theme_state("light", "blue")
    window = AddRemoveSparesSpacesWindow(theme_mode=theme_mode, accent_mode=accent_mode)
    try:
        window.ShowDialog()
    except Exception:
        window.Show()
    try:
        window.Activate()
    except Exception:
        pass


def _show_quick_modal(options):
    """Show quick action window and return selected action dict."""
    theme_mode, accent_mode = _load_theme_state("light", "blue")
    window = QuickPanelActionWindow(options=options, theme_mode=theme_mode, accent_mode=accent_mode)
    try:
        window.ShowDialog()
    except Exception:
        window.Show()
    return getattr(window, "result", None)


if __name__ == "__main__":
    active_doc = revit.doc
    quick_options = _get_quick_options(active_doc)
    if quick_options:
        quick_action = _show_quick_modal(quick_options)
        if quick_action:
            _run_quick_action(
                active_doc,
                quick_options,
                quick_action.get("action_type"),
                quick_action.get("mode"),
            )
    else:
        _show_modal()
