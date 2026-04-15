# -*- coding: utf-8 -*-
"""Batch Swap Circuits - staged planner (no apply transaction yet)."""

import copy
import os

import clr

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from System import Math
from System.Windows import DataObject, DragDrop, DragDropEffects
from System.Windows import FontWeights
from System.Windows import Visibility
from System.Windows.Controls import ListViewItem
from System.Windows.Documents import Run
from System.Windows.Input import MouseButtonState
from System.Windows.Media import Color, SolidColorBrush, VisualTreeHelper
from pyrevit import DB, forms, revit, script

TITLE = "Batch Swap Circuits"
_WINDOW_MARKER = "_ae_batch_swap_window"

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
from UIClasses import pathing as ui_pathing

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate workspace root for Batch Swap Circuits.", title=TITLE)
    raise SystemExit

from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_registry import OperationRegistry
from CEDElectrical.Application.services.operation_runner import OperationRunner
from CEDElectrical.Application.operations.panel_schedule_actions import (
    PanelScheduleAddSpareOperation,
    PanelScheduleAddSpaceOperation,
    PanelScheduleMoveCircuitInPanelOperation,
    PanelScheduleMoveCircuitToPanelOperation,
    PanelScheduleMoveCircuitToSpecificSlotOperation,
    PanelScheduleRemoveSpareOperation,
    PanelScheduleRemoveSpaceOperation,
)
from CEDElectrical.Model.panel_schedule_enums import PanelScheduleOperationKey as OpKey
from CEDElectrical.Model.panel_schedule_enums import PanelSpecialKind as SpecialKind
from CEDElectrical.Model.panel_schedule_enums import PanelStagedAction as StagedAction
from Snippets import revit_helpers
from UIClasses import resource_loader

UI_RESOURCES_ROOT = ui_pathing.resolve_ui_resources_root(LIB_ROOT)
LOGGER = script.get_logger()


def _build_panel_schedule_operation_runner():
    """Build operation runner for panel schedule actions."""
    registry = OperationRegistry()
    registry.register(PanelScheduleAddSpareOperation())
    registry.register(PanelScheduleAddSpaceOperation())
    registry.register(PanelScheduleRemoveSpareOperation())
    registry.register(PanelScheduleRemoveSpaceOperation())
    registry.register(PanelScheduleMoveCircuitToPanelOperation())
    registry.register(PanelScheduleMoveCircuitInPanelOperation())
    registry.register(PanelScheduleMoveCircuitToSpecificSlotOperation())
    return OperationRunner(registry)


def _load_theme_state(default_theme="light", default_accent="blue"):
    from UIClasses import load_theme_state_from_config

    return load_theme_state_from_config(
        default_theme=default_theme,
        default_accent=default_accent,
    )


def operation_key_for_action(action):
    """Resolve operation key for a staged placement action."""
    if StagedAction.is_add_spare(action):
        return OpKey.ADD_SPARE
    if StagedAction.is_add_space(action):
        return OpKey.ADD_SPACE
    if StagedAction.is_remove_spare(action):
        return OpKey.REMOVE_SPARE
    if StagedAction.is_remove_space(action):
        return OpKey.REMOVE_SPACE
    return OpKey.MOVE_TO_SPECIFIC_SLOT


def is_add_action(action):
    """Return True when staged action is add_spare/add_space."""
    return bool(StagedAction.is_add_spare(action) or StagedAction.is_add_space(action))


class PanelOptionItem(object):
    """Lightweight combo item wrapper for panel options."""

    def __init__(self, option):
        self.option = option
        self.panel_id = int(option.get("panel_id", 0) or 0)
        self.panel_name = str(option.get("panel_name", "Unnamed Panel") or "Unnamed Panel")
        self.part_type = str(option.get("part_type_name", "") or option.get("board_type", "Panelboard") or "Panelboard")
        self.dist_system_name = str(option.get("dist_system_name", "Unknown Dist. System") or "Unknown Dist. System")
        self.left_text = "{0} ({1})".format(self.panel_name, self.part_type)
        self.right_text = ""
        self._combo_text = self.left_text

    def __str__(self):
        return str(self._combo_text or self.left_text)

    def ToString(self):
        """Return text shown by WPF combo boxes."""
        return str(self._combo_text or self.left_text)

    def set_combo_mode(self, show_full):
        """Set current combo text to compact or full descriptor."""
        if bool(show_full):
            self.right_text = self.dist_system_name
            self._combo_text = "{0} | {1}".format(self.left_text, self.right_text)
        else:
            self.right_text = ""
            self._combo_text = self.left_text


class TemplateOptionItem(object):
    """Lightweight list item wrapper for schedule template choices."""

    def __init__(self, option):
        self.option = option
        self.template_id = int(option.get("template_id", 0) or 0)
        self.template_name = str(option.get("template_name", "Unnamed Template") or "Unnamed Template")
        self.display_name = self.template_name

    def __str__(self):
        return self.display_name

    def ToString(self):
        """Return text shown by WPF combo boxes."""
        return self.display_name


class SlotCellItem(object):
    """View model for one 1-pole slot cell in the left gutter."""

    def __init__(self, label, background, height=20):
        self.label = str(label or "")
        self.background = background
        self.height = int(max(12, height or 20))


class ChangeLogItem(object):
    """View model item for staged action history rows."""

    def __init__(self, text, background=None):
        self.text = str(text or "")
        self.background = background


class SwapRowItem(object):
    """View model for one staged row in a panel list."""

    def __init__(
        self,
        row,
        index,
        brushes,
        is_preview=False,
        show_divider=False,
        preview_slots=None,
        is_moving=False,
        allow_default_replace=False,
        display_column_lookup=None,
    ):
        kind = str(row.get("kind", "empty"))
        span = int(max(1, row.get("span", 1) or 1))
        is_regular = bool(row.get("is_regular_circuit", False))
        is_transferable = bool(row.get("transferable", False))
        is_editable = bool(row.get("is_editable", True))
        is_excess_slot = bool(row.get("is_excess_slot", False))
        is_staged = bool(row.get("is_staged", False))
        reason = str(row.get("transfer_reason", "") or "")
        covered_slots = [int(x) for x in list(row.get("covered_slots") or []) if int(x) > 0]
        if not covered_slots:
            slot_value = int(row.get("slot", 0) or 0)
            if slot_value > 0:
                covered_slots = [slot_value]

        self.row_key = row.get("row_key")
        self.slot = int(row.get("slot", 0) or 0)
        self.kind = kind
        self.span = span
        self.is_regular_circuit = is_regular
        self.is_transferable = is_transferable
        self.is_excess_slot = bool(is_excess_slot)
        self.is_draggable = bool(kind in ("circuit", "spare", "space") and is_editable)
        self.row_height = 20 * span
        self.row_opacity = 1.0 if self.is_draggable else 0.46
        is_default_spare = bool(kind == "spare" and bool(row.get("is_spare_removable", False)))
        is_default_space = bool(kind == "space" and bool(row.get("is_space_removable", False)))
        show_default_replace = bool(allow_default_replace and (is_default_spare or is_default_space))

        if kind == "empty":
            if bool(is_excess_slot):
                number_text = "N/A"
                load_text = "Exceeds equipment capacity"
                meta_text = "-"
                base_bg = brushes.get("slot_bg") or brushes.get("empty_bg")
                tooltip = "Unavailable slot: panel schedule template has more rows than device supports."
            else:
                number_text = "EMPTY"
                load_text = "Available slot"
                meta_text = "-"
                base_bg = brushes.get("empty_bg")
                tooltip = "Empty slot"
        elif kind == "spare":
            number_text = row.get("circuit_number") or "SPARE"
            load_text = "SPARE"
            meta_text = row.get("meta_text") or "-"
            base_bg = brushes.get("empty_bg") if show_default_replace else brushes.get("spare_bg")
            if is_default_spare:
                tooltip = "Default Spare. May be replaced by circuit."
            else:
                tooltip = "Spare row."
        elif kind == "space":
            number_text = row.get("circuit_number") or "SPACE"
            load_text = "SPACE"
            meta_text = row.get("meta_text") or "-"
            base_bg = brushes.get("empty_bg") if show_default_replace else brushes.get("space_bg")
            if is_default_space:
                tooltip = "Default Space. May be replaced by circuit."
            else:
                tooltip = "Space row."
        else:
            number_text = row.get("circuit_number") or "-"
            load_text = row.get("load_name") or "-"
            meta_text = row.get("meta_text") or "-"
            base_bg = brushes.get("circuit_bg")
            if not is_editable:
                tooltip = reason or "Owned by another user."
            elif is_transferable:
                tooltip = "Drag to reorder in this panel or move to the opposite panel."
            else:
                tooltip = "Reorder in this panel only. " + (reason or "Not compatible with opposite panel.")

        if is_preview and not bool(is_moving):
            row_bg = brushes.get("preview_bg") or base_bg
        elif is_staged and kind != "empty":
            row_bg = brushes.get("staged_bg") or base_bg
        else:
            row_bg = base_bg

        slot_values = list(covered_slots) if covered_slots else [0]
        preview_slot_set = set([int(x) for x in list(preview_slots or []) if int(x) > 0])
        slot_cells = []
        for slot_value in slot_values:
            label = str(slot_value) if int(slot_value or 0) > 0 else "-"
            display_col = 1
            try:
                if isinstance(display_column_lookup, dict):
                    display_col = int(display_column_lookup.get(int(slot_value or 0), 1) or 1)
            except Exception:
                display_col = 1
            is_secondary_column = bool(display_col == 2)
            cell_is_preview = bool(int(slot_value or 0) > 0 and int(slot_value) in preview_slot_set)
            if not preview_slot_set and bool(is_preview):
                cell_is_preview = True
            if cell_is_preview:
                slot_bg = brushes.get("slot_preview_col2_bg" if is_secondary_column else "slot_preview_col1_bg")
            else:
                slot_bg = brushes.get("slot_col2_bg" if is_secondary_column else "slot_col1_bg")
            slot_bg = slot_bg or brushes.get("slot_bg")
            slot_cells.append(SlotCellItem(label, slot_bg, height=20))
        self.slot_cells = slot_cells
        self.covered_slots = [int(x) for x in list(slot_values) if int(x) > 0]
        self.circuit_number_text = number_text
        self.load_text = load_text
        self.meta_text = meta_text
        self.indicator_text = str(row.get("indicator_text", "") or "")
        self.row_background = row_bg
        self.row_tooltip = tooltip
        self.slot_background = brushes.get("slot_bg")
        self.divider_visibility = Visibility.Visible if bool(show_divider) else Visibility.Collapsed
        self._index = index


class BatchSwapWindow(forms.WPFWindow):
    """Window controller for staged batch swap planning."""

    def __init__(self, theme_mode, accent_mode):
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "BatchSwapCircuitsWindow.xaml"))
        self._theme_mode = resource_loader.normalize_theme_mode(theme_mode, "light")
        self._accent_mode = resource_loader.normalize_accent_mode(accent_mode, "blue")
        self._panel_options = []
        self._panel_option_by_id = {}
        self._template_items_by_panel = {}
        self._working_rows_by_panel = {}
        self._all_panels_cache = []
        self._all_circuits_cache = []
        self._known_panel_ids = set()
        self._left_option = None
        self._right_option = None
        self._left_rows = []
        self._right_rows = []
        self._drag_start = None
        self._drag_payload = None
        self._suspend_panel_events = False
        self._brushes = {}
        self._left_preview_slots = set()
        self._right_preview_slots = set()
        self._preview_signature = None
        self._preview_moving_keys = set()
        self._slot_cells_cache = {}
        self._layout_context_cache = {}
        self._operation_history = []
        self._operation_seq = 0
        self._panel_schedule_runner = _build_panel_schedule_operation_runner()
        self._pending_click_selection = None
        self._drag_started = False
        self._temp_row_seed = -1
        self._current_username = ""
        try:
            self._current_username = str(getattr(revit.doc.Application, "Username", "") or "").strip().lower()
        except Exception:
            self._current_username = ""

        forms.WPFWindow.__init__(self, xaml_path)
        setattr(self, _WINDOW_MARKER, True)
        self._apply_theme()
        self._init_controls()
        self._load_panel_options()

    def _active_doc(self):
        """Return current active document."""
        return revit.doc

    def _apply_theme(self):
        """Apply configured theme and accent."""
        resource_loader.apply_theme(
            self,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )

    def _make_brush(self, argb_hex):
        """Create a fallback SolidColorBrush from ARGB hex."""
        try:
            text = str(argb_hex or "").strip().lstrip("#")
            if len(text) == 6:
                text = "FF" + text
            if len(text) != 8:
                return None
            a = int(text[0:2], 16)
            r = int(text[2:4], 16)
            g = int(text[4:6], 16)
            b = int(text[6:8], 16)
            return SolidColorBrush(Color.FromArgb(a, r, g, b))
        except Exception:
            return None

    def _init_controls(self):
        """Resolve controls and init local brush overrides."""
        self.LeftPanelCombo = self.FindName("LeftPanelCombo")
        self.RightPanelCombo = self.FindName("RightPanelCombo")
        self.LeftPanelHeaderText = self.FindName("LeftPanelHeaderText")
        self.RightPanelHeaderText = self.FindName("RightPanelHeaderText")
        self.LeftPanelMeta = self.FindName("LeftPanelMeta")
        self.RightPanelMeta = self.FindName("RightPanelMeta")
        self.LeftPanelTypeMeta = self.FindName("LeftPanelTypeMeta")
        self.RightPanelTypeMeta = self.FindName("RightPanelTypeMeta")
        self.LeftPanelCurrentConnectedText = self.FindName("LeftPanelCurrentConnectedText")
        self.LeftPanelCurrentDemandText = self.FindName("LeftPanelCurrentDemandText")
        self.RightPanelCurrentConnectedText = self.FindName("RightPanelCurrentConnectedText")
        self.RightPanelCurrentDemandText = self.FindName("RightPanelCurrentDemandText")
        self.LeftMissingSchedulePanel = self.FindName("LeftMissingSchedulePanel")
        self.RightMissingSchedulePanel = self.FindName("RightMissingSchedulePanel")
        self.LeftMissingScheduleText = self.FindName("LeftMissingScheduleText")
        self.RightMissingScheduleText = self.FindName("RightMissingScheduleText")
        self.LeftTemplateList = self.FindName("LeftTemplateList")
        self.RightTemplateList = self.FindName("RightTemplateList")
        self.LeftCreateScheduleButton = self.FindName("LeftCreateScheduleButton")
        self.RightCreateScheduleButton = self.FindName("RightCreateScheduleButton")
        self.LeftTemplateRequirementText = self.FindName("LeftTemplateRequirementText")
        self.RightTemplateRequirementText = self.FindName("RightTemplateRequirementText")
        self.LeftRowsList = self.FindName("LeftRowsList")
        self.RightRowsList = self.FindName("RightRowsList")
        self.AllowDiscardToggle = self.FindName("AllowDiscardToggle")
        self.MaintainGroupToggle = self.FindName("MaintainGroupToggle")
        self.SwapSummaryText = self.FindName("SwapSummaryText")
        self.ChangeLogList = self.FindName("ChangeLogList")
        self.UndoLastButton = self.FindName("UndoLastButton")
        self.UndoAllButton = self.FindName("UndoAllButton")
        self.ClearListButton = self.FindName("ClearListButton")
        self.ApplyButton = self.FindName("ApplyButton")
        self.AddSpareButton = self.FindName("AddSpareButton")
        self.AddSpaceButton = self.FindName("AddSpaceButton")
        self.RemoveSpecialButton = self.FindName("RemoveSpecialButton")
        self.AddPole1Radio = self.FindName("AddPole1Radio")
        self.AddPole2Radio = self.FindName("AddPole2Radio")
        self.AddPole3Radio = self.FindName("AddPole3Radio")
        self.SpareRatingTextBox = self.FindName("SpareRatingTextBox")
        self.SpareFrameTextBox = self.FindName("SpareFrameTextBox")
        if self.SpareRatingTextBox is not None:
            try:
                if not str(self.SpareRatingTextBox.Text or "").strip():
                    self.SpareRatingTextBox.Text = "20 A"
            except Exception:
                pass
        if self.SpareFrameTextBox is not None:
            try:
                if not str(self.SpareFrameTextBox.Text or "").strip():
                    self.SpareFrameTextBox.Text = "20 A"
            except Exception:
                pass
        self._normalize_amp_textbox(self.SpareRatingTextBox, 20)
        self._normalize_amp_textbox(self.SpareFrameTextBox, 20)

        base_item_bg = self._resource("CED.Brush.ListItemBackground")
        list_bg = self._resource("CED.Brush.ListBackground")
        readonly_bg = self._resource("CED.Brush.DataGridReadOnlyBackground")
        selected_bg = self._resource("CED.Brush.CircuitItemSelectedBackground")
        hover_bg = self._resource("CED.Brush.CircuitItemHoverBackground")
        changed_bg = self._resource("CED.Brush.DataGridChangedBackground")
        info_accent_bg = self._resource("CED.Brush.InfoAccentBackground")
        warning_bg = self._resource("CED.Brush.DataGridWarningBackground")
        error_bg = self._resource("CED.Brush.DataGridErrorBackground")
        status_staged_bg = self._resource("CED.Brush.StatusStagedBackground")
        status_warning_bg = self._resource("CED.Brush.StatusWarningBackground")
        status_success_bg = self._resource("CED.Brush.StatusSuccessBackground")
        status_error_bg = self._resource("CED.Brush.StatusErrorBackground")

        self._brushes["empty_bg"] = base_item_bg or list_bg
        self._brushes["spare_bg"] = base_item_bg or readonly_bg or list_bg
        self._brushes["space_bg"] = base_item_bg or readonly_bg or list_bg
        self._brushes["circuit_bg"] = base_item_bg or list_bg
        self._brushes["slot_bg"] = list_bg or base_item_bg
        self._brushes["slot_col1_bg"] = info_accent_bg or base_item_bg or self._brushes["slot_bg"]
        self._brushes["slot_col2_bg"] = warning_bg or readonly_bg or self._brushes["slot_bg"]
        self._brushes["staged_bg"] = status_staged_bg or changed_bg or selected_bg or base_item_bg
        self._brushes["preview_bg"] = selected_bg or hover_bg or base_item_bg
        self._brushes["slot_preview_bg"] = hover_bg or self._brushes["preview_bg"]
        self._brushes["slot_preview_col1_bg"] = selected_bg or self._brushes["slot_preview_bg"]
        self._brushes["slot_preview_col2_bg"] = info_accent_bg or self._brushes["slot_preview_bg"]
        self._brushes["log_pending_bg"] = base_item_bg or list_bg
        self._brushes["log_warning_bg"] = status_warning_bg or warning_bg or changed_bg or self._brushes["staged_bg"]
        self._brushes["log_success_bg"] = status_success_bg or self._resource("CED.Brush.ApplyBackground") or self._make_brush("FF2E7D32") or self._brushes["preview_bg"]
        self._brushes["log_failed_bg"] = status_error_bg or error_bg or warning_bg or self._brushes["staged_bg"]

    def _set_panel_combo_mode(self, show_full):
        """Toggle panel combo label text between compact and full descriptor."""
        for item in list(self._panel_options or []):
            try:
                item.set_combo_mode(bool(show_full))
            except Exception:
                continue
        for combo in (self.LeftPanelCombo, self.RightPanelCombo):
            if combo is None:
                continue
            try:
                combo.Items.Refresh()
            except Exception:
                pass

    def _resource(self, key):
        """Return WPF resource by key, or None."""
        try:
            value = self.TryFindResource(key)
            if value is not None:
                return value
        except Exception:
            pass
        try:
            return self.FindResource(key)
        except Exception:
            return None

    def _clone_rows(self, rows):
        """Return mutable clones of row dictionaries."""
        clones = []
        for row in list(rows or []):
            data = dict(row or {})
            data["covered_slots"] = [int(x) for x in list(data.get("covered_slots") or []) if int(x) > 0]
            data["slot_cells"] = list(data.get("slot_cells") or [])
            clones.append(data)
        return clones

    def _snapshot_panels(self, panel_ids):
        """Capture row snapshots for operation undo."""
        snapshots = {}
        for panel_id in list(panel_ids or []):
            pid = int(panel_id or 0)
            if pid <= 0:
                continue
            rows = self._working_rows_by_panel.get(pid)
            if rows is None:
                continue
            snapshots[pid] = self._clone_rows(rows)
        return snapshots

    def _slot_display_column_lookup(self, option):
        """Return map of slot->display column for one panel option."""
        lookup = {}
        if not option:
            return lookup
        try:
            max_slot = int(option.get("max_slot", 0) or 0)
        except Exception:
            max_slot = 0
        if max_slot <= 0:
            return lookup
        sort_mode = option.get("sort_mode", "panelboard")
        for slot in ps_repo.get_option_slot_order(option, include_excess=True):
            try:
                slot_value = int(slot)
                if slot_value <= 0:
                    continue
                lookup[slot_value] = int(ps_repo.get_slot_display_column(slot_value, max_slot, sort_mode))
            except Exception:
                continue
        return lookup

    def _load_panel_options(self):
        """Load panel options and initialize selection."""
        doc = self._active_doc()
        if doc is None:
            forms.alert("No active Revit document for Batch Swap.", title=TITLE, exitscript=True)
        self._refresh_model_caches(doc)

        options = ps_repo.collect_panel_equipment_options(
            doc,
            panels=self._all_panels_cache,
            include_without_schedule=True,
        )
        if len(options) < 2:
            forms.alert(
                "Need at least two panelboards/switchboards in the model for Batch Swap.",
                title=TITLE,
                exitscript=True,
            )
        options.sort(key=lambda x: (str(x.get("panel_name", "")), str(x.get("dist_system_name", ""))))

        self._panel_options = [PanelOptionItem(x) for x in options]
        self._panel_option_by_id = {item.panel_id: item for item in self._panel_options}

        self._suspend_panel_events = True
        try:
            self.LeftPanelCombo.ItemsSource = self._panel_options
            self.RightPanelCombo.ItemsSource = self._panel_options
            self.LeftPanelCombo.SelectedIndex = 0
            self.RightPanelCombo.SelectedIndex = 1 if len(self._panel_options) > 1 else 0
            self._set_panel_combo_mode(False)
        finally:
            self._suspend_panel_events = False

        self._reload_from_selected_panels()

    def _refresh_model_caches(self, doc=None):
        """Refresh panel/circuit caches used for row construction."""
        doc = doc or self._active_doc()
        self._template_items_by_panel = {}
        self._slot_cells_cache = {}
        self._layout_context_cache = {}
        if doc is None:
            self._all_panels_cache = []
            self._all_circuits_cache = []
            self._known_panel_ids = set()
            return
        self._all_panels_cache = list(ps_repo.get_all_panels(doc))
        self._known_panel_ids = set()
        for panel in list(self._all_panels_cache or []):
            try:
                self._known_panel_ids.add(int(revit_helpers.get_elementid_value(panel.Id)))
            except Exception:
                continue
        try:
            self._all_circuits_cache = list(
                DB.FilteredElementCollector(doc)
                .OfClass(ps_repo.DBE.ElectricalSystem)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            self._all_circuits_cache = []

    def _format_phase_text(self, phase_value):
        """Return compact phase text from distribution system phase enum."""
        if phase_value is None:
            return "-"
        try:
            if phase_value == ps_repo.DBE.ElectricalPhase.SinglePhase:
                return "1"
        except Exception:
            pass
        try:
            if phase_value == ps_repo.DBE.ElectricalPhase.ThreePhase:
                return "3"
        except Exception:
            pass
        try:
            text = str(phase_value).strip()
        except Exception:
            text = ""
        if not text:
            return "-"
        return text

    def _format_schedule_type_label(self, option):
        """Return user-facing panel schedule type text for meta line."""
        schedule_type = None
        try:
            schedule_type = (option or {}).get("schedule_type")
        except Exception:
            schedule_type = None
        try:
            if schedule_type == ps_repo.PSTYPE_DATA:
                return "Data Panel"
            if schedule_type == ps_repo.PSTYPE_BRANCH:
                return "Branch Panel"
            if schedule_type == ps_repo.PSTYPE_SWITCHBOARD:
                return "Switchboard"
        except Exception:
            pass
        return "Unknown"

    def _format_profile(self, option):
        """Format panel profile summary text for left-side metadata."""
        profile = (option or {}).get("profile") or {}
        lg = profile.get("lg_voltage")
        ll = profile.get("ll_voltage")
        phase = self._format_phase_text(profile.get("phase"))
        wires = profile.get("wire_count")
        wire_text = "-" if wires is None else str(int(wires))
        lg_text = "-" if lg is None else "{0:.0f}V".format(float(lg))
        ll_text = "-" if ll is None else "{0:.0f}V".format(float(ll))
        return "Ph: {0}  Wire: {1}  L-L: {2}  L-G: {3}".format(
            phase,
            wire_text,
            ll_text,
            lg_text,
        )

    def _format_amp_value(self, value):
        """Return compact amperes text from numeric-like value."""
        if value is None:
            return "-"
        try:
            numeric = float(value)
        except Exception:
            return "-"
        if abs(numeric - round(numeric)) < 0.05:
            return "{0:.0f}".format(numeric)
        return "{0:.1f}".format(numeric)

    def _set_rich_text(self, textblock, parts):
        """Apply mixed-color inline runs: descriptors secondary, values black."""
        if textblock is None:
            return
        try:
            textblock.Inlines.Clear()
        except Exception:
            try:
                textblock.Text = "".join([str(x[0] or "") for x in list(parts or [])])
            except Exception:
                pass
            return
        secondary = self._resource("CED.Brush.TextSecondary")
        for text, is_value in list(parts or []):
            run = Run(str(text or ""))
            try:
                if bool(is_value):
                    run.FontWeight = FontWeights.SemiBold
                    if secondary is not None:
                        run.Foreground = secondary
                elif secondary is not None:
                    run.Foreground = secondary
            except Exception:
                pass
            textblock.Inlines.Add(run)

    def _set_profile_rich_text(self, option, textblock):
        """Render distribution profile with descriptor/value colors."""
        profile = (option or {}).get("profile") or {}
        lg = profile.get("lg_voltage")
        ll = profile.get("ll_voltage")
        phase = self._format_phase_text(profile.get("phase"))
        wires = profile.get("wire_count")
        wire_text = "-" if wires is None else str(int(wires))
        lg_text = "-" if lg is None else "{0:.0f}V".format(float(lg))
        ll_text = "-" if ll is None else "{0:.0f}V".format(float(ll))
        self._set_rich_text(
            textblock,
            [
                ("Ph: ", False), (phase, True),
                ("  Wire: ", False), (wire_text, True),
                ("  L-L: ", False), (ll_text, True),
                ("  L-G: ", False), (lg_text, True),
            ],
        )

    def _set_currents_rich_text(self, option, connected_textblock, demand_textblock):
        """Render two-row current summary with aligned labels and black values."""
        model = (option or {}).get("equipment_model")
        conn = self._format_amp_value(getattr(model, "current_connected_total", None) if model is not None else None)
        dmd = self._format_amp_value(getattr(model, "current_demand_total", None) if model is not None else None)
        ia = self._format_amp_value(getattr(model, "branch_current_phase_a", None) if model is not None else None)
        ib = self._format_amp_value(getattr(model, "branch_current_phase_b", None) if model is not None else None)
        ic = self._format_amp_value(getattr(model, "branch_current_phase_c", None) if model is not None else None)

        def _amp_text(value_text):
            return "{0} A".format(value_text) if str(value_text or "-") != "-" else "-"

        self._set_rich_text(
            connected_textblock,
            [
                (_amp_text(conn), True),
                (" (Ia: ", False), (_amp_text(ia), True),
                (", Ib: ", False), (_amp_text(ib), True),
                (", Ic: ", False), (_amp_text(ic), True),
                (")", False),
            ],
        )
        self._set_rich_text(
            demand_textblock,
            [
                (_amp_text(dmd), True),
            ],
        )

    def _selected_option(self, combo):
        """Resolve selected panel option from combo."""
        item = getattr(combo, "SelectedItem", None)
        if item is None:
            return None
        return getattr(item, "option", None)

    def _selected_template_option(self, selector):
        """Resolve selected panel-schedule template option from selector control."""
        item = getattr(selector, "SelectedItem", None)
        if item is None:
            return None
        return getattr(item, "option", None)

    def _side_controls(self, side_name):
        """Return UI controls and state references for a side name."""
        if str(side_name).lower() == "left":
            return {
                "list_ctrl": self.LeftRowsList,
                "panel_combo": self.LeftPanelCombo,
                "panel_header": self.LeftPanelHeaderText,
                "missing_panel": self.LeftMissingSchedulePanel,
                "missing_text": self.LeftMissingScheduleText,
                "template_list": self.LeftTemplateList,
                "create_button": self.LeftCreateScheduleButton,
                "requirement_text": self.LeftTemplateRequirementText,
                "option": self._left_option,
            }
        return {
            "list_ctrl": self.RightRowsList,
            "panel_combo": self.RightPanelCombo,
            "panel_header": self.RightPanelHeaderText,
            "missing_panel": self.RightMissingSchedulePanel,
            "missing_text": self.RightMissingScheduleText,
            "template_list": self.RightTemplateList,
            "create_button": self.RightCreateScheduleButton,
            "requirement_text": self.RightTemplateRequirementText,
            "option": self._right_option,
        }

    def _has_schedule(self, option):
        """Return True when an option has a mapped panel schedule view."""
        if not option:
            return False
        if option.get("schedule_view") is None:
            return False
        return int(option.get("schedule_id", 0) or 0) > 0

    def _ensure_template_items(self, option):
        """Return cached compatible template items for a panel option."""
        panel_id = int((option or {}).get("panel_id", 0) or 0)
        if panel_id <= 0:
            return []
        cached = self._template_items_by_panel.get(panel_id)
        if cached is not None:
            return cached
        doc = self._active_doc()
        panel = (option or {}).get("panel")
        if doc is None or panel is None:
            self._template_items_by_panel[panel_id] = []
            return []
        try:
            template_options = ps_repo.get_compatible_panel_schedule_templates(
                doc,
                panel,
                probe_assignability=True,
            )
        except Exception as ex:
            LOGGER.warning("Batch Swap template probe failed for panel %s: %s", panel_id, ex)
            try:
                template_options = ps_repo.get_compatible_panel_schedule_templates(
                    doc,
                    panel,
                    probe_assignability=False,
                )
            except Exception:
                template_options = []
        if not template_options:
            try:
                template_options = ps_repo.get_compatible_panel_schedule_templates(
                    doc,
                    panel,
                    probe_assignability=False,
                )
            except Exception:
                template_options = []
        expected_config = (option or {}).get("panel_configuration")
        if expected_config is not None:
            filtered = []
            for item in list(template_options or []):
                try:
                    if item.get("panel_configuration") == expected_config:
                        filtered.append(item)
                except Exception:
                    continue
            if filtered:
                template_options = filtered
        items = [TemplateOptionItem(x) for x in list(template_options or [])]
        self._template_items_by_panel[panel_id] = items
        return items

    def _refresh_side_schedule_state(self, side_name, option):
        """Toggle list/template controls for one side based on schedule availability."""
        controls = self._side_controls(side_name)
        list_ctrl = controls.get("list_ctrl")
        panel_combo = controls.get("panel_combo")
        panel_header = controls.get("panel_header")
        missing_panel = controls.get("missing_panel")
        missing_text = controls.get("missing_text")
        template_list = controls.get("template_list")
        create_button = controls.get("create_button")
        requirement_text = controls.get("requirement_text")

        has_schedule = self._has_schedule(option)
        if list_ctrl is not None:
            list_ctrl.IsEnabled = bool(has_schedule)
            list_ctrl.Visibility = Visibility.Visible if bool(has_schedule) else Visibility.Collapsed

        if has_schedule:
            if panel_combo is not None:
                panel_combo.Visibility = Visibility.Visible
            if panel_header is not None:
                panel_header.Visibility = Visibility.Collapsed
            if missing_panel is not None:
                missing_panel.Visibility = Visibility.Collapsed
            if template_list is not None:
                template_list.Visibility = Visibility.Collapsed
                template_list.ItemsSource = []
                template_list.SelectedIndex = -1
            if create_button is not None:
                create_button.IsEnabled = False
            if requirement_text is not None:
                requirement_text.Text = ""
                requirement_text.Visibility = Visibility.Collapsed
            return

        if panel_combo is not None:
            panel_combo.Visibility = Visibility.Visible
        if panel_header is not None:
            panel_header.Visibility = Visibility.Collapsed

        if missing_panel is not None:
            missing_panel.Visibility = Visibility.Visible

        panel_name = str((option or {}).get("panel_name", "selected equipment") or "selected equipment")
        if missing_text is not None:
            missing_text.Text = "No panel schedule view found for {0}. Select a compatible template and create one.".format(panel_name)
        if requirement_text is not None:
            expected_type = str((option or {}).get("schedule_type_name", "Unknown") or "Unknown").strip()
            config_name = str((option or {}).get("panel_configuration_name", "Unknown") or "Unknown").strip()
            requirement_text.Text = "{0} | {1}".format(expected_type, config_name)
            requirement_text.Visibility = Visibility.Visible

        selected_template = self._selected_template_option(template_list) if template_list is not None else None
        selected_template_id = int((selected_template or {}).get("template_id", 0) or 0)
        items = self._ensure_template_items(option)
        if template_list is not None:
            template_list.Visibility = Visibility.Visible
            template_list.ItemsSource = items
            target_index = -1
            if selected_template_id > 0:
                for idx, item in enumerate(items):
                    if int(getattr(item, "template_id", 0) or 0) == selected_template_id:
                        target_index = idx
                        break
            if target_index < 0 and items:
                target_index = 0
            template_list.SelectedIndex = target_index
        if create_button is not None:
            create_button.IsEnabled = False
        self._update_create_schedule_button(side_name)

    def _ensure_distinct_panels(self):
        """Force left/right panel selection to different panels."""
        left = self._selected_option(self.LeftPanelCombo)
        right = self._selected_option(self.RightPanelCombo)
        if left is None or right is None:
            return
        if int(left.get("panel_id", 0)) != int(right.get("panel_id", 0)):
            return

        right_index = int(getattr(self.RightPanelCombo, "SelectedIndex", 0) or 0)
        for idx, item in enumerate(self._panel_options):
            option = getattr(item, "option", None) or {}
            if int(option.get("panel_id", 0)) != int(left.get("panel_id", 0)):
                self.RightPanelCombo.SelectedIndex = idx
                return
        self.RightPanelCombo.SelectedIndex = right_index

    def _get_or_load_rows(self, option):
        """Return mutable working rows for a panel, loading once from model."""
        if not self._has_schedule(option):
            return []
        panel_id = int(option.get("panel_id", 0) or 0)
        if panel_id <= 0:
            return []
        cached = self._working_rows_by_panel.get(panel_id)
        if cached is not None:
            return cached
        doc = self._active_doc()
        if doc is None:
            return []
        loaded = ps_repo.build_panel_rows(
            doc,
            option,
            panel_id_set=self._known_panel_ids,
            all_circuits=self._all_circuits_cache,
        )
        self._working_rows_by_panel[panel_id] = self._clone_rows(loaded)
        return self._working_rows_by_panel[panel_id]

    def _reload_from_selected_panels(self):
        """Rebind active panel selections to working row state."""
        doc = self._active_doc()
        if doc is None:
            self._set_status("No active document.")
            return

        self._ensure_distinct_panels()
        self._left_option = self._selected_option(self.LeftPanelCombo)
        self._right_option = self._selected_option(self.RightPanelCombo)
        if self._left_option is None or self._right_option is None:
            return

        self._left_rows = self._get_or_load_rows(self._left_option)
        self._right_rows = self._get_or_load_rows(self._right_option)
        self._refresh_side_schedule_state("left", self._left_option)
        self._refresh_side_schedule_state("right", self._right_option)
        self._clear_preview_slots(refresh=False)
        self._recompute_transferability()
        self._refresh_row_views()

        self._set_profile_rich_text(self._left_option, self.LeftPanelMeta)
        self._set_profile_rich_text(self._right_option, self.RightPanelMeta)
        self._set_currents_rich_text(
            self._left_option,
            self.LeftPanelCurrentConnectedText,
            self.LeftPanelCurrentDemandText,
        )
        self._set_currents_rich_text(
            self._right_option,
            self.RightPanelCurrentConnectedText,
            self.RightPanelCurrentDemandText,
        )
        if self.LeftPanelTypeMeta is not None:
            self.LeftPanelTypeMeta.Text = self._format_schedule_type_label(self._left_option)
        if self.RightPanelTypeMeta is not None:
            self.RightPanelTypeMeta.Text = self._format_schedule_type_label(self._right_option)
        if not self._has_schedule(self._left_option) or not self._has_schedule(self._right_option):
            self._set_status(
                "Select template(s) and create missing panel schedules before cross-panel planning. "
                "Staged operations: {0}".format(len(self._operation_history))
            )
        else:
            self._set_status(
                "Loaded panel rows. Staged operations: {0}".format(len(self._operation_history))
            )
        self._update_action_buttons_state()

    def _collect_staged_circuit_ids(self):
        """Return circuit ids currently involved in staged operations."""
        ids = set()
        for operation in list(self._operation_history or []):
            if str(operation.get("status", "pending")).lower() != "pending":
                continue
            for move in list(operation.get("placements") or []):
                try:
                    cid = int(move.get("circuit_id", 0) or 0)
                    if cid != 0:
                        ids.add(cid)
                except Exception:
                    continue
        return ids

    def _iter_pending_operations(self):
        """Yield staged operations that are still pending apply."""
        for operation in list(self._operation_history or []):
            if str(operation.get("status", "pending")).lower() == "pending":
                yield operation

    def _recompute_transferability(self):
        """Recompute cross-panel compatibility and row display metadata."""
        staged_ids = self._collect_staged_circuit_ids()
        for row in self._left_rows:
            row["is_editable"] = bool(self._is_row_editable(row))
            if bool(row.get("is_excess_slot", False)):
                ok, reason = False, "Exceeds equipment slot capacity."
            elif not bool(row.get("is_editable", True)):
                ok, reason = False, "Owned by {0}".format(str(row.get("edited_by", "") or "another user"))
            elif str(row.get("kind", "")).lower() in ("spare", "space"):
                ok, reason = True, ""
            else:
                ok, reason = ps_repo.evaluate_transferability(row, self._right_option)
            row["transferable"] = bool(ok)
            row["transfer_reason"] = reason
            row["is_staged"] = bool(int(row.get("circuit_id", 0) or 0) in staged_ids and str(row.get("kind", "")).lower() != "empty")
            row["meta_text"] = self._row_meta_text(row)
            row["indicator_text"] = self._row_indicator_text(row)

        for row in self._right_rows:
            row["is_editable"] = bool(self._is_row_editable(row))
            if bool(row.get("is_excess_slot", False)):
                ok, reason = False, "Exceeds equipment slot capacity."
            elif not bool(row.get("is_editable", True)):
                ok, reason = False, "Owned by {0}".format(str(row.get("edited_by", "") or "another user"))
            elif str(row.get("kind", "")).lower() in ("spare", "space"):
                ok, reason = True, ""
            else:
                ok, reason = ps_repo.evaluate_transferability(row, self._left_option)
            row["transferable"] = bool(ok)
            row["transfer_reason"] = reason
            row["is_staged"] = bool(int(row.get("circuit_id", 0) or 0) in staged_ids and str(row.get("kind", "")).lower() != "empty")
            row["meta_text"] = self._row_meta_text(row)
            row["indicator_text"] = self._row_indicator_text(row)

    def _row_meta_text(self, row):
        """Return concise metadata text for row body."""
        kind = str(row.get("kind", "") or "").lower()
        poles = int(max(1, row.get("poles", 1) or 1))
        if kind == "space":
            base_text = "-/{0}P".format(poles)
        elif kind in ("circuit", "spare"):
            base_text = "{0}/{1}P".format(row.get("rating_text", "-"), poles)
        else:
            base_text = "-"
        notes_text = str(row.get("schedule_notes_text", "") or "").strip()
        if notes_text:
            return "{0} [{1}]".format(base_text, notes_text)
        return base_text

    def _row_indicator_text(self, row):
        """Return short slot flags for grouped/locked/ownership states."""
        tags = []
        if bool(row.get("is_excess_slot", False)):
            tags.append("EXCESS")
        if bool(row.get("is_slot_grouped", False)):
            group_no = int(row.get("slot_group_number", 0) or 0)
            if group_no > 0:
                tags.append("GRP{0}".format(group_no))
            else:
                tags.append("GRP")
        if bool(row.get("is_slot_locked", False)):
            tags.append("LOCK")
        if not bool(row.get("is_editable", True)) and not bool(row.get("is_excess_slot", False)):
            tags.append("OWN")
        return " ".join(tags)

    def _sorted_rows(self, rows, option):
        """Sort rows by panel slot display order."""
        slot_order = ps_repo.get_option_slot_order(option, include_excess=True)
        index = {int(slot): i for i, slot in enumerate(slot_order)}
        return sorted(list(rows or []), key=lambda x: (index.get(int(x.get("slot", 0)), 999999), int(x.get("slot", 0))))

    def _refresh_row_views(self):
        """Rebuild list item sources and staged change log."""
        left_preview = set(self._left_preview_slots)
        right_preview = set(self._right_preview_slots)
        preview_keys = set(self._preview_moving_keys)
        allow_default_replace = bool(getattr(self.AllowDiscardToggle, "IsChecked", False))
        left_lookup = self._slot_display_column_lookup(self._left_option)
        right_lookup = self._slot_display_column_lookup(self._right_option)

        left_items = []
        for i, row in enumerate(self._sorted_rows(self._left_rows, self._left_option), 1):
            covered = set(ps_repo.get_row_covered_slots(row, option=self._left_option))
            is_preview = bool(covered.intersection(left_preview))
            is_moving = str(row.get("row_key", "") or "") in preview_keys
            left_items.append(
                SwapRowItem(
                    row,
                    i,
                    self._brushes,
                    is_preview=is_preview,
                    show_divider=False,
                    preview_slots=left_preview,
                    is_moving=is_moving,
                    allow_default_replace=allow_default_replace,
                    display_column_lookup=left_lookup,
                )
            )

        right_items = []
        for i, row in enumerate(self._sorted_rows(self._right_rows, self._right_option), 1):
            covered = set(ps_repo.get_row_covered_slots(row, option=self._right_option))
            is_preview = bool(covered.intersection(right_preview))
            is_moving = str(row.get("row_key", "") or "") in preview_keys
            right_items.append(
                SwapRowItem(
                    row,
                    i,
                    self._brushes,
                    is_preview=is_preview,
                    show_divider=False,
                    preview_slots=right_preview,
                    is_moving=is_moving,
                    allow_default_replace=allow_default_replace,
                    display_column_lookup=right_lookup,
                )
            )

        self.LeftRowsList.ItemsSource = left_items
        self.RightRowsList.ItemsSource = right_items
        self._rebuild_change_log()
        self._update_action_buttons_state()

    def _update_action_buttons_state(self):
        """Enable add/remove special-row buttons based on current selection."""
        can_add_spare = False
        can_add_space = False
        can_remove = False
        selected_poles = int(max(1, self._selected_add_poles()))
        for list_ctrl, rows in (
            (self.LeftRowsList, self._left_rows),
            (self.RightRowsList, self._right_rows),
        ):
            current_option = self._left_option if list_ctrl is self.LeftRowsList else self._right_option
            option_supports_poles = self._panel_accepts_add_poles(current_option, selected_poles)
            is_data_panel = False
            try:
                is_data_panel = bool((current_option or {}).get("schedule_type") == ps_repo.PSTYPE_DATA)
            except Exception:
                is_data_panel = False
            for row in self._selected_rows_from_list(list_ctrl, rows):
                kind = str(row.get("kind", "") or "").lower()
                if kind == "empty" and bool(row.get("is_editable", True)) and bool(option_supports_poles):
                    can_add_space = True
                    if not bool(is_data_panel):
                        can_add_spare = True
                elif kind in ("spare", "space") and bool(row.get("is_editable", True)):
                    can_remove = True
        if self.AddSpareButton is not None:
            self.AddSpareButton.IsEnabled = bool(can_add_spare)
        if self.AddSpaceButton is not None:
            self.AddSpaceButton.IsEnabled = bool(can_add_space)
        if self.RemoveSpecialButton is not None:
            self.RemoveSpecialButton.IsEnabled = bool(can_remove)
        self._update_create_schedule_button("left")
        self._update_create_schedule_button("right")

    def _panel_accepts_add_poles(self, option, poles):
        """Return True when panel option can accept add spare/space poles."""
        if option is None:
            return False
        poles_value = int(max(1, poles or 1))
        schedule_type = None
        try:
            schedule_type = option.get("schedule_type")
        except Exception:
            schedule_type = None
        is_data = False
        try:
            is_data = bool(schedule_type == ps_repo.PSTYPE_DATA)
        except Exception:
            is_data = False
        if bool(is_data):
            return poles_value == 1

        model = option.get("equipment_model")
        branch_options = []
        try:
            branch_options = list(getattr(model, "branch_circuit_options", None) or [])
        except Exception:
            branch_options = []
        if not branch_options and isinstance(model, dict):
            try:
                branch_options = list(model.get("branch_circuit_options") or [])
            except Exception:
                branch_options = []

        allowed_poles = set()
        for item in list(branch_options or []):
            try:
                value = int(item.get("poles", 0) or 0)
            except Exception:
                value = 0
            if value > 0:
                allowed_poles.add(value)

        if not allowed_poles:
            return poles_value == 1
        return poles_value in allowed_poles

    def _set_status(self, text):
        """Set status line text."""
        if self.SwapSummaryText is not None:
            self.SwapSummaryText.Text = str(text or "")

    def _resolve_rows_and_option(self, list_name):
        """Map list control name to backing rows/option."""
        name = str(list_name or "")
        if name == "LeftRowsList":
            return self._left_rows, self._left_option, "left"
        if name == "RightRowsList":
            return self._right_rows, self._right_option, "right"
        return None, None, None

    def _selected_items_for_list(self, list_ctrl):
        """Return selected view-model items from a list control."""
        try:
            return list(list_ctrl.SelectedItems or [])
        except Exception:
            return []

    def _selected_rows_from_list(self, list_ctrl, rows):
        """Return selected backing rows by row_key for a list."""
        row_map = {}
        for row in list(rows or []):
            row_map[str(row.get("row_key", "") or "")] = row
        selected_rows = []
        for item in self._selected_items_for_list(list_ctrl):
            key = str(getattr(item, "row_key", "") or "")
            row = row_map.get(key)
            if row is not None:
                selected_rows.append(row)
        return selected_rows

    def _next_temp_circuit_id(self):
        """Return unique temporary negative id for staged synthetic rows."""
        self._temp_row_seed -= 1
        return int(self._temp_row_seed)

    def _build_staged_special_row(self, option, slot, kind, spare_rating=None):
        """Build staged spare/space row for an empty slot."""
        slot_value = int(slot or 0)
        normalized_kind = "spare" if str(kind or "").strip().lower() == "spare" else "space"
        label = "SPARE" if normalized_kind == "spare" else "SPACE"
        row = ps_repo.build_empty_row(option, slot_value)
        row["kind"] = normalized_kind
        row["is_regular_circuit"] = False
        row["circuit"] = None
        row["circuit_id"] = self._next_temp_circuit_id()
        row["load_name"] = label
        row["poles"] = 1
        if normalized_kind == "spare" and int(spare_rating or 0) > 0:
            row["rating"] = float(int(spare_rating))
            row["rating_text"] = "{0}A".format(int(spare_rating))
        else:
            row["rating"] = None
            row["rating_text"] = "-"
        row["transferable"] = True
        row["transfer_reason"] = ""
        row["edited_by"] = ""
        row["is_editable"] = bool(row.get("is_valid_slot", True))
        row["circuit_number"] = ps_repo.predict_circuit_number(option, slot_value, poles=1)
        row["row_key"] = "panel:{0}|slot:{1}|{2}:temp:{3}".format(
            int(option.get("panel_id", 0) or 0),
            slot_value,
            normalized_kind,
            abs(int(row["circuit_id"])),
        )
        return row

    def _is_movable_row(self, row):
        """Return True for row kinds that can be moved/reordered."""
        kind = str(row.get("kind", "") or "").strip().lower()
        return bool(kind in ("circuit", "spare", "space") and bool(row.get("is_editable", True)))

    def _is_row_editable(self, row):
        """Return False when row is owned by another user."""
        if bool(row.get("is_excess_slot", False)):
            return False
        if not bool(row.get("is_valid_slot", True)):
            return False
        owner = str(row.get("edited_by", "") or "").strip()
        if not owner:
            return True
        if not self._current_username:
            return True
        return bool(owner.lower() == self._current_username)

    def _selected_add_poles(self):
        """Return selected pole count for staged add spare/space."""
        if bool(getattr(self.AddPole3Radio, "IsChecked", False)):
            return 3
        if bool(getattr(self.AddPole2Radio, "IsChecked", False)):
            return 2
        return 1

    def _parse_amp_text(self, text, default_value=20):
        """Parse an amp text entry and return normalized integer amps."""
        raw = str(text or "").strip().upper().replace(" ", "")
        if not raw:
            return int(default_value)
        if raw.endswith("A"):
            raw = raw[:-1]
        try:
            return int(round(float(raw)))
        except Exception:
            return None

    def _normalize_amp_textbox(self, text_box, default_value=20):
        """Normalize textbox value to '<N> A' format and return numeric amps."""
        if text_box is None:
            return int(default_value)
        parsed = self._parse_amp_text(getattr(text_box, "Text", ""), default_value=default_value)
        if parsed is None or int(parsed) <= 0:
            parsed = int(default_value)
        try:
            text_box.Text = "{0} A".format(int(parsed))
        except Exception:
            pass
        return int(parsed)

    def _selected_spare_rating(self):
        """Return spare rating from persistent textbox, or None on invalid input."""
        value = self._normalize_amp_textbox(self.SpareRatingTextBox, 20)
        if int(value) > 0:
            return int(value)
        forms.alert("Spare rating must be a positive integer.", title=TITLE)
        return None

    def _selected_spare_frame(self):
        """Return spare frame amps from persistent textbox, or None on invalid input."""
        value = self._normalize_amp_textbox(self.SpareFrameTextBox, 20)
        if int(value) > 0:
            return int(value)
        forms.alert("Spare frame must be a positive integer.", title=TITLE)
        return None

    def amp_input_focus(self, sender, args):
        """Select all text when amp textbox receives keyboard focus."""
        try:
            sender.SelectAll()
        except Exception:
            pass

    def amp_input_mouse_down(self, sender, args):
        """Focus textbox on click and select all for quick overwrite."""
        try:
            if sender is not None and not bool(sender.IsKeyboardFocusWithin):
                args.Handled = True
                sender.Focus()
        except Exception:
            pass

    def spare_rating_lost_focus(self, sender, args):
        """Normalize spare rating text when leaving field."""
        self._normalize_amp_textbox(self.SpareRatingTextBox, 20)

    def spare_frame_lost_focus(self, sender, args):
        """Normalize spare frame text when leaving field."""
        self._normalize_amp_textbox(self.SpareFrameTextBox, 20)

    def _expand_moving_keys_for_groups(self, source_rows, moving_keys, maintain_groups, source_option=None):
        """Expand selected row keys to include all rows in selected groups."""
        keys = [str(x) for x in list(moving_keys or []) if str(x or "").strip()]
        if not bool(maintain_groups):
            return list(dict.fromkeys(keys))

        source_map = {}
        selected_group_numbers = {}
        for row in list(source_rows or []):
            key = str(row.get("row_key", "") or "")
            source_map[key] = row
        for key in keys:
            row = source_map.get(key)
            if row is None:
                continue
            group_no = int(row.get("slot_group_number", 0) or 0)
            if group_no > 0:
                if group_no not in selected_group_numbers:
                    selected_group_numbers[group_no] = set()
                selected_group_numbers[group_no].add(int(row.get("slot", 0) or 0) % 2)
        if not selected_group_numbers:
            return list(dict.fromkeys(keys))

        expanded = list(keys)
        source_mode = str((source_option or {}).get("sort_mode", "")).strip().lower()
        is_panelboard = source_mode in ("panelboard", "panelboard_two_columns_across")
        for row in list(source_rows or []):
            if not self._is_movable_row(row):
                continue
            group_no = int(row.get("slot_group_number", 0) or 0)
            if group_no <= 0:
                continue
            if group_no not in selected_group_numbers:
                continue
            if is_panelboard:
                parity = int(row.get("slot", 0) or 0) % 2
                if parity not in selected_group_numbers.get(group_no, set()):
                    continue
            expanded.append(str(row.get("row_key", "") or ""))
        return list(dict.fromkeys([x for x in expanded if str(x or "").strip()]))

    def _record_operation(self, placements, before_snapshots):
        """Append one staged operation to history."""
        if not placements:
            return
        before = dict(before_snapshots or {})
        operation = {
            "seq": 0,
            "placements": list(placements),
            "before": before,
            "status": "pending",
            "message": "",
        }
        self._operation_history.append(operation)
        self._renumber_operations()

    def _renumber_operations(self):
        """Ensure sequence numbering is continuous after list edits."""
        for idx, operation in enumerate(list(self._operation_history or []), 1):
            operation["seq"] = int(idx)
        self._operation_seq = len(list(self._operation_history or []))

    def _collapse_pending_same_panel_moves(self, placements, before_snapshots):
        """Collapse prior pending in-panel reorders for circuits moved again.

        Cross-panel moves are not collapsed because intermediate steps may be
        required to preserve dependency order between staged sequences.
        """
        targets = {}
        for move in list(placements or []):
            if str(move.get("action", "move") or "move").lower() != "move":
                continue
            circuit_id = int(move.get("circuit_id", 0) or 0)
            if circuit_id <= 0:
                continue
            if circuit_id not in targets:
                targets[circuit_id] = move
        if not targets:
            return

        origin_by_circuit = {}
        updated_history = []
        for operation in list(self._operation_history or []):
            if str(operation.get("status", "pending") or "pending").lower() != "pending":
                updated_history.append(operation)
                continue
            collapse_circuits = set()
            for move in list(operation.get("placements") or []):
                action = str(move.get("action", "move") or "move").lower()
                if action != "move":
                    continue
                circuit_id = int(move.get("circuit_id", 0) or 0)
                target_move = targets.get(circuit_id)
                if target_move is None:
                    continue
                prior_from = int(move.get("from_panel_id", 0) or 0)
                prior_to = int(move.get("to_panel_id", 0) or 0)
                new_from = int(target_move.get("from_panel_id", 0) or 0)
                new_to = int(target_move.get("to_panel_id", 0) or 0)
                prior_is_inpanel = bool(prior_from > 0 and prior_from == prior_to)
                new_is_inpanel = bool(new_from > 0 and new_from == new_to)
                if not prior_is_inpanel or not new_is_inpanel:
                    continue
                if prior_from != new_from:
                    continue
                collapse_circuits.add(circuit_id)

            kept = []
            removed_any = False
            for move in list(operation.get("placements") or []):
                action = str(move.get("action", "move") or "move").lower()
                circuit_id = int(move.get("circuit_id", 0) or 0)
                linked_circuit_id = int(move.get("for_circuit_id", 0) or 0)
                if action == "move" and circuit_id in collapse_circuits:
                    if circuit_id not in origin_by_circuit:
                        origin_by_circuit[circuit_id] = dict(move)
                    removed_any = True
                    continue
                if (StagedAction.is_remove_spare(action) or StagedAction.is_remove_space(action)) and linked_circuit_id in collapse_circuits:
                    removed_any = True
                    continue
                kept.append(move)

            if removed_any:
                for panel_id, rows in dict(operation.get("before") or {}).items():
                    pid = int(panel_id or 0)
                    if pid <= 0 or pid in before_snapshots:
                        continue
                    before_snapshots[pid] = self._clone_rows(rows)

            if kept:
                operation["placements"] = kept
                updated_history.append(operation)
            elif not removed_any:
                updated_history.append(operation)

        for circuit_id, placement in targets.items():
            origin = origin_by_circuit.get(int(circuit_id))
            if not origin:
                continue
            placement["from_panel_id"] = int(origin.get("from_panel_id", placement.get("from_panel_id", 0)) or 0)
            placement["from_panel_name"] = str(origin.get("from_panel_name", placement.get("from_panel_name", "")) or "")
            placement["old_slot"] = int(origin.get("old_slot", placement.get("old_slot", 0)) or 0)
            placement["old_covered_slots"] = [int(x) for x in list(origin.get("old_covered_slots") or placement.get("old_covered_slots") or []) if int(x) > 0]
            placement["same_panel"] = bool(
                int(placement.get("from_panel_id", 0) or 0) == int(placement.get("to_panel_id", 0) or 0)
            )

        self._operation_history = updated_history
        self._renumber_operations()

    def _placement_preapply_warning(self, move):
        """Return pre-apply warning text for risky staged moves."""
        if str(move.get("action", "move") or "move").lower() != "move":
            return ""
        from_panel = int(move.get("from_panel_id", 0) or 0)
        to_panel = int(move.get("to_panel_id", 0) or 0)
        if from_panel <= 0 or to_panel <= 0 or from_panel == to_panel:
            return ""
        group_no = int(move.get("slot_group_number", 0) or 0)
        if group_no <= 0:
            return ""
        return "Grouped circuit will be ungrouped after cross-panel move."

    def _set_operation_status(self, operation, status, message=""):
        """Set operation status metadata in-place."""
        if operation is None:
            return
        operation["status"] = str(status or "pending").lower()
        operation["message"] = str(message or "")

    def _rebuild_change_log(self):
        """Render staged operation history list."""
        items = []
        for operation in self._operation_history:
            seq = int(operation.get("seq", 0))
            status = str(operation.get("status", "pending") or "pending").lower()
            if status == "completed":
                bg = self._brushes.get("log_success_bg")
                status_label = "APPLIED"
            elif status == "warning":
                bg = self._brushes.get("log_warning_bg")
                status_label = "WARNING"
            elif status == "failed":
                bg = self._brushes.get("log_failed_bg")
                status_label = "FAILED"
            else:
                bg = self._brushes.get("log_pending_bg")
                status_label = ""
            for move in list(operation.get("placements") or []):
                action = str(move.get("action", "move") or "move")
                old_slots = ",".join([str(x) for x in list(move.get("old_covered_slots") or [])]) or "-"
                new_slots = ",".join([str(x) for x in list(move.get("new_covered_slots") or [])]) or "-"
                load = move.get("load_name", "") or "-"
                from_panel = move.get("from_panel_name", "") or "Unknown Panel"
                to_panel = move.get("to_panel_name", "") or "Unknown Panel"
                target_condition = str(move.get("target_condition", "") or "").strip().lower()
                condition_suffix = " ({0})".format(target_condition) if target_condition else ""
                warning_text = self._placement_preapply_warning(move)
                if warning_text:
                    move_bg = self._brushes.get("log_warning_bg")
                else:
                    move_bg = bg
                if status_label:
                    prefix = "#{0} [{1}]".format(seq, status_label)
                else:
                    prefix = "#{0}".format(seq)
                if StagedAction.is_add_spare(action) or StagedAction.is_add_space(action):
                    text = "{0} ADD {1} -> {2}[{3}]{4}".format(prefix, load, to_panel, new_slots, condition_suffix)
                elif StagedAction.is_remove_spare(action) or StagedAction.is_remove_space(action):
                    text = "{0} REMOVE {1} {2}[{3}]".format(prefix, load, from_panel, old_slots)
                elif int(move.get("from_panel_id", 0)) == int(move.get("to_panel_id", 0)):
                    text = "{0} {1} {2}[{3}] -> {2}[{4}]".format(prefix, load, from_panel, old_slots, new_slots)
                else:
                    text = "{0} {1} {2}[{3}] -> {4}[{5}]{6}".format(
                        prefix, load, from_panel, old_slots, to_panel, new_slots, condition_suffix
                    )
                if warning_text:
                    text = "{0}  WARN: {1}".format(text, warning_text)
                items.append(ChangeLogItem(text, background=move_bg))
            if operation.get("message") and status in ("warning", "failed"):
                if status_label:
                    prefix = "#{0} [{1}]".format(seq, status_label)
                else:
                    prefix = "#{0}".format(seq)
                items.append(
                    ChangeLogItem(
                        "{0} {1}".format(prefix, operation.get("message", "")),
                        background=bg,
                    )
                )
        if not items:
            items = [ChangeLogItem("No staged changes.", background=self._brushes.get("log_pending_bg"))]
        if self.ChangeLogList is not None:
            self.ChangeLogList.ItemsSource = items
        pending_exists = any(True for _ in self._iter_pending_operations())
        applied_exists = any(
            str(op.get("status", "pending")).lower() in ("completed", "warning", "failed")
            for op in self._operation_history
        )
        if self.UndoLastButton is not None:
            self.UndoLastButton.IsEnabled = bool(pending_exists)
        if self.UndoAllButton is not None:
            self.UndoAllButton.IsEnabled = bool(pending_exists)
        if self.ApplyButton is not None:
            self.ApplyButton.IsEnabled = bool(pending_exists)
        if self.ClearListButton is not None:
            self.ClearListButton.IsEnabled = bool(applied_exists)

    def _clear_preview_slots(self, refresh=True):
        """Clear drag preview highlight state."""
        changed = bool(self._left_preview_slots or self._right_preview_slots)
        self._left_preview_slots = set()
        self._right_preview_slots = set()
        self._preview_signature = None
        self._preview_moving_keys = set()
        if changed and bool(refresh):
            self._refresh_row_views()

    def left_panel_changed(self, sender, args):
        """Handle left panel selection change."""
        if self._suspend_panel_events:
            return
        self._reload_from_selected_panels()

    def right_panel_changed(self, sender, args):
        """Handle right panel selection change."""
        if self._suspend_panel_events:
            return
        self._reload_from_selected_panels()

    def _update_create_schedule_button(self, side_name):
        """Enable create-schedule button only when a compatible template is selected."""
        controls = self._side_controls(side_name)
        option = controls.get("option")
        selector = controls.get("template_list")
        button = controls.get("create_button")
        if button is None:
            return
        if self._has_schedule(option):
            button.IsEnabled = False
            return
        selected = self._selected_template_option(selector)
        button.IsEnabled = bool(selected)

    def left_template_changed(self, sender, args):
        """Handle left missing-schedule template selection."""
        self._update_create_schedule_button("left")

    def right_template_changed(self, sender, args):
        """Handle right missing-schedule template selection."""
        self._update_create_schedule_button("right")

    def _create_schedule_for_side(self, side_name):
        """Create missing panel schedule view for selected side/template."""
        doc = self._active_doc()
        if doc is None:
            self._set_status("No active document.")
            return
        controls = self._side_controls(side_name)
        option = controls.get("option")
        selector = controls.get("template_list")
        if option is None:
            return
        if self._has_schedule(option):
            self._set_status("Selected panel already has a panel schedule view.")
            return
        template_option = self._selected_template_option(selector)
        if template_option is None:
            forms.alert("Select a compatible panel schedule template first.", title=TITLE)
            return

        panel_id = int(option.get("panel_id", 0) or 0)
        template_id = int(template_option.get("template_id", 0) or 0)
        if panel_id <= 0 or template_id <= 0:
            forms.alert("Invalid panel/template selection.", title=TITLE)
            return

        def _tx_action():
            schedule_view = ps_repo.create_panel_schedule_instance_view(
                doc,
                template_id,
                panel_id,
            )
            if schedule_view is None:
                raise Exception("Panel schedule creation returned no view.")
            ps_repo.attach_schedule_to_option(doc, option, schedule_view)

        try:
            self._run_transaction(doc, "Create Panel Schedule View", _tx_action)
        except Exception as ex:
            forms.alert("Failed to create panel schedule view:\n{0}".format(str(ex)), title=TITLE)
            self._set_status("Panel schedule creation failed.")
            return

        self._template_items_by_panel[int(panel_id)] = []
        self._working_rows_by_panel.pop(int(panel_id), None)
        self._refresh_model_caches(doc)
        self._reload_from_selected_panels()
        self._set_status(
            "Created panel schedule for {0}.".format(option.get("panel_name", "selected panel"))
        )

    def left_create_schedule_clicked(self, sender, args):
        """Create missing panel schedule view for the left panel."""
        self._create_schedule_for_side("left")

    def right_create_schedule_clicked(self, sender, args):
        """Create missing panel schedule view for the right panel."""
        self._create_schedule_for_side("right")

    def list_selection_changed(self, sender, args):
        """Refresh action-button enabled state when list selection changes."""
        self._update_action_buttons_state()

    def add_poles_changed(self, sender, args):
        """Refresh add/remove button state when add-pole mode changes."""
        self._update_action_buttons_state()

    def panel_dropdown_opened(self, sender, args):
        """Show full panel descriptors while dropdown is open."""
        self._set_panel_combo_mode(True)

    def panel_dropdown_closed(self, sender, args):
        """Collapse selected text to panel + part type when dropdown closes."""
        self._set_panel_combo_mode(False)

    def refresh_clicked(self, sender, args):
        """Refresh current view from working state."""
        self._reload_from_selected_panels()

    def undo_last_clicked(self, sender, args):
        """Undo most recent staged operation."""
        pending = [op for op in self._operation_history if str(op.get("status", "pending")).lower() == "pending"]
        if not pending:
            self._set_status("Nothing to undo.")
            return

        op = pending[-1]
        try:
            self._operation_history.remove(op)
        except Exception:
            pass
        before = dict(op.get("before") or {})
        for panel_id, rows in before.items():
            self._working_rows_by_panel[int(panel_id)] = self._clone_rows(rows)
        self._renumber_operations()
        self._reload_from_selected_panels()
        self._set_status("Undid operation #{0}.".format(int(op.get("seq", 0))))

    def undo_all_clicked(self, sender, args):
        """Undo all pending staged operations."""
        pending_count = len([op for op in self._operation_history if str(op.get("status", "pending")).lower() == "pending"])
        if pending_count <= 0:
            self._set_status("No pending staged actions to undo.")
            return
        self._operation_history = [op for op in self._operation_history if str(op.get("status", "pending")).lower() != "pending"]
        self._renumber_operations()
        self._working_rows_by_panel = {}
        self._refresh_model_caches()
        self._reload_from_selected_panels()
        self._rebuild_change_log()
        self._set_status("Undid {0} pending staged sequence(s).".format(int(pending_count)))

    def clear_list_clicked(self, sender, args):
        """Clear all applied/warning/failed history entries."""
        clearable_count = len(
            [op for op in self._operation_history if str(op.get("status", "pending")).lower() in ("completed", "warning", "failed")]
        )
        if clearable_count <= 0:
            self._set_status("No applied history to clear.")
            return
        self._operation_history = []
        self._renumber_operations()
        self._working_rows_by_panel = {}
        self._refresh_model_caches()
        self._reload_from_selected_panels()
        self._rebuild_change_log()
        self._set_status("Cleared {0} history sequence(s).".format(int(clearable_count)))

    def apply_clicked(self, sender, args):
        """Apply pending staged operations to the Revit model."""
        self._apply_pending_operations()

    def _option_for_panel_id(self, panel_id):
        """Return panel option dict by panel id."""
        item = self._panel_option_by_id.get(int(panel_id or 0))
        if item is None:
            return None
        return getattr(item, "option", None)

    def _element_by_id_value(self, doc, id_value):
        """Resolve element by integer id value."""
        value = int(id_value or 0)
        if value <= 0:
            return None
        try:
            return doc.GetElement(revit_helpers.elementid_from_value(value))
        except Exception:
            return None

    def _run_transaction(self, doc, name, action):
        """Run one scoped transaction with automatic rollback on exception."""
        with revit.Transaction(str(name or "BatchSwap Apply"), doc):
            return action()

    def _slot_cells(self, schedule_view, slot):
        """Return row/col cells for a schedule slot."""
        cache_key = self._schedule_slot_cache_key(schedule_view, slot)
        if cache_key is not None and cache_key in self._slot_cells_cache:
            return list(self._slot_cells_cache.get(cache_key) or [])
        slot_value = int(slot or 0)
        if slot_value <= 0:
            return []
        try:
            cells = list(ps_repo.get_cells_by_slot_number(schedule_view, slot_value) or [])
        except Exception:
            cells = []
        ordered = self._order_cells_for_slot(schedule_view, slot_value, cells)
        if cache_key is not None:
            self._slot_cells_cache[cache_key] = list(ordered)
        return ordered

    def _schedule_slot_cache_key(self, schedule_view, slot):
        """Return cache key tuple for slot-cell lookup."""
        if schedule_view is None:
            return None
        try:
            sid = int(ps_repo._idval(schedule_view.Id))
        except Exception:
            sid = 0
        slot_value = int(slot or 0)
        if sid <= 0 or slot_value <= 0:
            return None
        return (sid, slot_value)

    def _layout_context(self, schedule_view):
        """Return cached layout context and preferred move columns for one schedule."""
        if schedule_view is None:
            return {"max_slot": 0, "sort_mode": "panelboard", "preferred_cols": {}}
        try:
            sid = int(ps_repo._idval(schedule_view.Id))
        except Exception:
            sid = 0
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
                    try:
                        cid = int(self._cell_circuit_id(schedule_view, row, col))
                    except Exception:
                        cid = 0
                    if cid <= 0:
                        continue
                    bucket = col_counts.setdefault(display_col, {})
                    bucket[int(col)] = int(bucket.get(int(col), 0) or 0) + 1
        preferred_cols = {}
        for display_col, bucket in col_counts.items():
            ordered = [pair[0] for pair in sorted(bucket.items(), key=lambda x: (-int(x[1]), int(x[0])))]
            preferred_cols[int(display_col)] = [int(x) for x in ordered]
        context = {
            "max_slot": int(max_slot),
            "sort_mode": sort_mode,
            "preferred_cols": preferred_cols,
        }
        if sid > 0:
            self._layout_context_cache[sid] = dict(context)
        return context

    def _order_cells_for_slot(self, schedule_view, slot, cells):
        """Order slot cells by preferred columns derived from existing circuits."""
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
            try:
                priority[int(col)] = int(idx)
            except Exception:
                continue
        fallback_rank = 999999
        ordered.sort(key=lambda pair: (int(priority.get(int(pair[1]), fallback_rank)), int(pair[1]), int(pair[0])))
        return ordered

    def _cell_circuit_id(self, schedule_view, row, col):
        """Return circuit id at one schedule cell, or 0."""
        getter_id = getattr(schedule_view, "GetCircuitIdByCell", None)
        if getter_id is None:
            return 0
        try:
            cid = getter_id(int(row), int(col))
            if cid is None or cid == DB.ElementId.InvalidElementId:
                return 0
            return int(ps_repo._idval(cid))
        except Exception:
            return 0

    def _slot_is_locked(self, schedule_view, slot):
        """Best-effort slot lock query."""
        for row, col in self._slot_cells(schedule_view, slot):
            try:
                return bool(ps_repo._slot_is_locked(schedule_view, row, col))
            except Exception:
                continue
        return False

    def _set_slot_locked(self, schedule_view, slot, is_locked):
        """Best-effort set slot lock state."""
        cells = self._slot_cells(schedule_view, slot)
        setter = getattr(schedule_view, "SetLockSlot", None)
        if setter is None:
            return False
        for row, col in list(cells or []):
            try:
                setter(int(row), int(col), bool(is_locked))
                return True
            except Exception:
                continue
        return False

    def _unlock_slots_with_snapshot(self, schedule_view, slots):
        """Unlock slots and return original lock snapshot."""
        snapshot = {}
        for slot in sorted(set([int(x) for x in list(slots or []) if int(x) > 0])):
            locked = bool(self._slot_is_locked(schedule_view, slot))
            snapshot[int(slot)] = locked
            if locked:
                self._set_slot_locked(schedule_view, slot, False)
        return snapshot

    def _restore_slot_locks(self, schedule_view, snapshot):
        """Restore lock states captured by _unlock_slots_with_snapshot."""
        for slot, was_locked in dict(snapshot or {}).items():
            if bool(was_locked):
                self._set_slot_locked(schedule_view, int(slot), True)

    def _get_circuit_at_slot(self, doc, schedule_view, slot):
        """Return ElectricalSystem at a schedule slot, if any."""
        for row, col in self._slot_cells(schedule_view, slot):
            getter = getattr(schedule_view, "GetCircuitByCell", None)
            if getter:
                try:
                    circuit = getter(int(row), int(col))
                    if isinstance(circuit, ps_repo.DBE.ElectricalSystem):
                        return circuit
                except Exception:
                    pass
            getter_id = getattr(schedule_view, "GetCircuitIdByCell", None)
            if getter_id:
                try:
                    cid = getter_id(int(row), int(col))
                    if cid is None or cid == DB.ElementId.InvalidElementId:
                        continue
                    circuit = doc.GetElement(cid)
                    if isinstance(circuit, ps_repo.DBE.ElectricalSystem):
                        return circuit
                except Exception:
                    pass
        return None

    def _describe_slot_state(self, doc, option, slot, protected_circuit_id=0):
        """Return concise occupancy text for one target slot."""
        schedule_view = (option or {}).get("schedule_view")
        circuit = self._get_circuit_at_slot(doc, schedule_view, slot)
        if not isinstance(circuit, ps_repo.DBE.ElectricalSystem):
            return "empty"
        cid = int(ps_repo._idval(circuit.Id))
        if int(protected_circuit_id or 0) > 0 and cid == int(protected_circuit_id):
            return "self"
        kind = str(ps_repo._kind_from_circuit(circuit) or "circuit").lower()
        cnum = str(getattr(circuit, "CircuitNumber", "") or "")
        return "{0}:{1}".format(kind, cnum or cid)

    def _get_circuit_poles(self, circuit, fallback=1):
        """Return poles count for a circuit-like system."""
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
        """Return covered slots for a circuit under an option layout."""
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
        """Move a circuit slot to a target slot using best-known API signatures."""
        source_slot = int(from_slot or 0)
        target_slot = int(to_slot or 0)
        if source_slot <= 0 or target_slot <= 0 or source_slot == target_slot:
            return
        mover = getattr(schedule_view, "MoveSlotTo", None)
        if mover is None:
            raise Exception("PanelScheduleView.MoveSlotTo is unavailable.")

        src_cells = list(self._slot_cells(schedule_view, source_slot))
        dst_cells = list(self._slot_cells(schedule_view, target_slot))
        src_ordered = list(src_cells)
        if int(circuit_id or 0) > 0 and src_cells:
            owned = [cell for cell in src_cells if int(self._cell_circuit_id(schedule_view, cell[0], cell[1])) == int(circuit_id)]
            if owned:
                src_ordered = owned
        dst_ordered = list(dst_cells)

        attempts = []
        seen_attempts = set()
        def _add_attempt(args):
            key = tuple([int(x) for x in list(args or ())])
            if key in seen_attempts:
                return
            seen_attempts.add(key)
            attempts.append(key)

        for s_row, s_col in src_ordered:
            for d_row, d_col in dst_ordered:
                _add_attempt((int(s_row), int(s_col), int(d_row), int(d_col)))
        if not attempts:
            raise Exception("No valid source/target body cells resolved for MoveSlotTo.")

        LOGGER.info(
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
                LOGGER.info(
                    "MoveSlotTo success from_slot=%s to_slot=%s ckt=%s attempt=%s args=%s",
                    int(source_slot),
                    int(target_slot),
                    int(circuit_id or 0),
                    int(idx),
                    str(args),
                )
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
        """Select a new panel for circuit using ElectricalSystem.SelectPanel."""
        selector = getattr(circuit, "SelectPanel", None)
        if selector is None:
            raise Exception("ElectricalSystem.SelectPanel is unavailable.")
        result = selector(panel)
        if isinstance(result, bool) and not result:
            raise Exception("SelectPanel returned False.")

    def _add_special_to_slot(self, schedule_view, slot, kind):
        """Add spare/space at one slot using best-known API signatures."""
        action = str(kind or "").strip().lower()
        if action not in ("spare", "space"):
            raise Exception("Unsupported special kind: {0}".format(kind))
        method_names = ("AddSpare",) if action == "spare" else ("AddSpace",)
        cells = self._slot_cells(schedule_view, slot)
        errors = []
        for method_name in method_names:
            method = getattr(schedule_view, method_name, None)
            if method is None:
                continue
            attempts = []
            for row, col in list(cells or []):
                attempts.append((int(row), int(col)))
            attempts.append((int(slot),))
            for args in attempts:
                try:
                    result = method(*args)
                    if isinstance(result, bool) and not result:
                        continue
                    return
                except Exception as ex:
                    errors.append("{0}{1} -> {2}".format(method_name, tuple(args), str(ex)))
                    continue
        if errors:
            LOGGER.warning(
                "Add %s failed at slot %s attempts=%s details=%s",
                str(action).upper(),
                int(slot),
                int(len(errors)),
                " | ".join(list(errors or [])[:8]),
            )
        raise Exception("Could not add {0} at slot {1}.".format(action.upper(), int(slot)))

    def _set_circuit_poles(self, circuit, poles):
        """Set circuit poles count when writable API surface exists."""
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
        """Set circuit rating in amps when a writable API surface exists."""
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

    def _remove_special_circuits_in_slots(self, doc, target_option, slots, protected_circuit_id=0):
        """Delete spare/space circuits occupying target slots."""
        schedule_view = target_option.get("schedule_view")
        to_delete = {}
        covered_slots = set([int(x) for x in list(slots or []) if int(x) > 0])
        for slot in list(covered_slots):
            occupant = self._get_circuit_at_slot(doc, schedule_view, slot)
            if not isinstance(occupant, ps_repo.DBE.ElectricalSystem):
                continue
            occ_id = int(ps_repo._idval(occupant.Id))
            if occ_id <= 0 or occ_id == int(protected_circuit_id or 0):
                continue
            kind = str(ps_repo._kind_from_circuit(occupant) or "").lower()
            if kind not in ("spare", "space"):
                raise Exception("Target slot {0} is occupied by a non-spare/space circuit.".format(int(slot)))
            occ_slots = self._covered_slots_for_circuit(target_option, occupant, fallback_slot=slot, fallback_poles=1)
            to_delete[occ_id] = {"circuit": occupant, "slots": occ_slots}
            for occ_slot in occ_slots:
                covered_slots.add(int(occ_slot))

        for occ in to_delete.values():
            try:
                doc.Delete(occ["circuit"].Id)
            except Exception as ex:
                raise Exception("Failed removing target spare/space: {0}".format(ex))

    def _apply_add_placement_create(self, doc, placement):
        """Create one staged spare/space row at target slot."""
        target_panel_id = int(placement.get("to_panel_id", 0) or 0)
        target_option = self._option_for_panel_id(target_panel_id)
        if target_option is None:
            raise Exception("Missing target panel option for add operation.")
        schedule_view = target_option.get("schedule_view")
        kind = "spare" if str(placement.get("action", "")).lower().startswith("add_spare") else "space"
        slot_value = int(placement.get("new_slot", 0) or 0)
        if slot_value <= 0:
            raise Exception("Add operation has invalid target slot.")
        slots = [int(x) for x in list(placement.get("new_covered_slots") or []) if int(x) > 0]
        if not slots:
            slots = [slot_value] if slot_value > 0 else []
        if not slots:
            raise Exception("Add operation has no target slots.")

        snapshot = self._unlock_slots_with_snapshot(schedule_view, slots)
        placement["_add_lock_snapshot"] = dict(snapshot or {})
        placement["_add_slots"] = [int(x) for x in list(slots or [])]
        placement["_add_slot"] = int(slot_value)
        self._add_special_to_slot(schedule_view, int(slot_value), kind)
        doc.Regenerate()
        occupant = self._get_circuit_at_slot(doc, schedule_view, int(slot_value))
        if not isinstance(occupant, ps_repo.DBE.ElectricalSystem):
            raise Exception("Added {0} could not be resolved at slot {1}.".format(str(kind).upper(), int(slot_value)))
        placement["_added_circuit_id"] = int(ps_repo._idval(occupant.Id))

    def _apply_add_placement_finalize(self, doc, placement):
        """Finalize one staged add spare/space placement (unlock + poles/rating)."""
        target_panel_id = int(placement.get("to_panel_id", 0) or 0)
        target_option = self._option_for_panel_id(target_panel_id)
        if target_option is None:
            raise Exception("Missing target panel option for add finalize operation.")
        schedule_view = target_option.get("schedule_view")
        kind = "spare" if str(placement.get("action", "")).lower().startswith("add_spare") else "space"
        desired_poles = int(max(1, placement.get("poles", 1) or 1))
        spare_rating = int(placement.get("spare_rating", 0) or 0)

        slot_value = int(placement.get("_add_slot", placement.get("new_slot", 0)) or 0)
        if slot_value <= 0:
            raise Exception("Add finalize operation has invalid target slot.")
        slots = [int(x) for x in list(placement.get("_add_slots") or placement.get("new_covered_slots") or []) if int(x) > 0]
        if not slots:
            slots = [int(slot_value)]

        prior_snapshot = dict(placement.get("_add_lock_snapshot") or {})
        runtime_snapshot = self._unlock_slots_with_snapshot(schedule_view, slots)
        merged_snapshot = dict(prior_snapshot)
        for slot, was_locked in dict(runtime_snapshot or {}).items():
            if int(slot) not in merged_snapshot:
                merged_snapshot[int(slot)] = bool(was_locked)
            else:
                merged_snapshot[int(slot)] = bool(merged_snapshot[int(slot)] or bool(was_locked))

        try:
            circuit_id = int(placement.get("_added_circuit_id", 0) or 0)
            circuit = self._element_by_id_value(doc, circuit_id) if circuit_id > 0 else None
            if not isinstance(circuit, ps_repo.DBE.ElectricalSystem):
                circuit = self._get_circuit_at_slot(doc, schedule_view, int(slot_value))
            if not isinstance(circuit, ps_repo.DBE.ElectricalSystem):
                raise Exception(
                    "Could not resolve added {0} circuit for finalize at slot {1}.".format(
                        str(kind).upper(),
                        int(slot_value),
                    )
                )

            self._set_circuit_poles(circuit, desired_poles)
            if kind == "spare" and spare_rating > 0:
                self._set_circuit_rating(circuit, spare_rating)
            doc.Regenerate()
        finally:
            self._restore_slot_locks(schedule_view, merged_snapshot)

    def _apply_add_placement(self, doc, placement):
        """Apply one staged add spare/space placement in-process."""
        self._apply_add_placement_create(doc, placement)
        self._apply_add_placement_finalize(doc, placement)

    def _apply_remove_placement(self, doc, placement):
        """Apply one staged remove spare/space placement."""
        panel_id = int(placement.get("from_panel_id", 0) or 0)
        option = self._option_for_panel_id(panel_id)
        if option is None:
            raise Exception("Missing panel option for remove operation.")
        schedule_view = option.get("schedule_view")
        slots = [int(x) for x in list(placement.get("old_covered_slots") or []) if int(x) > 0]
        if not slots:
            slot_value = int(placement.get("old_slot", 0) or 0)
            slots = [slot_value] if slot_value > 0 else []
        if not slots:
            raise Exception("Remove operation has no source slots.")

        snapshot = self._unlock_slots_with_snapshot(schedule_view, slots)
        target_id = int(placement.get("circuit_id", 0) or 0)
        target = self._element_by_id_value(doc, target_id) if target_id > 0 else None
        deleted = False
        if isinstance(target, ps_repo.DBE.ElectricalSystem):
            kind = str(ps_repo._kind_from_circuit(target) or "").lower()
            if kind in ("spare", "space"):
                doc.Delete(target.Id)
                deleted = True
        if not deleted:
            removed = {}
            for slot in slots:
                occupant = self._get_circuit_at_slot(doc, schedule_view, slot)
                if not isinstance(occupant, ps_repo.DBE.ElectricalSystem):
                    continue
                kind = str(ps_repo._kind_from_circuit(occupant) or "").lower()
                if kind in ("spare", "space"):
                    removed[int(ps_repo._idval(occupant.Id))] = occupant
            for occupant in removed.values():
                doc.Delete(occupant.Id)
                deleted = True
        if not deleted:
            raise Exception("No removable spare/space found at target slots.")
        self._restore_slot_locks(schedule_view, snapshot)

    def _apply_move_placement(self, doc, placement):
        """Apply one staged move placement."""
        circuit_id = int(placement.get("circuit_id", 0) or 0)
        if circuit_id <= 0:
            raise Exception("Invalid circuit id in move operation.")
        circuit = self._element_by_id_value(doc, circuit_id)
        if not isinstance(circuit, ps_repo.DBE.ElectricalSystem):
            raise Exception("Circuit {0} could not be resolved.".format(circuit_id))

        target_panel_id = int(placement.get("to_panel_id", 0) or 0)
        target_option = self._option_for_panel_id(target_panel_id)
        if target_option is None:
            raise Exception("Missing target panel option for move operation.")
        target_schedule = target_option.get("schedule_view")
        target_slot = int(placement.get("new_slot", 0) or 0)
        if target_slot <= 0:
            raise Exception("Move operation has invalid target slot.")

        target_slots = [int(x) for x in list(placement.get("new_covered_slots") or []) if int(x) > 0]
        if not target_slots:
            poles_hint = int(max(1, len(list(placement.get("new_covered_slots") or [])) or int(placement.get("poles", 1) or 1)))
            target_slots = ps_repo.get_slot_span_slots_for_option(
                target_option,
                start_slot=int(target_slot),
                pole_count=int(poles_hint),
                require_valid=True,
            )
        if not target_slots:
            raise Exception("Move operation target slot exceeds equipment-supported slot capacity.")
        target_state_before = []
        for slot in list(target_slots or []):
            target_state_before.append("{0}:{1}".format(int(slot), self._describe_slot_state(doc, target_option, slot, protected_circuit_id=circuit_id)))

        current_panel = getattr(circuit, "BaseEquipment", None)
        current_panel_id = int(ps_repo._idval(getattr(current_panel, "Id", None)))
        current_option = self._option_for_panel_id(current_panel_id)
        if current_option is None:
            raise Exception("Current panel for circuit {0} is not in available options.".format(circuit_id))
        current_schedule = current_option.get("schedule_view")
        current_slots = self._covered_slots_for_circuit(
            current_option,
            circuit,
            fallback_slot=int(placement.get("old_slot", 0) or 0),
            fallback_poles=int(max(1, len(list(placement.get("old_covered_slots") or [])) or 1)),
        )

        source_lock_snapshot = self._unlock_slots_with_snapshot(current_schedule, current_slots)
        target_lock_snapshot = self._unlock_slots_with_snapshot(target_schedule, target_slots)
        source_was_locked = any(bool(x) for x in source_lock_snapshot.values())
        LOGGER.info(
            "Apply move ckt=%s from panel=%s slots=%s to panel=%s slot=%s targets=%s condition=%s",
            int(circuit_id),
            int(current_panel_id),
            ",".join([str(x) for x in list(current_slots or [])]) or "-",
            int(target_panel_id),
            int(target_slot),
            ",".join([str(x) for x in list(target_slots or [])]) or "-",
            ",".join(target_state_before) or "-",
        )

        same_panel = bool(current_panel_id == target_panel_id)
        if bool(placement.get("is_regular_circuit", True)):
            deleted_snapshot = self._remove_special_circuits_in_slots(
                doc,
                target_option,
                target_slots,
                protected_circuit_id=circuit_id,
            )
            for slot, was_locked in deleted_snapshot.items():
                if int(slot) not in target_lock_snapshot:
                    target_lock_snapshot[int(slot)] = bool(was_locked)
                else:
                    target_lock_snapshot[int(slot)] = bool(target_lock_snapshot[int(slot)] or bool(was_locked))

        if same_panel:
            current_start = int(ps_repo.get_circuit_start_slot(circuit) or 0)
            if current_start <= 0:
                raise Exception("Could not resolve current slot for circuit {0}.".format(circuit_id))
            try:
                self._move_slot_to(target_schedule, current_start, target_slot, circuit_id=circuit_id)
            except Exception as ex:
                dest_now = self._describe_slot_state(doc, target_option, target_slot, protected_circuit_id=circuit_id)
                raise Exception(
                    "MoveSlotTo failed for ckt {0}: {1} -> {2} (dest={3}) [{4}]".format(
                        int(circuit_id), int(current_start), int(target_slot), str(dest_now), str(ex)
                    )
                )
        else:
            target_panel = target_option.get("panel")
            if target_panel is None:
                raise Exception("Target panel element is unavailable.")
            self._select_panel_for_circuit(circuit, target_panel)
            doc.Regenerate()
            placed_start = int(ps_repo.get_circuit_start_slot(circuit) or 0)
            if placed_start <= 0:
                raise Exception("Circuit {0} has invalid placement after SelectPanel.".format(circuit_id))
            placed_slots = self._covered_slots_for_circuit(
                target_option,
                circuit,
                fallback_slot=placed_start,
                fallback_poles=int(max(1, len(target_slots))),
            )
            placed_lock_snapshot = self._unlock_slots_with_snapshot(target_schedule, placed_slots)
            for slot, was_locked in placed_lock_snapshot.items():
                if int(slot) not in target_lock_snapshot:
                    target_lock_snapshot[int(slot)] = bool(was_locked)
                else:
                    target_lock_snapshot[int(slot)] = bool(target_lock_snapshot[int(slot)] or bool(was_locked))
            if placed_start != target_slot:
                try:
                    self._move_slot_to(target_schedule, placed_start, target_slot, circuit_id=circuit_id)
                except Exception as ex:
                    dest_now = self._describe_slot_state(doc, target_option, target_slot, protected_circuit_id=circuit_id)
                    raise Exception(
                        "MoveSlotTo failed for ckt {0}: {1} -> {2} (dest={3}) [{4}]".format(
                            int(circuit_id), int(placed_start), int(target_slot), str(dest_now), str(ex)
                        )
                    )

        doc.Regenerate()
        final_slots = self._covered_slots_for_circuit(target_option, circuit, fallback_slot=target_slot, fallback_poles=len(target_slots))
        LOGGER.info(
            "Apply move result ckt=%s final_slots=%s target_panel=%s",
            int(circuit_id),
            ",".join([str(x) for x in list(final_slots or [])]) or "-",
            int(target_panel_id),
        )

        self._restore_slot_locks(current_schedule, source_lock_snapshot)
        self._restore_slot_locks(target_schedule, target_lock_snapshot)
        if source_was_locked:
            for slot in list(final_slots or []):
                self._set_slot_locked(target_schedule, int(slot), True)

        return None

    def _apply_placement(self, doc, placement):
        """Apply one placement entry through registered panel schedule operations."""
        action = placement.get("action", "")
        op_key = operation_key_for_action(action)

        option_lookup = {}
        for panel_id, item in dict(self._panel_option_by_id or {}).items():
            option = getattr(item, "option", None)
            if isinstance(option, dict):
                option_lookup[int(panel_id)] = option

        request = OperationRequest(
            operation_key=op_key,
            circuit_ids=[int(placement.get("circuit_id", 0) or 0)] if int(placement.get("circuit_id", 0) or 0) > 0 else [],
            source="batch_swap",
            options={
                "placement": placement,
                "panel_option_lookup": option_lookup,
            },
        )
        return self._panel_schedule_runner.run(request, doc)

    def _apply_sequence_operation(self, doc, operation):
        """Apply one staged sequence operation inside its own transaction group."""
        seq = int(operation.get("seq", 0) or 0)
        seq_tg = DB.TransactionGroup(doc, "Batch Swap Sequence #{0}".format(seq))
        seq_tg.Start()
        try:
            temp_id_map = {}
            for idx, placement in enumerate(list(operation.get("placements") or []), 1):
                action_name = StagedAction.normalize(placement.get("action", ""))
                effective = dict(placement)
                original_circuit_id = int(placement.get("circuit_id", 0) or 0)
                if original_circuit_id < 0 and original_circuit_id in temp_id_map:
                    effective["circuit_id"] = int(temp_id_map.get(original_circuit_id))
                linked_id = int(placement.get("for_circuit_id", 0) or 0)
                if linked_id < 0 and linked_id in temp_id_map:
                    effective["for_circuit_id"] = int(temp_id_map.get(linked_id))
                if action_name == StagedAction.MOVE and int(effective.get("circuit_id", 0) or 0) <= 0:
                    raise Exception(
                        "Move action for staged special could not resolve runtime circuit id (temp id: {0}).".format(
                            int(original_circuit_id)
                        )
                    )
                LOGGER.info(
                    "Apply sequence #%s step=%s action=%s ckt=%s from=%s[%s] to=%s[%s] condition=%s",
                    int(seq),
                    int(idx),
                    str(effective.get("action", "") or ""),
                    int(effective.get("circuit_id", 0) or 0),
                    str(effective.get("from_panel_name", "") or ""),
                    ",".join([str(x) for x in list(effective.get("old_covered_slots") or [])]) or "-",
                    str(effective.get("to_panel_name", "") or ""),
                    ",".join([str(x) for x in list(effective.get("new_covered_slots") or [])]) or "-",
                    str(effective.get("target_condition", "") or ""),
                )
                def _tx_action():
                    return self._apply_placement(doc, effective)

                tx_result = self._run_transaction(
                    doc,
                    "Batch Swap Seq #{0} - Step {1}".format(int(seq), int(idx)),
                    _tx_action,
                )
                if is_add_action(action_name) and int(original_circuit_id) < 0:
                    resolved_id = 0
                    if isinstance(tx_result, dict):
                        try:
                            resolved_id = int(tx_result.get("circuit_id", 0) or 0)
                        except Exception:
                            resolved_id = 0
                    if resolved_id > 0:
                        temp_id_map[int(original_circuit_id)] = int(resolved_id)

            seq_tg.Assimilate()
            return []
        except Exception as ex:
            LOGGER.warning("Apply sequence #{0} failed and rolled back: {1}".format(int(seq), str(ex)))
            try:
                seq_tg.RollBack()
            except Exception:
                pass
            raise

    def _apply_pending_operations(self):
        """Apply all pending operations in sequence order."""
        pending = sorted(list(self._iter_pending_operations()), key=lambda x: int(x.get("seq", 0)))
        if not pending:
            self._set_status("No pending staged actions to apply.")
            return
        LOGGER.info("Apply requested. pending_sequences=%s", len(pending))

        doc = self._active_doc()
        if doc is None:
            self._set_status("No active Revit document.")
            return

        outer_tg = DB.TransactionGroup(doc, "Batch Swap Circuits - Apply")
        success_count = 0
        failed_seq = 0
        failed_msg = ""
        executed_failure = False

        try:
            outer_tg.Start()
            for operation in pending:
                seq = int(operation.get("seq", 0) or 0)
                try:
                    seq_warnings = self._apply_sequence_operation(doc, operation) or []
                    if seq_warnings:
                        self._set_operation_status(operation, "warning", "Applied with warnings: {0}".format(" | ".join([str(x) for x in seq_warnings])))
                    else:
                        self._set_operation_status(operation, "completed", "")
                    success_count += 1
                except Exception as ex:
                    failed_seq = int(seq)
                    failed_msg = str(ex or "Sequence failed.")
                    LOGGER.warning("Apply failed at sequence #{0}: {1}".format(int(failed_seq), failed_msg))
                    self._set_operation_status(operation, "failed", failed_msg)
                    executed_failure = True
                    break

            if executed_failure:
                mark_remaining = False
                for operation in pending:
                    seq = int(operation.get("seq", 0) or 0)
                    if seq == failed_seq:
                        mark_remaining = True
                        continue
                    if mark_remaining and str(operation.get("status", "pending")).lower() == "pending":
                        self._set_operation_status(operation, "failed", "Not executed after failure at sequence #{0}.".format(failed_seq))

            if success_count > 0:
                outer_tg.Assimilate()
            else:
                outer_tg.RollBack()
        except Exception as ex:
            try:
                outer_tg.RollBack()
            except Exception:
                pass
            forms.alert(
                "Apply failed unexpectedly.\n\n{0}".format(str(ex)),
                title=TITLE,
            )
            self._rebuild_change_log()
            return

        self._working_rows_by_panel = {}
        self._refresh_model_caches(doc)
        self._reload_from_selected_panels()
        self._rebuild_change_log()

        if executed_failure:
            forms.alert(
                "Apply stopped at sequence #{0}.\n\nSuccessfully applied: {1}\nFailure: {2}".format(
                    int(failed_seq),
                    int(success_count),
                    failed_msg or "Unknown error.",
                ),
                title=TITLE,
            )
            self._set_status(
                "Apply stopped at sequence #{0}. Applied {1} sequence(s).".format(
                    int(failed_seq), int(success_count)
                )
            )
        else:
            self._set_status("Apply completed. Applied {0} sequence(s).".format(int(success_count)))

    def _stage_add_special(self, kind):
        """Stage adding spare/space rows into selected empty slots."""
        poles = int(max(1, self._selected_add_poles()))
        kind_value = SpecialKind.normalize(kind, None)
        if kind_value is None:
            self._set_status("Invalid special kind: {0}".format(str(kind or "")))
            return
        spare_rating = 0
        spare_frame = 0
        if kind_value == SpecialKind.SPARE:
            rating = self._selected_spare_rating()
            if rating is None:
                self._set_status("Add SPARE cancelled.")
                return
            spare_rating = int(rating)
            frame = self._selected_spare_frame()
            if frame is None:
                self._set_status("Add SPARE cancelled.")
                return
            spare_frame = int(frame)
        targets = []
        for list_ctrl, rows, option in (
            (self.LeftRowsList, self._left_rows, self._left_option),
            (self.RightRowsList, self._right_rows, self._right_option),
        ):
            for row in self._selected_rows_from_list(list_ctrl, rows):
                if str(row.get("kind", "") or "") != "empty":
                    continue
                if not bool(row.get("is_editable", True)):
                    continue
                targets.append((rows, option, row))
        if not targets:
            self._set_status("Select one or more EMPTY rows to add {0}.".format(str(kind_value).upper()))
            return

        touched_panels = set([int(option.get("panel_id", 0) or 0) for _, option, _ in targets])
        before = self._snapshot_panels(touched_panels)
        placements = []
        rejected = 0

        for rows, option, row in sorted(targets, key=lambda x: (int(x[1].get("panel_id", 0) or 0), int(x[2].get("slot", 0) or 0))):
            is_data_panel = False
            try:
                is_data_panel = bool((option or {}).get("schedule_type") == ps_repo.PSTYPE_DATA)
            except Exception:
                is_data_panel = False
            if bool(is_data_panel and kind_value != SpecialKind.SPACE):
                rejected += 1
                continue
            if not self._panel_accepts_add_poles(option, poles):
                rejected += 1
                continue
            slot_value = int(row.get("slot", 0) or 0)
            covered = ps_repo.get_slot_span_slots_for_option(
                option,
                start_slot=int(slot_value),
                pole_count=int(poles),
                require_valid=True,
            )
            if not covered:
                rejected += 1
                continue
            occupancy = self._slot_occupancy(rows, option)
            valid = True
            for slot in covered:
                occ = occupancy.get(int(slot))
                if occ is None or str(occ.get("kind", "")) != "empty":
                    valid = False
                    break
            if not valid:
                rejected += 1
                continue

            for occ_slot in covered:
                for existing in list(rows):
                    if str(existing.get("kind", "")) == "empty" and int(existing.get("slot", 0) or 0) == int(occ_slot):
                        self._remove_row_by_key(rows, existing.get("row_key"))
                        break

            staged = self._build_staged_special_row(option, slot_value, kind_value, spare_rating=spare_rating)
            staged["poles"] = int(poles)
            staged["spare_frame"] = int(spare_frame)
            staged["span"] = int(max(1, len(covered)))
            staged["covered_slots"] = [int(x) for x in covered]
            staged["circuit_number"] = ps_repo.predict_circuit_number(option, slot_value, poles=poles)
            rows.append(staged)
            placements.append(
                {
                    "action": StagedAction.ADD_SPARE if kind_value == SpecialKind.SPARE else StagedAction.ADD_SPACE,
                    "circuit_id": int(staged.get("circuit_id", 0)),
                    "circuit_number": staged.get("circuit_number", ""),
                    "load_name": staged.get("load_name", ""),
                    "target_condition": "empty",
                    "kind": str(staged.get("kind", "") or ""),
                    "is_regular_circuit": False,
                    "poles": int(max(1, staged.get("poles", 1) or 1)),
                    "slot_group_number": 0,
                    "from_panel_id": int(option.get("panel_id", 0) or 0),
                    "from_panel_name": option.get("panel_name", ""),
                    "to_panel_id": int(option.get("panel_id", 0) or 0),
                    "to_panel_name": option.get("panel_name", ""),
                    "old_slot": 0,
                    "old_covered_slots": [],
                    "new_slot": int(slot_value),
                    "new_covered_slots": [int(x) for x in covered],
                    "spare_rating": int(spare_rating),
                    "spare_frame": int(spare_frame),
                    "same_panel": True,
                }
            )

        for panel_id in touched_panels:
            item = self._panel_option_by_id.get(int(panel_id))
            if item is None:
                continue
            option = getattr(item, "option", None)
            rows = self._working_rows_by_panel.get(int(panel_id))
            if option is not None and rows is not None:
                self._normalize_rows(rows, option)

        self._record_operation(placements, before)
        self._recompute_transferability()
        self._refresh_row_views()
        if rejected:
            self._set_status(
                "Added {0} staged {1} row(s). {2} request(s) rejected (insufficient slot fit).".format(
                    len(placements), str(kind_value).upper(), int(rejected)
                )
            )
        else:
            self._set_status("Added {0} staged {1} row(s).".format(len(placements), str(kind_value).upper()))

    def add_spare_clicked(self, sender, args):
        """Add staged SPARE rows to selected empty slots."""
        self._stage_add_special(SpecialKind.SPARE)

    def add_space_clicked(self, sender, args):
        """Add staged SPACE rows to selected empty slots."""
        self._stage_add_special(SpecialKind.SPACE)

    def remove_special_clicked(self, sender, args):
        """Remove selected spare/space rows regardless of removable status."""
        targets = []
        for list_ctrl, rows, option in (
            (self.LeftRowsList, self._left_rows, self._left_option),
            (self.RightRowsList, self._right_rows, self._right_option),
        ):
            for row in self._selected_rows_from_list(list_ctrl, rows):
                kind = str(row.get("kind", "") or "").lower()
                if kind not in SpecialKind.all():
                    continue
                if not bool(row.get("is_editable", True)):
                    continue
                targets.append((rows, option, row))
        if not targets:
            self._set_status("Select one or more SPARE/SPACE rows to remove.")
            return

        touched_panels = set([int(option.get("panel_id", 0) or 0) for _, option, _ in targets])
        before = self._snapshot_panels(touched_panels)
        placements = []

        for rows, option, row in targets:
            kind = str(row.get("kind", "") or "").lower()
            covered = ps_repo.get_row_covered_slots(row, option=option)
            if not covered:
                covered = [int(row.get("slot", 0) or 0)]
            self._remove_row_by_key(rows, row.get("row_key"))
            for slot_value in covered:
                if int(slot_value or 0) <= 0:
                    continue
                rows.append(ps_repo.build_empty_row(option, int(slot_value)))
            placements.append(
                {
                    "action": StagedAction.REMOVE_SPARE if kind == SpecialKind.SPARE else StagedAction.REMOVE_SPACE,
                    "circuit_id": int(row.get("circuit_id", 0) or 0),
                    "circuit_number": row.get("circuit_number", ""),
                    "load_name": row.get("load_name", ""),
                    "target_condition": "",
                    "kind": str(row.get("kind", "") or ""),
                    "is_regular_circuit": False,
                    "poles": int(max(1, row.get("poles", 1) or 1)),
                    "slot_group_number": int(row.get("slot_group_number", 0) or 0),
                    "from_panel_id": int(option.get("panel_id", 0) or 0),
                    "from_panel_name": option.get("panel_name", ""),
                    "to_panel_id": int(option.get("panel_id", 0) or 0),
                    "to_panel_name": option.get("panel_name", ""),
                    "old_slot": int(row.get("slot", 0) or 0),
                    "old_covered_slots": [int(x) for x in covered],
                    "new_slot": 0,
                    "new_covered_slots": [],
                    "leave_unlocked": True,
                    "same_panel": True,
                }
            )

        for panel_id in touched_panels:
            item = self._panel_option_by_id.get(int(panel_id))
            if item is None:
                continue
            option = getattr(item, "option", None)
            rows = self._working_rows_by_panel.get(int(panel_id))
            if option is not None and rows is not None:
                self._normalize_rows(rows, option)

        self._record_operation(placements, before)
        self._recompute_transferability()
        self._refresh_row_views()
        self._set_status("Removed {0} staged SPARE/SPACE row(s).".format(len(placements)))

    def close_clicked(self, sender, args):
        """Close tool window."""
        self.Close()

    def list_preview_mouse_left_button_down(self, sender, args):
        """Track drag start point."""
        self._drag_started = False
        try:
            self._drag_start = args.GetPosition(sender)
        except Exception:
            self._drag_start = None
        self._drag_payload = None
        self._pending_click_selection = None

        clicked = self._drop_target_row(sender, args)
        if clicked is None:
            return
        clicked_key = str(getattr(clicked, "row_key", "") or "")
        if not clicked_key:
            return

        selected_items = []
        try:
            selected_items = list(sender.SelectedItems or [])
        except Exception:
            selected_items = []
        selected_keys = set([str(getattr(x, "row_key", "") or "") for x in selected_items])

        if len(selected_keys) > 1 and clicked_key in selected_keys:
            self._pending_click_selection = {
                "list_name": str(getattr(sender, "Name", "") or ""),
                "row_key": clicked_key,
            }
            try:
                args.Handled = True
            except Exception:
                pass

    def list_preview_mouse_left_button_up(self, sender, args):
        """Collapse preserved multi-selection to one row on click-release."""
        pending = dict(self._pending_click_selection or {})
        self._pending_click_selection = None
        if not pending:
            return
        if self._drag_started:
            return
        if str(getattr(sender, "Name", "") or "") != str(pending.get("list_name", "")):
            return
        target_key = str(pending.get("row_key", "") or "")
        if not target_key:
            return

        items = []
        try:
            items = list(sender.ItemsSource or [])
        except Exception:
            items = []
        target_item = None
        for item in items:
            if str(getattr(item, "row_key", "") or "") == target_key:
                target_item = item
                break
        if target_item is None:
            return

        try:
            sender.SelectedItems.Clear()
            sender.SelectedItem = target_item
        except Exception:
            pass

    def list_mouse_move(self, sender, args):
        """Begin drag operation for selected rows."""
        if args.LeftButton != MouseButtonState.Pressed:
            return
        if self._drag_start is None:
            return
        try:
            current = args.GetPosition(sender)
            dx = Math.Abs(current.X - self._drag_start.X)
            dy = Math.Abs(current.Y - self._drag_start.Y)
            if dx < 4 and dy < 4:
                return
        except Exception:
            return

        selected = []
        try:
            selected = list(sender.SelectedItems or [])
        except Exception:
            selected = []
        draggable = [x for x in selected if bool(getattr(x, "is_draggable", False))]
        if not draggable:
            return

        keys = [str(getattr(x, "row_key", "")) for x in draggable if getattr(x, "row_key", None)]
        if not keys:
            return
        keys = list(dict.fromkeys(keys))
        self._drag_started = True
        self._pending_click_selection = None

        source_name = str(getattr(sender, "Name", "") or "")
        self._drag_payload = {
            "source": source_name,
            "keys": keys,
        }
        try:
            data = DataObject("CED.BatchSwapRows", "rows")
            DragDrop.DoDragDrop(sender, data, DragDropEffects.Move)
        except Exception:
            pass
        self._drag_start = None
        self._drag_payload = None
        self._drag_started = False
        self._clear_preview_slots(refresh=True)

    def list_drag_leave(self, sender, args):
        """Keep preview state stable while drag transitions between lists."""
        return

    def _find_visual_ancestor(self, start, target_type):
        """Return nearest ancestor of requested type."""
        current = start
        while current is not None:
            if isinstance(current, target_type):
                return current
            try:
                current = VisualTreeHelper.GetParent(current)
            except Exception:
                return None
        return None

    def _drop_target_info(self, sender, args):
        """Resolve row and slot under drag/drop pointer using slot-cell hit position."""
        try:
            pos = args.GetPosition(sender)
            hit = sender.InputHitTest(pos)
            item = self._find_visual_ancestor(hit, ListViewItem)
            if item is None:
                return None, 0
            row_item = getattr(item, "DataContext", None)
            if row_item is None:
                return None, 0
            slot_value = int(getattr(row_item, "slot", 0) or 0)
            covered = [int(x) for x in list(getattr(row_item, "covered_slots", []) or []) if int(x) > 0]
            if not covered and slot_value > 0:
                covered = [slot_value]
            if covered:
                try:
                    local = args.GetPosition(item)
                    idx = int(local.Y / 20.0)
                    if idx < 0:
                        idx = 0
                    if idx >= len(covered):
                        idx = len(covered) - 1
                    cell_slot = int(covered[idx] or 0)
                    if cell_slot > 0:
                        slot_value = cell_slot
                except Exception:
                    pass
            return row_item, slot_value
        except Exception:
            return None, 0

    def _drop_target_row(self, sender, args):
        """Resolve row view-model under drag/drop pointer."""
        row_item, _ = self._drop_target_info(sender, args)
        return row_item

    def list_drag_over(self, sender, args):
        """Compute live drop preview and allowed effect."""
        can_drop = self._update_drop_preview(sender, args)
        args.Effects = DragDropEffects.Move if can_drop else getattr(DragDropEffects, "None")
        args.Handled = True

    def _update_drop_preview(self, sender, args):
        """Update target slot preview highlights for current drag."""
        payload = self._drag_payload or {}
        source_name = str(payload.get("source", "") or "")
        keys = list(payload.get("keys") or [])
        target_name = str(getattr(sender, "Name", "") or "")
        if not source_name or not target_name or not keys:
            self._clear_preview_slots(refresh=True)
            return False

        source_rows, source_option, _ = self._resolve_rows_and_option(source_name)
        target_rows, target_option, target_side = self._resolve_rows_and_option(target_name)
        if source_rows is None or target_rows is None or source_option is None or target_option is None:
            self._clear_preview_slots(refresh=True)
            return False

        target_row, target_slot = self._drop_target_info(sender, args)
        if target_row is None:
            return bool(self._left_preview_slots or self._right_preview_slots)
        if target_slot <= 0:
            return bool(self._left_preview_slots or self._right_preview_slots)
        allow_discard = bool(getattr(self.AllowDiscardToggle, "IsChecked", False))
        maintain_groups = bool(getattr(self.MaintainGroupToggle, "IsChecked", False))
        absolute_only = True
        same_list = bool(source_name == target_name)
        require_transferable = not same_list
        expanded_keys = self._expand_moving_keys_for_groups(
            source_rows,
            keys,
            maintain_groups,
            source_option=source_option,
        )
        self._preview_moving_keys = set([str(x) for x in list(expanded_keys or []) if str(x or "").strip()])
        if same_list:
            source_map = {str(r.get("row_key", "") or ""): r for r in list(source_rows or [])}
            selected_start_slots = set()
            for key in expanded_keys:
                row = source_map.get(str(key))
                if row is None:
                    continue
                selected_start_slots.add(int(row.get("slot", 0) or 0))
            if int(target_slot or 0) <= 0 or int(target_slot) in selected_start_slots:
                self._clear_preview_slots(refresh=True)
                return False

        signature = (
            source_name,
            target_name,
            tuple(sorted([str(x) for x in expanded_keys])),
            int(target_slot),
            bool(allow_discard),
            bool(absolute_only),
            bool(maintain_groups),
        )
        if signature == self._preview_signature:
            if target_side == "left":
                return bool(self._left_preview_slots)
            if target_side == "right":
                return bool(self._right_preview_slots)
            return False

        source_clone = self._clone_rows(source_rows)
        if same_list:
            target_clone = source_clone
        else:
            target_clone = self._clone_rows(target_rows)

        result = self._stage_transfer(
            source_rows=source_clone,
            source_option=source_option,
            target_rows=target_clone,
            target_option=target_option,
            moving_keys=expanded_keys,
            allow_discard=allow_discard,
            preferred_slot=target_slot,
            absolute_only=absolute_only,
            all_or_nothing=absolute_only,
            require_transferable=require_transferable,
            record_history=False,
        )
        preview_slots = set()
        for move in list(result.get("placements") or []):
            for slot in list(move.get("new_covered_slots") or []):
                preview_slots.add(int(slot))

        self._preview_signature = signature
        if target_side == "left":
            changed = (preview_slots != self._left_preview_slots) or bool(self._right_preview_slots)
            self._left_preview_slots = preview_slots
            self._right_preview_slots = set()
            if changed:
                self._refresh_row_views()
        elif target_side == "right":
            changed = (preview_slots != self._right_preview_slots) or bool(self._left_preview_slots)
            self._right_preview_slots = preview_slots
            self._left_preview_slots = set()
            if changed:
                self._refresh_row_views()
        else:
            self._clear_preview_slots(refresh=True)

        if absolute_only and int(result.get("moved", 0) or 0) < int(result.get("considered", 0) or 0):
            return False
        return bool(preview_slots) and int(result.get("moved", 0) or 0) > 0

    def list_drop(self, sender, args):
        """Apply staged drag/drop changes to working row state."""
        payload = self._drag_payload or {}
        source_name = str(payload.get("source", "") or "")
        keys = list(payload.get("keys") or [])
        target_name = str(getattr(sender, "Name", "") or "")
        target_row, target_slot = self._drop_target_info(sender, args)
        if target_row is None:
            self._clear_preview_slots(refresh=True)
            self._set_status("Drop ignored: target slot not resolved.")
            return
        if target_slot <= 0:
            self._clear_preview_slots(refresh=True)
            self._set_status("Drop ignored: invalid target slot.")
            return

        if not source_name or not target_name or not keys:
            self._clear_preview_slots(refresh=True)
            return

        source_rows, source_option, _ = self._resolve_rows_and_option(source_name)
        target_rows, target_option, _ = self._resolve_rows_and_option(target_name)
        if source_rows is None or target_rows is None or source_option is None or target_option is None:
            self._clear_preview_slots(refresh=True)
            return

        allow_discard = bool(getattr(self.AllowDiscardToggle, "IsChecked", False))
        maintain_groups = bool(getattr(self.MaintainGroupToggle, "IsChecked", False))
        absolute_only = True
        same_list = bool(source_name == target_name)
        require_transferable = not same_list
        expanded_keys = self._expand_moving_keys_for_groups(
            source_rows,
            keys,
            maintain_groups,
            source_option=source_option,
        )
        if same_list:
            source_map = {str(r.get("row_key", "") or ""): r for r in list(source_rows or [])}
            selected_start_slots = set()
            for key in expanded_keys:
                row = source_map.get(str(key))
                if row is None:
                    continue
                selected_start_slots.add(int(row.get("slot", 0) or 0))
            if int(target_slot or 0) <= 0 or int(target_slot) in selected_start_slots:
                self._clear_preview_slots(refresh=True)
                self._set_status("Drop ignored: selection returned to original position.")
                return

        result = self._stage_transfer(
            source_rows=source_rows,
            source_option=source_option,
            target_rows=target_rows,
            target_option=target_option,
            moving_keys=expanded_keys,
            allow_discard=allow_discard,
            preferred_slot=target_slot,
            absolute_only=absolute_only,
            all_or_nothing=absolute_only,
            require_transferable=require_transferable,
            record_history=True,
        )

        moved = int(result.get("moved", 0) or 0)
        rejected = int(result.get("rejected", 0) or 0)
        self._clear_preview_slots(refresh=False)
        self._recompute_transferability()
        self._refresh_row_views()
        self._drag_payload = None
        self._drag_started = False

        if moved <= 0:
            self._set_status("Drop blocked: selected circuits do not fit at the exact target slot.")
            return

        if rejected > 0:
            self._set_status("Staged {0} circuit(s). {1} rejected (exact target required).".format(moved, rejected))
        else:
            self._set_status("Staged {0} circuit(s).".format(moved))

    def _remove_row_by_key(self, rows, row_key):
        """Remove first row with matching row key."""
        target_key = str(row_key or "")
        for idx, row in enumerate(list(rows or [])):
            if str(row.get("row_key", "")) == target_key:
                try:
                    rows.pop(idx)
                except Exception:
                    pass
                return True
        return False

    def _next_slot_after(self, option, slot):
        """Return next slot in panel display order."""
        slot_order = ps_repo.get_option_slot_order(option, include_excess=False)
        current = int(slot or 0)
        if current not in slot_order:
            return 0
        idx = slot_order.index(current)
        if idx + 1 >= len(slot_order):
            return 0
        return int(slot_order[idx + 1])

    def _can_discard_row(self, row, allow_discard):
        """Return True when overlapping row can be discarded."""
        if not allow_discard:
            return False
        kind = str(row.get("kind", "empty"))
        if kind == "spare":
            return bool(row.get("is_spare_removable", False))
        if kind == "space":
            return bool(row.get("is_space_removable", False))
        return False

    def _clear_target_range(self, target_rows, covered_slots, allow_discard, target_option):
        """Clear overlapping rows in target slot range when allowed."""
        target_slots = set([int(x) for x in list(covered_slots or []) if int(x) > 0])
        if not target_slots:
            return []
        kept = []
        removed = []
        for row in list(target_rows or []):
            row_slots = set(ps_repo.get_row_covered_slots(row, option=target_option))
            overlaps = bool(row_slots.intersection(target_slots))
            if not overlaps:
                kept.append(row)
                continue
            kind = str(row.get("kind", "empty"))
            if kind == "empty":
                continue
            if self._can_discard_row(row, allow_discard):
                removed.append(copy.copy(row))
                continue
            kept.append(row)
        target_rows[:] = kept
        return removed

    def _slot_occupancy(self, rows, option):
        """Map slot -> row occupying that slot."""
        occupancy = {}
        for row in list(rows or []):
            covered_slots = ps_repo.get_row_covered_slots(row, option=option)
            if not covered_slots:
                covered_slots = [int(row.get("slot", 0) or 0)]
            for covered in covered_slots:
                if int(covered or 0) <= 0:
                    continue
                occupancy[int(covered)] = row
        return occupancy

    def _target_condition_label(self, target_rows, target_option, covered_slots):
        """Return summary label for current target occupancy in covered slots."""
        occupancy = self._slot_occupancy(target_rows, target_option)
        kinds = []
        for slot in [int(x) for x in list(covered_slots or []) if int(x) > 0]:
            row = occupancy.get(int(slot))
            kind = str((row or {}).get("kind", "unknown") or "unknown").lower()
            kinds.append(kind)
        if not kinds:
            return "unknown"
        uniq = sorted(set(kinds))
        if len(uniq) == 1:
            return uniq[0]
        if "circuit" in uniq:
            return "occupied"
        if "spare" in uniq and "space" in uniq:
            return "mixed spare/space"
        if "spare" in uniq and "empty" in uniq:
            return "mixed empty/spare"
        if "space" in uniq and "empty" in uniq:
            return "mixed empty/space"
        return "mixed"

    def _find_fit_slot_for_row(
        self,
        target_rows,
        target_option,
        moving_row,
        allow_discard,
        preferred_slot=0,
        absolute_only=False,
        temporary_row_keys=None,
    ):
        """Find first slot where moving row can fit."""
        slot_order = ps_repo.get_option_slot_order(target_option, include_excess=False)
        occupancy = self._slot_occupancy(target_rows, target_option)

        preferred = int(preferred_slot or 0)
        ordered_starts = list(slot_order)
        if absolute_only:
            if preferred not in ordered_starts:
                return None
            ordered_starts = [preferred]
        elif preferred in ordered_starts:
            idx = ordered_starts.index(preferred)
            ordered_starts = ordered_starts[idx:] + ordered_starts[:idx]

        poles = int(max(1, moving_row.get("poles", 1) or 1))
        for start in ordered_starts:
            covered_slots = ps_repo.get_slot_span_slots_for_option(
                target_option,
                start_slot=int(start),
                pole_count=int(poles),
                require_valid=True,
            )
            if not covered_slots:
                continue
            ok = True
            for slot in covered_slots:
                row = occupancy.get(slot)
                if row is None:
                    ok = False
                    break
                if temporary_row_keys and str(row.get("row_key", "") or "") in temporary_row_keys:
                    continue
                kind = str(row.get("kind", "empty"))
                if kind == "empty":
                    continue
                if self._can_discard_row(row, allow_discard):
                    continue
                ok = False
                break
            if ok:
                return int(start)
        return None

    def _normalize_rows(self, rows, option):
        """Normalize rows into unique occupants + explicit empties."""
        slot_order = ps_repo.get_option_slot_order(option, include_excess=True)
        slot_set = set(slot_order)

        occupants = []
        seen_slots = set()
        occupied_slots = set()
        for row in sorted(list(rows or []), key=lambda x: int(x.get("slot", 0))):
            covered_slots = [x for x in ps_repo.get_row_covered_slots(row, option=option) if x in slot_set]
            slot = int(covered_slots[0]) if covered_slots else int(row.get("slot", 0) or 0)
            kind = str(row.get("kind", "empty"))
            if slot not in slot_set:
                continue
            if kind == "empty":
                continue
            if slot in seen_slots:
                continue
            if any(int(x) in occupied_slots for x in covered_slots):
                continue
            seen_slots.add(slot)
            row["slot"] = slot
            row["covered_slots"] = list(covered_slots) if covered_slots else [slot]
            row["span"] = int(max(1, len(row["covered_slots"])))
            row["circuit_number"] = ps_repo.predict_circuit_number(
                option,
                row["slot"],
                poles=int(max(1, row.get("poles", 1) or 1)),
            )
            occupants.append(row)
            for covered in row["covered_slots"]:
                if int(covered) in slot_set:
                    occupied_slots.add(int(covered))

        normalized = list(occupants)
        for slot in slot_order:
            if slot in occupied_slots:
                continue
            normalized.append(ps_repo.build_empty_row(option, slot))

        order_index = {int(slot): i for i, slot in enumerate(slot_order)}
        normalized.sort(key=lambda x: (order_index.get(int(x.get("slot", 0)), 999999), int(x.get("slot", 0))))
        rows[:] = normalized

    def _stage_transfer(
        self,
        source_rows,
        source_option,
        target_rows,
        target_option,
        moving_keys,
        allow_discard,
        preferred_slot=0,
        absolute_only=False,
        all_or_nothing=False,
        require_transferable=True,
        record_history=True,
    ):
        """Stage transfer/reorder operation and return result summary."""
        source_panel_id = int((source_option or {}).get("panel_id", 0) or 0)
        target_panel_id = int((target_option or {}).get("panel_id", 0) or 0)
        same_panel = bool(source_panel_id == target_panel_id and source_rows is target_rows)

        before_snapshots = {}
        if record_history:
            before_snapshots = self._snapshot_panels([source_panel_id, target_panel_id])

        moving_rows = []
        source_map = {}
        for row in list(source_rows or []):
            source_map[str(row.get("row_key", ""))] = row
        for key in list(moving_keys or []):
            row = source_map.get(str(key))
            if row is None:
                continue
            if not self._is_movable_row(row):
                continue
            if require_transferable and str(row.get("kind", "")).lower() == "circuit" and not bool(row.get("transferable", False)):
                continue
            moving_rows.append(row)

        slot_order = ps_repo.get_option_slot_order(source_option, include_excess=True)
        order_index = {int(slot): idx for idx, slot in enumerate(list(slot_order or []))}
        moving_rows = sorted(
            moving_rows,
            key=lambda x: (
                order_index.get(int(x.get("slot", 0) or 0), 999999),
                int(x.get("slot", 0) or 0),
            ),
        )
        moving_key_set = set([str(row.get("row_key", "") or "") for row in list(moving_rows or [])])
        if not moving_rows:
            return {"moved": 0, "rejected": len(moving_keys or []), "placements": [], "considered": 0}
        if bool(all_or_nothing) and int(len(moving_rows)) < int(len(list(moving_keys or []))):
            return {
                "moved": 0,
                "rejected": int(len(list(moving_keys or []))),
                "placements": [],
                "considered": int(len(list(moving_keys or []))),
            }

        if bool(all_or_nothing):
            source_probe = self._clone_rows(source_rows)
            if source_rows is target_rows:
                target_probe = source_probe
            else:
                target_probe = self._clone_rows(target_rows)
            probe = self._stage_transfer(
                source_rows=source_probe,
                source_option=source_option,
                target_rows=target_probe,
                target_option=target_option,
                moving_keys=list(moving_keys or []),
                allow_discard=allow_discard,
                preferred_slot=preferred_slot,
                absolute_only=absolute_only,
                all_or_nothing=False,
                require_transferable=require_transferable,
                record_history=False,
            )
            considered = int(probe.get("considered", 0) or 0)
            moved_probe = int(probe.get("moved", 0) or 0)
            if considered <= 0 or moved_probe < considered:
                return {
                    "moved": 0,
                    "rejected": int(max(considered, len(moving_rows))),
                    "placements": [],
                    "considered": int(max(considered, len(moving_rows))),
                }

        moved_count = 0
        rejected_count = 0
        next_preferred_slot = int(preferred_slot or 0)
        placements = []

        for moving in moving_rows:
            fit_slot = self._find_fit_slot_for_row(
                target_rows=target_rows,
                target_option=target_option,
                moving_row=moving,
                allow_discard=allow_discard,
                preferred_slot=next_preferred_slot,
                absolute_only=absolute_only,
                temporary_row_keys=moving_key_set if same_panel else None,
            )
            if fit_slot is None:
                rejected_count += 1
                continue

            old_slot = int(moving.get("slot", 0) or 0)
            old_covered = ps_repo.get_row_covered_slots(moving, option=source_option)
            if not old_covered:
                old_covered = [old_slot] if old_slot > 0 else []

            self._remove_row_by_key(source_rows, moving.get("row_key"))
            for free_slot in old_covered:
                if int(free_slot or 0) <= 0:
                    continue
                source_rows.append(ps_repo.build_empty_row(source_option, free_slot))

            target_covered = ps_repo.get_slot_span_slots_for_option(
                target_option,
                start_slot=int(fit_slot),
                pole_count=int(max(1, moving.get("poles", 1) or 1)),
                require_valid=True,
            ) or [int(fit_slot)]
            target_condition = self._target_condition_label(
                target_rows=target_rows,
                target_option=target_option,
                covered_slots=target_covered,
            )

            removed_rows = self._clear_target_range(
                target_rows=target_rows,
                covered_slots=target_covered,
                allow_discard=allow_discard,
                target_option=target_option,
            )
            for removed in list(removed_rows or []):
                removed_kind = str(removed.get("kind", "") or "").strip().lower()
                if removed_kind not in SpecialKind.all():
                    continue
                removed_slots = ps_repo.get_row_covered_slots(removed, option=target_option)
                if not removed_slots:
                    removed_slot = int(removed.get("slot", 0) or 0)
                    if removed_slot > 0:
                        removed_slots = [removed_slot]
                placements.append(
                    {
                        "action": StagedAction.REMOVE_SPARE if removed_kind == SpecialKind.SPARE else StagedAction.REMOVE_SPACE,
                        "circuit_id": int(removed.get("circuit_id", 0) or 0),
                        "for_circuit_id": int(moving.get("circuit_id", 0) or 0),
                        "circuit_number": str(removed.get("circuit_number", "") or ""),
                        "load_name": str(removed.get("load_name", "") or removed_kind.upper()),
                        "target_condition": "replaced by move",
                        "kind": removed_kind,
                        "is_regular_circuit": False,
                        "poles": int(max(1, removed.get("poles", 1) or 1)),
                        "slot_group_number": int(removed.get("slot_group_number", 0) or 0),
                        "from_panel_id": target_panel_id,
                        "from_panel_name": str(target_option.get("panel_name", "") or ""),
                        "to_panel_id": target_panel_id,
                        "to_panel_name": str(target_option.get("panel_name", "") or ""),
                        "old_slot": int(removed_slots[0]) if removed_slots else int(removed.get("slot", 0) or 0),
                        "old_covered_slots": [int(x) for x in list(removed_slots or []) if int(x) > 0],
                        "new_slot": 0,
                        "new_covered_slots": [],
                        "leave_unlocked": True,
                        "same_panel": True,
                    }
                )

            placed = copy.copy(moving)
            placed["slot"] = int(target_covered[0])
            placed["covered_slots"] = [int(x) for x in target_covered]
            placed["span"] = int(max(1, len(target_covered)))
            placed["panel_id"] = target_panel_id
            placed["panel_name"] = target_option.get("panel_name", "")
            placed["circuit_number"] = ps_repo.predict_circuit_number(
                target_option,
                placed["slot"],
                poles=int(max(1, placed.get("poles", 1) or 1)),
            )
            placed["row_key"] = "panel:{0}|slot:{1}|ckt:{2}|seq:{3}".format(
                placed["panel_id"],
                placed["slot"],
                int(placed.get("circuit_id", 0)),
                moved_count + 1,
            )
            target_rows.append(placed)
            moved_count += 1
            next_preferred_slot = self._next_slot_after(target_option, placed["covered_slots"][-1])

            placements.append(
                {
                    "action": StagedAction.MOVE,
                    "circuit_id": int(placed.get("circuit_id", 0) or 0),
                    "circuit_number": str(moving.get("circuit_number", "") or ""),
                    "load_name": str(moving.get("load_name", "") or ""),
                    "target_condition": str(target_condition or ""),
                    "kind": str(moving.get("kind", "") or ""),
                    "is_regular_circuit": bool(moving.get("is_regular_circuit", False)),
                    "poles": int(max(1, moving.get("poles", 1) or 1)),
                    "slot_group_number": int(moving.get("slot_group_number", 0) or 0),
                    "from_panel_id": source_panel_id,
                    "from_panel_name": str(source_option.get("panel_name", "") or ""),
                    "to_panel_id": target_panel_id,
                    "to_panel_name": str(target_option.get("panel_name", "") or ""),
                    "old_slot": int(old_slot),
                    "old_covered_slots": [int(x) for x in old_covered],
                    "new_slot": int(placed["slot"]),
                    "new_covered_slots": [int(x) for x in placed["covered_slots"]],
                    "same_panel": bool(same_panel),
                }
            )

        if source_rows is target_rows:
            self._normalize_rows(source_rows, source_option)
        else:
            self._normalize_rows(source_rows, source_option)
            self._normalize_rows(target_rows, target_option)

        total_rejected = rejected_count + max(0, (len(moving_keys or []) - len(moving_rows)))
        result = {
            "moved": moved_count,
            "rejected": total_rejected,
            "placements": placements,
            "considered": int(len(moving_rows)),
        }

        if record_history and moved_count > 0:
            self._record_operation(placements, before_snapshots)
        return result


def _show_modal():
    """Show BatchSwap window modally to keep a valid API context."""
    theme_mode, accent_mode = _load_theme_state("light", "blue")
    window = BatchSwapWindow(theme_mode=theme_mode, accent_mode=accent_mode)
    try:
        window.ShowDialog()
    except Exception:
        window.Show()
    try:
        window.Activate()
    except Exception:
        pass


if __name__ == "__main__":
    _show_modal()
