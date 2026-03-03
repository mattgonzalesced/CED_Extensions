# -*- coding: utf-8 -*-

import imp
import json
import os
import sys

import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
from System import EventHandler
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility
from System.Windows.Controls import ContextMenu, MenuItem, Separator
from System.Windows.Media import BrushConverter
from pyrevit import forms, revit, DB, script

LIB_ROOT = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "..",
        "..",
        "..",
        "CEDLib.lib",
    )
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from CEDElectrical.Model.alerts import get_alert_definition
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Domain import settings_manager
from CEDElectrical.refdata.standard_ocp_table import BREAKER_FRAME_SWITCH_TABLE
from CEDElectrical.Infrastructure.Revit.external_events.circuit_operation_event import (
    CircuitOperationExternalEventGateway,
)


TITLE = "Circuit Browser"
PANEL_ID = "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2"
ALERT_DATA_PARAM = "Circuit Data_CED"
CALC_SETTINGS_PATH = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "Circuits2.stack",
        "Calculate Circuits.pushbutton",
        "config.py",
    )
)
CALC_SETTINGS_XAML_PATH = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "Circuits2.stack",
        "Calculate Circuits.pushbutton",
        "settings.xaml",
    )
)
HIDABLE_ALERT_IDS = {
    "Design.NonStandardOCPRating",
    "Design.BreakerLugSizeLimitOverride",
    "Design.BreakerLugQuantityLimitOverride",
    "Calculations.BreakerLugSizeLimit",
    "Calculations.BreakerLugQuantityLimit",
}
BLOCKED_BRANCH_TYPES = set(['N/A', 'SPACE', 'SPARE', 'CONDUIT ONLY'])
IG_BREAKER_ALLOWED_TYPES = set(['BRANCH', 'FEEDER', 'XFMR PRI', 'XFMR SEC'])

CIRCUIT_TYPE_TAG_STYLES = {
    "BRANCH": ("#DDEBFF", "#234A84"),
    "FEEDER": ("#D5F1E0", "#155C38"),
    "SPARE": ("#EAEAEA", "#444444"),
    "SPACE": ("#ECEFF3", "#52606D"),
    "XFMR PRI": ("#FFE6CC", "#8A4B08"),
    "XFMR SEC": ("#F7DEFF", "#6F2D86"),
    "CONDUIT ONLY": ("#E6F8FA", "#0E5966"),
    "N/A": ("#F0F0F0", "#666666"),
}


def _fmt_number(value, digits=1):
    try:
        return str(round(float(value), digits))
    except Exception:
        return "-"


def _lookup_param_value(element, name):
    try:
        param = element.LookupParameter(name)
        if not param:
            return None
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString()
        if st == DB.StorageType.Integer:
            return param.AsInteger()
        if st == DB.StorageType.Double:
            return param.AsDouble()
        if st == DB.StorageType.ElementId:
            return param.AsElementId()
    except Exception:
        return None
    return None


def _lookup_param_text(element, name):
    try:
        param = element.LookupParameter(name)
        if not param:
            return None
        value = param.AsString()
        if value is None:
            value = param.AsValueString()
        return value
    except Exception:
        return None


def _read_alert_payload(circuit):
    raw = _lookup_param_text(circuit, ALERT_DATA_PARAM)
    if not raw:
        return None
    try:
        return json.loads(raw)
    except Exception:
        return None


def _payload_alert_records(payload):
    if not isinstance(payload, dict):
        return []
    alerts = payload.get("alerts")
    if not isinstance(alerts, list):
        return []
    return alerts


def _payload_hidden_ids(payload):
    if not isinstance(payload, dict):
        return set()
    hidden = payload.get("hidden_definition_ids")
    if not isinstance(hidden, list):
        return set()
    return set([x for x in hidden if x])


def _alert_rows_from_payload(payload):
    hidden_ids = _payload_hidden_ids(payload)
    rows = []
    for item in _payload_alert_records(payload):
        if not isinstance(item, dict):
            continue
        definition_id = item.get("definition_id") or item.get("id")
        severity = (item.get("severity") or "NONE").upper()
        group = item.get("group") or "Other"
        definition = get_alert_definition(definition_id) if definition_id else None
        message_value = item.get("message")
        if message_value:
            text = str(message_value)
        elif definition:
            text = definition.GetDescriptionText()
        elif definition_id:
            text = definition_id
        else:
            text = "Unmapped alert"
        rows.append(
            AlertRow(
                severity,
                group,
                definition_id or "-",
                text,
                definition_id in hidden_ids if definition_id else False,
                definition_id in HIDABLE_ALERT_IDS if definition_id else False,
            )
        )
    return rows


def _alert_tooltip_from_payload(payload):
    items = _payload_alert_records(payload)
    if not items:
        return "", "Collapsed", 0, 0
    total = len(items)
    hidden_count = 0
    hidden_ids = _payload_hidden_ids(payload)
    for item in items:
        definition_id = (item or {}).get("definition_id") or (item or {}).get("id")
        if definition_id and definition_id in hidden_ids:
            hidden_count += 1
    return "Alert Count: {} | Hidden Count: {}".format(total, hidden_count), "Visible", total, hidden_count


def _derive_branch_type(circuit):
    if circuit.CircuitType == DBE.CircuitType.Space:
        return "SPACE"
    if circuit.CircuitType == DBE.CircuitType.Spare:
        return "SPARE"
    return "BRANCH"


def _is_yes(value):
    try:
        return int(value) == 1
    except Exception:
        return False


class CircuitListItem(object):
    def __init__(self, circuit):
        self.circuit = circuit
        self.is_checked = False

        self.panel = "No Panel"
        try:
            if circuit.BaseEquipment:
                self.panel = getattr(circuit.BaseEquipment, "Name", self.panel) or self.panel
        except Exception:
            pass

        self.circuit_number = getattr(circuit, "CircuitNumber", "") or ""
        self.load_name = getattr(circuit, "LoadName", "") or ""

        poles = "?"
        try:
            poles = str(getattr(circuit, "PolesNumber", "?") or "?")
        except Exception:
            pass

        rating_value = None
        if circuit.SystemType == DBE.ElectricalSystemType.PowerCircuit:
            try:
                rating_value = int(round(circuit.Rating, 0))
            except Exception:
                rating_value = None
        if rating_value is None:
            rating_value = _lookup_param_value(circuit, "CKT_Rating_CED")

        if circuit.CircuitType == DBE.CircuitType.Space:
            self.rating_poles = "/{}P".format(poles)
        elif rating_value is None:
            self.rating_poles = "- / {}P".format(poles)
        else:
            self.rating_poles = "{}A / {}P".format(int(round(float(rating_value), 0)), poles)

        self.device_line = "# Devices: {}".format(len(list(circuit.Elements or [])))

        load_current = _lookup_param_value(circuit, "Circuit Load Current_CED")
        self.load_line = "Load: {} A".format(_fmt_number(load_current, 1))
        self.load_line_color = "#384450"
        try:
            if rating_value is not None and load_current is not None and float(load_current) > float(rating_value):
                self.load_line_color = "#B32020"
        except Exception:
            self.load_line_color = "#384450"

        branch_type = _lookup_param_value(circuit, "CKT_Circuit Type_CEDT")
        if isinstance(branch_type, str):
            branch_type = branch_type.strip().upper()
        if not branch_type:
            branch_type = _derive_branch_type(circuit)

        self.branch_type = branch_type
        self.branch_type_line = "Circuit Type: {}".format(self.branch_type)

        conduit_wire = _lookup_param_value(circuit, "Conduit and Wire Size_CEDT")
        if conduit_wire is None:
            conduit_wire = "-"
        self.wire_line = "Conduit/Wire: {}".format(conduit_wire)

        tag_bg, tag_fg = CIRCUIT_TYPE_TAG_STYLES.get(self.branch_type, ("#E8EDF3", "#2F4356"))
        self.type_tag_text = self.branch_type
        self.type_tag_bg = tag_bg
        self.type_tag_fg = tag_fg

        user_override = _lookup_param_value(circuit, "CKT_User Override_CED")
        has_override = False
        try:
            has_override = int(user_override or 0) == 1
        except Exception:
            has_override = False
        self.has_override = has_override
        self.card_border_brush = "#7CB7FF" if has_override else "#AAB4C0"
        self.card_border_thickness = "2" if has_override else "1"

        neutral_qty = _lookup_param_value(circuit, "CKT_Wire Neutral Quantity_CED")
        ig_qty = _lookup_param_value(circuit, "CKT_Wire Isolated Ground Quantity_CED")
        self.neutral_badge_visibility = "Visible" if (neutral_qty or 0) > 0 else "Collapsed"
        self.ig_badge_visibility = "Visible" if (ig_qty or 0) > 0 else "Collapsed"

        payload = _read_alert_payload(circuit)
        self.alert_rows = _alert_rows_from_payload(payload)
        self.alert_summary, self.alert_visibility, self.alert_count, self.hidden_alert_count = _alert_tooltip_from_payload(payload)
        self.alert_bg = "#F9C846"
        self.alert_border = "#BFA23A"
        self.alert_text_color = "#5E4A00"
        if self.alert_count > 0 and self.hidden_alert_count == self.alert_count:
            self.alert_bg = "#00FFFFFF"
            self.alert_border = "#8F7300"
            self.alert_text_color = "#6F5700"

        self.search_name = "{} {} {} {} {} {}".format(
            self.panel,
            self.circuit_number,
            self.load_name,
            self.rating_poles,
            self.branch_type,
            conduit_wire,
        ).lower()


class AlertRow(object):
    def __init__(self, severity, group, definition_id, message, is_hidden=False, can_hide=True):
        self.severity = severity
        self.group = group
        self.definition_id = definition_id
        self.message = message
        self.is_hidden = bool(is_hidden)
        self.can_hide = bool(can_hide)


class CircuitAlertsWindow(forms.WPFWindow):
    def __init__(self, circuit_label, rows):
        xaml = os.path.abspath(os.path.join(os.path.dirname(__file__), "CircuitAlertsWindow.xaml"))
        forms.WPFWindow.__init__(self, xaml)
        self._rows = list(rows or [])
        self.updated_hidden_ids = None
        self.Topmost = True

        title_text = self.FindName("CircuitText")
        count_text = self.FindName("CountText")
        active_list = self.FindName("ActiveAlertsList")
        hidden_list = self.FindName("HiddenAlertsList")

        if title_text is not None:
            title_text.Text = circuit_label
        self._count_text = count_text
        self._active_list = active_list
        self._hidden_list = hidden_list
        self._tabs = self.FindName("AlertsTabs")
        self._hide_btn = self.FindName("HideTypeButton")
        self._unhide_btn = self.FindName("UnhideTypeButton")
        if self._tabs is not None:
            self._tabs.SelectionChanged += self.tabs_selection_changed
        self._refresh_lists()
        self._sync_action_buttons()

    def _refresh_lists(self):
        active = [x for x in self._rows if not x.is_hidden]
        hidden = [x for x in self._rows if x.is_hidden]

        if self._count_text is not None:
            self._count_text.Text = "Alerts: {} | Active: {} | Hidden: {}".format(len(self._rows), len(active), len(hidden))
        if self._active_list is not None:
            self._active_list.ItemsSource = ObservableCollection[AlertRow](active)
        if self._hidden_list is not None:
            self._hidden_list.ItemsSource = ObservableCollection[AlertRow](hidden)

    def _sync_action_buttons(self):
        if self._tabs is None:
            return
        selected_index = self._tabs.SelectedIndex
        if self._hide_btn is not None:
            self._hide_btn.Visibility = Visibility.Visible if selected_index == 0 else Visibility.Collapsed
        if self._unhide_btn is not None:
            self._unhide_btn.Visibility = Visibility.Visible if selected_index == 1 else Visibility.Collapsed

    def tabs_selection_changed(self, sender, args):
        self._sync_action_buttons()

    def hide_type_clicked(self, sender, args):
        if self._active_list is None:
            return
        row = getattr(self._active_list, "SelectedItem", None)
        if row is None:
            forms.alert("Select an active alert type first.", title="Circuit Alerts")
            return
        if not row.definition_id or row.definition_id == "-":
            forms.alert("Only mapped alert types can be hidden.", title="Circuit Alerts")
            return
        if not getattr(row, "can_hide", False):
            forms.alert("This alert type can not be hidden.", title="Circuit Alerts")
            return
        for r in self._rows:
            if r.definition_id == row.definition_id:
                r.is_hidden = True
        self._refresh_lists()

    def unhide_type_clicked(self, sender, args):
        if self._hidden_list is None:
            return
        row = getattr(self._hidden_list, "SelectedItem", None)
        if row is None:
            forms.alert("Select a hidden alert type first.", title="Circuit Alerts")
            return
        if not row.definition_id or row.definition_id == "-":
            return
        for r in self._rows:
            if r.definition_id == row.definition_id:
                r.is_hidden = False
        self._refresh_lists()

    def apply_clicked(self, sender, args):
        hidden_ids = sorted(
            list(
                {
                    x.definition_id
                    for x in self._rows
                    if x.is_hidden and x.can_hide and x.definition_id and x.definition_id != "-"
                }
            )
        )
        self.updated_hidden_ids = hidden_ids
        self.Close()

    def close_clicked(self, sender, args):
        self.Close()


class LockedRow(object):
    def __init__(self, circuit, circuit_owner, device_owner):
        self.circuit = circuit
        panel = ""
        number = ""
        try:
            text = str(circuit or "")
            if "-" in text:
                panel, number = text.split("-", 1)
        except Exception:
            panel = ""
            number = ""
        self.panel = panel
        self.number = number
        self.circuit_owner = circuit_owner
        self.device_owner = device_owner


class RuntimeAlertRow(object):
    def __init__(self, panel, number, severity, group, message):
        self.panel = panel
        self.number = number
        self.severity = severity
        self.group = group
        self.message = message


class CircuitRunSummaryWindow(forms.WPFWindow):
    def __init__(self, locked_rows, runtime_rows):
        xaml = os.path.abspath(os.path.join(os.path.dirname(__file__), "CircuitRunSummaryWindow.xaml"))
        forms.WPFWindow.__init__(self, xaml)
        locked = [
            LockedRow(x.get("circuit", ""), x.get("circuit_owner", "") or "-", x.get("device_owner", "") or "-")
            for x in list(locked_rows or [])
        ]
        runtime = [
            RuntimeAlertRow(
                x.get("panel", "") or "",
                x.get("number", "") or "",
                x.get("severity", "") or "",
                x.get("group", "") or "",
                x.get("message", "") or "",
            )
            for x in list(runtime_rows or [])
        ]

        locked_list = self.FindName("LockedList")
        runtime_list = self.FindName("RuntimeList")
        locked_count = self.FindName("LockedCountText")
        runtime_count = self.FindName("RuntimeCountText")

        if locked_list is not None:
            locked_list.ItemsSource = ObservableCollection[LockedRow](locked)
        if runtime_list is not None:
            runtime_list.ItemsSource = ObservableCollection[RuntimeAlertRow](runtime)
        if locked_count is not None:
            locked_count.Text = "Locked items: {}".format(len(locked))
        if runtime_count is not None:
            runtime_count.Text = "Runtime-only alerts: {}".format(len(runtime))

    def close_clicked(self, sender, args):
        self.Close()


class NeutralIGActionRow(object):
    def __init__(self, item, is_enabled, reason, current_qty, current_size, current_wire):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = item.circuit.Id.IntegerValue
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.is_enabled = bool(is_enabled)
        self.is_checked = bool(is_enabled)
        self.reason = reason or ""

        self.current_qty = current_qty
        self.current_size = current_size or ""
        self.current_wire = current_wire or ""

        self.new_qty = current_qty
        self.new_size = current_size or ""
        self.new_wire = "no change"
        self.new_wire_font_style = "Italic"
        self.is_changed = False


class NeutralIGActionWindow(forms.WPFWindow):
    def __init__(self, title, rows, preview_callback, apply_callback):
        xaml = os.path.abspath(os.path.join(os.path.dirname(__file__), "CircuitNeutralIGActionWindow.xaml"))
        forms.WPFWindow.__init__(self, xaml)
        self._rows = list(rows or [])
        self._preview_callback = preview_callback
        self._apply_callback = apply_callback
        self._mode = "add"
        self._is_syncing_checks = False
        self._is_ready = False
        self._show_unsupported = False

        title_tb = self.FindName("TitleText")
        if title_tb is not None:
            title_tb.Text = title

        self._grid = self.FindName("CircuitGrid")
        self._status = self.FindName("ChangedStatusText")
        self._show_unsupported_cb = self.FindName("ShowUnsupportedToggle")
        if self._show_unsupported_cb is not None:
            self._show_unsupported_cb.IsChecked = False
        self._apply_visibility_filter()
        self.Loaded += self.window_loaded
        self._refresh_status()

    def window_loaded(self, sender, args):
        self._is_ready = True

    def _apply_visibility_filter(self):
        if self._grid is None:
            return
        if self._show_unsupported:
            items = list(self._rows)
        else:
            items = [x for x in self._rows if bool(getattr(x, "is_enabled", False))]
        self._grid.ItemsSource = ObservableCollection[object](items)

    def _refresh_grid(self, refresh_items=False):
        if refresh_items:
            self._apply_visibility_filter()
        self._refresh_status()

    def _refresh_status(self):
        changed = len([x for x in self._rows if x.is_changed and x.is_enabled])
        if self._show_unsupported:
            total = len(self._rows)
        else:
            total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
        if self._status is not None:
            self._status.Text = "{} of {} circuits to be modified.".format(changed, total)

    def _apply_checkbox_to_selected(self, sender, state):
        if not self._is_ready or self._is_syncing_checks:
            return
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        targets = [row]
        selected = []
        try:
            selected = list(self._grid.SelectedItems or [])
        except Exception:
            selected = []
        if row in selected and len(selected) > 1:
            targets = selected
        self._is_syncing_checks = True
        try:
            for item in targets:
                if not getattr(item, "is_enabled", False):
                    item.is_checked = False
                else:
                    if bool(getattr(item, "is_checked", False)) != bool(state):
                        item.is_checked = bool(state)
        finally:
            self._is_syncing_checks = False
        self._refresh_status()

    def item_checked(self, sender, args):
        self._apply_checkbox_to_selected(sender, True)

    def item_unchecked(self, sender, args):
        self._apply_checkbox_to_selected(sender, False)

    def check_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = bool(row.is_enabled)
        finally:
            self._is_syncing_checks = False
        self._refresh_grid(refresh_items=True)

    def uncheck_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        self._refresh_grid(refresh_items=True)

    def add_clicked(self, sender, args):
        self._mode = "add"
        self._preview_callback(self._rows, "add")
        self._refresh_grid(refresh_items=True)

    def remove_clicked(self, sender, args):
        self._mode = "remove"
        self._preview_callback(self._rows, "remove")
        self._refresh_grid(refresh_items=True)

    def apply_clicked(self, sender, args):
        if self._apply_callback(self._rows, self._mode):
            self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def show_unsupported_toggled(self, sender, args):
        self._show_unsupported = bool(getattr(sender, "IsChecked", False))
        self._refresh_grid(refresh_items=True)


class BreakerActionRow(object):
    def __init__(self, item, is_enabled, reason, load_current, cur_rating, cur_frame, auto_rating, auto_frame_from_cur):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = item.circuit.Id.IntegerValue
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.is_enabled = bool(is_enabled)
        self.is_checked = bool(is_enabled)
        self.reason = reason or ""

        self._load_current_value = load_current
        self._current_rating_value = cur_rating
        self._current_frame_value = cur_frame
        self._autosized_rating_value = auto_rating
        self.current_load_current = _fmt_number(load_current, 2)
        self.current_rating = "-" if cur_rating is None else str(int(round(float(cur_rating), 0)))
        self.current_frame = "-" if cur_frame is None else str(int(round(float(cur_frame), 0)))
        self.autosized_rating = "-" if auto_rating is None else str(int(round(float(auto_rating), 0)))
        self.autosized_frame = "-" if auto_frame_from_cur is None else str(int(round(float(auto_frame_from_cur), 0)))
        self._auto_frame_from_cur = auto_frame_from_cur
        self._auto_frame_from_auto = auto_frame_from_cur
        self.new_rating = self.current_rating
        self.new_frame = self.current_frame
        self.is_changed = False

    def set_auto_frames(self, from_current, from_auto):
        self._auto_frame_from_cur = from_current
        self._auto_frame_from_auto = from_auto


class BreakerActionWindow(forms.WPFWindow):
    def __init__(self, rows, preview_apply_callback, apply_callback):
        xaml = os.path.abspath(os.path.join(os.path.dirname(__file__), "CircuitBreakerActionWindow.xaml"))
        forms.WPFWindow.__init__(self, xaml)
        self._rows = list(rows or [])
        self._preview_apply_callback = preview_apply_callback
        self._apply_callback = apply_callback
        self._is_syncing_checks = False
        self._is_ready = False
        self._show_unsupported = False

        self._grid = self.FindName("CircuitGrid")
        self._status = self.FindName("ChangedStatusText")
        self._set_breaker_cb = self.FindName("AutoBreakerCheck")
        self._set_frame_cb = self.FindName("AutoFrameCheck")
        self._show_unsupported_cb = self.FindName("ShowUnsupportedToggle")

        if self._show_unsupported_cb is not None:
            self._show_unsupported_cb.IsChecked = False
        self._apply_visibility_filter()
        self.Loaded += self.window_loaded
        self._refresh_status()

    def window_loaded(self, sender, args):
        self._is_ready = True

    def _apply_visibility_filter(self):
        if self._grid is None:
            return
        if self._show_unsupported:
            items = list(self._rows)
        else:
            items = [x for x in self._rows if bool(getattr(x, "is_enabled", False))]
        self._grid.ItemsSource = ObservableCollection[object](items)

    def _refresh_grid(self, refresh_items=False):
        if refresh_items:
            self._apply_visibility_filter()
        self._refresh_status()

    def _refresh_status(self):
        changed = len([x for x in self._rows if x.is_changed and x.is_enabled])
        if self._show_unsupported:
            total = len(self._rows)
        else:
            total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
        if self._status is not None:
            self._status.Text = "{} of {} circuits to be modified.".format(changed, total)

    def _apply_checkbox_to_selected(self, sender, state):
        if not self._is_ready or self._is_syncing_checks:
            return
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        targets = [row]
        selected = []
        try:
            selected = list(self._grid.SelectedItems or [])
        except Exception:
            selected = []
        if row in selected and len(selected) > 1:
            targets = selected
        self._is_syncing_checks = True
        try:
            for item in targets:
                if not getattr(item, "is_enabled", False):
                    item.is_checked = False
                else:
                    if bool(getattr(item, "is_checked", False)) != bool(state):
                        item.is_checked = bool(state)
        finally:
            self._is_syncing_checks = False
        self._refresh_status()

    def item_checked(self, sender, args):
        self._apply_checkbox_to_selected(sender, True)

    def item_unchecked(self, sender, args):
        self._apply_checkbox_to_selected(sender, False)

    def check_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = bool(row.is_enabled)
        finally:
            self._is_syncing_checks = False
        self._refresh_grid(refresh_items=True)

    def uncheck_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        self._refresh_grid(refresh_items=True)

    def apply_autosized_clicked(self, sender, args):
        set_breaker = bool(getattr(self._set_breaker_cb, "IsChecked", True))
        set_frame = bool(getattr(self._set_frame_cb, "IsChecked", True))
        self._preview_apply_callback(self._rows, set_breaker, set_frame)
        self._refresh_grid(refresh_items=True)

    def apply_clicked(self, sender, args):
        set_breaker = bool(getattr(self._set_breaker_cb, "IsChecked", True))
        set_frame = bool(getattr(self._set_frame_cb, "IsChecked", True))
        if self._apply_callback(self._rows, set_breaker, set_frame):
            self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def show_unsupported_toggled(self, sender, args):
        self._show_unsupported = bool(getattr(sender, "IsChecked", False))
        self._refresh_grid(refresh_items=True)


class CircuitBrowserPanel(forms.WPFPanel):
    panel_id = PANEL_ID
    panel_title = TITLE
    panel_source = os.path.abspath(os.path.join(os.path.dirname(__file__), "CircuitBrowserPanel.xaml"))

    _instance = None
    _operation_gateway = None

    def __init__(self):
        forms.WPFPanel.__init__(self)
        CircuitBrowserPanel._instance = self
        self._logger = script.get_logger()
        self._all_items = []
        self._is_card_view = False
        self._type_options = []
        self._active_type_filters = set()
        self._filter_menu = None
        self._warnings_only = False
        self._overrides_only = False
        self._checked_only = False
        self._actions_menu = None
        self._operation_gateway = self._get_operation_gateway()

        self._list = self.FindName("CircuitList")
        self._search = self.FindName("SearchBox")
        self._status = self.FindName("StatusText")
        self._doc_name_text = self.FindName("DocumentNameText")
        self._toggle = self.FindName("ToggleViewButton")
        self._filter_button = self.FindName("FilterButton")

        self._compact_template = self.FindResource("CompactTemplate")
        self._card_template = self.FindResource("CardTemplate")
        self._active_doc_key = None
        self._view_activated_handler = None

        self.Loaded += self.panel_loaded
        self._attach_view_activated_handler()
        self._safe_load_items()

    @classmethod
    def get_instance(cls):
        return cls._instance

    @classmethod
    def _get_operation_gateway(cls):
        if cls._operation_gateway is None:
            cls._operation_gateway = CircuitOperationExternalEventGateway(
                logger=script.get_logger(),
                alert_parameter_name=ALERT_DATA_PARAM,
            )
        return cls._operation_gateway

    def _set_status(self, text):
        if self._status is not None:
            self._status.Text = text

    def _on_operation_complete(self, status, request, result, error):
        if status == "error":
            self._set_status("Operation failed")
            forms.alert("Operation failed:\n\n{}".format(error), title=TITLE)
            return

        if not result:
            self._set_status("Operation finished")
            self._safe_load_items()
            return

        if result.get("status") == "ok":
            self._set_status("Calculated {} circuits".format(result.get("updated_circuits", 0)))
            self._show_run_summary_if_needed(result)
            self._safe_load_items()
            return

        reason = result.get("reason", "unknown")
        self._set_status("Operation cancelled ({})".format(reason))

    def _on_alert_visibility_saved(self, status, request, result, error):
        if status == "error":
            self._set_status("Failed to update alert visibility")
            forms.alert("Failed to update hidden alert types:\n\n{}".format(error), title=TITLE)
            return
        if result and result.get("status") == "ok":
            self._set_status("Updated alert visibility")
            self._safe_load_items()
            return
        self._set_status("Alert visibility update cancelled")

    def _show_run_summary_if_needed(self, result):
        locked_rows = list(result.get("locked_rows") or [])
        runtime_rows = list(result.get("runtime_alert_rows") or [])
        if not locked_rows and not runtime_rows:
            return
        window = CircuitRunSummaryWindow(locked_rows, runtime_rows)
        window.ShowDialog()

    def _has_active_doc(self):
        try:
            return revit.doc is not None and revit.uidoc is not None
        except Exception:
            return False

    def _get_active_doc(self):
        try:
            uidoc = __revit__.ActiveUIDocument
            if uidoc:
                return uidoc.Document
        except Exception:
            pass
        try:
            return revit.doc
        except Exception:
            return None

    def _doc_key(self, doc):
        if doc is None:
            return None
        try:
            path = doc.PathName or ""
        except Exception:
            path = ""
        try:
            title = doc.Title or ""
        except Exception:
            title = ""
        return "{}|{}".format(path, title)

    def _set_doc_banner(self, doc):
        if self._doc_name_text is None:
            return
        if doc is None:
            self._doc_name_text.Text = "Document: -"
            return
        try:
            title = doc.Title or "-"
        except Exception:
            title = "-"
        self._doc_name_text.Text = "Document: {}".format(title)

    def _attach_view_activated_handler(self):
        if self._view_activated_handler is not None:
            return
        try:
            self._view_activated_handler = EventHandler[ViewActivatedEventArgs](self._on_view_activated)
            __revit__.ViewActivated += self._view_activated_handler
        except Exception:
            self._view_activated_handler = None

    def _on_view_activated(self, sender, args):
        try:
            doc = getattr(args, "Document", None)
        except Exception:
            doc = None

        if doc is None:
            self._safe_load_items()
            return

        key = self._doc_key(doc)
        if key != self._active_doc_key:
            self._safe_load_items()

    def _safe_load_items(self):
        doc = self._get_active_doc()
        self._active_doc_key = self._doc_key(doc)
        self._set_doc_banner(doc)
        if doc is None:
            self._all_items = []
            self._list.ItemsSource = ObservableCollection[CircuitListItem]([])
            self._set_status("Open a model document to load circuits.")
            return
        self._load_items(doc)

    def refresh_on_open(self):
        self._safe_load_items()

    def _rebuild_filter_options(self):
        old_options = set(self._type_options or [])
        old_active = set(self._active_type_filters or [])
        all_types = sorted(list({x.branch_type for x in self._all_items}))
        self._type_options = all_types
        all_set = set(all_types)

        if not old_options or not old_active:
            self._active_type_filters = set(all_types)
        elif old_active == old_options:
            self._active_type_filters = set(all_types)
        else:
            hidden_types = old_options.difference(old_active)
            self._active_type_filters = all_set.difference(hidden_types)
            if not self._active_type_filters:
                self._active_type_filters = set(all_types)
        self._update_filter_button_style()

    def _update_filter_button_style(self):
        if self._filter_button is None:
            return

        converter = BrushConverter()
        is_filtered = (
            self._warnings_only
            or self._overrides_only
            or self._checked_only
            or (set(self._type_options) != set(self._active_type_filters))
        )
        if is_filtered:
            self._filter_button.Background = converter.ConvertFrom("#1F6FB2")
            self._filter_button.Foreground = converter.ConvertFrom("#FFFFFF")
            self._filter_button.Content = "Filter*"
        else:
            self._filter_button.Background = converter.ConvertFrom("#E6E9EE")
            self._filter_button.Foreground = converter.ConvertFrom("#1F2E3D")
            self._filter_button.Content = "Filter"

    def _refresh_list(self):
        query = ""
        try:
            query = (self._search.Text or "").strip().lower()
        except Exception:
            query = ""

        items = list(self._all_items)
        if self._warnings_only:
            items = [x for x in items if int(getattr(x, "alert_count", 0) or 0) > 0]
        elif self._overrides_only:
            items = [x for x in items if bool(getattr(x, "has_override", False))]
        else:
            items = [x for x in items if x.branch_type in self._active_type_filters]
        if self._checked_only:
            items = [x for x in items if bool(getattr(x, "is_checked", False))]
        if query:
            items = [x for x in items if query in x.search_name]

        self._list.ItemsSource = ObservableCollection[CircuitListItem](items)
        self._set_status("Showing {} of {} circuits".format(len(items), len(self._all_items)))

    def _load_items(self, doc):
        self._set_status("Loading circuits...")

        circuits = list(
            DB.FilteredElementCollector(doc)
            .OfClass(DBE.ElectricalSystem)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        circuits.sort(key=lambda c: (
            (getattr(getattr(c, "BaseEquipment", None), "Name", "") or ""),
            (getattr(c, "StartSlot", 0) or 0),
            (getattr(c, "LoadName", "") or "")
        ))

        self._all_items = [CircuitListItem(c) for c in circuits]
        self._rebuild_filter_options()
        self._refresh_list()

    def _target_items(self):
        checked = [x for x in self._all_items if x.is_checked]
        if checked:
            return checked

        selected = []
        try:
            selected = list(self._list.SelectedItems)
        except Exception:
            selected = []
        return selected

    def _set_revit_selection(self, elements):
        uidoc = revit.uidoc
        ids = List[DB.ElementId]()
        for el in elements:
            try:
                if el is not None and hasattr(el, "Id"):
                    ids.Add(el.Id)
            except Exception:
                continue
        uidoc.Selection.SetElementIds(ids)

    def _apply_check_state(self, sender, state):
        clicked_item = None
        try:
            clicked_item = sender.DataContext
        except Exception:
            clicked_item = None

        if clicked_item is None:
            return

        selected = []
        try:
            selected = list(self._list.SelectedItems)
        except Exception:
            selected = []

        targets = [clicked_item]
        if clicked_item in selected and len(selected) > 1:
            targets = selected

        for item in targets:
            item.is_checked = bool(state)

        try:
            self._list.Items.Refresh()
        except Exception:
            pass

    def _build_filter_menu(self):
        menu = ContextMenu()

        warn_item = MenuItem()
        warn_item.Header = "Circuits With Alerts"
        warn_item.IsCheckable = True
        warn_item.IsChecked = self._warnings_only
        warn_item.StaysOpenOnClick = True
        warn_item.Checked += self.filter_warnings_toggled
        warn_item.Unchecked += self.filter_warnings_toggled
        menu.Items.Add(warn_item)

        overrides_item = MenuItem()
        overrides_item.Header = "User Overrides"
        overrides_item.IsCheckable = True
        overrides_item.IsChecked = self._overrides_only
        overrides_item.StaysOpenOnClick = True
        overrides_item.Checked += self.filter_overrides_toggled
        overrides_item.Unchecked += self.filter_overrides_toggled
        menu.Items.Add(overrides_item)

        checked_item = MenuItem()
        checked_item.Header = "Checked Circuits Only"
        checked_item.IsCheckable = True
        checked_item.IsChecked = self._checked_only
        checked_item.StaysOpenOnClick = True
        checked_item.Checked += self.filter_checked_toggled
        checked_item.Unchecked += self.filter_checked_toggled
        menu.Items.Add(checked_item)
        menu.Items.Add(Separator())

        all_item = MenuItem()
        all_item.Header = "Select All"
        all_item.Click += self.filter_select_all_clicked
        all_item.IsEnabled = not self._warnings_only and not self._overrides_only
        menu.Items.Add(all_item)

        none_item = MenuItem()
        none_item.Header = "Clear All"
        none_item.Click += self.filter_clear_all_clicked
        none_item.IsEnabled = not self._warnings_only and not self._overrides_only
        menu.Items.Add(none_item)

        menu.Items.Add(Separator())

        for ctype in self._type_options:
            mi = MenuItem()
            mi.Header = ctype
            mi.IsCheckable = True
            mi.IsChecked = ctype in self._active_type_filters
            mi.StaysOpenOnClick = True
            mi.Tag = ctype
            mi.IsEnabled = not self._warnings_only and not self._overrides_only
            mi.Checked += self.filter_type_toggled
            mi.Unchecked += self.filter_type_toggled
            menu.Items.Add(mi)

        self._filter_menu = menu
        if self._filter_button is not None:
            self._filter_button.ContextMenu = menu

    def panel_loaded(self, sender, args):
        self._safe_load_items()

    def search_changed(self, sender, args):
        self._refresh_list()

    def refresh_clicked(self, sender, args):
        self._safe_load_items()

    def filter_button_clicked(self, sender, args):
        self._build_filter_menu()
        if self._filter_menu is not None:
            self._filter_menu.PlacementTarget = self._filter_button
            self._filter_menu.IsOpen = True

    def _build_actions_menu(self):
        menu = ContextMenu()

        neutral_item = MenuItem()
        neutral_item.Header = "Add/Remove Neutral"
        neutral_item.Click += self.action_neutral_clicked
        menu.Items.Add(neutral_item)

        ig_item = MenuItem()
        ig_item.Header = "Add/Remove IG"
        ig_item.Click += self.action_ig_clicked
        menu.Items.Add(ig_item)

        breaker_item = MenuItem()
        breaker_item.Header = "Auto Size Breaker"
        breaker_item.Click += self.action_breaker_clicked
        menu.Items.Add(breaker_item)

        self._actions_menu = menu

    def actions_button_clicked(self, sender, args):
        self._build_actions_menu()
        if self._actions_menu is not None:
            self._actions_menu.PlacementTarget = sender
            self._actions_menu.IsOpen = True

    def _collect_action_targets(self):
        # Actions run only on explicitly checked circuits to avoid
        # selection/checkbox ambiguity for users.
        return [x for x in self._all_items if getattr(x, "is_checked", False)]

    def _is_locked_circuit(self, circuit):
        doc = self._get_active_doc()
        if doc is None or not getattr(doc, "IsWorkshared", False):
            return False, ""
        try:
            status = DB.WorksharingUtils.GetCheckoutStatus(doc, circuit.Id)
            if status != DB.CheckoutStatus.OwnedByOtherUser:
                return False, ""
            owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, circuit.Id).Owner or ""
            return True, owner
        except Exception:
            return False, ""

    def _param_int(self, element, name, default=0):
        value = _lookup_param_value(element, name)
        try:
            return int(value or 0)
        except Exception:
            return default

    def _param_text(self, element, name, default=""):
        value = _lookup_param_value(element, name)
        if value is None:
            return default
        try:
            return str(value)
        except Exception:
            return default

    def _wire_size_string(self, circuit):
        text = self._param_text(circuit, "Wire Size_CEDT", "")
        if text:
            return text
        return self._param_text(circuit, "CKT_Wire Size_CEDT", "")

    def _simulate_branch(self, circuit, include_neutral=None, include_ig=None):
        settings = settings_manager.load_circuit_settings(circuit.Document)
        branch = CircuitBranch(circuit, settings=settings)
        if include_neutral is not None:
            branch._include_neutral = bool(include_neutral)
        if include_ig is not None:
            branch._include_isolated_ground = bool(include_ig)
        # Include flags affect structural quantities; recompute before sizing preview.
        branch._setup_structural_quantities()
        branch.calculate_hot_wire_size()
        branch.calculate_neutral_wire_size()
        branch.calculate_ground_wire_size()
        branch.calculate_isolated_ground_wire_size()
        branch.calculate_conduit_size()
        return branch

    def _next_ocp_size(self, amps):
        try:
            value = float(amps)
        except Exception:
            return None
        keys = sorted([int(k) for k in BREAKER_FRAME_SWITCH_TABLE.keys()])
        for k in keys:
            if k >= value:
                return k
        return keys[-1] if keys else None

    def _frame_for_rating(self, rating):
        size = self._next_ocp_size(rating)
        if size is None:
            return None
        record = BREAKER_FRAME_SWITCH_TABLE.get(size) or {}
        return record.get("frame")

    def _raise_action_operation(self, operation_key, circuit_ids, options):
        if self._operation_gateway.is_busy():
            forms.alert("An operation is already running. Please wait.", title=TITLE)
            return False
        self._set_status("Applying action...")
        self._operation_gateway.raise_operation(
            operation_key=operation_key,
            circuit_ids=list(circuit_ids or []),
            source="pane",
            options=dict(options or {}),
            callback=self._on_operation_complete,
        )
        return True

    def _build_neutral_rows(self, targets):
        rows = []
        for item in targets:
            circuit = item.circuit
            btype = (item.branch_type or "").upper()
            reason = ""
            is_enabled = True

            if btype in BLOCKED_BRANCH_TYPES:
                reason = "Unsupported type: {}".format(btype)
                is_enabled = False
            elif btype != "BRANCH":
                reason = "Neutral action only supports BRANCH circuits."
                is_enabled = False
            else:
                poles = 0
                try:
                    poles = int(getattr(circuit, "PolesNumber", 0) or 0)
                except Exception:
                    poles = 0
                if poles <= 1:
                    reason = "Blocked: 1-pole circuits require neutral."
                    is_enabled = False

            locked, owner = self._is_locked_circuit(circuit)
            if locked:
                reason = "Locked by {}".format(owner or "another user")
                is_enabled = False

            current_qty = self._param_int(circuit, "CKT_Wire Neutral Quantity_CED", 0)
            current_size = self._param_text(circuit, "CKT_Wire Neutral Size_CEDT", "")
            current_wire = self._wire_size_string(circuit)
            row = NeutralIGActionRow(item, is_enabled, reason, current_qty, current_size, current_wire)
            if not is_enabled:
                row.is_checked = False
            rows.append(row)
        return rows

    def _preview_neutral_rows(self, rows, mode):
        include_neutral = mode == "add"
        for row in rows:
            row.new_qty = row.current_qty
            row.new_size = row.current_size
            row.new_wire = "no change"
            row.new_wire_font_style = "Italic"
            row.is_changed = False

            if not row.is_enabled or not row.is_checked:
                continue

            try:
                branch = self._simulate_branch(row.circuit, include_neutral=include_neutral)
                new_qty = int(branch.neutral_wire_quantity or 0)
                new_size = branch.neutral_wire_size or ""
                new_wire = branch.get_wire_size_callout() or ""
            except Exception:
                continue

            changed = (new_qty != int(row.current_qty or 0)) or (str(new_size or "") != str(row.current_size or "")) or (
                str(new_wire or "") != str(row.current_wire or "")
            )
            if changed:
                row.new_qty = new_qty
                row.new_size = new_size
                row.new_wire = new_wire or "-"
                row.new_wire_font_style = "Normal"
                row.is_changed = True

    def _apply_neutral_rows(self, rows, mode):
        changed_ids = [x.circuit_id for x in rows if x.is_enabled and x.is_checked and x.is_changed]
        if not changed_ids:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False
        return self._raise_action_operation(
            "set_neutral_and_recalculate",
            changed_ids,
            {"mode": mode, "show_output": False},
        )

    def _build_ig_rows(self, targets):
        rows = []
        for item in targets:
            circuit = item.circuit
            btype = (item.branch_type or "").upper()
            reason = ""
            is_enabled = True

            if btype in BLOCKED_BRANCH_TYPES:
                reason = "Unsupported type: {}".format(btype)
                is_enabled = False
            elif btype not in IG_BREAKER_ALLOWED_TYPES:
                reason = "IG action supports BRANCH/FEEDER/XFMR circuits."
                is_enabled = False

            if is_enabled:
                override_yes = self._param_int(circuit, "CKT_User Override_CED", 0) == 1
                ground_size = self._param_text(circuit, "CKT_Wire Ground Size_CEDT", "")
                if override_yes and ground_size.strip() == "-":
                    reason = "Blocked: user override with cleared ground size."
                    is_enabled = False

            locked, owner = self._is_locked_circuit(circuit)
            if locked:
                reason = "Locked by {}".format(owner or "another user")
                is_enabled = False

            current_qty = self._param_int(circuit, "CKT_Wire Isolated Ground Quantity_CED", 0)
            current_size = self._param_text(circuit, "CKT_Wire Isolated Ground Size_CEDT", "")
            current_wire = self._wire_size_string(circuit)
            row = NeutralIGActionRow(item, is_enabled, reason, current_qty, current_size, current_wire)
            if not is_enabled:
                row.is_checked = False
            rows.append(row)
        return rows

    def _preview_ig_rows(self, rows, mode):
        include_ig = mode == "add"
        for row in rows:
            row.new_qty = row.current_qty
            row.new_size = row.current_size
            row.new_wire = "no change"
            row.new_wire_font_style = "Italic"
            row.is_changed = False

            if not row.is_enabled or not row.is_checked:
                continue

            try:
                branch = self._simulate_branch(row.circuit, include_ig=include_ig)
                new_qty = int(branch.isolated_ground_wire_quantity or 0)
                new_size = branch.isolated_ground_wire_size or ""
                new_wire = branch.get_wire_size_callout() or ""
            except Exception:
                continue

            changed = (new_qty != int(row.current_qty or 0)) or (str(new_size or "") != str(row.current_size or "")) or (
                str(new_wire or "") != str(row.current_wire or "")
            )
            if changed:
                row.new_qty = new_qty
                row.new_size = new_size
                row.new_wire = new_wire or "-"
                row.new_wire_font_style = "Normal"
                row.is_changed = True

    def _apply_ig_rows(self, rows, mode):
        changed_ids = [x.circuit_id for x in rows if x.is_enabled and x.is_checked and x.is_changed]
        if not changed_ids:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False
        return self._raise_action_operation(
            "set_ig_and_recalculate",
            changed_ids,
            {"mode": mode, "show_output": False},
        )

    def _build_breaker_rows(self, targets):
        rows = []
        for item in targets:
            circuit = item.circuit
            btype = (item.branch_type or "").upper()
            reason = ""
            is_enabled = True

            if btype in BLOCKED_BRANCH_TYPES:
                reason = "Unsupported type: {}".format(btype)
                is_enabled = False
            elif btype not in IG_BREAKER_ALLOWED_TYPES:
                reason = "Breaker action supports BRANCH/FEEDER/XFMR circuits."
                is_enabled = False

            locked, owner = self._is_locked_circuit(circuit)
            if locked:
                reason = "Locked by {}".format(owner or "another user")
                is_enabled = False

            load_current = _lookup_param_value(circuit, "Circuit Load Current_CED")
            cur_rating = _lookup_param_value(circuit, "CKT_Rating_CED")
            if cur_rating is None:
                try:
                    cur_rating = circuit.Rating
                except Exception:
                    cur_rating = None
            cur_frame = _lookup_param_value(circuit, "CKT_Frame_CED")
            if cur_frame is None:
                try:
                    cur_frame = circuit.Frame
                except Exception:
                    cur_frame = None

            auto_rating = self._next_ocp_size((float(load_current) * 1.25) if load_current is not None else None)
            if auto_rating is None:
                auto_rating = self._next_ocp_size(cur_rating)
            auto_frame_current = self._frame_for_rating(cur_rating)
            auto_frame_auto = self._frame_for_rating(auto_rating)

            row = BreakerActionRow(item, is_enabled, reason, load_current, cur_rating, cur_frame, auto_rating, auto_frame_current)
            row.set_auto_frames(auto_frame_current, auto_frame_auto)
            row.autosized_frame = "-" if auto_frame_auto is None else str(int(round(float(auto_frame_auto), 0)))
            if not is_enabled:
                row.is_checked = False
            rows.append(row)
        return rows

    def _preview_breaker_rows(self, rows, set_breaker, set_frame):
        for row in rows:
            frame_preview = row._auto_frame_from_auto if set_breaker else row._auto_frame_from_cur
            row.autosized_frame = "-" if frame_preview is None else str(int(round(float(frame_preview), 0)))
            row.new_rating = row.current_rating
            row.new_frame = row.current_frame
            row.is_changed = False
            if not row.is_enabled or not row.is_checked:
                continue

            if set_breaker:
                row.new_rating = row.autosized_rating
            if set_frame:
                row.new_frame = row.autosized_frame

            row.is_changed = (str(row.new_rating) != str(row.current_rating)) or (str(row.new_frame) != str(row.current_frame))

    def _apply_breaker_rows(self, rows, set_breaker, set_frame):
        if not set_breaker and not set_frame:
            forms.alert("Choose at least one option: Autosize Breakers or Autosize Frames.", title=TITLE)
            return False

        updates = []
        for row in rows:
            if not (row.is_enabled and row.is_checked and row.is_changed):
                continue
            rating_val = None
            frame_val = None
            try:
                rating_val = float(row.new_rating)
            except Exception:
                rating_val = None
            try:
                frame_val = float(row.new_frame)
            except Exception:
                frame_val = None
            updates.append(
                {
                    "circuit_id": row.circuit_id,
                    "rating": rating_val,
                    "frame": frame_val,
                    "set_rating": bool(set_breaker),
                    "set_frame": bool(set_frame),
                }
            )

        if not updates:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False

        return self._raise_action_operation(
            "autosize_breaker_and_recalculate",
            [x.get("circuit_id") for x in updates],
            {"updates": updates, "show_output": False},
        )

    def action_neutral_clicked(self, sender, args):
        targets = self._collect_action_targets()
        if not targets:
            forms.alert("Check one or more circuits first.", title=TITLE)
            return
        if len(targets) > 300:
            choice = forms.alert(
                "{} circuits will be loaded for this action.\n\nContinue?".format(len(targets)),
                title="Large Action Selection",
                options=["Continue", "Cancel"],
            )
            if choice != "Continue":
                return
        rows = self._build_neutral_rows(targets)
        window = NeutralIGActionWindow(
            "Add/Remove Neutral",
            rows,
            preview_callback=self._preview_neutral_rows,
            apply_callback=self._apply_neutral_rows,
        )
        window.ShowDialog()

    def action_ig_clicked(self, sender, args):
        targets = self._collect_action_targets()
        if not targets:
            forms.alert("Check one or more circuits first.", title=TITLE)
            return
        if len(targets) > 300:
            choice = forms.alert(
                "{} circuits will be loaded for this action.\n\nContinue?".format(len(targets)),
                title="Large Action Selection",
                options=["Continue", "Cancel"],
            )
            if choice != "Continue":
                return
        rows = self._build_ig_rows(targets)
        window = NeutralIGActionWindow(
            "Add/Remove Isolated Ground",
            rows,
            preview_callback=self._preview_ig_rows,
            apply_callback=self._apply_ig_rows,
        )
        window.ShowDialog()

    def action_breaker_clicked(self, sender, args):
        targets = self._collect_action_targets()
        if not targets:
            forms.alert("Check one or more circuits first.", title=TITLE)
            return
        if len(targets) > 300:
            choice = forms.alert(
                "{} circuits will be loaded for this action.\n\nContinue?".format(len(targets)),
                title="Large Action Selection",
                options=["Continue", "Cancel"],
            )
            if choice != "Continue":
                return
        rows = self._build_breaker_rows(targets)
        window = BreakerActionWindow(
            rows,
            preview_apply_callback=self._preview_breaker_rows,
            apply_callback=self._apply_breaker_rows,
        )
        window.ShowDialog()

    def filter_select_all_clicked(self, sender, args):
        self._active_type_filters = set(self._type_options)
        self._update_filter_button_style()
        self._refresh_list()
        self._build_filter_menu()

    def filter_clear_all_clicked(self, sender, args):
        self._active_type_filters = set()
        self._update_filter_button_style()
        self._refresh_list()
        self._build_filter_menu()

    def filter_type_toggled(self, sender, args):
        ctype = getattr(sender, "Tag", None)
        if not ctype:
            return

        if sender.IsChecked:
            self._active_type_filters.add(ctype)
        else:
            if ctype in self._active_type_filters:
                self._active_type_filters.remove(ctype)

        self._update_filter_button_style()
        self._refresh_list()

    def filter_warnings_toggled(self, sender, args):
        self._warnings_only = bool(getattr(sender, "IsChecked", False))
        if self._warnings_only:
            self._overrides_only = False
        self._update_filter_button_style()
        self._refresh_list()
        self._build_filter_menu()

    def filter_overrides_toggled(self, sender, args):
        self._overrides_only = bool(getattr(sender, "IsChecked", False))
        if self._overrides_only:
            self._warnings_only = False
        self._update_filter_button_style()
        self._refresh_list()
        self._build_filter_menu()

    def filter_checked_toggled(self, sender, args):
        self._checked_only = bool(getattr(sender, "IsChecked", False))
        self._update_filter_button_style()
        self._refresh_list()
        self._build_filter_menu()

    def toggle_view_clicked(self, sender, args):
        self._is_card_view = not self._is_card_view
        if self._is_card_view:
            self._list.ItemTemplate = self._card_template
            self._toggle.Content = "Compact"
        else:
            self._list.ItemTemplate = self._compact_template
            self._toggle.Content = "Card"

    def selection_changed(self, sender, args):
        selected_count = 0
        try:
            selected_count = len(list(self._list.SelectedItems))
        except Exception:
            selected_count = 0
        checked_count = len([x for x in self._all_items if x.is_checked])
        self._set_status("Checked: {} | Selected rows: {}".format(checked_count, selected_count))

    def item_checked(self, sender, args):
        self._apply_check_state(sender, True)
        self.selection_changed(None, None)

    def item_unchecked(self, sender, args):
        self._apply_check_state(sender, False)
        self.selection_changed(None, None)

    def select_equipment_clicked(self, sender, args):
        targets = []
        for item in self._target_items():
            try:
                base_eq = item.circuit.BaseEquipment
                if base_eq:
                    targets.append(base_eq)
            except Exception:
                continue
        self._set_revit_selection(targets)

    def select_circuits_clicked(self, sender, args):
        self._set_revit_selection([x.circuit for x in self._target_items()])

    def select_downstream_clicked(self, sender, args):
        targets = []
        for item in self._target_items():
            try:
                for el in item.circuit.Elements:
                    targets.append(el)
            except Exception:
                continue
        self._set_revit_selection(targets)

    def clear_revit_selection_clicked(self, sender, args):
        revit.uidoc.Selection.SetElementIds(List[DB.ElementId]())

    def check_all_clicked(self, sender, args):
        try:
            items = list(self._list.ItemsSource or [])
        except Exception:
            items = []
        for item in items:
            item.is_checked = True
        try:
            self._list.Items.Refresh()
        except Exception:
            pass
        self.selection_changed(None, None)

    def uncheck_all_clicked(self, sender, args):
        try:
            items = list(self._list.ItemsSource or [])
        except Exception:
            items = []
        for item in items:
            item.is_checked = False
        try:
            self._list.Items.Refresh()
        except Exception:
            pass
        self.selection_changed(None, None)

    def calculate_selected_clicked(self, sender, args):
        if not self._has_active_doc():
            forms.alert("Open a model document first.", title=TITLE)
            return
        items = self._target_items()
        circuit_ids = [x.circuit.Id.IntegerValue for x in items if getattr(x, "circuit", None)]
        if not circuit_ids:
            forms.alert("No circuits selected.", title=TITLE)
            return
        if self._operation_gateway.is_busy():
            forms.alert("An operation is already running. Please wait.", title=TITLE)
            return
        self._set_status("Calculating selected circuits...")
        raised = self._operation_gateway.raise_operation(
            operation_key="calculate_circuits",
            circuit_ids=circuit_ids,
            source="pane",
            options={"show_output": False},
            callback=self._on_operation_complete,
        )
        if not raised:
            self._set_status("Unable to queue operation")

    def calculate_all_clicked(self, sender, args):
        if not self._has_active_doc():
            forms.alert("Open a model document first.", title=TITLE)
            return
        circuit_ids = [x.circuit.Id.IntegerValue for x in self._all_items if getattr(x, "circuit", None)]
        if not circuit_ids:
            forms.alert("No circuits available.", title=TITLE)
            return
        if self._operation_gateway.is_busy():
            forms.alert("An operation is already running. Please wait.", title=TITLE)
            return
        self._set_status("Calculating all circuits...")
        raised = self._operation_gateway.raise_operation(
            operation_key="calculate_circuits",
            circuit_ids=circuit_ids,
            source="pane",
            options={"show_output": False},
            callback=self._on_operation_complete,
        )
        if not raised:
            self._set_status("Unable to queue operation")

    def calculate_settings_clicked(self, sender, args):
        if not os.path.exists(CALC_SETTINGS_PATH):
            forms.alert("Calculate Circuits settings file not found:\n\n{}".format(CALC_SETTINGS_PATH), title=TITLE)
            return
        if not os.path.exists(CALC_SETTINGS_XAML_PATH):
            forms.alert("Calculate Circuits settings XAML not found:\n\n{}".format(CALC_SETTINGS_XAML_PATH), title=TITLE)
            return
        try:
            module = imp.load_source("ced_calculate_circuits_config", CALC_SETTINGS_PATH)
            try:
                module.XAML_PATH = CALC_SETTINGS_XAML_PATH
            except Exception:
                pass
            window_cls = getattr(module, "CircuitSettingsWindow", None)
            if window_cls is None:
                forms.alert("CircuitSettingsWindow was not found in config script.", title=TITLE)
                return
            window = window_cls()
            try:
                window.ShowDialog()
            except Exception:
                window.show_dialog()
        except Exception as ex:
            forms.alert("Failed to open Calculate Circuits settings:\n\n{}".format(ex), title=TITLE)

    def alert_tag_clicked(self, sender, args):
        item = None
        try:
            item = sender.DataContext
        except Exception:
            item = None
        if not item:
            return

        rows = list(getattr(item, "alert_rows", []) or [])
        if not rows:
            forms.alert("No alerts stored for this circuit.", title=TITLE)
            return

        header = "{} / {} - {}".format(item.panel, item.circuit_number, item.load_name)
        try:
            window = CircuitAlertsWindow(header, rows)
            try:
                window.ShowDialog()
            except Exception:
                window.Show()
                try:
                    window.Activate()
                except Exception:
                    pass
        except Exception as ex:
            forms.alert("Failed to open alerts window:\n\n{}".format(ex), title=TITLE)
            return
        hidden_ids = getattr(window, "updated_hidden_ids", None)
        if hidden_ids is None:
            return

        if self._operation_gateway.is_busy():
            forms.alert("An operation is already running. Please wait.", title=TITLE)
            return
        self._set_status("Updating alert visibility...")
        self._operation_gateway.raise_operation(
            operation_key="set_hidden_alert_types",
            circuit_ids=[item.circuit.Id.IntegerValue],
            source="pane",
            options={"hidden_definition_ids": list(hidden_ids)},
            callback=self._on_alert_visibility_saved,
        )


def ensure_panel_visible():
    try:
        if not forms.is_registered_dockable_panel(CircuitBrowserPanel):
            forms.register_dockable_panel(CircuitBrowserPanel, default_visible=False)
    except Exception:
        pass

    try:
        forms.open_dockable_panel(CircuitBrowserPanel.panel_id)
    except Exception:
        try:
            forms.open_dockable_panel(CircuitBrowserPanel)
        except Exception:
            pass

    panel = CircuitBrowserPanel.get_instance()
    if panel is not None:
        panel.refresh_on_open()
