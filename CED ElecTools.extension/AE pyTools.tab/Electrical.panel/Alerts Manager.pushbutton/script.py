# -*- coding: utf-8 -*-

import os

<<<<<<< HEAD
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Windows import Application
=======
import clr

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Windows import Application, WindowState
>>>>>>> main
from System.Windows.Controls import Button, DataGridRow
from System.Windows.Media import VisualTreeHelper
from pyrevit import forms, revit, script

from Snippets import revit_helpers
from UIClasses import pathing as ui_pathing

TITLE = "Alerts Manager"
ALERT_DATA_PARAM = "Circuit Data_CED"
<<<<<<< HEAD
THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_THEME_KEY = "theme_mode"
THEME_CONFIG_ACCENT_KEY = "accent_mode"
VALID_THEME_MODES = ("light", "dark", "dark_alt")
VALID_ACCENT_MODES = ("blue", "neutral")
=======
>>>>>>> main
_WINDOW_MARKER = "_ae_alerts_browser_window"

def _idval(item):
    return int(revit_helpers.get_elementid_value(item))


def _idfrom(value):
    return revit_helpers.elementid_from_value(value)


def _find_visual_ancestor(start, target_type):
    current = start
    while current is not None:
        if isinstance(current, target_type):
            return current
        try:
            current = VisualTreeHelper.GetParent(current)
        except Exception:
            return None
    return None


def _is_descendant_of_control(start, control):
    current = start
    while current is not None:
        if current == control:
            return True
        try:
            current = VisualTreeHelper.GetParent(current)
        except Exception:
            return False
    return False


<<<<<<< HEAD
def _normalize_theme_mode(value, fallback="light"):
    mode = str(value or fallback).strip().lower()
    return mode if mode in VALID_THEME_MODES else fallback


def _normalize_accent_mode(value, fallback="blue"):
    mode = str(value or fallback).strip().lower()
    return mode if mode in VALID_ACCENT_MODES else fallback


=======
>>>>>>> main
def _load_theme_state_from_config(default_theme="light", default_accent="blue"):
    from UIClasses import load_theme_state_from_config

    return load_theme_state_from_config(
<<<<<<< HEAD
        section_name=THEME_CONFIG_SECTION,
        theme_key_name=THEME_CONFIG_THEME_KEY,
        accent_key_name=THEME_CONFIG_ACCENT_KEY,
=======
>>>>>>> main
        default_theme=default_theme,
        default_accent=default_accent,
    )


THIS_DIR = os.path.abspath(os.path.dirname(__file__))
LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate workspace root for Alerts Manager.", title=TITLE)
    raise SystemExit

from CEDElectrical.Infrastructure.Revit.repositories.revit_circuit_repository import RevitCircuitRepository
from Snippets.circuit_ui_actions import (
    clear_revit_selection,
    collect_circuit_targets,
    set_revit_selection,
)
from UIClasses import resource_loader
from alerts_browser_services import build_snapshot
from alerts_browser_services import recalculate_and_snapshot
from alerts_browser_services import update_hidden_alert_types_and_snapshot

UI_RESOURCES_ROOT = ui_pathing.resolve_ui_resources_root(LIB_ROOT)
_LOCK_REPOSITORY = RevitCircuitRepository()
_LOGGER = script.get_logger()


def _active_doc():
    doc = getattr(revit, "doc", None)
    if doc is not None:
        return doc
    try:
        uidoc = __revit__.ActiveUIDocument
        return uidoc.Document if uidoc else None
    except Exception:
        return getattr(revit, "doc", None)


class AlertsBrowserExternalEventGateway(object):
    def __init__(self, logger=None):
        self._logger = logger
        self._pending = None
        self._handler = _AlertsBrowserExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def _is_event_pending(self):
        try:
            return bool(self._event.IsPending)
        except Exception:
            return False

    def is_busy(self):
        return self._pending is not None or self._is_event_pending()

    def _raise(self, op_name, payload=None, callback=None):
        if self._pending is not None or self._is_event_pending():
            return False
        self._pending = {
            "op": str(op_name or ""),
            "payload": dict(payload or {}),
            "callback": callback,
        }
        try:
            self._event.Raise()
            return True
        except Exception as ex:
            self._pending = None
            if self._logger:
                self._logger.warning("Alerts Manager ExternalEvent raise failed: %s", ex)
            return False

    def raise_refresh(self, callback):
        return self._raise("refresh", callback=callback)

    def raise_recalculate(self, circuit_id, callback):
        return self._raise("recalculate", payload={"circuit_id": int(circuit_id)}, callback=callback)

    def raise_select(self, mode, circuit_id, callback=None):
        payload = {
            "mode": str(mode or ""),
            "circuit_id": int(circuit_id or 0),
        }
        return self._raise("select", payload=payload, callback=callback)

    def raise_set_hidden(self, circuit_id, hidden_definition_ids, callback):
        payload = {
            "circuit_id": int(circuit_id or 0),
            "hidden_definition_ids": list(hidden_definition_ids or []),
        }
        return self._raise("set_hidden", payload=payload, callback=callback)

    def _consume_pending(self):
        pending = self._pending
        self._pending = None
        return pending


class _AlertsBrowserExternalEventHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, application):  # noqa: N802
        pending = self._gateway._consume_pending()
        if not pending:
            return
        op_name = pending.get("op")
        payload = pending.get("payload") or {}
        callback = pending.get("callback")
        status = "ok"
        result = None
        error = None
        try:
            uidoc = application.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if op_name == "refresh":
                result = build_snapshot(doc, ALERT_DATA_PARAM, _idval, _LOCK_REPOSITORY)
            elif op_name == "recalculate":
                if doc is None:
                    raise Exception("No active document.")
                circuit_id = int(payload.get("circuit_id") or 0)
                if circuit_id <= 0:
                    raise Exception("No circuit selected.")
                result = recalculate_and_snapshot(doc, circuit_id, ALERT_DATA_PARAM, _idval, _LOCK_REPOSITORY)
            elif op_name == "select":
                if uidoc is None or doc is None:
                    raise Exception("No active document.")
                mode = str(payload.get("mode") or "").strip().lower()
                circuit_id = int(payload.get("circuit_id") or 0)
                if mode == "clear":
                    clear_revit_selection(uidoc=uidoc)
                else:
                    if circuit_id <= 0:
                        raise Exception("No circuit selected.")
                    circuit = doc.GetElement(_idfrom(circuit_id))
                    targets = collect_circuit_targets(circuit, mode)
                    set_revit_selection(targets, uidoc=uidoc)
                result = {"selected": True}
            elif op_name == "set_hidden":
                if doc is None:
                    raise Exception("No active document.")
                circuit_id = int(payload.get("circuit_id") or 0)
                if circuit_id <= 0:
                    raise Exception("No circuit selected.")
                hidden_definition_ids = list(payload.get("hidden_definition_ids") or [])
                result = update_hidden_alert_types_and_snapshot(
                    doc,
                    circuit_id,
                    hidden_definition_ids,
                    ALERT_DATA_PARAM,
                    _idval,
                    _LOCK_REPOSITORY,
                )
            else:
                raise Exception("Unknown operation: {}".format(op_name))
        except Exception as ex:
            status = "error"
            error = ex
            if self._gateway._logger:
                self._gateway._logger.exception("Alerts Manager external operation failed: %s", ex)

        if callback:
            try:
                callback(status, op_name, result, error)
            except Exception:
                pass

    def GetName(self):  # noqa: N802
        return "CED Alerts Manager External Event"


class AlertsBrowserWindow(forms.WPFWindow):
    def __init__(self, theme_mode, accent_mode, snapshot, gateway):
        xaml = os.path.abspath(os.path.join(THIS_DIR, "AlertsBrowserWindow.xaml"))
<<<<<<< HEAD
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
=======
        self._theme_mode = resource_loader.normalize_theme_mode(theme_mode, "light")
        self._accent_mode = resource_loader.normalize_accent_mode(accent_mode, "blue")
>>>>>>> main
        self._gateway = gateway
        self._items = []
        self._doc_title = "-"
        forms.WPFWindow.__init__(self, xaml)
<<<<<<< HEAD
        self._apply_theme()
        setattr(self, _WINDOW_MARKER, True)
=======
        # Use CLR Tag for cross-runtime singleton detection.
        try:
            self.Tag = _WINDOW_MARKER
        except Exception:
            pass
        self._apply_theme()
>>>>>>> main

        self._circuit_list = self.FindName("CircuitList")
        self._active_list = self.FindName("ActiveAlertsList")
        self._hidden_list = self.FindName("HiddenAlertsList")
        self._document_text = self.FindName("DocumentText")
        self._count_text = self.FindName("CircuitCountText")
        self._selected_circuit_text = self.FindName("SelectedCircuitText")
        self._selected_counts_text = self.FindName("SelectedCountsText")
        self._status_text = self.FindName("StatusText")
        self._refresh_button = self.FindName("RefreshButton")
        self._show_hidden_toggle = self.FindName("ShowHiddenToggle")
        self._tabs = self.FindName("AlertsTabs")
        self._hide_unhide_button = self.FindName("HideUnhideButton")
        if self._tabs is not None:
            self._tabs.SelectionChanged += self.tabs_selection_changed
        self._apply_snapshot(snapshot, preferred_circuit_id=None)

    def _apply_theme(self):
        resource_loader.apply_theme(
            self,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )

    def _set_status(self, text):
        if self._status_text is None:
            return
        try:
            self._status_text.Text = str(text or "")
        except Exception:
            pass

    def _current_tab_index(self):
        if self._tabs is None:
            return 0
        try:
            return int(getattr(self._tabs, "SelectedIndex", 0) or 0)
        except Exception:
            return 0

    def _selected_alert_row(self):
        idx = self._current_tab_index()
        target_list = self._hidden_list if idx == 1 else self._active_list
        if target_list is None:
            return None
        try:
            return getattr(target_list, "SelectedItem", None)
        except Exception:
            return None

    def _build_hidden_ids_for_item(self, item):
        if item is None:
            return []
        hidden_ids = set()
        for row in list(getattr(item, "rows", []) or []):
            definition_id = str(getattr(row, "definition_id", "") or "").strip()
            if not definition_id or definition_id == "-":
                continue
            if not bool(getattr(row, "can_hide", False)):
                continue
            if bool(getattr(row, "is_hidden", False)):
                hidden_ids.add(definition_id)
        return sorted(list(hidden_ids))

    def _sync_hide_unhide_state(self):
        btn = self._hide_unhide_button
        if btn is None:
            return
        hidden_tab = self._current_tab_index() == 1
        btn.Content = "Unhide Type" if hidden_tab else "Hide Type"

        if self._gateway is not None and self._gateway.is_busy():
            btn.IsEnabled = False
            btn.ToolTip = "Operation is running..."
            return

        item = self._selected_item()
        if item is None:
            btn.IsEnabled = False
            btn.ToolTip = "Select a circuit first."
            return

        row = self._selected_alert_row()
        if row is None:
            btn.IsEnabled = False
            btn.ToolTip = "Select an alert type first."
            return

        definition_id = str(getattr(row, "definition_id", "") or "").strip()
        if not definition_id or definition_id == "-":
            btn.IsEnabled = False
            btn.ToolTip = "Only mapped alert types can be changed."
            return

        if (not hidden_tab) and (not bool(getattr(row, "can_hide", False))):
            btn.IsEnabled = False
            btn.ToolTip = "This alert type can not be hidden."
            return

        btn.IsEnabled = True
        btn.ToolTip = "Toggle visibility for this alert type on the selected circuit."

    def _selected_item(self):
        try:
            return getattr(self._circuit_list, "SelectedItem", None)
        except Exception:
            return None

    def _find_item_by_id(self, circuit_id):
        target = int(circuit_id or 0)
        if target <= 0:
            return None
        for item in list(self._items or []):
            try:
                if int(getattr(item, "circuit_id", 0)) == target:
                    return item
            except Exception:
                continue
        return None

    def _show_hidden_enabled(self):
        if self._show_hidden_toggle is None:
            return True
        try:
            return bool(getattr(self._show_hidden_toggle, "IsChecked", True))
        except Exception:
            return True

    def _visible_items(self):
        items = list(self._items or [])
        if self._show_hidden_enabled():
            return items
        visible = []
        for item in items:
            active_count = int(getattr(item, "active_count", 0) or 0)
            hidden_count = int(getattr(item, "hidden_count", 0) or 0)
            if active_count <= 0 and hidden_count > 0:
                continue
            visible.append(item)
        return visible

    def _apply_snapshot(self, snapshot, preferred_circuit_id=None):
        data = dict(snapshot or {})
        doc_title = str(data.get("doc_title") or "-")
        self._doc_title = doc_title
        items = list(data.get("items") or [])
        self._items = items
        visible_items = self._visible_items()
        if self._document_text is not None:
            self._document_text.Text = "Document: {}".format(doc_title)
        if self._count_text is not None:
            self._count_text.Text = "{} circuits with alerts".format(len(visible_items))
        if self._circuit_list is not None:
            self._circuit_list.ItemsSource = list(visible_items)
        selected = self._find_item_by_id(preferred_circuit_id)
        if selected is None:
            selected = self._selected_item()
            if selected not in visible_items:
                selected = visible_items[0] if visible_items else None
        if self._circuit_list is not None:
            try:
                self._circuit_list.SelectedItem = selected
            except Exception:
                pass
        self._set_selected(selected)

    def show_hidden_toggled(self, sender, args):
        selected = self._selected_item()
        preferred_id = int(getattr(selected, "circuit_id", 0) or 0) if selected is not None else 0
        snapshot = {
            "doc_title": self._doc_title,
            "items": list(self._items or []),
        }
        self._apply_snapshot(snapshot, preferred_circuit_id=preferred_id)

    def _set_selected(self, item):
        if item is None:
            if self._selected_circuit_text is not None:
                self._selected_circuit_text.Text = "Select a circuit with alerts"
            if self._selected_counts_text is not None:
                self._selected_counts_text.Text = "Alerts: 0"
            if self._active_list is not None:
                self._active_list.ItemsSource = []
                self._active_list.SelectedItem = None
            if self._hidden_list is not None:
                self._hidden_list.ItemsSource = []
                self._hidden_list.SelectedItem = None
            self._update_refresh_state(None)
            self._sync_hide_unhide_state()
            return
        if self._selected_circuit_text is not None:
            self._selected_circuit_text.Text = "{} - {}".format(item.panel_ckt_text, item.load_name or "-")
        if self._selected_counts_text is not None:
            self._selected_counts_text.Text = item.counts_text
        if self._active_list is not None:
            self._active_list.ItemsSource = list(item.active_rows or [])
            self._active_list.SelectedItem = None
        if self._hidden_list is not None:
            self._hidden_list.ItemsSource = list(item.hidden_rows or [])
            self._hidden_list.SelectedItem = None
        self._update_refresh_state(item)
        self._sync_hide_unhide_state()

    def _update_refresh_state(self, item):
        if self._refresh_button is None:
            return
        if self._gateway is not None and self._gateway.is_busy():
            self._refresh_button.IsEnabled = False
            self._refresh_button.ToolTip = "Operation is running..."
            self._sync_hide_unhide_state()
            return
        if item is None:
            self._refresh_button.IsEnabled = False
            self._refresh_button.ToolTip = "Select a circuit first."
            self._sync_hide_unhide_state()
            return
        if bool(getattr(item, "recalc_blocked", False)):
            self._refresh_button.IsEnabled = False
            reason = getattr(item, "recalc_block_reason", "") or "Calculation blocked by ownership constraints."
            self._refresh_button.ToolTip = reason
            self._sync_hide_unhide_state()
            return
        self._refresh_button.IsEnabled = True
        self._refresh_button.ToolTip = "Recalculate selected circuit and refresh alerts."
        self._sync_hide_unhide_state()

    def _handle_external_complete(self, status, op_name, result, error):
        if status == "error":
            self._set_status("Operation failed")
            forms.alert("Alerts Manager operation failed:\n\n{}".format(error), title=TITLE)
            self._update_refresh_state(self._selected_item())
            return
        if op_name == "refresh":
            self._apply_snapshot(result or {}, preferred_circuit_id=None)
            self._set_status("Refreshed alerts.")
            return
        if op_name == "recalculate":
            data = dict(result or {})
            operation_result = dict(data.get("operation_result") or {})
            snapshot = data.get("snapshot") or {}
            circuit_id = data.get("circuit_id")
            self._apply_snapshot(snapshot, preferred_circuit_id=circuit_id)
            if operation_result.get("status") == "ok":
                self._set_status("Recalculated selected circuit.")
            else:
                reason = operation_result.get("reason") or "cancelled"
                self._set_status("Recalculate cancelled ({})".format(reason))
            return
        if op_name == "set_hidden":
            data = dict(result or {})
            operation_result = dict(data.get("operation_result") or {})
            snapshot = data.get("snapshot") or {}
            circuit_id = data.get("circuit_id")
            self._apply_snapshot(snapshot, preferred_circuit_id=circuit_id)
            if operation_result.get("status") == "ok":
                self._set_status("Updated hidden alert types.")
            else:
                reason = operation_result.get("reason") or "cancelled"
                self._set_status("Hide/unhide cancelled ({})".format(reason))
            return
        if op_name == "select":
            self._set_status("Selection updated.")
            self._update_refresh_state(self._selected_item())
            return

    def _raise_select(self, mode):
        item = self._selected_item()
        circuit_id = int(getattr(item, "circuit_id", 0) or 0) if item is not None else 0
        if mode != "clear" and circuit_id <= 0:
            self._set_status("No circuit selected.")
            return
        if self._gateway is None:
            return
        raised = self._gateway.raise_select(mode, circuit_id, callback=self._handle_external_complete)
        if not raised:
            self._set_status("Unable to queue selection operation.")
            self._update_refresh_state(item)

    def circuit_selection_changed(self, sender, args):
        self._set_selected(self._selected_item())

    def tabs_selection_changed(self, sender, args):
        self._sync_hide_unhide_state()

    def _clear_alert_grid_selection(self):
        try:
            if self._active_list is not None:
                self._active_list.SelectedItem = None
            if self._hidden_list is not None:
                self._hidden_list.SelectedItem = None
        except Exception:
            pass

    def window_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if self._active_list is not None and _is_descendant_of_control(source, self._active_list):
            return
        if self._hidden_list is not None and _is_descendant_of_control(source, self._hidden_list):
            return
        self._clear_alert_grid_selection()

    def alerts_list_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, DataGridRow) is not None:
            return
        try:
            sender.SelectedItem = None
        except Exception:
            pass
        self._sync_hide_unhide_state()

    def alerts_grid_selection_changed(self, sender, args):
        if sender == self._active_list and self._active_list is not None:
            if getattr(self._active_list, "SelectedItem", None) is not None and self._hidden_list is not None:
                self._hidden_list.SelectedItem = None
        elif sender == self._hidden_list and self._hidden_list is not None:
            if getattr(self._hidden_list, "SelectedItem", None) is not None and self._active_list is not None:
                self._active_list.SelectedItem = None
        self._sync_hide_unhide_state()

    def select_equipment_clicked(self, sender, args):
        self._raise_select("panel")

    def select_circuit_clicked(self, sender, args):
        self._raise_select("circuit")

    def select_downstream_clicked(self, sender, args):
        self._raise_select("device")

    def clear_selection_clicked(self, sender, args):
        self._raise_select("clear")

    def hide_unhide_clicked(self, sender, args):
        item = self._selected_item()
        if item is None:
            self._set_status("No circuit selected.")
            self._sync_hide_unhide_state()
            return
        row = self._selected_alert_row()
        if row is None:
            self._set_status("No alert type selected.")
            self._sync_hide_unhide_state()
            return
        definition_id = str(getattr(row, "definition_id", "") or "").strip()
        if not definition_id or definition_id == "-":
            self._set_status("Only mapped alert types can be changed.")
            self._sync_hide_unhide_state()
            return

        hidden_tab = self._current_tab_index() == 1
        if (not hidden_tab) and (not bool(getattr(row, "can_hide", False))):
            self._set_status("This alert type can not be hidden.")
            self._sync_hide_unhide_state()
            return

        for item_row in list(getattr(item, "rows", []) or []):
            if str(getattr(item_row, "definition_id", "") or "").strip() != definition_id:
                continue
            item_row.is_hidden = not hidden_tab

        hidden_ids = self._build_hidden_ids_for_item(item)
        if self._gateway is None:
            return
        raised = self._gateway.raise_set_hidden(
            getattr(item, "circuit_id", 0),
            hidden_ids,
            callback=self._handle_external_complete,
        )
        if not raised:
            self._set_status("Unable to queue hide/unhide update.")
            self._sync_hide_unhide_state()
            return
        self._set_status("Updating hidden alert types...")
        self._sync_hide_unhide_state()

    def refresh_clicked(self, sender, args):
        selected = self._selected_item()
        self._update_refresh_state(selected)
        if selected is None:
            self._set_status("No circuit selected.")
            return
        if bool(getattr(selected, "recalc_blocked", False)):
            self._set_status(getattr(selected, "recalc_block_reason", "") or "Calculation blocked.")
            return
        if self._gateway is None:
            return
        raised = self._gateway.raise_recalculate(
            getattr(selected, "circuit_id", 0),
            callback=self._handle_external_complete,
        )
        if not raised:
            self._set_status("Unable to queue recalculate.")
            self._update_refresh_state(selected)
            return
        self._set_status("Recalculating selected circuit...")
        self._update_refresh_state(selected)

    def close_clicked(self, sender, args):
        self.Close()


def _find_existing_window():
    app = Application.Current
    if app is None:
        return None
    try:
        windows = list(app.Windows)
    except Exception:
        windows = []
    for win in windows:
        try:
<<<<<<< HEAD
            if bool(getattr(win, _WINDOW_MARKER, False)):
=======
            tag = str(getattr(win, "Tag", "") or "")
            if tag == _WINDOW_MARKER:
                return win
            if str(getattr(win, "Title", "") or "") == TITLE:
>>>>>>> main
                return win
        except Exception:
            continue
    return None


<<<<<<< HEAD
def _show_or_focus_window():
    existing = _find_existing_window()
    if existing is not None:
        try:
            theme_mode, accent_mode = _load_theme_state_from_config(
                default_theme=getattr(existing, "_theme_mode", "light"),
                default_accent=getattr(existing, "_accent_mode", "blue"),
            )
            existing._theme_mode = theme_mode
            existing._accent_mode = accent_mode
            if hasattr(existing, "_apply_theme"):
                existing._apply_theme()
            existing.Show()
            existing.Activate()
            return
        except Exception:
            pass
=======
def _focus_existing_window(existing):
    try:
        theme_mode, accent_mode = _load_theme_state_from_config(
            default_theme=getattr(existing, "_theme_mode", "light"),
            default_accent=getattr(existing, "_accent_mode", "blue"),
        )
        existing._theme_mode = resource_loader.normalize_theme_mode(theme_mode, getattr(existing, "_theme_mode", "light"))
        existing._accent_mode = resource_loader.normalize_accent_mode(accent_mode, getattr(existing, "_accent_mode", "blue"))
        if hasattr(existing, "_apply_theme"):
            existing._apply_theme()
    except Exception as ex:
        try:
            _LOGGER.debug("Alerts Manager focus theme sync failed: %s", ex)
        except Exception:
            pass
    try:
        if getattr(existing, "WindowState", None) == WindowState.Minimized:
            existing.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        existing.Show()
    except Exception:
        pass
    try:
        existing.Activate()
    except Exception:
        pass
    try:
        existing.Focus()
    except Exception:
        pass


def _show_or_focus_window():
    existing = _find_existing_window()
    if existing is not None:
        _focus_existing_window(existing)
        return
>>>>>>> main
    theme_mode, accent_mode = _load_theme_state_from_config("light", "blue")
    snapshot = build_snapshot(_active_doc(), ALERT_DATA_PARAM, _idval, _LOCK_REPOSITORY)
    gateway = AlertsBrowserExternalEventGateway(logger=_LOGGER)
    window = AlertsBrowserWindow(
        theme_mode=theme_mode,
        accent_mode=accent_mode,
        snapshot=snapshot,
        gateway=gateway,
    )
    window.Show()
    try:
        window.Activate()
    except Exception:
        pass


_show_or_focus_window()

