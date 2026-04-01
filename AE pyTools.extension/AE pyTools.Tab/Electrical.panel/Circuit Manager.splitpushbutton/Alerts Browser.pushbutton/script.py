# -*- coding: utf-8 -*-

import os
import sys

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Windows import Application
from System.Windows.Controls import Button, DataGridRow
from System.Windows.Media import VisualTreeHelper
from pyrevit import forms, revit, script

from Snippets import revit_helpers

TITLE = "Alerts Browser"
ALERT_DATA_PARAM = "Circuit Data_CED"
THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_THEME_KEY = "theme_mode"
THEME_CONFIG_ACCENT_KEY = "accent_mode"
VALID_THEME_MODES = ("light", "dark", "dark_alt")
VALID_ACCENT_MODES = ("blue", "red", "green", "neutral")
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


def _normalize_theme_mode(value, fallback="light"):
    mode = str(value or fallback).strip().lower()
    return mode if mode in VALID_THEME_MODES else fallback


def _normalize_accent_mode(value, fallback="blue"):
    mode = str(value or fallback).strip().lower()
    return mode if mode in VALID_ACCENT_MODES else fallback


def _load_theme_state_from_config(default_theme="light", default_accent="blue"):
    from UIClasses import load_theme_state_from_config

    return load_theme_state_from_config(
        section_name=THEME_CONFIG_SECTION,
        theme_key_name=THEME_CONFIG_THEME_KEY,
        accent_key_name=THEME_CONFIG_ACCENT_KEY,
        default_theme=default_theme,
        default_accent=default_accent,
    )


def _find_workspace_root(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        if os.path.isdir(os.path.join(current, "CEDLib.lib")):
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


THIS_DIR = os.path.abspath(os.path.dirname(__file__))
WORKSPACE_ROOT = _find_workspace_root(THIS_DIR)
if not WORKSPACE_ROOT:
    forms.alert("Could not locate workspace root for Alerts Browser.", title=TITLE)
    raise SystemExit

LIB_ROOT = os.path.abspath(os.path.join(WORKSPACE_ROOT, "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from CEDElectrical.Infrastructure.Revit.repositories.revit_circuit_repository import RevitCircuitRepository
from Snippets.circuit_ui_actions import (
    clear_revit_selection,
    collect_circuit_targets,
    set_revit_selection,
)
from UIClasses import pathing as ui_pathing
from UIClasses import resource_loader
from alerts_browser_services import build_snapshot
from alerts_browser_services import recalculate_and_snapshot

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
                self._logger.warning("Alerts Browser ExternalEvent raise failed: %s", ex)
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
            else:
                raise Exception("Unknown operation: {}".format(op_name))
        except Exception as ex:
            status = "error"
            error = ex
            if self._gateway._logger:
                self._gateway._logger.exception("Alerts Browser external operation failed: %s", ex)

        if callback:
            try:
                callback(status, op_name, result, error)
            except Exception:
                pass

    def GetName(self):  # noqa: N802
        return "CED Alerts Browser External Event"


class AlertsBrowserWindow(forms.WPFWindow):
    def __init__(self, theme_mode, accent_mode, snapshot, gateway):
        xaml = os.path.abspath(os.path.join(THIS_DIR, "AlertsBrowserWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        self._gateway = gateway
        self._items = []
        forms.WPFWindow.__init__(self, xaml)
        self._apply_theme()
        setattr(self, _WINDOW_MARKER, True)

        self._circuit_list = self.FindName("CircuitList")
        self._active_list = self.FindName("ActiveAlertsList")
        self._hidden_list = self.FindName("HiddenAlertsList")
        self._document_text = self.FindName("DocumentText")
        self._count_text = self.FindName("CircuitCountText")
        self._selected_circuit_text = self.FindName("SelectedCircuitText")
        self._selected_counts_text = self.FindName("SelectedCountsText")
        self._status_text = self.FindName("StatusText")
        self._refresh_button = self.FindName("RefreshButton")
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

    def _apply_snapshot(self, snapshot, preferred_circuit_id=None):
        data = dict(snapshot or {})
        doc_title = str(data.get("doc_title") or "-")
        items = list(data.get("items") or [])
        self._items = items
        if self._document_text is not None:
            self._document_text.Text = "Document: {}".format(doc_title)
        if self._count_text is not None:
            self._count_text.Text = "{} circuits with alerts".format(len(items))
        if self._circuit_list is not None:
            self._circuit_list.ItemsSource = list(items)
        selected = self._find_item_by_id(preferred_circuit_id)
        if selected is None:
            selected = self._selected_item()
            if selected not in items:
                selected = items[0] if items else None
        if self._circuit_list is not None:
            try:
                self._circuit_list.SelectedItem = selected
            except Exception:
                pass
        self._set_selected(selected)

    def _set_selected(self, item):
        if item is None:
            if self._selected_circuit_text is not None:
                self._selected_circuit_text.Text = "Select a circuit with alerts"
            if self._selected_counts_text is not None:
                self._selected_counts_text.Text = "Alerts: 0"
            if self._active_list is not None:
                self._active_list.ItemsSource = []
            if self._hidden_list is not None:
                self._hidden_list.ItemsSource = []
            self._update_refresh_state(None)
            return
        if self._selected_circuit_text is not None:
            self._selected_circuit_text.Text = "{} - {}".format(item.panel_ckt_text, item.load_name or "-")
        if self._selected_counts_text is not None:
            self._selected_counts_text.Text = item.counts_text
        if self._active_list is not None:
            self._active_list.ItemsSource = list(item.active_rows or [])
        if self._hidden_list is not None:
            self._hidden_list.ItemsSource = list(item.hidden_rows or [])
        self._update_refresh_state(item)

    def _update_refresh_state(self, item):
        if self._refresh_button is None:
            return
        if self._gateway is not None and self._gateway.is_busy():
            self._refresh_button.IsEnabled = False
            self._refresh_button.ToolTip = "Operation is running..."
            return
        if item is None:
            self._refresh_button.IsEnabled = False
            self._refresh_button.ToolTip = "Select a circuit first."
            return
        if bool(getattr(item, "recalc_blocked", False)):
            self._refresh_button.IsEnabled = False
            reason = getattr(item, "recalc_block_reason", "") or "Calculation blocked by ownership constraints."
            self._refresh_button.ToolTip = reason
            return
        self._refresh_button.IsEnabled = True
        self._refresh_button.ToolTip = "Recalculate selected circuit and refresh alerts."

    def _handle_external_complete(self, status, op_name, result, error):
        if status == "error":
            self._set_status("Operation failed")
            forms.alert("Alerts Browser operation failed:\n\n{}".format(error), title=TITLE)
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

    def select_equipment_clicked(self, sender, args):
        self._raise_select("panel")

    def select_circuit_clicked(self, sender, args):
        self._raise_select("circuit")

    def select_downstream_clicked(self, sender, args):
        self._raise_select("device")

    def clear_selection_clicked(self, sender, args):
        self._raise_select("clear")

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
            if bool(getattr(win, _WINDOW_MARKER, False)):
                return win
        except Exception:
            continue
    return None


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

