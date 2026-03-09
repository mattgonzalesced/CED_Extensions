# -*- coding: utf-8 -*-

import json
import os
import sys

import Autodesk.Revit.DB.Electrical as DBE
from System import Uri
from System.Windows import ResourceDictionary
from System.Windows.Media import BrushConverter
from pyrevit import forms, revit, DB

TITLE = "Alerts Browser"
ALERT_DATA_PARAM = "Circuit Data_CED"
HIDABLE_ALERT_IDS = {
    "Design.NonStandardOCPRating",
    "Design.BreakerLugSizeLimitOverride",
    "Design.BreakerLugQuantityLimitOverride",
    "Calculations.BreakerLugSizeLimit",
    "Calculations.BreakerLugQuantityLimit",
}
ACCENT_BRUSH_MAP = {
    "blue": {"light": "#0459A4", "dark": "#3B8ED8"},
    "red": {"light": "#BE202F", "dark": "#D15762"},
    "green": {"light": "#43A047", "dark": "#58B95E"},
    "neutral": {"light": "#5F6F82", "dark": "#9DB1C8"},
}


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

from CEDElectrical.Model.alerts import get_alert_definition

THEME_LIGHT_PATH = os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Themes", "CEDTheme.Light.xaml"))
THEME_DARK_PATH = os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Themes", "CEDTheme.Dark.xaml"))
BASE_RESOURCE_PATHS = (
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Themes", "CED.Sizes.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Themes", "CED.Colors.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Themes", "CED.Brushes.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Styles", "ButtonStyles.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Styles", "TextStyles.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Styles", "InputStyles.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Styles", "ListStyles.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Styles", "BadgeStyles.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Icons", "Icons.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Templates", "ListItems.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Templates", "Cards.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Templates", "Badges.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Templates", "DataGrids.xaml")),
    os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources", "Controls", "SearchBox.xaml")),
)


def _normalize_path(path):
    try:
        return os.path.abspath(path).replace("\\", "/").lower()
    except Exception:
        return str(path).replace("\\", "/").lower()


def _resource_source_path(dictionary):
    source = getattr(dictionary, "Source", None)
    if source is None:
        return ""
    try:
        return _normalize_path(source.LocalPath)
    except Exception:
        return _normalize_path(str(source))


def _load_resource_dictionary(path):
    if not path or not os.path.exists(path):
        return None
    try:
        dictionary = ResourceDictionary()
        dictionary.Source = Uri(path)
        return dictionary
    except Exception:
        return None


def _ensure_base_resources(owner):
    resources = getattr(owner, "Resources", None)
    if resources is None:
        return False
    merged = resources.MergedDictionaries
    existing = set([_resource_source_path(x) for x in list(merged) if _resource_source_path(x)])
    for path in BASE_RESOURCE_PATHS:
        normalized = _normalize_path(path)
        if not normalized or normalized in existing:
            continue
        dictionary = _load_resource_dictionary(path)
        if dictionary is None:
            continue
        merged.Add(dictionary)
        existing.add(normalized)
    return True


def _is_theme_resource_dictionary(dictionary):
    try:
        source_text = str(getattr(dictionary, "Source", "") or "").replace("\\", "/").lower()
    except Exception:
        source_text = ""
    return source_text.endswith("/cedtheme.light.xaml") or source_text.endswith("/cedtheme.dark.xaml")


def _to_brush(value, fallback=None):
    converter = BrushConverter()
    try:
        return converter.ConvertFrom(value)
    except Exception:
        if fallback is None:
            return None
        try:
            return converter.ConvertFrom(fallback)
        except Exception:
            return None


def _try_apply_accent(owner):
    resources = getattr(owner, "Resources", None)
    if resources is None:
        return False
    theme_mode = str(getattr(owner, "_theme_mode", "light") or "light").lower()
    accent_mode = str(getattr(owner, "_accent_mode", "blue") or "blue").lower()
    palette = ACCENT_BRUSH_MAP.get(accent_mode) or ACCENT_BRUSH_MAP.get("blue") or {}
    accent_hex = palette.get("dark") if theme_mode == "dark" else palette.get("light")
    accent_brush = _to_brush(accent_hex, "#0459A4")
    if accent_brush is None:
        return False
    resources["CED.Brush.Accent"] = accent_brush
    return True


def _try_apply_theme(owner):
    _ensure_base_resources(owner)
    mode = str(getattr(owner, "_theme_mode", "light") or "light").lower()
    theme_path = THEME_DARK_PATH if mode == "dark" else THEME_LIGHT_PATH
    dictionary = _load_resource_dictionary(theme_path)
    if dictionary is None and mode == "dark":
        dictionary = _load_resource_dictionary(THEME_LIGHT_PATH)
    if dictionary is None:
        return False
    resources = getattr(owner, "Resources", None)
    if resources is None:
        return False
    merged = resources.MergedDictionaries
    for existing in list(merged):
        if _is_theme_resource_dictionary(existing):
            merged.Remove(existing)
    merged.Add(dictionary)
    _try_apply_accent(owner)
    return True


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
    return alerts if isinstance(alerts, list) else []


def _payload_hidden_ids(payload):
    if not isinstance(payload, dict):
        return set()
    hidden = payload.get("hidden_definition_ids")
    if not isinstance(hidden, list):
        return set()
    return set([x for x in hidden if x])


class AlertRow(object):
    def __init__(self, severity, group, definition_id, message, is_hidden=False, can_hide=True):
        self.severity = str(severity or "NONE")
        self.group = str(group or "Other")
        self.definition_id = str(definition_id or "-")
        self.message = str(message or "")
        self.is_hidden = bool(is_hidden)
        self.can_hide = bool(can_hide)


def _alert_rows_from_payload(payload):
    hidden_ids = _payload_hidden_ids(payload)
    rows = []
    for item in _payload_alert_records(payload):
        if not isinstance(item, dict):
            continue
        definition_id = item.get("definition_id") or item.get("id")
        severity = str(item.get("severity") or "NONE").upper()
        group = str(item.get("group") or "Other")
        definition = get_alert_definition(definition_id) if definition_id else None
        message_value = item.get("message")
        if message_value:
            if isinstance(message_value, (dict, list)):
                try:
                    text = json.dumps(message_value, ensure_ascii=False)
                except Exception:
                    text = str(message_value)
            else:
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


def _active_doc():
    doc = getattr(revit, "doc", None)
    if doc is not None:
        return doc
    try:
        return __revit__.ActiveUIDocument.Document
    except Exception:
        return None


class AlertCircuitItem(object):
    def __init__(self, circuit, rows):
        self.circuit = circuit
        self.panel = "No Panel"
        try:
            if circuit.BaseEquipment:
                self.panel = getattr(circuit.BaseEquipment, "Name", self.panel) or self.panel
        except Exception:
            pass
        self.circuit_number = getattr(circuit, "CircuitNumber", "") or ""
        self.load_name = getattr(circuit, "LoadName", "") or ""
        self.panel_ckt_text = "{} / {}".format(self.panel or "-", self.circuit_number or "-")
        self.rows = list(rows or [])
        self.active_rows = [x for x in self.rows if not bool(getattr(x, "is_hidden", False))]
        self.hidden_rows = [x for x in self.rows if bool(getattr(x, "is_hidden", False))]
        self.total_count = len(self.rows)
        self.active_count = len(self.active_rows)
        self.hidden_count = len(self.hidden_rows)
        self.counts_text = "Alerts: {} | Active: {} | Hidden: {}".format(
            self.total_count,
            self.active_count,
            self.hidden_count,
        )


class AlertsBrowserWindow(forms.WPFWindow):
    def __init__(self, theme_mode="light", accent_mode="blue"):
        xaml = os.path.abspath(os.path.join(THIS_DIR, "AlertsBrowserWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        forms.WPFWindow.__init__(self, xaml)
        _try_apply_theme(self)
        self._items = []
        self._circuit_list = self.FindName("CircuitList")
        self._active_list = self.FindName("ActiveAlertsList")
        self._hidden_list = self.FindName("HiddenAlertsList")
        self._document_text = self.FindName("DocumentText")
        self._count_text = self.FindName("CircuitCountText")
        self._selected_circuit_text = self.FindName("SelectedCircuitText")
        self._selected_counts_text = self.FindName("SelectedCountsText")
        self._load_items()

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
            return
        if self._selected_circuit_text is not None:
            self._selected_circuit_text.Text = "{} - {}".format(item.panel_ckt_text, item.load_name or "-")
        if self._selected_counts_text is not None:
            self._selected_counts_text.Text = item.counts_text
        if self._active_list is not None:
            self._active_list.ItemsSource = list(item.active_rows or [])
        if self._hidden_list is not None:
            self._hidden_list.ItemsSource = list(item.hidden_rows or [])

    def _load_items(self):
        doc = _active_doc()
        if doc is None:
            forms.alert("Open a model document first.", title=TITLE)
            self.Close()
            return

        if self._document_text is not None:
            try:
                self._document_text.Text = "Document: {}".format(doc.Title or "-")
            except Exception:
                self._document_text.Text = "Document: -"

        circuits = list(
            DB.FilteredElementCollector(doc)
            .OfClass(DBE.ElectricalSystem)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        circuits.sort(
            key=lambda c: (
                (getattr(getattr(c, "BaseEquipment", None), "Name", "") or ""),
                (getattr(c, "StartSlot", 0) or 0),
                (getattr(c, "LoadName", "") or ""),
            )
        )

        items = []
        for circuit in circuits:
            rows = _alert_rows_from_payload(_read_alert_payload(circuit))
            if not rows:
                continue
            items.append(AlertCircuitItem(circuit, rows))

        self._items = items
        if self._circuit_list is not None:
            self._circuit_list.ItemsSource = list(items)

        if self._count_text is not None:
            self._count_text.Text = "{} circuits with alerts".format(len(items))

        if self._circuit_list is not None and items:
            try:
                self._circuit_list.SelectedIndex = 0
            except Exception:
                pass
        else:
            self._set_selected(None)

    def circuit_selection_changed(self, sender, args):
        selected = None
        try:
            selected = getattr(self._circuit_list, "SelectedItem", None)
        except Exception:
            selected = None
        self._set_selected(selected)

    def refresh_clicked(self, sender, args):
        self._load_items()

    def close_clicked(self, sender, args):
        self.Close()


def _current_browser_theme_accent():
    module = sys.modules.get("ced_circuit_browser_panel")
    if module is None:
        return "light", "blue"
    panel_cls = getattr(module, "CircuitBrowserPanel", None)
    if panel_cls is None or not hasattr(panel_cls, "get_instance"):
        return "light", "blue"
    panel = panel_cls.get_instance()
    if panel is None:
        return "light", "blue"
    return (
        str(getattr(panel, "_theme_mode", "light") or "light"),
        str(getattr(panel, "_accent_mode", "blue") or "blue"),
    )


theme_mode, accent_mode = _current_browser_theme_accent()
AlertsBrowserWindow(theme_mode=theme_mode, accent_mode=accent_mode).ShowDialog()
