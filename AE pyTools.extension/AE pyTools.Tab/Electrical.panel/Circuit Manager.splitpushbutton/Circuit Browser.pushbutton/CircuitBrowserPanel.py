# -*- coding: utf-8 -*-

import imp
import json
import os
import sys

import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.DB.Events import (
    DocumentOpeningEventArgs,
    DocumentOpenedEventArgs,
    DocumentClosingEventArgs,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
from System import EventHandler
from System import Uri
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import ResourceDictionary
from System.Windows import Visibility
from System.Windows.Controls import (
    ContextMenu,
    MenuItem,
    Separator,
    DataGridRow,
    ListViewItem,
    Button,
    DataGridTextColumn,
)
from System.Windows.Media import BrushConverter, VisualTreeHelper
from pyrevit import forms, revit, DB, script

_THIS_DIR = os.path.abspath(os.path.dirname(__file__))


def _find_workspace_root(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        if os.path.isdir(os.path.join(current, "CEDLib.lib")):
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


def _find_named_ancestor(start_dir, folder_name):
    current = os.path.abspath(start_dir)
    target = (folder_name or "").strip().lower()
    while True:
        if os.path.basename(current).lower() == target:
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


_WORKSPACE_ROOT = _find_workspace_root(_THIS_DIR)
if _WORKSPACE_ROOT:
    LIB_ROOT = os.path.abspath(os.path.join(_WORKSPACE_ROOT, "CEDLib.lib"))
else:
    LIB_ROOT = os.path.abspath(os.path.join(_THIS_DIR, "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from CEDElectrical.Model.alerts import get_alert_definition
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Domain import settings_manager
from CEDElectrical.refdata.standard_ocp_table import BREAKER_FRAME_SWITCH_TABLE
from CEDElectrical.Infrastructure.Revit.external_events.circuit_operation_event import (
    CircuitOperationExternalEventGateway,
)
from UIClasses.revit_theme_bridge import DOCK_PANE_FRAME_DARK, DOCK_PANE_FRAME_LIGHT, RevitThemeBridge


TITLE = "Circuit Browser"
PANEL_ID = "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2"
ALERT_DATA_PARAM = "Circuit Data_CED"
_ELECTRICAL_PANEL_ROOT = _find_named_ancestor(_THIS_DIR, "Electrical.panel") or os.path.abspath(
    os.path.join(_THIS_DIR, "..", "..")
)
CALC_SETTINGS_PATH = os.path.abspath(
    os.path.join(
        _ELECTRICAL_PANEL_ROOT,
        "Circuits2.stack",
        "Calculate Circuits.pushbutton",
        "config.py",
    )
)
CALC_SETTINGS_XAML_PATH = os.path.abspath(
    os.path.join(
        _ELECTRICAL_PANEL_ROOT,
        "Circuits2.stack",
        "Calculate Circuits.pushbutton",
        "settings.xaml",
    )
)
THEME_LIGHT_PATH = os.path.abspath(
    os.path.join(
        LIB_ROOT,
        "UIClasses",
        "Resources",
        "Themes",
        "CEDTheme.Light.xaml",
    )
)
THEME_DARK_PATH = os.path.abspath(
    os.path.join(
        LIB_ROOT,
        "UIClasses",
        "Resources",
        "Themes",
        "CEDTheme.Dark.xaml",
    )
)
CURRENT_THEME_MODE = "light"
CURRENT_ACCENT_MODE = "blue"
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
HIDABLE_ALERT_IDS = {
    "Design.NonStandardOCPRating",
    "Design.BreakerLugSizeLimitOverride",
    "Design.BreakerLugQuantityLimitOverride",
    "Calculations.BreakerLugSizeLimit",
    "Calculations.BreakerLugQuantityLimit",
}
BLOCKED_BRANCH_TYPES = set(['N/A', 'SPACE', 'SPARE', 'CONDUIT ONLY'])
IG_BREAKER_ALLOWED_TYPES = set(['BRANCH', 'FEEDER', 'XFMR PRI', 'XFMR SEC'])

ACCENT_BRUSH_MAP = {
    "blue": {"light": "#0459A4", "dark": "#3B8ED8"},
    "red": {"light": "#BE202F", "dark": "#D15762"},
    "green": {"light": "#43A047", "dark": "#58B95E"},
    "neutral": {"light": "#5F6F82", "dark": "#9DB1C8"},
}
CIRCUIT_TYPE_TAG_STYLES = {
    "BRANCH": ("CED.Brush.BadgeStd01Background", "CED.Brush.BadgeStd01Text"),
    "FEEDER": ("CED.Brush.BadgeStd02Background", "CED.Brush.BadgeStd02Text"),
    "SPARE": ("CED.Brush.BadgeStd03Background", "CED.Brush.BadgeStd03Text"),
    "SPACE": ("CED.Brush.BadgeStd04Background", "CED.Brush.BadgeStd04Text"),
    "XFMR PRI": ("CED.Brush.BadgeStd05Background", "CED.Brush.BadgeStd05Text"),
    "XFMR SEC": ("CED.Brush.BadgeStd06Background", "CED.Brush.BadgeStd06Text"),
    "CONDUIT ONLY": ("CED.Brush.BadgeStd07Background", "CED.Brush.BadgeStd07Text"),
    "N/A": ("CED.Brush.BadgeStd08Background", "CED.Brush.BadgeStd08Text"),
}
OCP_TABLE_KEYS = sorted([int(k) for k in BREAKER_FRAME_SWITCH_TABLE.keys()])
_DOC_SENTINEL = object()


def _fmt_number(value, digits=1):
    try:
        return str(round(float(value), digits))
    except Exception:
        return "-"


def _fmt_amp(value, digits=0):
    try:
        numeric = float(value)
    except Exception:
        return "-"
    if digits <= 0:
        return "{} A".format(int(round(numeric, 0)))
    return "{} A".format(round(numeric, digits))


def _parse_whole_amps(value):
    if value is None:
        return None
    text = str(value).strip().upper()
    if not text or text == "-":
        return None
    text = text.replace("AMPS", "").replace("AMP", "").replace("A", "").strip()
    if not text:
        return None
    try:
        numeric = float(text)
    except Exception:
        return None
    rounded = int(round(numeric, 0))
    if abs(numeric - rounded) > 0.000001:
        return None
    if rounded <= 0:
        return None
    return rounded


def _safe_float(value):
    try:
        return float(value)
    except Exception:
        return None


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


def _normalize_path(path):
    try:
        return os.path.abspath(path).replace("\\", "/").lower()
    except Exception:
        try:
            return str(path).replace("\\", "/").lower()
        except Exception:
            return ""


def _resource_source_path(dictionary):
    source = getattr(dictionary, "Source", None)
    if source is None:
        return ""
    try:
        local_path = source.LocalPath
        if local_path:
            return _normalize_path(local_path)
    except Exception:
        pass
    try:
        return _normalize_path(str(source))
    except Exception:
        return ""


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
    try:
        resources = getattr(owner, "Resources", None)
        if resources is None:
            return False
        merged = resources.MergedDictionaries
        existing = set()
        for dictionary in list(merged):
            source_path = _resource_source_path(dictionary)
            if source_path:
                existing.add(source_path)
        added = False
        for path in BASE_RESOURCE_PATHS:
            normalized = _normalize_path(path)
            if not normalized or normalized in existing:
                continue
            dictionary = _load_resource_dictionary(path)
            if dictionary is None:
                continue
            merged.Add(dictionary)
            existing.add(normalized)
            added = True
        return added
    except Exception:
        return False


def _load_theme_dictionary(theme_mode="light"):
    mode = (theme_mode or "light").strip().lower()
    candidates = (THEME_DARK_PATH, THEME_LIGHT_PATH) if mode == "dark" else (THEME_LIGHT_PATH,)
    for path in candidates:
        dictionary = _load_resource_dictionary(path)
        if dictionary is not None:
            return dictionary
    return None


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
    try:
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
    except Exception:
        return False


def _try_apply_theme(owner):
    try:
        _ensure_base_resources(owner)
        theme_mode = getattr(owner, "_theme_mode", "light")
        dictionary = _load_theme_dictionary(theme_mode)
        if dictionary is None:
            return False
        resources = getattr(owner, "Resources", None)
        if resources is None:
            return False
        merged = resources.MergedDictionaries
        previous = getattr(owner, "_ced_theme_dictionary", None)
        try:
            if previous is not None and previous in merged:
                merged.Remove(previous)
        except Exception:
            pass
        try:
            for existing in list(merged):
                if _is_theme_resource_dictionary(existing):
                    merged.Remove(existing)
        except Exception:
            pass
        try:
            merged.Add(dictionary)
        except Exception:
            merged.Add(dictionary)
        owner._ced_theme_dictionary = dictionary
        _try_apply_accent(owner)
        return True
    except Exception:
        return False


def _try_find_resource(owner, key):
    if owner is None or not key:
        return None
    try:
        resource = owner.TryFindResource(key)
        if resource is not None:
            return resource
    except Exception:
        pass
    try:
        return owner.FindResource(key)
    except Exception:
        return None


def _set_if_resource(owner, target, property_name, key):
    if target is None:
        return
    resource = _try_find_resource(owner, key)
    if resource is None:
        return
    try:
        setattr(target, property_name, resource)
    except Exception:
        pass


def _column_at(grid, index):
    try:
        if grid is None:
            return None
        columns = grid.Columns
        if columns is None or index < 0 or index >= columns.Count:
            return None
        return columns[index]
    except Exception:
        return None


def _apply_grid_styles_neutral_ig(owner, grid):
    if grid is None:
        return
    _set_if_resource(owner, grid, "Style", "CED.DataGrid.Base")
    _set_if_resource(owner, grid, "ColumnHeaderStyle", "HeaderWrapCenter")
    _set_if_resource(owner, grid, "RowStyle", "CED.DataGrid.RowDisabledAware")

    _set_if_resource(owner, _column_at(grid, 0), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 1), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 2), "CellStyle", "DividerThickReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 3), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 4), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 5), "CellStyle", "DividerThickReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 6), "CellStyle", "NewQtyCell")
    _set_if_resource(owner, _column_at(grid, 7), "CellStyle", "NewSizeCell")
    _set_if_resource(owner, _column_at(grid, 8), "CellStyle", "NewWireCell")
    _set_if_resource(owner, _column_at(grid, 9), "CellStyle", "ReadonlyCell")

    for index in (2, 3, 4, 6, 7):
        column = _column_at(grid, index)
        if isinstance(column, DataGridTextColumn):
            _set_if_resource(owner, column, "ElementStyle", "CenterTextCell")
    remarks_column = _column_at(grid, 9)
    if isinstance(remarks_column, DataGridTextColumn):
        _set_if_resource(owner, remarks_column, "ElementStyle", "RemarksTextCell")


def _apply_grid_styles_breaker(owner, grid):
    if grid is None:
        return
    _set_if_resource(owner, grid, "Style", "CED.DataGrid.Base")
    _set_if_resource(owner, grid, "ColumnHeaderStyle", "HeaderWrapCenterTall")
    _set_if_resource(owner, grid, "RowStyle", "CED.DataGrid.RowDisabledAware")

    _set_if_resource(owner, _column_at(grid, 0), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 1), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 2), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 3), "CellStyle", "DividerThickReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 4), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 5), "CellStyle", "DividerThickReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 6), "CellStyle", "NewValueCell")
    _set_if_resource(owner, _column_at(grid, 7), "CellStyle", "NewFrameCell")
    _set_if_resource(owner, _column_at(grid, 8), "CellStyle", "ReadonlyCell")

    for index in (2, 3, 4, 5, 6, 7):
        column = _column_at(grid, index)
        if isinstance(column, DataGridTextColumn):
            _set_if_resource(owner, column, "ElementStyle", "CenterTextCell")
    for index in (6, 7):
        column = _column_at(grid, index)
        if isinstance(column, DataGridTextColumn):
            _set_if_resource(owner, column, "EditingElementStyle", "CenterTextEdit")
    remarks_column = _column_at(grid, 8)
    if isinstance(remarks_column, DataGridTextColumn):
        _set_if_resource(owner, remarks_column, "ElementStyle", "RemarksTextCell")


def _apply_grid_styles_mark_existing(owner, grid):
    if grid is None:
        return
    _set_if_resource(owner, grid, "Style", "CED.DataGrid.Base")
    _set_if_resource(owner, grid, "ColumnHeaderStyle", "HeaderWrapCenterTall")
    _set_if_resource(owner, grid, "RowStyle", "CED.DataGrid.RowDisabledAware")

    _set_if_resource(owner, _column_at(grid, 0), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 1), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 2), "CellStyle", "DividerThickReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 3), "CellStyle", "ReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 4), "CellStyle", "DividerThickReadonlyCell")
    _set_if_resource(owner, _column_at(grid, 5), "CellStyle", "NewNotesCell")
    _set_if_resource(owner, _column_at(grid, 6), "CellStyle", "NewWireStringCell")
    _set_if_resource(owner, _column_at(grid, 7), "CellStyle", "ReadonlyCell")

    for index in (2, 3, 5):
        column = _column_at(grid, index)
        if isinstance(column, DataGridTextColumn):
            _set_if_resource(owner, column, "ElementStyle", "CenterTextCell")
    remarks_column = _column_at(grid, 7)
    if isinstance(remarks_column, DataGridTextColumn):
        _set_if_resource(owner, remarks_column, "ElementStyle", "RemarksTextCell")


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


def _lookup_schedule_notes_text(circuit):
    try:
        param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
    except Exception:
        param = None
    if param is not None:
        try:
            value = param.AsString()
            if value is None:
                value = param.AsValueString()
            if value:
                return value
        except Exception:
            pass
    fallback = _lookup_param_text(circuit, "CKT_Schedule Notes_CEDT")
    return fallback or ""


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

        tag_bg, tag_fg = CIRCUIT_TYPE_TAG_STYLES.get(
            self.branch_type,
            ("CED.Brush.BadgeStd04Background", "CED.Brush.BadgeStd04Text"),
        )
        self.type_tag_text = self.branch_type
        self.type_tag_bg_key = tag_bg
        self.type_tag_fg_key = tag_fg
        self.type_tag_bg = tag_bg
        self.type_tag_fg = tag_fg
        self.show_type_tag = True
        self.type_tag_visibility = "Visible"

        user_override = _lookup_param_value(circuit, "CKT_User Override_CED")
        has_override = False
        try:
            has_override = int(user_override or 0) == 1
        except Exception:
            has_override = False
        self.has_override = has_override
        self.override_badge_visibility = "Visible" if has_override else "Collapsed"

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
        self.severity = str(severity or "NONE")
        self.group = str(group or "Other")
        self.definition_id = str(definition_id or "-")
        self.message = str(message or "")
        self.is_hidden = bool(is_hidden)
        self.can_hide = bool(can_hide)


class CircuitAlertsWindow(forms.WPFWindow):
    def __init__(self, circuit_label, rows, theme_mode="light", accent_mode="blue"):
        xaml = os.path.abspath(os.path.join(_THIS_DIR, "CircuitAlertsWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        forms.WPFWindow.__init__(self, xaml)
        _try_apply_theme(self)
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
        try:
            if self._active_list is not None:
                self._active_list.SelectedItem = None
            if self._hidden_list is not None:
                self._hidden_list.SelectedItem = None
        except Exception:
            pass

    def alerts_list_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, ListViewItem) is not None:
            return
        try:
            sender.SelectedItem = None
        except Exception:
            pass

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
    def __init__(self, circuit, load_name):
        self.circuit = str(circuit or "-")
        self.load_name = str(load_name or "-")


class RuntimeAlertRow(object):
    def __init__(self, circuit, load_name):
        self.circuit = str(circuit or "-")
        self.load_name = str(load_name or "-")


class CircuitRunSummaryWindow(forms.WPFWindow):
    def __init__(self, locked_rows, runtime_rows, theme_mode="light", accent_mode="blue"):
        xaml = os.path.abspath(os.path.join(_THIS_DIR, "CircuitRunSummaryWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        forms.WPFWindow.__init__(self, xaml)
        _try_apply_theme(self)
        locked = [
            LockedRow(
                x.get("circuit", "") or "-",
                x.get("load_name", "") or "-",
            )
            for x in list(locked_rows or [])
        ]

        def _runtime_circuit_text(raw):
            panel = str(raw.get("panel", "") or "").strip()
            number = str(raw.get("number", "") or "").strip()
            if panel and number:
                return "{}-{}".format(panel, number)
            if panel:
                return panel
            if number:
                return number
            return "-"

        runtime = [
            RuntimeAlertRow(
                _runtime_circuit_text(x),
                x.get("load_name", "") or "-",
            )
            for x in list(runtime_rows or [])
        ]

        locked_list = self.FindName("LockedList")
        runtime_list = self.FindName("RuntimeList")
        self._locked_list = locked_list
        self._runtime_list = runtime_list
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

    def window_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if self._runtime_list is not None and _is_descendant_of_control(source, self._runtime_list):
            return
        if self._locked_list is not None and _is_descendant_of_control(source, self._locked_list):
            return
        try:
            if self._runtime_list is not None:
                self._runtime_list.SelectedItem = None
            if self._locked_list is not None:
                self._locked_list.SelectedItem = None
        except Exception:
            pass

    def summary_list_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, ListViewItem) is not None:
            return
        if _find_visual_ancestor(source, DataGridRow) is not None:
            return
        try:
            sender.SelectedItem = None
        except Exception:
            pass


class NeutralIGActionRow(object):
    def __init__(self, item, is_enabled, reason, current_qty, current_size, current_wire):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = item.circuit.Id.IntegerValue
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.panel_ckt_text = "{} / {}".format(item.panel or "-", item.circuit_number or "-")
        self.branch_type = item.branch_type
        self.is_enabled = bool(is_enabled)
        self.is_checked = False
        self.reason = reason or ""

        self.current_qty = current_qty
        self.current_size = current_size or ""
        self.current_wire = current_wire or ""

        self.new_qty = current_qty
        self.new_size = current_size or ""
        self.new_wire = "no change"
        self.new_wire_font_style = "Italic"
        self.new_qty_changed = False
        self.new_size_changed = False
        self.new_wire_changed = False
        self.is_changed = False
        self.remarks = ""
        self.recompute_state()

    def recompute_state(self):
        self.is_changed = bool(self.new_qty_changed or self.new_size_changed or self.new_wire_changed)
        if not self.is_enabled:
            self.remarks = "Blocked - {}".format(self.reason or "Unsupported")
        elif self.is_changed:
            self.remarks = "Will be modified"
        else:
            self.remarks = "No change"


class NeutralIGActionWindow(forms.WPFWindow):
    def __init__(self, title, rows, preview_callback, apply_callback, theme_mode="light", accent_mode="blue"):
        xaml = os.path.abspath(os.path.join(_THIS_DIR, "CircuitNeutralIGActionWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        forms.WPFWindow.__init__(self, xaml)
        _try_apply_theme(self)
        self._rows = list(rows or [])
        self._preview_callback = preview_callback
        self._apply_callback = apply_callback
        self._mode = "add"
        self._is_syncing_checks = False
        self._suppress_check_events = False
        self._is_ready = False
        self._show_unsupported = True
        self._last_selected_rows = []

        title_tb = self.FindName("TitleText")
        if title_tb is not None:
            title_tb.Text = title

        self._grid = self.FindName("CircuitGrid")
        _apply_grid_styles_neutral_ig(self, self._grid)
        self._status = self.FindName("ChangedStatusText")
        self._checked_status = self.FindName("CheckedStatusText")
        self._add_btn = self.FindName("AddButton")
        self._remove_btn = self.FindName("RemoveButton")
        self._show_unsupported_cb = self.FindName("ShowUnsupportedToggle")
        if self._show_unsupported_cb is not None:
            self._show_unsupported_cb.IsChecked = True
        for row in self._rows:
            row.is_checked = False
        self._apply_visibility_filter()
        self.Loaded += self.window_loaded
        self._refresh_grid()

    def window_loaded(self, sender, args):
        self._is_ready = True
        self._sync_action_buttons()

    def _apply_visibility_filter(self):
        if self._grid is None:
            return
        if self._show_unsupported:
            items = list(self._rows)
        else:
            items = [x for x in self._rows if bool(getattr(x, "is_enabled", False))]
        self._grid.ItemsSource = items

    def _refresh_grid(self, refresh_items=False):
        if refresh_items:
            self._apply_visibility_filter()
        elif self._grid is not None:
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass
        self._refresh_status()

    def _refresh_status(self):
        changed = len([x for x in self._rows if x.is_changed and x.is_enabled and x.is_checked])
        if self._show_unsupported:
            total = len(self._rows)
        else:
            total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
        if self._status is not None:
            self._status.Text = "{} of {} circuits to be modified.".format(changed, total)
        if self._checked_status is not None:
            checkable_total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
            checked = len([x for x in self._rows if x.is_enabled and x.is_checked])
            self._checked_status.Text = "{} of {} checked".format(checked, checkable_total)
        self._sync_action_buttons()

    def _sync_action_buttons(self):
        checked_count = len([x for x in self._rows if x.is_enabled and x.is_checked])
        if self._add_btn is not None:
            self._add_btn.IsEnabled = checked_count > 0
        if self._remove_btn is not None:
            self._remove_btn.IsEnabled = checked_count > 0

    def _apply_checkbox_to_selected(self, sender, state):
        if not self._is_ready or self._is_syncing_checks or self._suppress_check_events:
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
        if len(selected) > 1 and row in selected:
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
        self._suppress_check_events = True
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def item_checked(self, sender, args):
        self._apply_checkbox_to_selected(sender, True)

    def item_unchecked(self, sender, args):
        self._apply_checkbox_to_selected(sender, False)

    def item_checkbox_clicked(self, sender, args):
        self._apply_checkbox_to_selected(sender, bool(getattr(sender, "IsChecked", False)))

    def check_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for row in self._rows:
                row.is_checked = bool(row.is_enabled)
        finally:
            self._is_syncing_checks = False
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def uncheck_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def add_clicked(self, sender, args):
        self._mode = "add"
        self._preview_callback(self._rows, "add")
        self._refresh_grid(refresh_items=False)

    def remove_clicked(self, sender, args):
        self._mode = "remove"
        self._preview_callback(self._rows, "remove")
        self._refresh_grid(refresh_items=False)

    def apply_clicked(self, sender, args):
        if self._apply_callback(self._rows, self._mode):
            self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def show_unsupported_toggled(self, sender, args):
        self._show_unsupported = bool(getattr(sender, "IsChecked", False))
        self._refresh_grid(refresh_items=True)

    def grid_selection_changed(self, sender, args):
        try:
            selected = list(self._grid.SelectedItems or [])
        except Exception:
            selected = []
        if selected:
            self._last_selected_rows = selected

    def _clear_grid_selection(self):
        if self._grid is None:
            return
        try:
            self._grid.UnselectAll()
        except Exception:
            pass

    def window_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if self._grid is None or source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if not _is_descendant_of_control(source, self._grid):
            self._clear_grid_selection()

    def grid_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, DataGridRow) is None:
            self._clear_grid_selection()


class BreakerActionRow(object):
    def __init__(self, item, is_enabled, reason, load_current, cur_rating, cur_frame, auto_rating, auto_frame_from_cur):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = item.circuit.Id.IntegerValue
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.panel_ckt_text = "{} / {}".format(item.panel or "-", item.circuit_number or "-")
        self.branch_type = item.branch_type
        self.is_enabled = bool(is_enabled)
        self.is_checked = False
        self.reason = reason or ""

        self._load_current_value = load_current
        self._current_rating_value = cur_rating
        self._current_frame_value = cur_frame
        self._autosized_rating_value = auto_rating
        self.current_load_current = _fmt_amp(load_current, 2)
        self.current_rating = _fmt_amp(cur_rating, 0)
        self.current_frame = _fmt_amp(cur_frame, 0)
        self.autosized_rating = _fmt_amp(auto_rating, 0)
        self.autosized_frame = _fmt_amp(auto_frame_from_cur, 0)
        self._auto_frame_from_cur = auto_frame_from_cur
        self._auto_frame_from_auto = auto_frame_from_cur
        self.new_rating = self.current_rating
        self.new_frame = self.current_frame
        self.new_rating_changed = False
        self.new_frame_changed = False
        self.new_frame_warning = False
        self.is_changed = False
        self.rating_warning_brush = "#00FFFFFF"
        self.rating_warning_tooltip = ""
        self.rating_warning_visibility = "Collapsed"
        self.rating_warning_level = 0
        self.remarks = ""
        self.recompute_state()

    def set_auto_frames(self, from_current, from_auto):
        self._auto_frame_from_cur = from_current
        self._auto_frame_from_auto = from_auto

    def _warning_data(self):
        rating_value = _parse_whole_amps(self.new_rating)
        if rating_value is None:
            return 0, None, "", "Collapsed", ""
        try:
            load_value = float(self._load_current_value)
        except Exception:
            return 0, None, "", "Collapsed", ""
        if rating_value <= 0:
            return 0, None, "", "Collapsed", ""
        pct = float(load_value) / float(rating_value)
        if pct > 1.0:
            return 2, "#BE202F", "Load exceeds breaker rating", "Visible", "Load exceeds breaker rating"
        if pct > 0.90:
            return 1, "#F08A00", "Load greater than 90% of breaker rating", "Visible", "Load > 90% of rating"
        if pct > 0.80:
            return 0, "#D8B300", "Load greater than 80% of breaker rating", "Visible", "Load > 80% of rating"
        return 0, None, "", "Collapsed", ""

    def recompute_state(self):
        self.new_rating_changed = _parse_whole_amps(self.new_rating) != _parse_whole_amps(self.current_rating)
        self.new_frame_changed = _parse_whole_amps(self.new_frame) != _parse_whole_amps(self.current_frame)
        self.is_changed = self.new_rating_changed or self.new_frame_changed

        warning_level, warning_brush, warning_tooltip, warning_vis, warning_short = self._warning_data()
        self.rating_warning_level = int(warning_level or 0)
        self.rating_warning_brush = warning_brush or "#00FFFFFF"
        self.rating_warning_tooltip = warning_tooltip
        self.rating_warning_visibility = warning_vis

        rating_value = _parse_whole_amps(self.new_rating)
        frame_value = _parse_whole_amps(self.new_frame)
        self.new_frame_warning = bool(
            rating_value is not None and frame_value is not None and int(frame_value) < int(rating_value)
        )

        if not self.is_enabled:
            self.remarks = "Blocked - {}".format(self.reason or "Unsupported")
        else:
            lines = []
            if self.new_rating_changed or warning_short:
                breaker_line = "Breaker will be modified" if self.new_rating_changed else "No change"
                if warning_short:
                    breaker_line += " - WARNING: {}".format(warning_short)
                lines.append(breaker_line)
            if self.new_frame_changed or self.new_frame_warning:
                frame_line = "Frame will be modified" if self.new_frame_changed else "No change"
                if self.new_frame_warning:
                    frame_line += " - WARNING: Frame is below breaker rating"
                lines.append(frame_line)
            if not lines:
                lines.append("No change")
            self.remarks = "\n".join(lines[:2])


class BreakerActionWindow(forms.WPFWindow):
    def __init__(self, rows, preview_apply_callback, apply_callback, theme_mode="light", accent_mode="blue"):
        xaml = os.path.abspath(os.path.join(_THIS_DIR, "CircuitBreakerActionWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        forms.WPFWindow.__init__(self, xaml)
        _try_apply_theme(self)
        self._rows = list(rows or [])
        self._preview_apply_callback = preview_apply_callback
        self._apply_callback = apply_callback
        self._is_syncing_checks = False
        self._suppress_check_events = False
        self._is_ready = False
        self._show_unsupported = True
        self._last_selected_rows = []

        self._grid = self.FindName("CircuitGrid")
        _apply_grid_styles_breaker(self, self._grid)
        self._status = self.FindName("ChangedStatusText")
        self._checked_status = self.FindName("CheckedStatusText")
        self._set_breaker_cb = self.FindName("AutoBreakerCheck")
        self._set_frame_cb = self.FindName("AutoFrameCheck")
        self._allow_15a_cb = self.FindName("Allow15AToggle")
        self._upsize_only_cb = self.FindName("UpsizeOnlyToggle")
        self._show_unsupported_cb = self.FindName("ShowUnsupportedToggle")
        self._show_changed_only_cb = self.FindName("ShowChangedOnlyToggle")
        self._autosize_btn = self.FindName("AutosizePreviewButton")
        self._reset_btn = self.FindName("ResetButton")
        self._apply_btn = self.FindName("ApplyButton")
        self._row_filter_mode = "all"

        if self._show_unsupported_cb is not None:
            self._show_unsupported_cb.IsChecked = True
        if self._show_changed_only_cb is not None:
            self._show_changed_only_cb.IsChecked = False
        for row in self._rows:
            row.is_checked = False
        self.Loaded += self.window_loaded
        self._refresh_status()
        self._sync_button_states()

    def window_loaded(self, sender, args):
        self._is_ready = True
        self._apply_visibility_filter()
        self._refresh_status()
        self._sync_button_states()

    def _apply_visibility_filter(self):
        if self._grid is None:
            return
        if self._show_unsupported:
            items = list(self._rows)
        else:
            items = [x for x in self._rows if bool(getattr(x, "is_enabled", False))]
        if self._row_filter_mode == "changed":
            items = [x for x in items if bool(getattr(x, "is_changed", False))]
        self._grid.ItemsSource = items

    def _refresh_grid(self, refresh_items=False):
        if refresh_items:
            self._suppress_check_events = True
            try:
                self._apply_visibility_filter()
            finally:
                self._suppress_check_events = False
        elif self._grid is not None:
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass
        self._refresh_status()
        self._sync_button_states()

    def _refresh_status(self):
        changed = len([x for x in self._rows if x.is_changed and x.is_enabled])
        total = len(self._rows)
        if self._status is not None:
            self._status.Text = "{} of {} circuits to be modified".format(changed, total)
        if self._checked_status is not None:
            checkable_total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
            checked = len([x for x in self._rows if x.is_enabled and x.is_checked])
            self._checked_status.Text = "{} of {} checked".format(checked, checkable_total)

    def _selected_rows(self):
        try:
            return list(self._grid.SelectedItems or [])
        except Exception:
            return []

    def _sync_button_states(self):
        checked_count = len([x for x in self._rows if x.is_enabled and x.is_checked])
        selected_count = len(self._selected_rows())
        changed_count = len([x for x in self._rows if x.is_enabled and x.is_changed])
        if self._autosize_btn is not None:
            self._autosize_btn.IsEnabled = checked_count > 0
        if self._reset_btn is not None:
            self._reset_btn.IsEnabled = selected_count > 0
        if self._apply_btn is not None:
            self._apply_btn.IsEnabled = changed_count > 0

    def _apply_checkbox_to_selected(self, sender, state):
        if not self._is_ready or self._is_syncing_checks or self._suppress_check_events:
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
        if len(selected) > 1 and row in selected:
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
        self._suppress_check_events = True
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def item_checked(self, sender, args):
        self._apply_checkbox_to_selected(sender, True)

    def item_unchecked(self, sender, args):
        self._apply_checkbox_to_selected(sender, False)

    def item_checkbox_clicked(self, sender, args):
        self._apply_checkbox_to_selected(sender, bool(getattr(sender, "IsChecked", False)))

    def check_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for row in self._rows:
                row.is_checked = bool(row.is_enabled)
        finally:
            self._is_syncing_checks = False
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False
        self._refresh_status()
        self._sync_button_states()

    def uncheck_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False
        self._refresh_status()
        self._sync_button_states()

    def apply_autosized_clicked(self, sender, args):
        set_breaker = bool(getattr(self._set_breaker_cb, "IsChecked", True))
        set_frame = bool(getattr(self._set_frame_cb, "IsChecked", True))
        allow_15a = bool(getattr(self._allow_15a_cb, "IsChecked", False))
        upsize_only = bool(getattr(self._upsize_only_cb, "IsChecked", False))
        self._preview_apply_callback(self._rows, set_breaker, set_frame, allow_15a, upsize_only)
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        self._refresh_grid(refresh_items=True)

    def apply_clicked(self, sender, args):
        set_breaker = bool(getattr(self._set_breaker_cb, "IsChecked", True))
        set_frame = bool(getattr(self._set_frame_cb, "IsChecked", True))
        allow_15a = bool(getattr(self._allow_15a_cb, "IsChecked", False))
        upsize_only = bool(getattr(self._upsize_only_cb, "IsChecked", False))
        if self._apply_callback(self._rows, set_breaker, set_frame, allow_15a, upsize_only):
            self.Close()

    def reset_clicked(self, sender, args):
        for row in self._selected_rows():
            if not row.is_enabled:
                continue
            row.new_rating = row.current_rating
            row.new_frame = row.current_frame
            row.recompute_state()
        self._refresh_grid(refresh_items=True)

    def cancel_clicked(self, sender, args):
        self.Close()

    def show_unsupported_toggled(self, sender, args):
        self._show_unsupported = bool(getattr(sender, "IsChecked", False))
        self._refresh_grid(refresh_items=True)

    def row_filter_toggled(self, sender, args):
        self._row_filter_mode = "changed" if bool(getattr(sender, "IsChecked", False)) else "all"
        self._refresh_grid(refresh_items=True)

    def grid_selection_changed(self, sender, args):
        try:
            selected = list(self._grid.SelectedItems or [])
        except Exception:
            selected = []
        if len(selected) > 1:
            self._last_selected_rows = selected
        self._sync_button_states()

    def breaker_cell_edit_ending(self, sender, args):
        try:
            row = getattr(getattr(args, "Row", None), "Item", None)
            if row is None:
                return
            column = getattr(args, "Column", None)
            sort_path = getattr(column, "SortMemberPath", "")
            if sort_path not in ("new_rating", "new_frame"):
                return
            editor = getattr(args, "EditingElement", None)
            text_value = getattr(editor, "Text", None)
            if text_value is None:
                text_value = getattr(row, sort_path, None)
            parsed = _parse_whole_amps(text_value)
            current_value = getattr(row, sort_path, "-")
            if parsed is None:
                if editor is not None:
                    try:
                        editor.Text = str(current_value)
                    except Exception:
                        pass
                setattr(row, sort_path, str(current_value))
                row.recompute_state()
                self._refresh_grid(refresh_items=True)
                return
            formatted = _fmt_amp(parsed, 0)
            setattr(row, sort_path, formatted)
            if editor is not None:
                editor.Text = formatted
            row.recompute_state()
            self._refresh_grid(refresh_items=True)
        except Exception:
            self._refresh_grid(refresh_items=True)

    def _clear_grid_selection(self):
        if self._grid is None:
            return
        try:
            self._grid.UnselectAll()
        except Exception:
            pass

    def window_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if self._grid is None or source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if not _is_descendant_of_control(source, self._grid):
            self._clear_grid_selection()

    def grid_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, DataGridRow) is None:
            self._clear_grid_selection()
            self._sync_button_states()


class MarkExistingActionRow(object):
    def __init__(self, item, is_enabled, reason, current_notes, current_wire):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = item.circuit.Id.IntegerValue
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.panel_ckt_text = "{} / {}".format(item.panel or "-", item.circuit_number or "-")
        self.branch_type = item.branch_type
        self.is_enabled = bool(is_enabled)
        self.is_checked = False
        self.reason = reason or ""

        self.current_notes = current_notes or ""
        self.current_wire = current_wire or "-"

        self.new_notes = self.current_notes
        self.new_wire = self.current_wire
        self.new_notes_changed = False
        self.new_wire_changed = False
        self.is_changed = False
        self.remarks = ""
        self.recompute_state()

    def recompute_state(self):
        self.new_notes_changed = str(self.new_notes or "") != str(self.current_notes or "")
        self.new_wire_changed = str(self.new_wire or "") != str(self.current_wire or "")
        self.is_changed = bool(self.new_notes_changed or self.new_wire_changed)
        if not self.is_enabled:
            self.remarks = "Blocked - {}".format(self.reason or "Unsupported")
        elif self.is_changed:
            self.remarks = "Will be modified"
        else:
            self.remarks = "No change"


class MarkExistingActionWindow(forms.WPFWindow):
    def __init__(self, rows, preview_callback, apply_callback, theme_mode="light", accent_mode="blue"):
        xaml = os.path.abspath(os.path.join(_THIS_DIR, "CircuitMarkExistingActionWindow.xaml"))
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        forms.WPFWindow.__init__(self, xaml)
        _try_apply_theme(self)
        self._rows = list(rows or [])
        self._preview_callback = preview_callback
        self._apply_callback = apply_callback
        self._is_syncing_checks = False
        self._suppress_check_events = False
        self._is_ready = False
        self._show_unsupported = True
        self._last_selected_rows = []

        self._grid = self.FindName("CircuitGrid")
        _apply_grid_styles_mark_existing(self, self._grid)
        self._status = self.FindName("ChangedStatusText")
        self._checked_status = self.FindName("CheckedStatusText")
        self._show_unsupported_cb = self.FindName("ShowUnsupportedToggle")
        self._set_notes_cb = self.FindName("SetNotesCheck")
        self._clear_wire_cb = self.FindName("ClearWireCheck")
        self._clear_conduit_cb = self.FindName("ClearConduitCheck")
        self._set_existing_btn = self.FindName("SetExistingButton")
        self._set_new_btn = self.FindName("SetNewButton")
        self._apply_btn = self.FindName("ApplyButton")
        self._last_mode = "existing"

        if self._show_unsupported_cb is not None:
            self._show_unsupported_cb.IsChecked = True
        if self._set_notes_cb is not None:
            self._set_notes_cb.IsChecked = True
        if self._clear_wire_cb is not None:
            self._clear_wire_cb.IsChecked = False
        if self._clear_conduit_cb is not None:
            self._clear_conduit_cb.IsChecked = False

        for row in self._rows:
            row.is_checked = False

        self._apply_visibility_filter()
        self._sync_option_controls()
        self.Loaded += self.window_loaded
        self._refresh_grid(refresh_items=False)

    def window_loaded(self, sender, args):
        self._is_ready = True
        self._sync_button_states()

    def _apply_visibility_filter(self):
        if self._grid is None:
            return
        if self._show_unsupported:
            items = list(self._rows)
        else:
            items = [x for x in self._rows if bool(getattr(x, "is_enabled", False))]
        self._grid.ItemsSource = items

    def _refresh_grid(self, refresh_items=False):
        if refresh_items:
            self._apply_visibility_filter()
        elif self._grid is not None:
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass
        self._refresh_status()

    def _refresh_status(self):
        changed = len([x for x in self._rows if x.is_enabled and x.is_checked and x.is_changed])
        total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))]) if not self._show_unsupported else len(self._rows)
        if self._status is not None:
            self._status.Text = "{} of {} circuits to be modified".format(changed, total)
        if self._checked_status is not None:
            checkable_total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
            checked = len([x for x in self._rows if x.is_enabled and x.is_checked])
            self._checked_status.Text = "{} of {} checked".format(checked, checkable_total)
        self._sync_button_states()

    def _sync_option_controls(self):
        pass

    def _sync_button_states(self):
        checked_count = len([x for x in self._rows if x.is_enabled and x.is_checked])
        changed_count = len([x for x in self._rows if x.is_enabled and x.is_checked and x.is_changed])
        if self._set_existing_btn is not None:
            self._set_existing_btn.IsEnabled = checked_count > 0
        if self._set_new_btn is not None:
            self._set_new_btn.IsEnabled = checked_count > 0
        if self._apply_btn is not None:
            self._apply_btn.IsEnabled = changed_count > 0

    def _apply_checkbox_to_selected(self, sender, state):
        if not self._is_ready or self._is_syncing_checks or self._suppress_check_events:
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
        if len(selected) > 1 and row in selected:
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
        self._suppress_check_events = True
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def _option_values(self):
        set_notes = bool(getattr(self._set_notes_cb, "IsChecked", True))
        clear_wire = bool(getattr(self._clear_wire_cb, "IsChecked", False))
        clear_conduit = bool(getattr(self._clear_conduit_cb, "IsChecked", False))
        return set_notes, clear_wire, clear_conduit

    def options_changed(self, sender, args):
        self._sync_option_controls()
        self._sync_button_states()

    def item_checked(self, sender, args):
        self._apply_checkbox_to_selected(sender, True)

    def item_unchecked(self, sender, args):
        self._apply_checkbox_to_selected(sender, False)

    def item_checkbox_clicked(self, sender, args):
        self._apply_checkbox_to_selected(sender, bool(getattr(sender, "IsChecked", False)))

    def check_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for row in self._rows:
                row.is_checked = bool(row.is_enabled)
        finally:
            self._is_syncing_checks = False
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def uncheck_all_clicked(self, sender, args):
        self._is_syncing_checks = True
        self._suppress_check_events = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        try:
            self._refresh_grid(refresh_items=False)
        finally:
            self._suppress_check_events = False

    def set_existing_clicked(self, sender, args):
        set_notes, clear_wire, clear_conduit = self._option_values()
        self._last_mode = "existing"
        self._preview_callback(self._rows, "existing", set_notes, clear_wire, clear_conduit)
        self._refresh_grid(refresh_items=False)

    def set_new_clicked(self, sender, args):
        set_notes, clear_wire, clear_conduit = self._option_values()
        self._last_mode = "new"
        self._preview_callback(self._rows, "new", set_notes, clear_wire, clear_conduit)
        self._refresh_grid(refresh_items=False)

    def apply_clicked(self, sender, args):
        set_notes, clear_wire, clear_conduit = self._option_values()
        if self._apply_callback(self._rows, self._last_mode, set_notes, clear_wire, clear_conduit):
            self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def show_unsupported_toggled(self, sender, args):
        self._show_unsupported = bool(getattr(sender, "IsChecked", False))
        self._refresh_grid(refresh_items=True)

    def grid_selection_changed(self, sender, args):
        try:
            selected = list(self._grid.SelectedItems or [])
        except Exception:
            selected = []
        if selected:
            self._last_selected_rows = selected

    def _clear_grid_selection(self):
        if self._grid is None:
            return
        try:
            self._grid.UnselectAll()
        except Exception:
            pass

    def window_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if self._grid is None or source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if not _is_descendant_of_control(source, self._grid):
            self._clear_grid_selection()

    def grid_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, DataGridRow) is None:
            self._clear_grid_selection()
            self._sync_button_states()


class CalculateSettingsExternalEventGateway(object):
    """Opens Calculate Circuits settings inside valid Revit API context."""

    def __init__(self, logger=None):
        self.logger = logger
        self._pending = None
        self._handler = _CalculateSettingsHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def is_busy(self):
        return self._pending is not None

    def raise_open(self, callback=None):
        if self._pending is not None:
            return False
        self._pending = {"callback": callback}
        self._event.Raise()
        return True

    def _consume_pending(self):
        pending = self._pending
        self._pending = None
        return pending


class _CalculateSettingsHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, application):
        pending = self._gateway._consume_pending()
        if not pending:
            return

        callback = pending.get("callback")
        status = "ok"
        error = None
        try:
            if not os.path.exists(CALC_SETTINGS_PATH):
                raise Exception("Calculate Circuits settings file not found:\n\n{}".format(CALC_SETTINGS_PATH))
            if not os.path.exists(CALC_SETTINGS_XAML_PATH):
                raise Exception("Calculate Circuits settings XAML not found:\n\n{}".format(CALC_SETTINGS_XAML_PATH))
            module = imp.load_source("ced_calculate_circuits_config", CALC_SETTINGS_PATH)
            try:
                module.XAML_PATH = CALC_SETTINGS_XAML_PATH
            except Exception:
                pass
            window_cls = getattr(module, "CircuitSettingsWindow", None)
            if window_cls is None:
                raise Exception("CircuitSettingsWindow was not found in config script.")
            window = window_cls()
            try:
                window.show_dialog()
            except Exception:
                window.ShowDialog()
        except Exception as ex:
            status = "error"
            error = ex
            if self._gateway.logger:
                self._gateway.logger.exception("Failed to open Calculate Circuits settings in API context: %s", ex)

        if callback:
            try:
                callback(status, error)
            except Exception:
                pass

    def GetName(self):
        return "CED Calculate Settings External Event"


class CircuitBrowserPanel(forms.WPFPanel):
    panel_id = PANEL_ID
    panel_title = TITLE
    panel_source = os.path.abspath(os.path.join(_THIS_DIR, "CircuitBrowserPanel.xaml"))

    _instance = None
    _operation_gateway = None

    def __init__(self):
        forms.WPFPanel.__init__(self)
        self._theme_mode = CURRENT_THEME_MODE
        self._accent_mode = CURRENT_ACCENT_MODE
        _try_apply_theme(self)
        CircuitBrowserPanel._instance = self
        self._logger = script.get_logger()
        self._theme_bridge = None
        self._dock_frame_host = None
        self._all_items = []
        self._is_card_view = False
        self._type_options = []
        self._active_type_filters = set()
        self._filter_menu = None
        self._warnings_only = False
        self._overrides_only = False
        self._checked_only = False
        self._actions_menu = None
        self._browser_options_menu = None
        self._compact_show_type_badges = True
        self._use_surface_item_states = True
        self._last_selected_items = []
        self._operation_gateway = self._get_operation_gateway()
        self._settings_gateway = CalculateSettingsExternalEventGateway(logger=self._logger)

        self._list = self.FindName("CircuitList")
        self._search = self.FindName("SearchBox")
        self._search_placeholder = self.FindName("SearchPlaceholderText")
        self._search_clear = self.FindName("ClearSearchButton")
        self._status = self.FindName("StatusText")
        self._doc_name_text = self.FindName("DocumentNameText")
        self._toggle = self.FindName("ToggleViewButton")
        self._filter_button = self.FindName("FilterButton")
        self._browser_options_button = self.FindName("BrowserOptionsButton")
        self._filter_active_mark = self.FindName("FilterActiveMark")
        self._toggle_list_icon = self.FindName("ToggleListIcon")
        self._toggle_card_icon = self.FindName("ToggleCardIcon")
        self._dock_frame_host = self.FindName("DockFrameHost")
        self._apply_revit_frame_background(is_dark=False)
        self._surface_item_style = _try_find_resource(self, "CED.ListViewItem.SurfaceBehavior")
        self._apply_list_interaction_mode()
        self._update_search_chrome()
        self._update_toggle_button_visual()

        self._compact_template = self.FindResource("CompactTemplate")
        self._card_template = self.FindResource("CardTemplate")
        self._active_doc_key = None
        self._loaded_doc_key = None
        self._doc_is_opening = False
        self._view_activated_handler = None
        self._doc_opening_handler = None
        self._doc_opened_handler = None
        self._doc_closing_handler = None

        self.Loaded += self.panel_loaded
        self.Unloaded += self.panel_unloaded
        self.IsVisibleChanged += self.panel_visibility_changed
        self._attach_document_lifecycle_handlers()
        self._attach_view_activated_handler()

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

    def _apply_revit_frame_background(self, is_dark):
        color_hex = DOCK_PANE_FRAME_DARK if bool(is_dark) else DOCK_PANE_FRAME_LIGHT
        brush = _to_brush(color_hex, DOCK_PANE_FRAME_LIGHT)
        if brush is None:
            return
        try:
            self.Background = brush
        except Exception:
            pass
        try:
            if self._dock_frame_host is not None:
                self._dock_frame_host.Background = brush
        except Exception:
            pass

    def _ensure_theme_bridge(self):
        if self._theme_bridge is None:
            self._theme_bridge = RevitThemeBridge(
                uiapp=__revit__,
                on_theme_changed=self._apply_revit_frame_background,
                logger=self._logger,
            )
        self._theme_bridge.attach()

    def _detach_theme_bridge(self):
        if self._theme_bridge is None:
            return
        self._theme_bridge.detach()

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
            try:
                self._show_run_summary_if_needed(result)
            except Exception as ex:
                self._logger.warning("Run summary window failed: %s", ex)
            self._safe_load_items()
            return

        reason = result.get("reason", "unknown")
        self._set_status("Operation cancelled ({})".format(reason))
        self._safe_load_items()

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
        window = CircuitRunSummaryWindow(
            locked_rows,
            runtime_rows,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )
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

    def _is_pane_visible(self):
        try:
            return bool(self.IsVisible)
        except Exception:
            return False

    def _attach_document_lifecycle_handlers(self):
        app = None
        try:
            app = getattr(__revit__, "Application", None)
        except Exception:
            app = None
        if app is None:
            return

        if self._doc_opening_handler is None:
            try:
                handler = EventHandler[DocumentOpeningEventArgs](self._on_document_opening)
                app.DocumentOpening += handler
                self._doc_opening_handler = handler
            except Exception as ex:
                self._logger.warning("Circuit Browser failed to attach DocumentOpening handler: %s", ex)
                self._doc_opening_handler = None

        if self._doc_opened_handler is None:
            try:
                handler = EventHandler[DocumentOpenedEventArgs](self._on_document_opened)
                app.DocumentOpened += handler
                self._doc_opened_handler = handler
            except Exception as ex:
                self._logger.warning("Circuit Browser failed to attach DocumentOpened handler: %s", ex)
                self._doc_opened_handler = None

        if self._doc_closing_handler is None:
            try:
                handler = EventHandler[DocumentClosingEventArgs](self._on_document_closing)
                app.DocumentClosing += handler
                self._doc_closing_handler = handler
            except Exception as ex:
                self._logger.warning("Circuit Browser failed to attach DocumentClosing handler: %s", ex)
                self._doc_closing_handler = None

    def _attach_view_activated_handler(self):
        if self._view_activated_handler is not None:
            return
        try:
            self._view_activated_handler = EventHandler[ViewActivatedEventArgs](self._on_view_activated)
            __revit__.ViewActivated += self._view_activated_handler
        except Exception as ex:
            self._logger.warning("Circuit Browser failed to attach ViewActivated handler: %s", ex)
            self._view_activated_handler = None

    def _detach_event_handlers(self):
        app = None
        try:
            app = getattr(__revit__, "Application", None)
        except Exception:
            app = None

        try:
            if self._view_activated_handler is not None:
                __revit__.ViewActivated -= self._view_activated_handler
        except Exception:
            pass
        self._view_activated_handler = None

        try:
            if app is not None and self._doc_opening_handler is not None:
                app.DocumentOpening -= self._doc_opening_handler
        except Exception:
            pass
        self._doc_opening_handler = None

        try:
            if app is not None and self._doc_opened_handler is not None:
                app.DocumentOpened -= self._doc_opened_handler
        except Exception:
            pass
        self._doc_opened_handler = None

        try:
            if app is not None and self._doc_closing_handler is not None:
                app.DocumentClosing -= self._doc_closing_handler
        except Exception:
            pass
        self._doc_closing_handler = None

    def _on_document_opening(self, sender, args):
        self._doc_is_opening = True

    def _on_document_opened(self, sender, args):
        self._doc_is_opening = False
        doc = None
        try:
            doc = getattr(args, "Document", None)
        except Exception:
            doc = None
        key = self._doc_key(doc)
        self._active_doc_key = key
        if not self._is_pane_visible():
            return
        if key != self._loaded_doc_key:
            self._safe_load_items(doc_override=doc)
        else:
            self._set_doc_banner(doc)

    def _on_document_closing(self, sender, args):
        doc = None
        try:
            doc = getattr(args, "Document", None)
        except Exception:
            doc = None
        key = self._doc_key(doc)
        if key is None:
            return
        if key == self._active_doc_key:
            self._active_doc_key = None
        if key == self._loaded_doc_key:
            self._loaded_doc_key = None

    def _on_view_activated(self, sender, args):
        if not self._is_pane_visible():
            return
        try:
            doc = getattr(args, "Document", None)
        except Exception:
            doc = None
        if doc is None:
            return
        self._doc_is_opening = False

        key = self._doc_key(doc)
        if key != self._active_doc_key:
            self._active_doc_key = key
            if key != self._loaded_doc_key:
                self._safe_load_items(doc_override=doc)
                return

        self._set_doc_banner(doc)

    def _safe_load_items(self, doc_override=_DOC_SENTINEL):
        if self._doc_is_opening:
            return
        doc = self._get_active_doc() if doc_override is _DOC_SENTINEL else doc_override
        self._active_doc_key = self._doc_key(doc)
        self._set_doc_banner(doc)
        if doc is None:
            self._all_items = []
            self._list.ItemsSource = ObservableCollection[CircuitListItem]([])
            self._loaded_doc_key = None
            self._set_status("Open a model document to load circuits.")
            return
        self._load_items(doc)
        self._loaded_doc_key = self._active_doc_key

    def refresh_on_open(self):
        self._ensure_theme_bridge()
        if self._doc_is_opening:
            return
        doc = self._get_active_doc()
        key = self._doc_key(doc)
        self._active_doc_key = key
        if key != self._loaded_doc_key:
            self._safe_load_items(doc_override=doc)
        else:
            self._set_doc_banner(doc)

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
        is_filtered = (
            self._warnings_only
            or self._overrides_only
            or self._checked_only
            or (set(self._type_options) != set(self._active_type_filters))
        )
        primary = _try_find_resource(self, "CED.Brush.Accent")
        button_bg = _try_find_resource(self, "CED.Brush.ButtonDefaultBackground")
        foreground_on_accent = _try_find_resource(self, "CED.Brush.ButtonForegroundOnAccent")
        if is_filtered:
            if primary is not None:
                self._filter_button.Background = primary
                self._filter_button.BorderBrush = primary
            if foreground_on_accent is not None:
                self._filter_button.Foreground = foreground_on_accent
            if self._filter_active_mark is not None:
                self._filter_active_mark.Visibility = Visibility.Visible
        else:
            if button_bg is not None:
                self._filter_button.Background = button_bg
            if primary is not None:
                self._filter_button.BorderBrush = primary
                self._filter_button.Foreground = primary
            if self._filter_active_mark is not None:
                self._filter_active_mark.Visibility = Visibility.Collapsed

    def _update_search_chrome(self):
        text = ""
        try:
            text = (self._search.Text or "")
        except Exception:
            text = ""
        has_text = bool(text.strip())
        has_focus = False
        try:
            has_focus = bool(self._search.IsKeyboardFocused)
        except Exception:
            has_focus = False
        if self._search_placeholder is not None:
            self._search_placeholder.Visibility = Visibility.Collapsed if (has_text or has_focus) else Visibility.Visible
        if self._search_clear is not None:
            self._search_clear.Visibility = Visibility.Visible if has_text else Visibility.Collapsed

    def _update_toggle_button_visual(self):
        if self._toggle is None:
            return
        if self._is_card_view:
            if self._toggle_list_icon is not None:
                self._toggle_list_icon.Visibility = Visibility.Collapsed
            if self._toggle_card_icon is not None:
                self._toggle_card_icon.Visibility = Visibility.Visible
            self._toggle.ToolTip = "Switch to compact display"
        else:
            if self._toggle_list_icon is not None:
                self._toggle_list_icon.Visibility = Visibility.Visible
            if self._toggle_card_icon is not None:
                self._toggle_card_icon.Visibility = Visibility.Collapsed
            self._toggle.ToolTip = "Switch to card display"

    def _apply_list_interaction_mode(self):
        if self._list is None:
            return
        if self._use_surface_item_states:
            self._list.Tag = "surface"
            if self._surface_item_style is not None:
                self._list.ItemContainerStyle = self._surface_item_style
        else:
            self._list.Tag = None
            self._list.ItemContainerStyle = None
        try:
            self._list.Items.Refresh()
        except Exception:
            pass

    def _apply_type_tag_brushes(self):
        for item in list(self._all_items or []):
            bg_key = getattr(item, "type_tag_bg_key", None)
            fg_key = getattr(item, "type_tag_fg_key", None)
            bg_resource = _try_find_resource(self, bg_key) if bg_key else None
            fg_resource = _try_find_resource(self, fg_key) if fg_key else None
            item.type_tag_bg = bg_resource if bg_resource is not None else "#ECEFF3"
            item.type_tag_fg = fg_resource if fg_resource is not None else "#52606D"

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
        for item in items:
            item.show_type_tag = bool(self._compact_show_type_badges)
            item.type_tag_visibility = "Visible" if item.show_type_tag else "Collapsed"

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
        self._apply_type_tag_brushes()
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

    def _validate_circuit_item(self, item, doc):
        circuit = getattr(item, "circuit", None)
        if circuit is None or doc is None:
            return False
        try:
            if hasattr(circuit, "IsValidObject") and not circuit.IsValidObject:
                return False
        except Exception:
            return False
        try:
            element_id = circuit.Id
        except Exception:
            return False
        if element_id is None:
            return False
        try:
            live = doc.GetElement(element_id)
        except Exception:
            return False
        if live is None:
            return False
        if not isinstance(live, DBE.ElectricalSystem):
            return False
        try:
            item.circuit = live
        except Exception:
            pass
        return True

    def _prune_stale_items(self, target_items=None):
        doc = self._get_active_doc()
        if doc is None:
            return [], 0

        target_ids = None
        if target_items is not None:
            target_ids = set()
            for item in list(target_items or []):
                try:
                    cid = item.circuit.Id.IntegerValue
                except Exception:
                    cid = None
                if cid is not None:
                    target_ids.add(cid)

        valid_all = []
        valid_targets = []
        removed_count = 0
        for item in list(self._all_items or []):
            if not self._validate_circuit_item(item, doc):
                removed_count += 1
                continue
            valid_all.append(item)
            if target_ids is None:
                continue
            try:
                cid = item.circuit.Id.IntegerValue
            except Exception:
                cid = None
            if cid in target_ids:
                valid_targets.append(item)

        if removed_count:
            self._all_items = valid_all
            self._rebuild_filter_options()
            self._refresh_list()
            self.selection_changed(None, None)
            self._logger.debug("Circuit Browser pruned %s stale rows.", removed_count)

        if target_ids is None:
            return list(valid_all), removed_count
        return valid_targets, removed_count

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
        elif len(self._last_selected_items) > 1 and clicked_item in self._last_selected_items:
            targets = list(self._last_selected_items)

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

        reset_item = MenuItem()
        reset_item.Header = "Reset Filters"
        reset_item.Click += self.filter_reset_clicked
        menu.Items.Add(reset_item)

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
        self._ensure_theme_bridge()
        if self._view_activated_handler is None or self._doc_opening_handler is None:
            self._attach_document_lifecycle_handlers()
            self._attach_view_activated_handler()
        if not self._is_pane_visible() or self._doc_is_opening:
            return
        doc = self._get_active_doc()
        key = self._doc_key(doc)
        self._active_doc_key = key
        if key != self._loaded_doc_key:
            self._safe_load_items(doc_override=doc)
        else:
            self._set_doc_banner(doc)

    def panel_unloaded(self, sender, args):
        self._detach_theme_bridge()
        self._detach_event_handlers()

    def panel_visibility_changed(self, sender, args):
        if not self._is_pane_visible():
            self._detach_theme_bridge()
            return
        if self._doc_is_opening:
            return
        self._ensure_theme_bridge()
        if self._view_activated_handler is None or self._doc_opening_handler is None:
            self._attach_document_lifecycle_handlers()
            self._attach_view_activated_handler()
        doc = self._get_active_doc()
        key = self._doc_key(doc)
        self._active_doc_key = key
        if key != self._loaded_doc_key:
            self._safe_load_items(doc_override=doc)
        else:
            self._set_doc_banner(doc)

    def search_changed(self, sender, args):
        self._update_search_chrome()
        self._refresh_list()

    def search_got_focus(self, sender, args):
        self._update_search_chrome()

    def search_lost_focus(self, sender, args):
        self._update_search_chrome()

    def clear_search_clicked(self, sender, args):
        try:
            if self._search is not None:
                self._search.Text = ""
                self._search.Focus()
        except Exception:
            pass
        self._update_search_chrome()
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

        mark_existing_item = MenuItem()
        mark_existing_item.Header = "Mark as New/Existing"
        mark_existing_item.Click += self.action_mark_existing_clicked
        menu.Items.Add(mark_existing_item)

        self._actions_menu = menu

    def actions_button_clicked(self, sender, args):
        self._build_actions_menu()
        if self._actions_menu is not None:
            self._actions_menu.PlacementTarget = sender
            self._actions_menu.IsOpen = True

    def _build_browser_options_menu(self):
        menu = ContextMenu()

        theme_menu = MenuItem()
        theme_menu.Header = "Theme"
        light_item = MenuItem()
        light_item.Header = "Light"
        light_item.IsCheckable = True
        light_item.IsChecked = (self._theme_mode != "dark")
        light_item.Tag = "light"
        light_item.Click += self.browser_theme_clicked
        theme_menu.Items.Add(light_item)
        dark_item = MenuItem()
        dark_item.Header = "Dark"
        dark_item.IsCheckable = True
        dark_item.IsChecked = (self._theme_mode == "dark")
        dark_item.Tag = "dark"
        dark_item.Click += self.browser_theme_clicked
        theme_menu.Items.Add(dark_item)
        menu.Items.Add(theme_menu)

        accent_menu = MenuItem()
        accent_menu.Header = "Accent Color"
        for accent_mode, accent_label in (
            ("blue", "Blue"),
            ("red", "Red"),
            ("green", "Green"),
            ("neutral", "Neutral"),
        ):
            item = MenuItem()
            item.Header = accent_label
            item.IsCheckable = True
            item.IsChecked = (self._accent_mode == accent_mode)
            item.Tag = accent_mode
            item.Click += self.browser_accent_clicked
            accent_menu.Items.Add(item)
        menu.Items.Add(accent_menu)

        display_menu = MenuItem()
        display_menu.Header = "Display Mode"
        compact_item = MenuItem()
        compact_item.Header = "Compact"
        compact_item.IsCheckable = True
        compact_item.IsChecked = not self._is_card_view
        compact_item.Tag = "compact"
        compact_item.Click += self.browser_display_mode_clicked
        display_menu.Items.Add(compact_item)
        card_item = MenuItem()
        card_item.Header = "Card"
        card_item.IsCheckable = True
        card_item.IsChecked = self._is_card_view
        card_item.Tag = "card"
        card_item.Click += self.browser_display_mode_clicked
        display_menu.Items.Add(card_item)
        menu.Items.Add(display_menu)

        compact_menu = MenuItem()
        compact_menu.Header = "List View"
        badges_item = MenuItem()
        badges_item.Header = "Show Circuit Type Badges"
        badges_item.IsCheckable = True
        badges_item.IsChecked = bool(self._compact_show_type_badges)
        badges_item.Click += self.toggle_compact_badges_clicked
        compact_menu.Items.Add(badges_item)
        menu.Items.Add(compact_menu)

        self._browser_options_menu = menu

    def browser_options_clicked(self, sender, args):
        self._build_browser_options_menu()
        if self._browser_options_menu is not None:
            self._browser_options_menu.PlacementTarget = sender if sender is not None else self._browser_options_button
            self._browser_options_menu.IsOpen = True

    def browser_theme_clicked(self, sender, args):
        global CURRENT_THEME_MODE
        mode = str(getattr(sender, "Tag", "light")).lower()
        if mode not in ("light", "dark"):
            return
        if self._theme_mode == mode:
            return
        self._theme_mode = mode
        CURRENT_THEME_MODE = mode
        _try_apply_theme(self)
        self._surface_item_style = _try_find_resource(self, "CED.ListViewItem.SurfaceBehavior")
        self._apply_type_tag_brushes()
        self._apply_list_interaction_mode()
        self._update_filter_button_style()
        self._update_search_chrome()
        self._update_toggle_button_visual()
        self._refresh_list()

    def browser_accent_clicked(self, sender, args):
        global CURRENT_ACCENT_MODE
        mode = str(getattr(sender, "Tag", "blue")).lower()
        if mode not in ACCENT_BRUSH_MAP:
            return
        if self._accent_mode == mode:
            return
        self._accent_mode = mode
        CURRENT_ACCENT_MODE = mode
        _try_apply_theme(self)
        self._apply_type_tag_brushes()
        self._update_filter_button_style()
        self._refresh_list()

    def browser_display_mode_clicked(self, sender, args):
        mode = str(getattr(sender, "Tag", "")).lower()
        if mode == "compact":
            self._set_card_view(False)
            return
        if mode == "card":
            self._set_card_view(True)

    def _collect_action_targets(self):
        # Actions run only on explicitly checked circuits to avoid
        # selection/checkbox ambiguity for users.
        return [x for x in self._all_items if getattr(x, "is_checked", False)]

    def _is_locked_circuit(self, circuit, doc=None):
        if doc is None:
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

    def _wire_hot_size_string(self, circuit):
        text = self._param_text(circuit, "CKT_Wire Hot Size_CEDT", "")
        if text:
            return text
        text = self._param_text(circuit, "Wire Hot Size_CEDT", "")
        if text:
            return text
        return self._wire_size_string(circuit) or "-"

    def _conduit_size_string(self, circuit):
        text = self._param_text(circuit, "Conduit Size_CEDT", "")
        return text or "-"

    def _conduit_wire_size_string(self, circuit):
        text = self._param_text(circuit, "Conduit and Wire Size_CEDT", "")
        if text:
            return text
        conduit = self._conduit_size_string(circuit)
        wire = self._wire_hot_size_string(circuit)
        if conduit == "-" and wire == "-":
            return "-"
        return "{} / {}".format(conduit, wire)

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

    def _next_ocp_size(self, amps, min_ocp=20):
        try:
            value = float(amps)
        except Exception:
            return None
        try:
            value = max(value, float(min_ocp or 0))
        except Exception:
            pass
        for k in OCP_TABLE_KEYS:
            if k >= value:
                return k
        return OCP_TABLE_KEYS[-1] if OCP_TABLE_KEYS else None

    def _frame_for_rating(self, rating, min_ocp=20):
        size = self._next_ocp_size(rating, min_ocp=min_ocp)
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
            if not row.is_enabled:
                row.is_checked = False
                row.recompute_state()
                continue
            if not row.is_checked:
                row.recompute_state()
                continue
            row.new_qty = row.current_qty
            row.new_size = row.current_size
            row.new_wire = "no change"
            row.new_wire_font_style = "Italic"
            row.new_qty_changed = False
            row.new_size_changed = False
            row.new_wire_changed = False

            try:
                branch = self._simulate_branch(row.circuit, include_neutral=include_neutral)
                new_qty = int(branch.neutral_wire_quantity or 0)
                new_size = branch.neutral_wire_size or ""
                new_wire = branch.get_wire_size_callout() or ""
            except Exception:
                row.recompute_state()
                continue

            changed = (new_qty != int(row.current_qty or 0)) or (str(new_size or "") != str(row.current_size or "")) or (
                str(new_wire or "") != str(row.current_wire or "")
            )
            if changed:
                row.new_qty = new_qty
                row.new_size = new_size
                row.new_wire = new_wire or "-"
                row.new_wire_font_style = "Normal"
                row.new_qty_changed = new_qty != int(row.current_qty or 0)
                row.new_size_changed = str(new_size or "") != str(row.current_size or "")
                row.new_wire_changed = str(new_wire or "") != str(row.current_wire or "")
            row.recompute_state()

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
            if not row.is_enabled:
                row.is_checked = False
                row.recompute_state()
                continue
            if not row.is_checked:
                row.recompute_state()
                continue
            row.new_qty = row.current_qty
            row.new_size = row.current_size
            row.new_wire = "no change"
            row.new_wire_font_style = "Italic"
            row.new_qty_changed = False
            row.new_size_changed = False
            row.new_wire_changed = False

            try:
                branch = self._simulate_branch(row.circuit, include_ig=include_ig)
                new_qty = int(branch.isolated_ground_wire_quantity or 0)
                new_size = branch.isolated_ground_wire_size or ""
                new_wire = branch.get_wire_size_callout() or ""
            except Exception:
                row.recompute_state()
                continue

            changed = (new_qty != int(row.current_qty or 0)) or (str(new_size or "") != str(row.current_size or "")) or (
                str(new_wire or "") != str(row.current_wire or "")
            )
            if changed:
                row.new_qty = new_qty
                row.new_size = new_size
                row.new_wire = new_wire or "-"
                row.new_wire_font_style = "Normal"
                row.new_qty_changed = new_qty != int(row.current_qty or 0)
                row.new_size_changed = str(new_size or "") != str(row.current_size or "")
                row.new_wire_changed = str(new_wire or "") != str(row.current_wire or "")
            row.recompute_state()

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
        doc = self._get_active_doc()
        is_workshared = bool(getattr(doc, "IsWorkshared", False)) if doc is not None else False
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

            locked, owner = self._is_locked_circuit(circuit, doc=doc) if is_workshared else (False, "")
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

            # Defer autosize calculations to preview/apply so the window can open faster.
            row = BreakerActionRow(item, is_enabled, reason, load_current, cur_rating, cur_frame, None, None)
            row.set_auto_frames(None, None)
            row.autosized_frame = "-"
            if not is_enabled:
                row.is_checked = False
            rows.append(row)
        return rows

    def _preview_breaker_rows(self, rows, set_breaker, set_frame, allow_15a=False, upsize_only=False):
        min_ocp = 15 if allow_15a else 20
        for row in rows:
            load_value = _safe_float(row._load_current_value)
            auto_rating = self._next_ocp_size((load_value * 1.25) if load_value is not None else None, min_ocp=min_ocp)
            if auto_rating is None:
                auto_rating = self._next_ocp_size(row._current_rating_value, min_ocp=min_ocp)
            row.autosized_rating = _fmt_amp(auto_rating, 0)
            row._autosized_rating_value = auto_rating
            row.autosized_frame = _fmt_amp(self._frame_for_rating(auto_rating, min_ocp=min_ocp), 0)
            if not row.is_enabled or not row.is_checked:
                row.recompute_state()
                continue

            if set_breaker:
                breaker_value = auto_rating
                if upsize_only:
                    current_value = _parse_whole_amps(row.current_rating)
                    if current_value is not None:
                        breaker_value = max(int(current_value), int(auto_rating or current_value))
                row.new_rating = _fmt_amp(breaker_value, 0)
            if set_frame:
                frame_basis_rating = _parse_whole_amps(row.new_rating)
                if frame_basis_rating is not None:
                    row.new_frame = _fmt_amp(self._frame_for_rating(frame_basis_rating, min_ocp=min_ocp), 0)

            row.recompute_state()

    def _apply_breaker_rows(self, rows, set_breaker, set_frame, allow_15a=False, upsize_only=False):
        updates = []
        invalid_rows = []
        for row in rows:
            if not (row.is_enabled and row.is_changed):
                continue
            rating_val = _parse_whole_amps(row.new_rating) if row.new_rating_changed else None
            frame_val = _parse_whole_amps(row.new_frame) if row.new_frame_changed else None
            if row.new_rating_changed and rating_val is None:
                invalid_rows.append(row.panel_ckt_text)
                continue
            if row.new_frame_changed and frame_val is None:
                invalid_rows.append(row.panel_ckt_text)
                continue
            updates.append(
                {
                    "circuit_id": row.circuit_id,
                    "rating": float(rating_val) if rating_val is not None else None,
                    "frame": float(frame_val) if frame_val is not None else None,
                    "set_rating": bool(row.new_rating_changed),
                    "set_frame": bool(row.new_frame_changed),
                }
            )

        if invalid_rows:
            forms.alert(
                "New breaker/frame values must be whole-number amps.\n\nFix rows like:\n{}".format("\n".join(invalid_rows[:8])),
                title=TITLE,
            )
            return False

        if not updates:
            forms.alert("No staged breaker/frame changes found.", title=TITLE)
            return False

        return self._raise_action_operation(
            "autosize_breaker_and_recalculate",
            [x.get("circuit_id") for x in updates],
            {"updates": updates, "show_output": False},
        )

    def _build_mark_existing_rows(self, targets):
        rows = []
        doc = self._get_active_doc()
        is_workshared = bool(getattr(doc, "IsWorkshared", False)) if doc is not None else False
        for item in targets:
            circuit = item.circuit
            reason = ""
            is_enabled = True
            locked, owner = self._is_locked_circuit(circuit, doc=doc) if is_workshared else (False, "")
            if locked:
                reason = "Locked by {}".format(owner or "another user")
                is_enabled = False

            current_notes = _lookup_schedule_notes_text(circuit)
            current_wire = self._conduit_wire_size_string(circuit)
            row = MarkExistingActionRow(
                item,
                is_enabled,
                reason,
                current_notes,
                current_wire,
            )
            if not is_enabled:
                row.is_checked = False
            rows.append(row)
        return rows

    def _preview_mark_existing_rows(self, rows, mode, set_notes, clear_wire, clear_conduit):
        mode_text = str(mode or "existing").lower()
        is_new_mode = mode_text == "new"
        for row in rows:
            row.new_notes = row.current_notes
            row.new_wire = row.current_wire
            if not row.is_enabled:
                row.is_checked = False
                row.recompute_state()
                continue
            if not row.is_checked:
                row.recompute_state()
                continue
            if set_notes:
                row.new_notes = "" if is_new_mode else "EX"
            if is_new_mode:
                row.new_wire = "Auto Calculated"
            else:
                row.new_wire = "Auto Calculated" if (clear_wire or clear_conduit) else row.current_wire
            row.recompute_state()

    def _apply_mark_existing_rows(self, rows, mode, set_notes, clear_wire, clear_conduit):
        changed_ids = [x.circuit_id for x in rows if x.is_enabled and x.is_checked and x.is_changed]
        if not changed_ids:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False
        mode_text = str(mode or "existing").lower()
        return self._raise_action_operation(
            "mark_existing_and_recalculate",
            changed_ids,
            {
                "mode": mode_text if mode_text in ("existing", "new") else "existing",
                "set_notes": bool(set_notes),
                "clear_wire": bool(clear_wire),
                "clear_conduit": bool(clear_conduit),
                "show_output": False,
            },
        )

    def action_mark_existing_clicked(self, sender, args):
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
        try:
            rows = self._build_mark_existing_rows(targets)
            window = MarkExistingActionWindow(
                rows,
                preview_callback=self._preview_mark_existing_rows,
                apply_callback=self._apply_mark_existing_rows,
                theme_mode=self._theme_mode,
                accent_mode=self._accent_mode,
            )
            window.ShowDialog()
        except Exception as ex:
            forms.alert("Failed to open Mark as New/Existing window:\n\n{}".format(ex), title=TITLE)

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
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
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
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
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
        try:
            rows = self._build_breaker_rows(targets)
            window = BreakerActionWindow(
                rows,
                preview_apply_callback=self._preview_breaker_rows,
                apply_callback=self._apply_breaker_rows,
                theme_mode=self._theme_mode,
                accent_mode=self._accent_mode,
            )
            window.ShowDialog()
        except Exception as ex:
            forms.alert("Failed to open breaker autosize window:\n\n{}".format(ex), title=TITLE)

    def _reset_filters(self):
        self._active_type_filters = set(self._type_options)
        self._warnings_only = False
        self._overrides_only = False
        self._checked_only = False
        self._update_filter_button_style()
        self._refresh_list()
        self._build_filter_menu()

    def filter_reset_clicked(self, sender, args):
        self._reset_filters()

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

    def _set_card_view(self, is_card_view):
        self._is_card_view = bool(is_card_view)
        if self._is_card_view:
            self._list.ItemTemplate = self._card_template
        else:
            self._list.ItemTemplate = self._compact_template
        self._update_toggle_button_visual()

    def toggle_view_clicked(self, sender, args):
        self._set_card_view(not self._is_card_view)

    def panel_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None or self._list is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if not _is_descendant_of_control(source, self._list):
            self._clear_list_selection()

    def list_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if _find_visual_ancestor(source, ListViewItem) is None:
            self._clear_list_selection()

    def list_preview_mouse_right_button_up(self, sender, args):
        if self._is_card_view:
            return
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        menu = ContextMenu()
        toggle = MenuItem()
        toggle.Header = "Show Circuit Type Badges"
        toggle.IsCheckable = True
        toggle.IsChecked = bool(self._compact_show_type_badges)
        toggle.Click += self.toggle_compact_badges_clicked
        menu.Items.Add(toggle)
        menu.PlacementTarget = self._list
        menu.IsOpen = True
        args.Handled = True

    def toggle_compact_badges_clicked(self, sender, args):
        self._compact_show_type_badges = bool(getattr(sender, "IsChecked", True))
        self._refresh_list()

    def _clear_list_selection(self):
        if self._list is None:
            return
        try:
            self._list.SelectedItems.Clear()
        except Exception:
            try:
                self._list.SelectedItem = None
            except Exception:
                pass
        self.selection_changed(None, None)

    def selection_changed(self, sender, args):
        selected_count = 0
        selected_items = []
        try:
            selected_items = list(self._list.SelectedItems)
            selected_count = len(selected_items)
        except Exception:
            selected_count = 0
        self._last_selected_items = selected_items
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
        items = list(self._all_items or [])
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
        items, removed_count = self._prune_stale_items(items)
        if removed_count:
            self._set_status("Removed {} deleted circuits from list.".format(removed_count))
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
        items, removed_count = self._prune_stale_items()
        if removed_count:
            self._set_status("Removed {} deleted circuits from list.".format(removed_count))
        circuit_ids = [x.circuit.Id.IntegerValue for x in items if getattr(x, "circuit", None)]
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
        if not self._has_active_doc():
            forms.alert("Open a model document first.", title=TITLE)
            return
        if self._settings_gateway.is_busy():
            forms.alert("Settings window request is already running.", title=TITLE)
            return
        self._set_status("Opening Calculate Circuits settings...")
        raised = self._settings_gateway.raise_open(callback=self._on_settings_open_complete)
        if not raised:
            self._set_status("Unable to queue settings window")

    def _on_settings_open_complete(self, status, error):
        if status == "error":
            self._set_status("Failed to open settings")
            forms.alert("Failed to open Calculate Circuits settings:\n\n{}".format(error), title=TITLE)
            return
        self._set_status("Closed Calculate Circuits settings")

    def _set_alert_hover_state(self, sender, is_hovered):
        item_container = _find_visual_ancestor(sender, ListViewItem)
        if item_container is None:
            return
        try:
            item_container.Tag = "alert-hover" if is_hovered else None
        except Exception:
            pass

    def alert_button_mouse_enter(self, sender, args):
        self._set_alert_hover_state(sender, True)

    def alert_button_mouse_leave(self, sender, args):
        self._set_alert_hover_state(sender, False)

    def alert_tag_clicked(self, sender, args):
        self._set_alert_hover_state(sender, False)
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
            window = CircuitAlertsWindow(
                header,
                rows,
                theme_mode=self._theme_mode,
                accent_mode=self._accent_mode,
            )
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
