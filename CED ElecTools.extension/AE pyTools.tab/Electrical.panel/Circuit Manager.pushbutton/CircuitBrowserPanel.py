# -*- coding: utf-8 -*-

import imp
import json
import os

<<<<<<< HEAD
=======
import clr
>>>>>>> main
import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.DB.Events import (
    DocumentOpenedEventArgs,
    DocumentClosedEventArgs,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
from System import EventHandler, Action
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
<<<<<<< HEAD
=======

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

>>>>>>> main
from System.Windows import Visibility
from System.Windows.Controls import (
    ContextMenu,
    MenuItem,
    Separator,
    DataGridRow,
    ListViewItem,
    Button,
    DataGridTextColumn,
    ScrollViewer,
    ScrollBarVisibility,
)
<<<<<<< HEAD
from System.Windows.Input import Keyboard, ModifierKeys
=======
from System.Windows.Input import Keyboard, ModifierKeys, Key
>>>>>>> main
from System.Windows.Media import BrushConverter, Stretch, VisualTreeHelper
from System.Windows.Shapes import Path as ShapePath
from pyrevit import forms, revit, DB, script, HOST_APP

_THIS_DIR = os.path.abspath(os.path.dirname(__file__))
from Snippets import revit_helpers

from CEDElectrical.Model.alerts import get_alert_definition
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_runner import build_default_runner
from CEDElectrical.Domain import settings_manager
from CEDElectrical.refdata.standard_ocp_table import BREAKER_FRAME_SWITCH_TABLE
from CEDElectrical.Infrastructure.Revit.external_events.circuit_operation_event import (
    CircuitOperationExternalEventGateway,
)
from CEDElectrical.ui.circuit_properties_editor import CircuitPropertiesEditorWindow
from CEDElectrical.Infrastructure.Revit.repositories.revit_circuit_repository import RevitCircuitRepository
from Snippets.circuit_ui_actions import (
    clear_revit_selection,
    collect_circuit_targets,
    format_writeback_lock_reason,
    set_revit_selection,
)
from Snippets._elecutils import (
    get_all_panels,
    get_compatible_panels,
    get_panel_dist_system,
    move_circuits_to_panel,
)
from UIClasses import Resources as UIResources
from UIClasses import pathing as ui_pathing
from UIClasses import resource_loader
from UIClasses.revit_theme_bridge import DOCK_PANE_FRAME_DARK, DOCK_PANE_FRAME_LIGHT, RevitThemeBridge

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(_THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    raise Exception("Could not locate workspace root for Circuit Manager.")

TITLE = "Circuit Manager"
PANEL_ID = "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2"
ALERT_DATA_PARAM = "Circuit Data_CED"
UI_RESOURCES_ROOT = (
    UIResources.get_resources_root()
    or ui_pathing.resolve_ui_resources_root(LIB_ROOT)
    or os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources"))
)
_ELECTRICAL_PANEL_ROOT = ui_pathing.find_named_ancestor(_THIS_DIR, "Electrical.panel") or os.path.abspath(
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
CURRENT_THEME_MODE = "light"
CURRENT_ACCENT_MODE = "blue"
THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_THEME_KEY = "theme_mode"
THEME_CONFIG_ACCENT_KEY = "accent_mode"
HIDABLE_ALERT_IDS = {
    "Design.NonStandardOCPRating",
    "Design.BreakerLugSizeLimitOverride",
    "Design.BreakerLugQuantityLimitOverride",
    "Design.CircuitLoadsNull",
    "Calculations.BreakerLugSizeLimit",
    "Calculations.BreakerLugQuantityLimit",
}
BLOCKED_BRANCH_TYPES = set(['N/A', 'SPACE', 'SPARE', 'CONDUIT ONLY'])
IG_BREAKER_ALLOWED_TYPES = set(['BRANCH', 'FEEDER', 'XFMR PRI', 'XFMR SEC'])

ACCENT_BRUSH_KEY_MAP = dict(resource_loader.ACCENT_BRUSH_KEY_MAP)
VALID_THEME_MODES = ("light", "dark", "dark_alt")
VALID_ACCENT_MODES = ("blue", "neutral")
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
CIRCUIT_TYPE_TAG_COMPACT_TEXT = {
    "BRANCH": "BR",
    "FEEDER": "FDR",
    "XFMR PRI": "XP",
    "XFMR SEC": "XS",
    "CONDUIT ONLY": "CO",
    "N/A": "NA",
    "SPARE": "SPR",
    "SPACE": "SPC",
}
OCP_TABLE_KEYS = sorted([int(k) for k in BREAKER_FRAME_SWITCH_TABLE.keys()])
_DOC_SENTINEL = object()
DEFAULT_HIDDEN_TYPE_FILTERS = set(["SPARE", "SPACE"])


def _normalize_theme_mode(value, fallback="light"):
    mode = str(value or fallback).strip().lower()
    return mode if mode in VALID_THEME_MODES else fallback


def _elid_value(item):
    return int(revit_helpers.get_elementid_value(item))


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


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


def _save_theme_state_to_config(theme_mode, accent_mode):
    try:
        cfg = script.get_config(THEME_CONFIG_SECTION)
        if cfg is None:
            return
        cfg.set_option(THEME_CONFIG_THEME_KEY, _normalize_theme_mode(theme_mode, "light"))
        cfg.set_option(THEME_CONFIG_ACCENT_KEY, _normalize_accent_mode(accent_mode, "blue"))
        script.save_config()
    except Exception:
        pass


CURRENT_THEME_MODE, CURRENT_ACCENT_MODE = _load_theme_state_from_config(
    CURRENT_THEME_MODE,
    CURRENT_ACCENT_MODE,
)


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


def _clip_with_ellipsis(value, max_chars):
    text = str(value or "")
    try:
        limit = int(max_chars or 0)
    except Exception:
        limit = 0
    if limit <= 3 or len(text) <= limit:
        return text
    return "{}...".format(text[: max(0, limit - 3)])


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


def _find_visual_descendant(start, target_type):
    if start is None:
        return None
    queue = [start]
    while queue:
        current = queue.pop(0)
        if isinstance(current, target_type):
            return current
        try:
            child_count = int(VisualTreeHelper.GetChildrenCount(current) or 0)
        except Exception:
            child_count = 0
        for idx in range(child_count):
            try:
                queue.append(VisualTreeHelper.GetChild(current, idx))
            except Exception:
                continue
    return None


def _find_visual_descendants(start, target_type):
    if start is None:
        return []
    results = []
    queue = [start]
    while queue:
        current = queue.pop(0)
        try:
            child_count = int(VisualTreeHelper.GetChildrenCount(current) or 0)
        except Exception:
            child_count = 0
        for idx in range(child_count):
            try:
                child = VisualTreeHelper.GetChild(current, idx)
            except Exception:
                continue
            if isinstance(child, target_type):
                results.append(child)
            queue.append(child)
    return results


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


def _try_apply_theme(owner):
    return bool(
        resource_loader.apply_theme(
            owner,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=getattr(owner, "_theme_mode", "light"),
            accent_mode=getattr(owner, "_accent_mode", "blue"),
        )
    )


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


def _menu_icon(owner, geometry_key, size=12):
    geometry = _try_find_resource(owner, geometry_key)
    if geometry is None:
        return None
    fill = _try_find_resource(owner, "CED.Brush.PrimaryText")
    if fill is None:
        fill = _try_find_resource(owner, "CED.Brush.NeutralDark")
    icon = ShapePath()
    icon.Data = geometry
    icon.Width = size
    icon.Height = size
    icon.Stretch = Stretch.Uniform
    if fill is not None:
        icon.Fill = fill
    return icon


def _invoke_later(owner, callback):
    if callback is None:
        return False
    try:
        dispatcher = getattr(owner, "Dispatcher", None)
        if dispatcher is not None:
            dispatcher.BeginInvoke(Action(callback))
            return True
    except Exception:
        pass
    try:
        callback()
        return True
    except Exception:
        return False


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


def _build_sync_lock_tooltip(circuit_owner, device_owner):
    lines = [
        "Calculation blocked by ownership constraints. Try again later.",
    ]
    owner_lines = []
    circuit_text = str(circuit_owner or "").strip()
    device_text = str(device_owner or "").strip()
    if circuit_text:
        owner_lines.append("Circuit Owner: {}".format(circuit_text))
    if device_text:
        owner_lines.append("Device Owner(s): {}".format(device_text))
    if owner_lines:
        lines.append("")
        lines.extend(owner_lines)
    return "\n".join(lines)


def _sync_lock_state_from_payload(payload):
    if not isinstance(payload, dict):
        return False, ""
    sync_lock = payload.get("sync_lock")
    if not isinstance(sync_lock, dict):
        return False, ""
    if not bool(sync_lock.get("blocked", False)):
        return False, ""
    tooltip = _build_sync_lock_tooltip(
        sync_lock.get("circuit_owner"),
        sync_lock.get("device_owner"),
    )
    return True, tooltip


def _sync_lock_state_from_row(row):
    if not isinstance(row, dict):
        return False, ""
    if not bool(row.get("blocked", False)):
        return False, ""
    tooltip = _build_sync_lock_tooltip(
        row.get("circuit_owner"),
        row.get("device_owner"),
    )
    return True, tooltip


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
    def __init__(self, circuit, session_sync_state=None):
        self.circuit = circuit
        self.circuit_id = _elid_value(circuit.Id)
        self.is_checked = False

        self.panel = "-"
        try:
            if circuit.BaseEquipment:
                self.panel = getattr(circuit.BaseEquipment, "Name", self.panel) or self.panel
        except Exception:
            pass

        self.circuit_number = getattr(circuit, "CircuitNumber", "") or ""
        self.load_name = getattr(circuit, "LoadName", "") or ""
        self.load_name_display = self.load_name
        self.load_name_visibility = "Visible"

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

        if circuit.CircuitType == DBE.CircuitType.Space:
            self.rating_poles = "/{}P".format(poles)
        elif rating_value is None:
            self.rating_poles = "- / {}P".format(poles)
        else:
            self.rating_poles = "{}A / {}P".format(int(round(float(rating_value), 0)), poles)

        device_count = 0
        try:
            count_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER_OF_ELEMENTS_PARAM)
            if count_param is not None:
                device_count = int(count_param.AsInteger() or 0)
        except Exception:
            device_count = 0
        self.device_line = "# Devices: {}".format(device_count)

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
        self.wire_line = "Wire: {}".format(conduit_wire)

        tag_bg, tag_fg = CIRCUIT_TYPE_TAG_STYLES.get(
            self.branch_type,
            ("CED.Brush.BadgeStd04Background", "CED.Brush.BadgeStd04Text"),
        )
        self.type_tag_text = self.branch_type
        compact_tag = CIRCUIT_TYPE_TAG_COMPACT_TEXT.get(self.branch_type)
        if compact_tag is None:
            compact_tag = str(self.branch_type or "-")[:3].upper()
        self.type_tag_text_compact = compact_tag
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

        sync_blocked, sync_tooltip = _sync_lock_state_from_payload(payload)
        if not sync_blocked:
            sync_blocked, sync_tooltip = _sync_lock_state_from_row(session_sync_state)
        self.sync_blocked = bool(sync_blocked)
        self.sync_lock_badge_visibility = "Visible" if sync_blocked else "Collapsed"
        self.sync_lock_tooltip = sync_tooltip
        self.item_max_width = 100000.0

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
        if _find_visual_ancestor(source, DataGridRow) is not None:
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
    def __init__(self, circuit, load_name, circuit_owner, device_owner):
        self.circuit = str(circuit or "-")
        self.load_name = str(load_name or "-")
        self.circuit_owner = str(circuit_owner or "-")
        self.device_owner = str(device_owner or "-")


class RuntimeAlertRow(object):
    def __init__(self, circuit, load_name, group, definition_id, message):
        self.circuit = str(circuit or "-")
        self.load_name = str(load_name or "-")
        self.group = str(group or "Other")
        self.definition_id = str(definition_id or "-")
        self.message = str(message or "")


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
                x.get("circuit_owner", "") or "-",
                x.get("device_owner", "") or "-",
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
                x.get("group", "") or "Other",
                x.get("definition_id", "") or "-",
                x.get("message", "") or "",
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

    def readonly_grid_selection_changed(self, sender, args):
        try:
            sender.SelectedItem = None
        except Exception:
            pass


class NeutralIGActionRow(object):
    def __init__(self, item, is_enabled, reason, current_qty, current_size, current_wire):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = _elid_value(item.circuit.Id)
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.panel_ckt_text = "{}/{}".format(item.panel or "-", item.circuit_number or "-")
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
<<<<<<< HEAD
=======
        self.target_include = None
>>>>>>> main
        self.is_changed = False
        self.remarks = ""
        self.recompute_state()

    def recompute_state(self):
<<<<<<< HEAD
        self.is_changed = bool(self.new_qty_changed or self.new_size_changed or self.new_wire_changed)
=======
        # Neutral / IG actions are staged strictly off quantity transitions.
        self.is_changed = bool(self.new_qty_changed)
>>>>>>> main
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
<<<<<<< HEAD
=======
        self._reset_selected_btn = self.FindName("ResetSelectedButton")
>>>>>>> main
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
<<<<<<< HEAD
        changed = len([x for x in self._rows if x.is_changed and x.is_enabled and x.is_checked])
        if self._show_unsupported:
            total = len(self._rows)
        else:
            total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
=======
        changed = len([x for x in self._rows if x.is_changed and x.is_enabled])
        total = len(self._rows)
>>>>>>> main
        if self._status is not None:
            self._status.Text = "{} of {} circuits to be modified.".format(changed, total)
        if self._checked_status is not None:
            checkable_total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
            checked = len([x for x in self._rows if x.is_enabled and x.is_checked])
            self._checked_status.Text = "{} of {} checked".format(checked, checkable_total)
        self._sync_action_buttons()

<<<<<<< HEAD
    def _sync_action_buttons(self):
        checked_count = len([x for x in self._rows if x.is_enabled and x.is_checked])
=======
    def _selected_rows(self):
        if self._grid is None:
            return []
        try:
            return list(self._grid.SelectedItems or [])
        except Exception:
            return []

    def _sync_action_buttons(self):
        checked_count = len([x for x in self._rows if x.is_enabled and x.is_checked])
        selected_changed = len([x for x in self._selected_rows() if x.is_enabled and x.is_changed])
>>>>>>> main
        if self._add_btn is not None:
            self._add_btn.IsEnabled = checked_count > 0
        if self._remove_btn is not None:
            self._remove_btn.IsEnabled = checked_count > 0
<<<<<<< HEAD
=======
        if self._reset_selected_btn is not None:
            self._reset_selected_btn.IsEnabled = selected_changed > 0
>>>>>>> main

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
<<<<<<< HEAD
=======
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
>>>>>>> main
        self._refresh_grid(refresh_items=False)

    def remove_clicked(self, sender, args):
        self._mode = "remove"
        self._preview_callback(self._rows, "remove")
<<<<<<< HEAD
=======
        self._is_syncing_checks = True
        try:
            for row in self._rows:
                row.is_checked = False
        finally:
            self._is_syncing_checks = False
        self._refresh_grid(refresh_items=False)

    def reset_selected_clicked(self, sender, args):
        targets = self._selected_rows()
        if not targets:
            self._sync_action_buttons()
            return
        for row in targets:
            if not bool(getattr(row, "is_enabled", False)):
                continue
            row.new_qty = row.current_qty
            row.new_size = row.current_size
            row.new_wire = "no change"
            row.new_wire_font_style = "Italic"
            row.new_qty_changed = False
            row.new_size_changed = False
            row.new_wire_changed = False
            row.target_include = None
            row.recompute_state()
>>>>>>> main
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
<<<<<<< HEAD
=======
        self._sync_action_buttons()
>>>>>>> main

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
<<<<<<< HEAD
=======
            self._sync_action_buttons()
>>>>>>> main


class BreakerActionRow(object):
    def __init__(self, item, is_enabled, reason, load_current, cur_rating, cur_frame, auto_rating, auto_frame_from_cur):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = _elid_value(item.circuit.Id)
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.panel_ckt_text = "{}/{}".format(item.panel or "-", item.circuit_number or "-")
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
        pct_text = "{}%".format(int(round(pct * 100.0, 0)))
        if pct >= 1.0:
            return 2, "#BE202F", "Load = {} of breaker rating".format(pct_text), "Visible", "Load = {} of rating".format(pct_text)
        if pct >= 0.90:
            return 1, "#F08A00", "Load = {} of breaker rating".format(pct_text), "Visible", "Load = {} of rating".format(pct_text)
        if pct >= 0.80:
            return 0, "#D8B300", "Load = {} of breaker rating".format(pct_text), "Visible", "Load = {} of rating".format(pct_text)
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
        self._max_load_pct_tb = self.FindName("MaxLoadPercentTextBox")
        self._show_unsupported_cb = self.FindName("ShowUnsupportedToggle")
        self._show_changed_only_cb = self.FindName("ShowChangedOnlyToggle")
        self._autosize_btn = self.FindName("AutosizePreviewButton")
        self._reset_btn = self.FindName("ResetButton")
        self._apply_btn = self.FindName("ApplyButton")
        self._row_filter_mode = "all"
        self._max_load_percent = 80

        if self._show_unsupported_cb is not None:
            self._show_unsupported_cb.IsChecked = True
        if self._show_changed_only_cb is not None:
            self._show_changed_only_cb.IsChecked = False
        if self._max_load_pct_tb is not None:
            self._max_load_pct_tb.Text = self._format_max_load_percent(self._max_load_percent)
            self._max_load_pct_tb.PreviewMouseLeftButtonDown += self.max_load_percent_preview_mouse_down
            self._max_load_pct_tb.GotKeyboardFocus += self.max_load_percent_got_focus
            self._max_load_pct_tb.LostFocus += self.max_load_percent_lost_focus
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

    def _format_max_load_percent(self, value):
        try:
            pct = int(round(float(value)))
        except Exception:
            pct = 80
        if pct < 50:
            pct = 50
        elif pct > 100:
            pct = 100
        return "{} %".format(pct)

    def _parse_max_load_percent(self, value):
        text = str(value or "").strip().replace("%", "").strip()
        try:
            pct = float(text)
        except Exception:
            pct = 80.0
        if pct < 50.0:
            pct = 50.0
        elif pct > 100.0:
            pct = 100.0
        return int(round(pct))

    def _get_max_load_percent(self):
        pct = self._parse_max_load_percent(getattr(self._max_load_pct_tb, "Text", None))
        if self._max_load_pct_tb is not None:
            self._max_load_pct_tb.Text = self._format_max_load_percent(pct)
            self._max_load_pct_tb.CaretIndex = len(self._max_load_pct_tb.Text or "")
        self._max_load_percent = pct
        return pct

    def max_load_percent_preview_mouse_down(self, sender, args):
        if sender is None:
            return
        try:
            if not sender.IsKeyboardFocusWithin:
                sender.Focus()
                args.Handled = True
                return
            sender.SelectAll()
            args.Handled = True
        except Exception:
            pass

    def max_load_percent_got_focus(self, sender, args):
        try:
            sender.SelectAll()
        except Exception:
            pass

    def max_load_percent_lost_focus(self, sender, args):
        self._get_max_load_percent()

    def apply_autosized_clicked(self, sender, args):
        set_breaker = bool(getattr(self._set_breaker_cb, "IsChecked", True))
        set_frame = bool(getattr(self._set_frame_cb, "IsChecked", True))
        allow_15a = bool(getattr(self._allow_15a_cb, "IsChecked", False))
        upsize_only = bool(getattr(self._upsize_only_cb, "IsChecked", False))
        max_load_percent = self._get_max_load_percent()
        self._preview_apply_callback(self._rows, set_breaker, set_frame, allow_15a, upsize_only, max_load_percent)
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
        max_load_percent = self._get_max_load_percent()
        if self._apply_callback(self._rows, set_breaker, set_frame, allow_15a, upsize_only, max_load_percent):
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
    def __init__(
        self,
        item,
        is_enabled,
        reason,
        current_notes,
        current_wire,
        current_sets,
        current_conduit_size,
        current_wire_size,
    ):
        self.item = item
        self.circuit = item.circuit
        self.circuit_id = _elid_value(item.circuit.Id)
        self.panel = item.panel
        self.circuit_number = item.circuit_number
        self.load_name = item.load_name
        self.panel_ckt_text = "{}/{}".format(item.panel or "-", item.circuit_number or "-")
        self.branch_type = item.branch_type
        self.is_enabled = bool(is_enabled)
        self.is_checked = False
        self.reason = reason or ""

        self.current_notes = current_notes or ""
        self.current_wire = current_wire or "-"
        self.current_sets = int(current_sets or 0)
        self.current_conduit_size = str(current_conduit_size or "").strip()
        self.current_wire_size = str(current_wire_size or "").strip()

        self.new_notes = self.current_notes
        self.new_wire = self.current_wire
        self.action_mode = "existing"
        self.action_set_notes = True
        self.preview_mode = "existing"
        self.preview_clear_wire = False
        self.preview_clear_conduit = False
        self.new_notes_changed = False
        self.new_wire_changed = False
        self.is_changed = False
        self.remarks = ""
        self.recompute_state()

    @staticmethod
    def _split_conduit_wire(text):
        raw = str(text or "").strip()
        if not raw or raw == "-":
            return "-", "-"
        parts = [x.strip() for x in raw.split("/") if x is not None]
        if len(parts) >= 2:
            conduit = parts[0] or "-"
            wire = parts[1] or "-"
            return conduit, wire
        value = parts[0] if parts else raw
        return value or "-", value or "-"

    def _existing_wire_only_display(self):
        wire = str(self.current_wire_size or "").strip()
        return wire or "-"

    def _existing_conduit_only_display(self):
        conduit = str(self.current_conduit_size or "").strip()
        if not conduit:
            return "-"
        sets = int(self.current_sets or 0)
        sets_text = str(sets if sets > 0 else 1)
        return "({}) {}C".format(sets_text, conduit)

    def recompute_state(self):
        self.new_notes_changed = str(self.new_notes or "") != str(self.current_notes or "")
        self.new_wire_changed = str(self.new_wire or "") != str(self.current_wire or "")
        self.is_changed = bool(self.new_notes_changed or self.new_wire_changed)
        if not self.is_enabled:
            self.remarks = "Blocked - {}".format(self.reason or "Unsupported")
        elif self.is_changed:
            mode_text = str(self.preview_mode or "existing").strip().lower()
            if mode_text == "new":
                self.remarks = "Will be modified - Conduit and Wire will recalculate"
            else:
                suffix = ""
                if bool(self.preview_clear_wire) and bool(self.preview_clear_conduit):
                    suffix = "Wire and Conduit Cleared"
                elif bool(self.preview_clear_wire):
                    suffix = "Wire Cleared"
                elif bool(self.preview_clear_conduit):
                    suffix = "Conduit Cleared"
                self.remarks = "Will be modified"
                if suffix:
                    self.remarks = "{} - {}".format(self.remarks, suffix)
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
        self._reset_selected_btn = self.FindName("ResetSelectedButton")
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
        changed = len([x for x in self._rows if x.is_enabled and x.is_changed])
        total = len([x for x in self._rows if bool(getattr(x, "is_enabled", False))])
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
        changed_count = len([x for x in self._rows if x.is_enabled and x.is_changed])
        selected_changed = len([x for x in self._selected_rows() if x.is_enabled and x.is_changed])
        if self._set_existing_btn is not None:
            self._set_existing_btn.IsEnabled = checked_count > 0
        if self._set_new_btn is not None:
            self._set_new_btn.IsEnabled = checked_count > 0
        if self._reset_selected_btn is not None:
            self._reset_selected_btn.IsEnabled = selected_changed > 0
        if self._apply_btn is not None:
            self._apply_btn.IsEnabled = changed_count > 0

    def _selected_rows(self):
        if self._grid is None:
            return []
        try:
            return list(self._grid.SelectedItems or [])
        except Exception:
            return []

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
        target_rows = [x for x in self._rows if x.is_enabled and x.is_checked]
        if not target_rows:
            self._sync_button_states()
            return
        self._preview_callback(target_rows, "existing", set_notes, clear_wire, clear_conduit)
        for row in target_rows:
            row.is_checked = False
        self._refresh_grid(refresh_items=False)

    def set_new_clicked(self, sender, args):
        set_notes, _, _ = self._option_values()
        self._last_mode = "new"
        target_rows = [x for x in self._rows if x.is_enabled and x.is_checked]
        if not target_rows:
            self._sync_button_states()
            return
        self._preview_callback(target_rows, "new", set_notes, False, False)
        for row in target_rows:
            row.is_checked = False
        self._refresh_grid(refresh_items=False)

    def reset_selected_clicked(self, sender, args):
        targets = self._selected_rows()
        if not targets:
            self._sync_button_states()
            return
        for row in targets:
            if not bool(getattr(row, "is_enabled", False)):
                continue
            row.new_notes = row.current_notes
            row.new_wire = row.current_wire
            row.action_mode = "existing"
            row.action_set_notes = True
            row.preview_mode = "existing"
            row.preview_clear_wire = False
            row.preview_clear_conduit = False
            row.is_checked = False
            row.recompute_state()
        self._refresh_grid(refresh_items=False)

    def apply_clicked(self, sender, args):
        set_notes, clear_wire, clear_conduit = self._option_values()
        if str(self._last_mode or "").strip().lower() == "new":
            clear_wire = False
            clear_conduit = False
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
        self._sync_button_states()

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
        if self._pending is not None:
            return True
        try:
            return bool(self._event.IsPending)
        except Exception:
            return False

    def raise_open(self, callback=None):
        if self._pending is not None:
            return False
        try:
            if bool(self._event.IsPending):
                return False
        except Exception:
            pass
        self._pending = {"callback": callback}
        try:
            self._event.Raise()
            return True
        except Exception as ex:
            self._pending = None
            if self.logger:
                self.logger.warning("Failed to raise Calculate Settings ExternalEvent: %s", ex)
            return False

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


class _BufferedMoveOutput(object):
    """Captures move-tool output so successful runs stay silent."""

    def __init__(self):
        self._events = []

    def linkify(self, element_id):
        return _elid_value(element_id)

    def print_md(self, text):
        self._events.append(("md", str(text or "")))

    def print_table(self, table_data, columns):
        self._events.append((
            "table",
            list(table_data or []),
            list(columns or []),
        ))

    def flush_to(self, output):
        if output is None:
            return
        for event in list(self._events or []):
            kind = event[0]
            if kind == "md":
                output.print_md(event[1])
                continue
            if kind == "table":
                output.print_table(event[1], event[2])


class MoveCircuitsExternalEventGateway(object):
    """Runs move-selected-circuits in API context from the modeless browser pane."""

    def __init__(self, logger=None, alert_parameter_name=ALERT_DATA_PARAM):
        self.logger = logger
        self.alert_parameter_name = alert_parameter_name
        self._pending = None
        self._handler = _MoveCircuitsHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def is_busy(self):
        if self._pending is not None:
            return True
        try:
            return bool(self._event.IsPending)
        except Exception:
            return False

    def raise_move(self, circuit_ids, target_panel_id, callback=None):
        if self._pending is not None:
            return False
        try:
            if bool(self._event.IsPending):
                return False
        except Exception:
            pass
        self._pending = {
            "circuit_ids": [int(x) for x in list(circuit_ids or []) if int(x or 0) > 0],
            "target_panel_id": int(target_panel_id or 0),
            "callback": callback,
        }
        try:
            self._event.Raise()
            return True
        except Exception as ex:
            self._pending = None
            if self.logger:
                self.logger.warning("Failed to raise Move Circuits ExternalEvent: %s", ex)
            return False

    def _consume_pending(self):
        pending = self._pending
        self._pending = None
        return pending


class _MoveCircuitsHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, application):
        pending = self._gateway._consume_pending()
        if not pending:
            return

        callback = pending.get("callback")
        status = "ok"
        error = None
        payload = {}
        try:
            uidoc = application.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if doc is None:
                raise Exception("No active Revit document available.")

            target_panel_id = int(pending.get("target_panel_id", 0) or 0)
            if target_panel_id <= 0:
                raise Exception("Target panel selection is invalid.")

            target_panel = doc.GetElement(_elid_from_value(target_panel_id))
            if target_panel is None:
                raise Exception("Target panel was not found in the active document.")

            circuits = []
            pre_on_target_ids = set()
            for cid in list(pending.get("circuit_ids") or []):
                circuit = doc.GetElement(_elid_from_value(int(cid)))
                if not isinstance(circuit, DBE.ElectricalSystem):
                    continue
                circuits.append(circuit)
                base_equipment = getattr(circuit, "BaseEquipment", None)
                if base_equipment is None:
                    continue
                if _elid_value(getattr(base_equipment, "Id", None)) == target_panel_id:
                    pre_on_target_ids.add(_elid_value(circuit.Id))

            if not circuits:
                raise Exception("No valid circuits were found to move.")

            buffered_output = _BufferedMoveOutput()
            move_result = move_circuits_to_panel(circuits, target_panel, doc, buffered_output)

            moved_ids = []
            for circuit in list(circuits or []):
                cid = _elid_value(circuit.Id)
                if cid in pre_on_target_ids:
                    continue
                base_equipment = getattr(circuit, "BaseEquipment", None)
                if base_equipment is None:
                    continue
                if _elid_value(getattr(base_equipment, "Id", None)) == target_panel_id:
                    moved_ids.append(cid)

            recalc_result = None
            recalc_error = None
            moved_ids = sorted(list(set([int(x) for x in list(moved_ids or []) if int(x) > 0])))
            if moved_ids:
                try:
                    runner = build_default_runner(alert_parameter_name=self._gateway.alert_parameter_name)
                    recalc_request = OperationRequest(
                        operation_key="calculate_circuits",
                        circuit_ids=moved_ids,
                        source="pane_move",
                        options={"show_output": False},
                    )
                    recalc_result = runner.run(recalc_request, doc)
                except Exception as ex:
                    recalc_error = ex
                    if self._gateway.logger:
                        self._gateway.logger.exception("Move succeeded but recalculation failed: %s", ex)

            payload = {
                "move_result": move_result,
                "buffered_output": buffered_output,
                "moved_ids": moved_ids,
                "recalc_result": recalc_result,
                "recalc_error": recalc_error,
            }
        except Exception as ex:
            status = "error"
            error = ex
            if self._gateway.logger:
                self._gateway.logger.exception("Move Circuits ExternalEvent failed: %s", ex)

        if callback:
            try:
                callback(status, payload, error)
            except Exception as cb_ex:
                if self._gateway.logger:
                    self._gateway.logger.exception("Move Circuits callback failed: %s", cb_ex)

    def GetName(self):
        return "CED Move Circuits External Event"


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
        self._item_index = {}
        self._session_sync_lock_map = {}
        self._is_card_view = False
        self._type_options = []
        self._active_type_filters = set()
        self._filter_menu = None
        self._filter_type_menu_items = {}
        self._filter_warn_item = None
        self._filter_warn_show_all_item = None
        self._filter_warn_active_only_item = None
        self._filter_overrides_item = None
        self._filter_syncblocked_item = None
        self._filter_checked_item = None
        self._suppress_filter_toggle_events = False
        self._warnings_only = False
        self._warnings_active_only = False
        self._overrides_only = False
        self._syncblocked_only = False
        self._checked_only = False
        self._actions_menu = None
        self._browser_options_menu = None
        self._browser_theme_items = {}
        self._browser_accent_items = {}
        self._browser_display_items = {}
        self._browser_compress_item = None
        self._compact_show_type_badges = True
        self._use_surface_item_states = True
        self._last_selected_items = []
        self._last_visible_ids = []
        self._type_tag_brush_cache = {}
        self._visible_items = ObservableCollection[CircuitListItem]()
        self._operation_gateway = self._get_operation_gateway()
        self._settings_gateway = CalculateSettingsExternalEventGateway(logger=self._logger)
        self._move_gateway = MoveCircuitsExternalEventGateway(
            logger=self._logger,
            alert_parameter_name=ALERT_DATA_PARAM,
        )
        self._lock_repository = RevitCircuitRepository()

        self._list = self.FindName("CircuitList")
        self._list_scrollviewer = None
        self._list_scrollviewer_hooked = False
        self._uniform_item_width = 0.0
        self._compress_item_width = False
        self._compress_hide_load_name = False
        self._skip_width_measure_on_next_refresh = False
        self._browser_compress_item = None
        self._applying_scroll_policy = False
        self._edit_properties_reselect_ids = []
        if self._list is not None:
            self._list.ItemsSource = self._visible_items
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
        self._events_attached = False
        self._view_activated_handler = None
        self._doc_opened_handler = None
        self._doc_closed_handler = None

        self.Loaded += self.panel_loaded
        self.Unloaded += self.panel_unloaded
        self.IsVisibleChanged += self.panel_visibility_changed
        self._ensure_theme_bridge()

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
            uiapp = self._get_uiapp()
            self._theme_bridge = RevitThemeBridge(
                uiapp=uiapp,
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
            self._session_sync_lock_map = {}
            self._set_status("Operation finished")
            self._safe_load_items()
            return

        self._update_session_sync_lock_map(result)

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

    def _update_session_sync_lock_map(self, result):
        rows = list((result or {}).get("locked_rows") or [])
        mapping = {}
        for row in rows:
            try:
                circuit_id = int((row or {}).get("circuit_id") or 0)
            except Exception:
                circuit_id = 0
            if circuit_id <= 0:
                continue
            mapping[circuit_id] = {
                "blocked": True,
                "circuit_owner": (row or {}).get("circuit_owner") or "",
                "device_owner": (row or {}).get("device_owner") or "",
            }
        self._session_sync_lock_map = mapping

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
        return getattr(revit, "doc", None) is not None and getattr(revit, "uidoc", None) is not None

    def _get_active_doc(self):
        try:
            uidoc = __revit__.ActiveUIDocument
            if uidoc:
                return uidoc.Document
        except Exception:
            pass
        return getattr(revit, "doc", None)

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

    def _get_uiapp(self):
        try:
            uiapp = getattr(HOST_APP, "uiapp", None)
            if uiapp is not None:
                return uiapp
        except Exception:
            pass
        try:
            return __revit__
        except Exception:
            return None

    def _get_revit_app(self):
        uiapp = self._get_uiapp()
        if uiapp is None:
            return None
        try:
            return getattr(uiapp, "Application", None)
        except Exception:
            return None

    def _has_api_context(self):
        try:
            return bool(HOST_APP.has_api_context)
        except Exception:
            return True

    def _attach_event_handlers(self):
        if self._events_attached:
            return True
        if not self._has_api_context():
            return False
        uiapp = self._get_uiapp()
        app = self._get_revit_app()
        if uiapp is None or app is None:
            return False
        try:
            if self._doc_opened_handler is None:
                self._doc_opened_handler = EventHandler[DocumentOpenedEventArgs](self._on_document_opened)
                app.DocumentOpened += self._doc_opened_handler
            if self._doc_closed_handler is None:
                self._doc_closed_handler = EventHandler[DocumentClosedEventArgs](self._on_document_closed)
                app.DocumentClosed += self._doc_closed_handler
            if self._view_activated_handler is None:
                self._view_activated_handler = EventHandler[ViewActivatedEventArgs](self._on_view_activated)
                uiapp.ViewActivated += self._view_activated_handler
            self._events_attached = True
            return True
        except Exception as ex:
            self._logger.warning("Circuit Browser failed to attach lifecycle handlers: %s", ex)
            self._detach_event_handlers()
            return False

    def _detach_event_handlers(self):
        uiapp = self._get_uiapp()
        app = self._get_revit_app()

        try:
            if uiapp is not None and self._view_activated_handler is not None:
                uiapp.ViewActivated -= self._view_activated_handler
        except Exception:
            pass
        self._view_activated_handler = None

        try:
            if app is not None and self._doc_opened_handler is not None:
                app.DocumentOpened -= self._doc_opened_handler
        except Exception:
            pass
        self._doc_opened_handler = None

        try:
            if app is not None and self._doc_closed_handler is not None:
                app.DocumentClosed -= self._doc_closed_handler
        except Exception:
            pass
        self._doc_closed_handler = None
        self._events_attached = False

    def _on_document_opened(self, sender, args):
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

    def _on_document_closed(self, sender, args):
        self._loaded_doc_key = None
        doc = self._get_active_doc()
        self._active_doc_key = self._doc_key(doc)
        if not self._is_pane_visible():
            return
        self._safe_load_items(doc_override=doc)

    def _on_view_activated(self, sender, args):
        if not self._is_pane_visible():
            return
        try:
            doc = getattr(args, "Document", None)
        except Exception:
            doc = None
        if doc is None:
            return

        key = self._doc_key(doc)
        if key != self._active_doc_key:
            self._active_doc_key = key
            if key != self._loaded_doc_key:
                self._safe_load_items(doc_override=doc)
                return

        self._set_doc_banner(doc)

    def _safe_load_items(self, doc_override=_DOC_SENTINEL, fast=False):
        doc = self._get_active_doc() if doc_override is _DOC_SENTINEL else doc_override
        self._active_doc_key = self._doc_key(doc)
        self._set_doc_banner(doc)
        if doc is None:
            self._all_items = []
            self._item_index = {}
            self._last_visible_ids = []
            try:
                self._visible_items.Clear()
            except Exception:
                self._visible_items = ObservableCollection[CircuitListItem]()
                if self._list is not None:
                    self._list.ItemsSource = self._visible_items
            self._loaded_doc_key = None
            self._set_status("Open a model document to load circuits.")
            return
        self._load_items(doc, fast=bool(fast))
        self._loaded_doc_key = self._active_doc_key

    def refresh_on_open(self):
        self._attach_event_handlers()
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
        default_active = self._default_active_type_filters(all_types)

        if not old_options or not old_active:
            self._active_type_filters = set(default_active)
        elif old_active == old_options:
            self._active_type_filters = set(all_types)
        else:
            hidden_types = old_options.difference(old_active)
            self._active_type_filters = all_set.difference(hidden_types)
            if not self._active_type_filters:
                self._active_type_filters = set(default_active)
        self._update_filter_button_style()

    def _default_active_type_filters(self, type_options=None):
        options = list(type_options if type_options is not None else (self._type_options or []))
        default_set = set()
        for ctype in options:
            key = str(ctype or "").strip().upper()
            if key in DEFAULT_HIDDEN_TYPE_FILTERS:
                continue
            default_set.add(ctype)
        if default_set:
            return default_set
        return set(options)

    def _update_filter_button_style(self):
        if self._filter_button is None:
            return
        default_active = set(self._default_active_type_filters())
        is_filtered = (
            self._warnings_only
            or self._overrides_only
            or self._syncblocked_only
            or self._checked_only
            or (default_active != set(self._active_type_filters))
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

    def _get_list_scrollviewer(self):
        viewer = _find_visual_descendant(self._list, ScrollViewer)
        if viewer is None:
            return self._list_scrollviewer
        if viewer is not self._list_scrollviewer:
            old = self._list_scrollviewer
            if old is not None and bool(self._list_scrollviewer_hooked):
                try:
                    old.ScrollChanged -= self.list_scroll_changed
                except Exception:
                    pass
            self._list_scrollviewer = viewer
            self._list_scrollviewer_hooked = False
        if not bool(self._list_scrollviewer_hooked):
            try:
                viewer.ScrollChanged += self.list_scroll_changed
                self._list_scrollviewer_hooked = True
            except Exception:
                pass
        self._list_scrollviewer = viewer
        return self._list_scrollviewer

    def _detach_list_scrollviewer(self):
        if self._list_scrollviewer is not None and bool(self._list_scrollviewer_hooked):
            try:
                self._list_scrollviewer.ScrollChanged -= self.list_scroll_changed
            except Exception:
                pass
        self._list_scrollviewer_hooked = False
        self._list_scrollviewer = None

    def _compute_uniform_item_width(self):
        viewer = self._get_list_scrollviewer()
        if viewer is None or self._list is None:
            return 0.0
        try:
            self._list.UpdateLayout()
        except Exception:
            pass
        try:
            viewer.UpdateLayout()
        except Exception:
            pass
        try:
            extent_width = float(getattr(viewer, "ExtentWidth", 0.0) or 0.0)
        except Exception:
            extent_width = 0.0
        try:
            viewport_width = float(getattr(viewer, "ViewportWidth", 0.0) or 0.0)
        except Exception:
            viewport_width = 0.0
        realized_max = 0.0
        for row in _find_visual_descendants(self._list, ListViewItem):
            if row is None:
                continue
            try:
                width = float(getattr(row, "ActualWidth", 0.0) or 0.0)
            except Exception:
                width = 0.0
            if width > realized_max:
                realized_max = width
        base_width = max(extent_width, viewport_width, realized_max)
        return float(base_width)

    def _compute_compressed_item_width(self):
        viewer = self._get_list_scrollviewer()
        if viewer is not None:
            try:
                viewport = float(getattr(viewer, "ViewportWidth", 0.0) or 0.0)
            except Exception:
                viewport = 0.0
        else:
            viewport = 0.0
        if viewport <= 0.0:
            try:
                viewport = float(getattr(self._list, "ActualWidth", 0.0) or 0.0)
            except Exception:
                viewport = 0.0
        if viewport <= 0.0:
            return 0.0
        return max(0.0, float(viewport - 2.0))

    def _compute_load_name_char_limit(self):
        viewport = self._compute_list_viewport_width()
        if viewport <= 0.0:
            return 24
        if bool(self._is_card_view):
            fixed_budget = 214.0
        else:
            fixed_budget = 196.0 if bool(self._compact_show_type_badges) else 166.0
        available = max(24.0, float(viewport - fixed_budget))
        approx_chars = int(available / 6.7)
        return max(6, min(approx_chars, 96))

    def _compute_list_viewport_width(self):
        viewer = self._get_list_scrollviewer()
        if viewer is not None:
            try:
                viewport = float(getattr(viewer, "ViewportWidth", 0.0) or 0.0)
            except Exception:
                viewport = 0.0
        else:
            viewport = 0.0
        if viewport <= 0.0:
            try:
                viewport = float(getattr(self._list, "ActualWidth", 0.0) or 0.0)
            except Exception:
                viewport = 0.0
        return max(0.0, float(viewport))

    def _use_compact_compress_mode(self):
        return bool(self._compress_item_width)

    def _load_name_hide_thresholds(self):
        if bool(self._is_card_view):
            hide = 250.0
        else:
            hide = 300.0 if bool(self._compact_show_type_badges) else 260.0
        return float(hide), float(hide + 36.0)

    def _should_hide_load_name_for_width(self, viewport_width):
        try:
            viewport = float(viewport_width or 0.0)
        except Exception:
            viewport = 0.0
        if viewport <= 0.0:
            return bool(self._compress_hide_load_name)
        hide_threshold, show_threshold = self._load_name_hide_thresholds()
        if bool(self._compress_hide_load_name):
            self._compress_hide_load_name = viewport < show_threshold
        else:
            self._compress_hide_load_name = viewport <= hide_threshold
        return bool(self._compress_hide_load_name)

    def _reset_horizontal_offset_for_compress(self, viewer=None):
        if not self._use_compact_compress_mode():
            return
        if viewer is None:
            viewer = self._get_list_scrollviewer()
        if viewer is None:
            return
        try:
            if float(getattr(viewer, "HorizontalOffset", 0.0) or 0.0) != 0.0:
                viewer.ScrollToHorizontalOffset(0.0)
        except Exception:
            pass

    def _apply_horizontal_scroll_policy(self, viewer):
        if viewer is None:
            return
        desired = ScrollBarVisibility.Disabled if self._use_compact_compress_mode() else ScrollBarVisibility.Auto
        try:
            if getattr(viewer, "HorizontalScrollBarVisibility", ScrollBarVisibility.Auto) != desired:
                viewer.HorizontalScrollBarVisibility = desired
        except Exception:
            pass
        if self._use_compact_compress_mode():
            self._reset_horizontal_offset_for_compress(viewer)

    def _apply_uniform_item_width_to_realized_rows(self):
        if self._list is None:
            return
        if self._use_compact_compress_mode():
            return
        width_value = float(self._uniform_item_width or 0.0)
        use_uniform = width_value > 0.0
        for row in _find_visual_descendants(self._list, ListViewItem):
            if row is None:
                continue
            if use_uniform:
                try:
                    row.MinWidth = width_value
                except Exception:
                    pass
                try:
                    row.Width = width_value
                except Exception:
                    pass
            else:
                try:
                    row.MinWidth = 0.0
                except Exception:
                    pass
                try:
                    row.Width = float("nan")
                except Exception:
                    pass

    def _apply_item_width_mode(self, items):
        records = list(items or [])
        use_compress = self._use_compact_compress_mode()
        hide_load_name = False
        char_limit = 0
        max_item_width = 100000.0
        if use_compress:
            viewport_width = self._compute_list_viewport_width()
            hide_load_name = self._should_hide_load_name_for_width(viewport_width)
            char_limit = 0 if bool(self._is_card_view) else self._compute_load_name_char_limit()
            max_item_width = max(0.0, float(viewport_width - 2.0))
        else:
            self._compress_hide_load_name = False
        for item in records:
            full_name = str(getattr(item, "load_name", "") or "")
            if use_compress and hide_load_name:
                item.load_name_display = ""
                item.load_name_visibility = "Collapsed"
            else:
                item.load_name_display = _clip_with_ellipsis(full_name, char_limit) if (use_compress and char_limit > 0) else full_name
                item.load_name_visibility = "Visible"
            item.item_max_width = max_item_width if use_compress else 100000.0

    def list_scroll_changed(self, sender, args):
        viewer = sender if isinstance(sender, ScrollViewer) else self._get_list_scrollviewer()
        if bool(self._applying_scroll_policy):
            return
        if self._use_compact_compress_mode():
            try:
                viewport_change = abs(float(getattr(args, "ViewportWidthChange", 0.0) or 0.0)) > 0.0
            except Exception:
                viewport_change = False
            self._applying_scroll_policy = True
            try:
                if viewport_change:
                    self._apply_item_width_mode(self._visible_items)
                    self._refresh_visible_items()
                self._uniform_item_width = 0.0
                self._apply_uniform_item_width_to_realized_rows()
                self._apply_horizontal_scroll_policy(viewer)
                self._reset_horizontal_offset_for_compress(viewer)
            finally:
                self._applying_scroll_policy = False
            return
        self._uniform_item_width = self._compute_uniform_item_width()
        self._apply_uniform_item_width_to_realized_rows()

    def list_preview_mouse_wheel(self, sender, args):
        none_mod = getattr(ModifierKeys, "None")
        try:
            modifiers = Keyboard.Modifiers
        except Exception:
            modifiers = none_mod
        if (modifiers & ModifierKeys.Shift) != ModifierKeys.Shift:
            return
        if self._use_compact_compress_mode():
            return
        viewer = self._get_list_scrollviewer()
        if viewer is None:
            return
        try:
            delta = int(getattr(args, "Delta", 0) or 0)
        except Exception:
            delta = 0
        if delta == 0:
            return
        step = 40.0
        try:
            current = float(viewer.HorizontalOffset)
        except Exception:
            current = 0.0
        if delta > 0:
            target = max(0.0, current - step)
        else:
            target = current + step
        try:
            viewer.ScrollToHorizontalOffset(float(target))
            args.Handled = True
        except Exception:
            pass

    def _clear_type_tag_brush_cache(self):
        self._type_tag_brush_cache = {}

    def _resolve_type_tag_brush(self, resource_key, fallback):
        if not resource_key:
            return fallback
        if resource_key in self._type_tag_brush_cache:
            return self._type_tag_brush_cache[resource_key]
        resolved = _try_find_resource(self, resource_key)
        value = resolved if resolved is not None else fallback
        self._type_tag_brush_cache[resource_key] = value
        return value

    def _apply_type_tag_brush(self, item):
        if item is None:
            return
        bg_key = getattr(item, "type_tag_bg_key", None)
        fg_key = getattr(item, "type_tag_fg_key", None)
        item.type_tag_bg = self._resolve_type_tag_brush(bg_key, "#ECEFF3")
        item.type_tag_fg = self._resolve_type_tag_brush(fg_key, "#52606D")

    def _apply_type_tag_brushes(self):
        for item in list(self._all_items or []):
            self._apply_type_tag_brush(item)

    def _refresh_visible_items(self):
        try:
            if self._list is not None:
                self._list.Items.Refresh()
                return
        except Exception:
            pass
        self._refresh_list()

    def _set_visible_items(self, items):
        items = list(items or [])
        visible_ids = [int(getattr(item, "circuit_id", 0) or 0) for item in items]
        same_ids = visible_ids == list(self._last_visible_ids or [])
        same_refs = False
        if same_ids:
            try:
                current_items = list(self._visible_items or [])
            except Exception:
                current_items = []
            if len(current_items) == len(items):
                same_refs = True
                for idx, item in enumerate(items):
                    if current_items[idx] is not item:
                        same_refs = False
                        break
        if same_ids and same_refs:
            try:
                if self._list is not None:
                    self._list.Items.Refresh()
            except Exception:
                pass
            return
        self._last_visible_ids = list(visible_ids)
        try:
            self._visible_items.Clear()
            for item in items:
                self._visible_items.Add(item)
        except Exception:
            self._visible_items = ObservableCollection[CircuitListItem](items)
            if self._list is not None:
                self._list.ItemsSource = self._visible_items

    def _refresh_list(self):
        query = ""
        try:
            query = (self._search.Text or "").strip().lower()
        except Exception:
            query = ""

        items = list(self._all_items)
        if self._warnings_only:
            items = [x for x in items if int(getattr(x, "alert_count", 0) or 0) > 0]
            if self._warnings_active_only:
                items = [
                    x for x in items
                    if (int(getattr(x, "alert_count", 0) or 0) - int(getattr(x, "hidden_alert_count", 0) or 0)) > 0
                ]
        elif self._overrides_only:
            items = [x for x in items if bool(getattr(x, "has_override", False))]
        elif self._syncblocked_only:
            items = [x for x in items if bool(getattr(x, "sync_blocked", False))]
        elif self._checked_only:
            items = [x for x in items if bool(getattr(x, "is_checked", False))]
        else:
            items = [x for x in items if x.branch_type in self._active_type_filters]
        if query:
            items = [x for x in items if query in x.search_name]
        for item in items:
            item.show_type_tag = bool(self._compact_show_type_badges)
            item.type_tag_visibility = "Visible" if item.show_type_tag else "Collapsed"
        self._apply_item_width_mode(items)

        self._set_visible_items(items)
        viewer = self._get_list_scrollviewer()
        self._apply_horizontal_scroll_policy(viewer)
        if self._use_compact_compress_mode():
            self._uniform_item_width = 0.0
            self._apply_uniform_item_width_to_realized_rows()
            self._reset_horizontal_offset_for_compress(viewer)
        else:
            if bool(self._skip_width_measure_on_next_refresh):
                self._skip_width_measure_on_next_refresh = False
                self._uniform_item_width = 0.0
            else:
                self._uniform_item_width = self._compute_uniform_item_width()
            self._apply_uniform_item_width_to_realized_rows()
        self._set_status("Showing {} of {} circuits".format(len(items), len(self._all_items)))

    def list_size_changed(self, sender, args):
        if not self._use_compact_compress_mode():
            return
        if bool(self._applying_scroll_policy):
            return
        self._applying_scroll_policy = True
        try:
            self._apply_item_width_mode(self._visible_items)
            self._refresh_visible_items()
            viewer = self._get_list_scrollviewer()
            self._apply_horizontal_scroll_policy(viewer)
            self._apply_uniform_item_width_to_realized_rows()
            self._reset_horizontal_offset_for_compress(viewer)
        finally:
            self._applying_scroll_policy = False

    def _collect_sorted_circuits(self, doc):
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
        return circuits

    def _session_state_for_circuit(self, circuit):
        try:
            return self._session_sync_lock_map.get(_elid_value(circuit.Id))
        except Exception:
            return None

    def _load_items_full(self, circuits):
        checked_ids = set([int(item.circuit_id) for item in list(self._all_items or []) if bool(getattr(item, "is_checked", False))])
        refreshed_items = []
        refreshed_index = {}
        for circuit in list(circuits or []):
            item = CircuitListItem(circuit, session_sync_state=self._session_state_for_circuit(circuit))
            item.is_checked = int(item.circuit_id) in checked_ids
            self._apply_type_tag_brush(item)
            refreshed_items.append(item)
            refreshed_index[int(item.circuit_id)] = item
        self._all_items = refreshed_items
        self._item_index = refreshed_index
        self._rebuild_filter_options()
        self._refresh_list()

    def _load_items_fast(self, circuits):
        existing_index = dict(self._item_index or {})
        refreshed_items = []
        refreshed_index = {}
        added = 0
        removed = 0
        for circuit in list(circuits or []):
            cid = int(_elid_value(circuit.Id))
            item = existing_index.pop(cid, None)
            if item is None:
                item = CircuitListItem(circuit, session_sync_state=self._session_state_for_circuit(circuit))
                self._apply_type_tag_brush(item)
                added += 1
            else:
                item.circuit = circuit
            refreshed_items.append(item)
            refreshed_index[cid] = item
        removed = len(existing_index)
        self._all_items = refreshed_items
        self._item_index = refreshed_index
        self._rebuild_filter_options()
        self._refresh_list()
        if added or removed:
            self._logger.debug("Circuit Browser fast refresh: +%s / -%s", added, removed)

    def _load_items(self, doc, fast=False):
        self._set_status("Loading circuits...")
        circuits = self._collect_sorted_circuits(doc)
        if bool(fast):
            self._load_items_fast(circuits)
            return
        self._load_items_full(circuits)

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
                    cid = _elid_value(item.circuit.Id)
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
                cid = _elid_value(item.circuit.Id)
            except Exception:
                cid = None
            if cid in target_ids:
                valid_targets.append(item)

        if removed_count:
            self._all_items = valid_all
            self._item_index = {int(getattr(x, "circuit_id", 0) or 0): x for x in list(valid_all or [])}
            self._rebuild_filter_options()
            self._refresh_list()
            self.selection_changed(None, None)
            self._logger.debug("Circuit Browser pruned %s stale rows.", removed_count)

        if target_ids is None:
            return list(valid_all), removed_count
        return valid_targets, removed_count

    def _set_revit_selection(self, elements):
        set_revit_selection(elements, uidoc=revit.uidoc)

<<<<<<< HEAD
=======
    def _show_and_select_revit_targets(self, elements):
        uidoc = getattr(revit, "uidoc", None)
        if uidoc is None:
            return
        ids = List[DB.ElementId]()
        seen = set()
        for element in list(elements or []):
            element_id = getattr(element, "Id", None)
            if element_id is None:
                continue
            value = _elid_value(element_id)
            if value <= 0 or value in seen:
                continue
            seen.add(value)
            ids.Add(element_id)
        if ids.Count <= 0:
            return
        try:
            uidoc.ShowElements(ids)
        except Exception:
            pass
        set_revit_selection(elements, uidoc=uidoc)

>>>>>>> main
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

        changed_any = False
        for item in targets:
            next_state = bool(state)
            if bool(getattr(item, "is_checked", False)) == next_state:
                continue
            item.is_checked = next_state
            changed_any = True

        if not changed_any:
            return

        # Only force list refresh when visibility/content must change.
        if self._checked_only:
            self._refresh_list()
            return
        if len(targets) > 1:
            try:
                self._list.Items.Refresh()
            except Exception:
                pass

    def _build_filter_menu(self):
        menu = ContextMenu()
        _set_if_resource(self, menu, "Style", "CED.ContextMenu.Base")
        self._filter_type_menu_items = {}
        self._filter_warn_item = None
        self._filter_warn_show_all_item = None
        self._filter_warn_active_only_item = None
        self._filter_overrides_item = None
        self._filter_syncblocked_item = None
        self._filter_checked_item = None

        warn_item = MenuItem()
        _set_if_resource(self, warn_item, "Style", "CED.MenuItem.Base")
        warn_item.Header = "Circuits With Alerts"
        warn_item.Icon = _menu_icon(self, "CED.Icon.AlertBoxOutline")
        # Parent is submenu-only for behavior; child items control filtering.
        warn_item.IsCheckable = True
        warn_item.IsChecked = bool(self._warnings_only)
        warn_item.StaysOpenOnClick = True
        warn_item.Click += self.filter_warnings_parent_clicked
        warn_item.MouseEnter += self.filter_warnings_parent_mouse_enter
        warn_show_all_item = MenuItem()
        _set_if_resource(self, warn_show_all_item, "Style", "CED.MenuItem.Base")
        warn_show_all_item.Header = "Show All"
        warn_show_all_item.IsCheckable = True
        warn_show_all_item.IsChecked = bool(self._warnings_only and (not self._warnings_active_only))
        warn_show_all_item.StaysOpenOnClick = True
        warn_show_all_item.Click += self.filter_warnings_show_all_clicked
        warn_item.Items.Add(warn_show_all_item)
        self._filter_warn_show_all_item = warn_show_all_item

        warn_active_item = MenuItem()
        _set_if_resource(self, warn_active_item, "Style", "CED.MenuItem.Base")
        warn_active_item.Header = "Active Alerts Only"
        warn_active_item.IsCheckable = True
        warn_active_item.IsChecked = bool(self._warnings_only and self._warnings_active_only)
        warn_active_item.StaysOpenOnClick = True
        warn_active_item.Click += self.filter_warnings_active_only_clicked
        warn_item.Items.Add(warn_active_item)
        self._filter_warn_active_only_item = warn_active_item
        menu.Items.Add(warn_item)
        self._filter_warn_item = warn_item

        overrides_item = MenuItem()
        _set_if_resource(self, overrides_item, "Style", "CED.MenuItem.Base")
        overrides_item.Header = "User Overrides"
        overrides_item.Icon = _menu_icon(self, "CED.Icon.Account")
        overrides_item.IsCheckable = True
        overrides_item.IsChecked = self._overrides_only
        overrides_item.StaysOpenOnClick = True
        overrides_item.Checked += self.filter_overrides_toggled
        overrides_item.Unchecked += self.filter_overrides_toggled
        menu.Items.Add(overrides_item)
        self._filter_overrides_item = overrides_item

        syncblocked_item = MenuItem()
        _set_if_resource(self, syncblocked_item, "Style", "CED.MenuItem.Base")
        syncblocked_item.Header = "Failed Calculations"
        syncblocked_item.Icon = _menu_icon(self, "CED.Icon.SyncAlert")
        syncblocked_item.IsCheckable = True
        syncblocked_item.IsChecked = self._syncblocked_only
        syncblocked_item.StaysOpenOnClick = True
        syncblocked_item.Checked += self.filter_sync_blocked_toggled
        syncblocked_item.Unchecked += self.filter_sync_blocked_toggled
        menu.Items.Add(syncblocked_item)
        self._filter_syncblocked_item = syncblocked_item

        checked_item = MenuItem()
        _set_if_resource(self, checked_item, "Style", "CED.MenuItem.Base")
        checked_item.Header = "Checked Circuits Only"
        checked_item.Icon = _menu_icon(self, "CED.Icon.CheckCircleOutline")
        checked_item.IsCheckable = True
        checked_item.IsChecked = self._checked_only
        checked_item.StaysOpenOnClick = True
        checked_item.Checked += self.filter_checked_toggled
        checked_item.Unchecked += self.filter_checked_toggled
        menu.Items.Add(checked_item)
        self._filter_checked_item = checked_item
        sep_top = Separator()
        _set_if_resource(self, sep_top, "Style", "CED.Separator.Menu")
        menu.Items.Add(sep_top)

        reset_item = MenuItem()
        _set_if_resource(self, reset_item, "Style", "CED.MenuItem.Base")
        reset_item.Header = "Reset Filters"
        reset_item.StaysOpenOnClick = False
        reset_item.Click += self.filter_reset_clicked
        menu.Items.Add(reset_item)

        sep_mid = Separator()
        _set_if_resource(self, sep_mid, "Style", "CED.Separator.Menu")
        menu.Items.Add(sep_mid)

        for ctype in self._type_options:
            mi = MenuItem()
            _set_if_resource(self, mi, "Style", "CED.MenuItem.Base")
            mi.Header = ctype
            mi.IsCheckable = True
            mi.IsChecked = ctype in self._active_type_filters
            mi.StaysOpenOnClick = True
            mi.Tag = ctype
            mi.IsEnabled = not (
                self._warnings_only
                or self._overrides_only
                or self._syncblocked_only
                or self._checked_only
            )
            mi.Checked += self.filter_type_toggled
            mi.Unchecked += self.filter_type_toggled
            menu.Items.Add(mi)
            self._filter_type_menu_items[ctype] = mi

        self._filter_menu = menu
        if self._filter_button is not None:
            self._filter_button.ContextMenu = menu
        self._sync_filter_menu_state()

    def _sync_filter_menu_state(self):
        self._suppress_filter_toggle_events = True
        try:
            if self._filter_warn_item is not None:
                self._filter_warn_item.IsChecked = bool(self._warnings_only)
            if self._filter_warn_show_all_item is not None:
                self._filter_warn_show_all_item.IsChecked = bool(self._warnings_only and (not self._warnings_active_only))
                self._filter_warn_show_all_item.IsEnabled = True
            if self._filter_warn_active_only_item is not None:
                self._filter_warn_active_only_item.IsChecked = bool(self._warnings_only and self._warnings_active_only)
                self._filter_warn_active_only_item.IsEnabled = True
            if self._filter_overrides_item is not None:
                self._filter_overrides_item.IsChecked = bool(self._overrides_only)
            if self._filter_syncblocked_item is not None:
                self._filter_syncblocked_item.IsChecked = bool(self._syncblocked_only)
            if self._filter_checked_item is not None:
                self._filter_checked_item.IsChecked = bool(self._checked_only)

            disable_types = bool(
                self._warnings_only
                or self._overrides_only
                or self._syncblocked_only
                or self._checked_only
            )
            for ctype, item in (self._filter_type_menu_items or {}).items():
                if item is None:
                    continue
                item.IsEnabled = not disable_types
                item.IsChecked = ctype in self._active_type_filters
        finally:
            self._suppress_filter_toggle_events = False

    def _set_exclusive_filter(self, mode_name, is_checked):
        checked = bool(is_checked)
        if checked:
            self._warnings_only = mode_name == "warnings"
            self._overrides_only = mode_name == "overrides"
            self._syncblocked_only = mode_name == "syncblocked"
            self._checked_only = mode_name == "checked"
            self._active_type_filters = set(self._type_options)
            return
        if mode_name == "warnings":
            self._warnings_only = False
        elif mode_name == "overrides":
            self._overrides_only = False
        elif mode_name == "syncblocked":
            self._syncblocked_only = False
        elif mode_name == "checked":
            self._checked_only = False

    def panel_loaded(self, sender, args):
        if not self._is_pane_visible():
            return
        self._sync_theme_from_config(apply_if_changed=True)
        self._get_list_scrollviewer()
        self._attach_event_handlers()
        doc = self._get_active_doc()
        key = self._doc_key(doc)
        self._active_doc_key = key
        if key != self._loaded_doc_key:
            self._safe_load_items(doc_override=doc)
        else:
            self._set_doc_banner(doc)

    def panel_unloaded(self, sender, args):
        self._detach_list_scrollviewer()
        self._detach_event_handlers()

    def panel_visibility_changed(self, sender, args):
        if not self._is_pane_visible():
            self._detach_event_handlers()
            return
        self._sync_theme_from_config(apply_if_changed=True)
        self._attach_event_handlers()
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
        # Manual refresh should reflect latest circuit metadata (name/number/panel/type).
        self._skip_width_measure_on_next_refresh = True
        self._sync_theme_from_config(apply_if_changed=False)
        self._safe_load_items()

    def filter_button_clicked(self, sender, args):
        self._build_filter_menu()
        if self._filter_menu is not None:
            self._filter_menu.PlacementTarget = self._filter_button
            self._filter_menu.IsOpen = True

    def _build_actions_menu(self):
        menu = ContextMenu()
        _set_if_resource(self, menu, "Style", "CED.ContextMenu.Base")

        neutral_item = MenuItem()
        _set_if_resource(self, neutral_item, "Style", "CED.MenuItem.Base")
        neutral_item.Header = "Add/Remove Neutral"
        neutral_item.Click += self.action_neutral_clicked
        menu.Items.Add(neutral_item)

        ig_item = MenuItem()
        _set_if_resource(self, ig_item, "Style", "CED.MenuItem.Base")
        ig_item.Header = "Add/Remove IG"
        ig_item.Click += self.action_ig_clicked
        menu.Items.Add(ig_item)

        breaker_item = MenuItem()
        _set_if_resource(self, breaker_item, "Style", "CED.MenuItem.Base")
        breaker_item.Header = "Auto Size Breaker"
        breaker_item.Click += self.action_breaker_clicked
        menu.Items.Add(breaker_item)

        mark_existing_item = MenuItem()
        _set_if_resource(self, mark_existing_item, "Style", "CED.MenuItem.Base")
        mark_existing_item.Header = "Mark as New/Existing"
        mark_existing_item.Click += self.action_mark_existing_clicked
        menu.Items.Add(mark_existing_item)

        edit_properties_item = MenuItem()
        _set_if_resource(self, edit_properties_item, "Style", "CED.MenuItem.Base")
        edit_properties_item.Header = "Edit Circuit Properties"
        edit_properties_item.Click += self.action_edit_properties_clicked
        menu.Items.Add(edit_properties_item)

        sep = Separator()
        _set_if_resource(self, sep, "Style", "CED.Separator.Menu")
        menu.Items.Add(sep)

        move_item = MenuItem()
        _set_if_resource(self, move_item, "Style", "CED.MenuItem.Base")
        move_item.Header = "Move Selected Circuits"
        move_item.Click += self.action_move_checked_clicked
        menu.Items.Add(move_item)

        self._actions_menu = menu

    def actions_button_clicked(self, sender, args):
        self._build_actions_menu()
        if self._actions_menu is not None:
            self._actions_menu.PlacementTarget = sender
            self._actions_menu.IsOpen = True

    def _build_browser_options_menu(self):
        menu = ContextMenu()
        menu.StaysOpen = True
        _set_if_resource(self, menu, "Style", "CED.ContextMenu.Base")
        self._browser_theme_items = {}
        self._browser_accent_items = {}
        self._browser_display_items = {}

        theme_menu = MenuItem()
        _set_if_resource(self, theme_menu, "Style", "CED.MenuItem.Base")
        theme_menu.Header = "Theme"
        light_item = MenuItem()
        _set_if_resource(self, light_item, "Style", "CED.MenuItem.Base")
        light_item.Header = "Light"
        light_item.IsCheckable = True
        light_item.IsChecked = (self._theme_mode == "light")
        light_item.StaysOpenOnClick = True
        light_item.Tag = "light"
        light_item.Click += self.browser_theme_clicked
        theme_menu.Items.Add(light_item)
        self._browser_theme_items["light"] = light_item
        dark_item = MenuItem()
        _set_if_resource(self, dark_item, "Style", "CED.MenuItem.Base")
        dark_item.Header = "Dark"
        dark_item.IsCheckable = True
        dark_item.IsChecked = (self._theme_mode == "dark")
        dark_item.StaysOpenOnClick = True
        dark_item.Tag = "dark"
        dark_item.Click += self.browser_theme_clicked
        theme_menu.Items.Add(dark_item)
        self._browser_theme_items["dark"] = dark_item
        dark_alt_item = MenuItem()
        _set_if_resource(self, dark_alt_item, "Style", "CED.MenuItem.Base")
        dark_alt_item.Header = "Dark Alt"
        dark_alt_item.IsCheckable = True
        dark_alt_item.IsChecked = (self._theme_mode == "dark_alt")
        dark_alt_item.StaysOpenOnClick = True
        dark_alt_item.Tag = "dark_alt"
        dark_alt_item.Click += self.browser_theme_clicked
        theme_menu.Items.Add(dark_alt_item)
        self._browser_theme_items["dark_alt"] = dark_alt_item
        menu.Items.Add(theme_menu)

        accent_menu = MenuItem()
        _set_if_resource(self, accent_menu, "Style", "CED.MenuItem.Base")
        accent_menu.Header = "Accent Color"
        for accent_mode, accent_label in (
            ("blue", "Blue"),
            ("neutral", "Neutral"),
        ):
            item = MenuItem()
            _set_if_resource(self, item, "Style", "CED.MenuItem.Base")
            item.Header = accent_label
            item.IsCheckable = True
            item.IsChecked = (self._accent_mode == accent_mode)
            item.StaysOpenOnClick = True
            item.Tag = accent_mode
            item.Click += self.browser_accent_clicked
            accent_menu.Items.Add(item)
            self._browser_accent_items[accent_mode] = item
        menu.Items.Add(accent_menu)

        display_menu = MenuItem()
        _set_if_resource(self, display_menu, "Style", "CED.MenuItem.Base")
        display_menu.Header = "Display Mode"
        compact_item = MenuItem()
        _set_if_resource(self, compact_item, "Style", "CED.MenuItem.Base")
        compact_item.Header = "Compact"
        compact_item.IsCheckable = True
        compact_item.IsChecked = not self._is_card_view
        compact_item.StaysOpenOnClick = True
        compact_item.Tag = "compact"
        compact_item.Click += self.browser_display_mode_clicked
        display_menu.Items.Add(compact_item)
        self._browser_display_items["compact"] = compact_item
        card_item = MenuItem()
        _set_if_resource(self, card_item, "Style", "CED.MenuItem.Base")
        card_item.Header = "Card"
        card_item.IsCheckable = True
        card_item.IsChecked = self._is_card_view
        card_item.StaysOpenOnClick = True
        card_item.Tag = "card"
        card_item.Click += self.browser_display_mode_clicked
        display_menu.Items.Add(card_item)
        self._browser_display_items["card"] = card_item
        menu.Items.Add(display_menu)

        compact_menu = MenuItem()
        _set_if_resource(self, compact_menu, "Style", "CED.MenuItem.Base")
        compact_menu.Header = "List View"
        badges_item = MenuItem()
        _set_if_resource(self, badges_item, "Style", "CED.MenuItem.Base")
        badges_item.Header = "Show Circuit Type Badges"
        badges_item.IsCheckable = True
        badges_item.IsChecked = bool(self._compact_show_type_badges)
        badges_item.StaysOpenOnClick = True
        badges_item.Click += self.toggle_compact_badges_clicked
        compact_menu.Items.Add(badges_item)
        compress_item = MenuItem()
        _set_if_resource(self, compress_item, "Style", "CED.MenuItem.Base")
        compress_item.Header = "Compress Item Width"
        compress_item.IsCheckable = True
        compress_item.IsChecked = bool(self._compress_item_width)
        compress_item.StaysOpenOnClick = True
        compress_item.Click += self.toggle_compress_item_width_clicked
        compact_menu.Items.Add(compress_item)
        self._browser_compress_item = compress_item
        menu.Items.Add(compact_menu)

        self._browser_options_menu = menu
        self._sync_browser_options_menu_state()

    def _sync_browser_options_menu_state(self):
        selected_display = "card" if self._is_card_view else "compact"
        for mode, item in (self._browser_theme_items or {}).items():
            if item is not None:
                item.IsChecked = (mode == self._theme_mode)
        for mode, item in (self._browser_accent_items or {}).items():
            if item is not None:
                item.IsChecked = (mode == self._accent_mode)
        for mode, item in (self._browser_display_items or {}).items():
            if item is not None:
                item.IsChecked = (mode == selected_display)
        if self._browser_compress_item is not None:
            self._browser_compress_item.IsChecked = bool(self._compress_item_width)

    def _apply_theme_visual_state(self, refresh_visible_items=True):
        _try_apply_theme(self)
        self._surface_item_style = _try_find_resource(self, "CED.ListViewItem.SurfaceBehavior")
        self._clear_type_tag_brush_cache()
        self._apply_type_tag_brushes()
        self._apply_list_interaction_mode()
        self._update_filter_button_style()
        self._update_search_chrome()
        self._update_toggle_button_visual()
        if bool(refresh_visible_items):
            self._refresh_visible_items()
        self._sync_browser_options_menu_state()

    def _sync_theme_from_config(self, apply_if_changed=True):
        global CURRENT_THEME_MODE
        global CURRENT_ACCENT_MODE
        cfg_theme, cfg_accent = _load_theme_state_from_config(self._theme_mode, self._accent_mode)
        cfg_theme = _normalize_theme_mode(cfg_theme, self._theme_mode)
        cfg_accent = _normalize_accent_mode(cfg_accent, self._accent_mode)
        changed = (cfg_theme != self._theme_mode) or (cfg_accent != self._accent_mode)
        self._theme_mode = cfg_theme
        self._accent_mode = cfg_accent
        CURRENT_THEME_MODE = cfg_theme
        CURRENT_ACCENT_MODE = cfg_accent
        if changed and bool(apply_if_changed):
            self._apply_theme_visual_state(refresh_visible_items=True)
        else:
            self._sync_browser_options_menu_state()
        return changed

    def browser_options_clicked(self, sender, args):
        self._build_browser_options_menu()
        if self._browser_options_menu is not None:
            self._browser_options_menu.PlacementTarget = sender if sender is not None else self._browser_options_button
            self._browser_options_menu.IsOpen = True

    def browser_theme_clicked(self, sender, args):
        global CURRENT_THEME_MODE
        mode = str(getattr(sender, "Tag", "light")).lower()
        if mode not in VALID_THEME_MODES:
            return
        if self._theme_mode == mode:
            self._sync_browser_options_menu_state()
            return
        self._theme_mode = mode
        CURRENT_THEME_MODE = mode
        _save_theme_state_to_config(self._theme_mode, self._accent_mode)
        self._sync_browser_options_menu_state()
        def _apply_theme_change():
            self._apply_theme_visual_state(refresh_visible_items=True)
        _invoke_later(self, _apply_theme_change)

    def browser_accent_clicked(self, sender, args):
        global CURRENT_ACCENT_MODE
        mode = str(getattr(sender, "Tag", "blue")).lower()
        if mode not in ACCENT_BRUSH_KEY_MAP:
            return
        if self._accent_mode == mode:
            self._sync_browser_options_menu_state()
            return
        self._accent_mode = mode
        CURRENT_ACCENT_MODE = mode
        _save_theme_state_to_config(self._theme_mode, self._accent_mode)
        self._sync_browser_options_menu_state()
        def _apply_accent_change():
            self._apply_theme_visual_state(refresh_visible_items=True)
        _invoke_later(self, _apply_accent_change)

    def browser_display_mode_clicked(self, sender, args):
        mode = str(getattr(sender, "Tag", "")).lower()
        if mode == "compact":
            self._set_card_view(False)
            self._sync_browser_options_menu_state()
            return
        if mode == "card":
            self._set_card_view(True)
            self._sync_browser_options_menu_state()

    def _collect_action_targets(self):
        # Actions run only on explicitly checked circuits to avoid
        # selection/checkbox ambiguity for users.
        return [x for x in self._all_items if getattr(x, "is_checked", False)]

    def _collect_selected_row_targets(self):
        try:
            return list(self._list.SelectedItems or [])
        except Exception:
            return []

    def _collect_selected_row_ids(self):
        ids = []
        for item in list(self._collect_selected_row_targets() or []):
            try:
                cid = int(getattr(item, "circuit_id", 0) or 0)
            except Exception:
                cid = 0
            if cid > 0:
                ids.append(cid)
        return ids

    def _reselect_rows_by_ids(self, circuit_ids):
        if self._list is None:
            return
        target_ids = set()
        for raw in list(circuit_ids or []):
            try:
                cid = int(raw or 0)
            except Exception:
                cid = 0
            if cid > 0:
                target_ids.add(cid)
        if not target_ids:
            self.selection_changed(None, None)
            return
        try:
            self._list.SelectedItems.Clear()
        except Exception:
            pass
        matches = []
        for item in list(self._visible_items or []):
            try:
                cid = int(getattr(item, "circuit_id", 0) or 0)
            except Exception:
                cid = 0
            if cid in target_ids:
                matches.append(item)
        for item in list(matches or []):
            try:
                self._list.SelectedItems.Add(item)
            except Exception:
                try:
                    if self._list.SelectedItem is None:
                        self._list.SelectedItem = item
                except Exception:
                    pass
        self.selection_changed(None, None)

    def _refresh_items_by_circuit_ids(self, circuit_ids, reselect_ids=None):
        doc = self._get_active_doc()
        if doc is None:
            return
        target_ids = set()
        for raw in list(circuit_ids or []):
            try:
                cid = int(raw or 0)
            except Exception:
                cid = 0
            if cid > 0:
                target_ids.add(cid)
        if not target_ids:
            self._reselect_rows_by_ids(reselect_ids or [])
            return

        updated = False
        refreshed_items = []
        refreshed_index = {}
        for item in list(self._all_items or []):
            try:
                cid = int(getattr(item, "circuit_id", 0) or 0)
            except Exception:
                cid = 0
            if cid <= 0:
                continue
            if cid not in target_ids:
                refreshed_items.append(item)
                refreshed_index[cid] = item
                continue
            try:
                live = doc.GetElement(_elid_from_value(cid))
            except Exception:
                live = None
            if not isinstance(live, DBE.ElectricalSystem):
                updated = True
                continue
            replacement = CircuitListItem(live, session_sync_state=self._session_state_for_circuit(live))
            replacement.is_checked = bool(getattr(item, "is_checked", False))
            self._apply_type_tag_brush(replacement)
            refreshed_items.append(replacement)
            refreshed_index[cid] = replacement
            updated = True

        if not updated:
            self._reselect_rows_by_ids(reselect_ids or [])
            return

        self._all_items = refreshed_items
        self._item_index = refreshed_index
        self._rebuild_filter_options()
        self._refresh_list()
        self._reselect_rows_by_ids(reselect_ids or [])

    def _sorted_display_panels(self, panels, doc):
        valid = []
        for panel in list(panels or []):
            panel_data = get_panel_dist_system(panel, doc)
            dist_name = (panel_data or {}).get("dist_system_name")
            if not dist_name:
                continue
            if dist_name == "Unnamed Distribution System":
                continue
            valid.append((str(getattr(panel, "Name", "") or ""), str(dist_name), panel))
        valid.sort(key=lambda x: (x[0], x[1]))
        return [panel for _, _, panel in valid]

    def _format_panel_display(self, panel, doc):
        panel_data = get_panel_dist_system(panel, doc)
        dist_name = (panel_data or {}).get("dist_system_name") or "Unknown Dist. System"
        return "{} - {} (ID: {})".format(
            getattr(panel, "Name", "Unnamed Panel") or "Unnamed Panel",
            dist_name,
            _elid_value(getattr(panel, "Id", None)),
        )

    def _prompt_for_move_target_panel(self, panels, doc):
        sorted_panels = self._sorted_display_panels(panels, doc)
        panel_map = {}
        for panel in sorted_panels:
            panel_map[self._format_panel_display(panel, doc)] = panel
        if not panel_map:
            return None
        selected_display = forms.SelectFromList.show(
            sorted(panel_map.keys()),
            title="Select Target Panel",
            prompt="Choose the target panel to move the circuits to:",
            multiselect=False,
        )
        if not selected_display:
            return None
        return panel_map.get(selected_display)

    def _print_move_results(self, output, move_result):
        if isinstance(move_result, dict):
            moved_rows = list(move_result.get("moved") or [])
            failed_rows = list(move_result.get("failed") or [])
            skipped_rows = list(move_result.get("skipped") or [])
            partial = bool(move_result.get("partial", False))
            fallback_used = bool(move_result.get("fallback_used", False))
        else:
            moved_rows = list(move_result or [])
            failed_rows = []
            skipped_rows = []
            partial = False
            fallback_used = False

        if output is not None:
            if partial:
                output.print_md("**Partial move accepted.**")
            elif fallback_used:
                output.print_md("**Circuits transferred successfully (with default SPARE/SPACE replacement workflow).**")
            else:
                output.print_md("**Circuits transferred successfully.**")
            if partial:
                result_rows = []
                for row in list(moved_rows or []):
                    cid = row[0] if len(row) > 0 else "-"
                    prev = row[1] if len(row) > 1 else "-"
                    result_rows.append([cid, prev, "Moved"])
                for row in list(failed_rows or []):
                    cid = row[0] if len(row) > 0 else "-"
                    prev = row[1] if len(row) > 1 else "-"
                    reason = row[2] if len(row) > 2 else "Move failed."
                    result_rows.append([cid, prev, "Failed: {0}".format(reason)])
                for row in list(skipped_rows or []):
                    cid = row[0] if len(row) > 0 else "-"
                    prev = row[1] if len(row) > 1 else "-"
                    reason = row[2] if len(row) > 2 else "Skipped."
                    result_rows.append([cid, prev, "Skipped: {0}".format(reason)])
                output.print_table(result_rows, ["Circuit ID", "Previous Circuit", "Result"])

        if failed_rows:
            return "Moved {} circuits ({} failed)".format(len(moved_rows), len(failed_rows))
        if skipped_rows:
            return "Moved {} circuits ({} skipped)".format(len(moved_rows), len(skipped_rows))
        return "Moved {} circuits".format(len(moved_rows))

    def _on_move_selected_circuits_complete(self, status, payload, error):
        payload = dict(payload or {})
        if status == "error":
            self._set_status("Move failed")
            forms.alert("Failed to move circuits:\n\n{}".format(error), title=TITLE)
            self._safe_load_items()
            return

        move_result = payload.get("move_result")
        moved_rows = []
        failed_rows = []
        if isinstance(move_result, dict):
            moved_rows = list(move_result.get("moved") or [])
            failed_rows = list(move_result.get("failed") or [])
            partial = bool(move_result.get("partial", False))
        else:
            moved_rows = list(move_result or [])
            partial = False

        show_output = bool(partial)
        output = None
        if show_output:
            output = script.get_output()
            try:
                output.show()
            except Exception:
                pass
            buffered_output = payload.get("buffered_output")
            if buffered_output is not None:
                try:
                    buffered_output.flush_to(output)
                except Exception:
                    pass

        summary = self._print_move_results(output, move_result)
        recalc_error = payload.get("recalc_error")
        recalc_result = payload.get("recalc_result")
        moved_ids = list(payload.get("moved_ids") or [])
        if recalc_error is not None:
            self._set_status("{} | Recalc failed".format(summary))
            forms.alert("Circuits moved, but recalculation failed:\n\n{}".format(recalc_error), title=TITLE)
        else:
            recalced_count = 0
            if isinstance(recalc_result, dict):
                try:
                    recalced_count = int(recalc_result.get("updated_circuits", 0) or 0)
                except Exception:
                    recalced_count = 0
            if moved_ids and recalced_count > 0:
                self._set_status("{} | Recalculated {}".format(summary, recalced_count))
            else:
                self._set_status(summary)
        self._safe_load_items()

    def _run_move_selected_circuits(self, targets, checked_only):
        if not self._has_active_doc():
            forms.alert("Open a model document first.", title=TITLE)
            return

        targets = list(targets or [])
        if not targets:
            forms.alert(
                "Check one or more circuits first." if checked_only else "Select one or more rows first.",
                title=TITLE,
            )
            return

        targets, removed_count = self._prune_stale_items(targets)
        if removed_count:
            self._set_status("Removed {} deleted circuits from list.".format(removed_count))

        deduped_circuits = []
        seen_ids = set()
        for item in list(targets or []):
            circuit = getattr(item, "circuit", None)
            if circuit is None:
                continue
            cid = _elid_value(circuit.Id)
            if cid in seen_ids:
                continue
            seen_ids.add(cid)
            deduped_circuits.append(circuit)

        if not deduped_circuits:
            forms.alert("No valid circuits selected.", title=TITLE)
            return

        doc = self._get_active_doc()
        all_panels = list(get_all_panels(doc) or [])
        if not all_panels:
            forms.alert("No panels found in this model.", title=TITLE)
            return

        compatible_panels = []
        for idx, circuit in enumerate(deduped_circuits):
            circuit_panels = list(get_compatible_panels(circuit, all_panels, doc) or [])
            if idx == 0:
                compatible_panels = list(circuit_panels)
                continue
            allowed_ids = set([_elid_value(panel.Id) for panel in list(circuit_panels or [])])
            compatible_panels = [
                panel for panel in list(compatible_panels or [])
                if _elid_value(panel.Id) in allowed_ids
            ]
            if not compatible_panels:
                break

        if not compatible_panels:
            forms.alert("No compatible target panel found for the selected circuits.", title=TITLE)
            return

        target_panel = self._prompt_for_move_target_panel(compatible_panels, doc)
        if target_panel is None:
            self._set_status("Move cancelled")
            return

        if self._move_gateway.is_busy() or self._operation_gateway.is_busy():
            forms.alert("An operation is already running. Please wait.", title=TITLE)
            return

        self._set_status("Moving {} circuits...".format(len(deduped_circuits)))
        raised = self._move_gateway.raise_move(
            circuit_ids=[_elid_value(circuit.Id) for circuit in deduped_circuits],
            target_panel_id=_elid_value(target_panel.Id),
            callback=self._on_move_selected_circuits_complete,
        )
        if not raised:
            self._set_status("Unable to queue move operation")

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

    def _format_writeback_lock_reason(self, row):
        return format_writeback_lock_reason(row)

    def _build_writeback_lock_map(self, circuits, doc=None, settings=None):
        if doc is None:
            doc = self._get_active_doc()
        if doc is None or not getattr(doc, "IsWorkshared", False):
            return {}
        circuit_list = [c for c in list(circuits or []) if c is not None]
        if not circuit_list:
            return {}
        if settings is None:
            try:
                settings = settings_manager.load_circuit_settings(doc)
            except Exception:
                settings = None
        if settings is None:
            return {}
        try:
            _, _, locked_rows = self._lock_repository.partition_locked_elements(
                doc,
                circuit_list,
                settings,
                collect_all_device_owners=False,
            )
        except Exception:
            return {}
        lock_map = {}
        for row in list(locked_rows or []):
            try:
                cid = int((row or {}).get("circuit_id") or 0)
            except Exception:
                cid = 0
            if cid <= 0:
                continue
            lock_map[cid] = row
        return lock_map

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

<<<<<<< HEAD
=======
    def _resolve_action_branch_type(self, circuit, fallback_type="", settings=None):
        try:
            use_settings = settings
            if use_settings is None:
                use_settings = settings_manager.load_circuit_settings(circuit.Document)
            branch = CircuitBranch(circuit, settings=use_settings)
            return str(getattr(branch, "branch_type", "") or "").strip().upper()
        except Exception:
            return str(fallback_type or "").strip().upper()

    def _validate_include_param(self, circuit, param_name):
        param = None
        try:
            param = circuit.LookupParameter(param_name)
        except Exception:
            param = None
        if not param:
            return False, "Missing parameter: {}".format(param_name)
        try:
            if param.StorageType != DB.StorageType.Integer:
                return False, "Invalid parameter type: {}".format(param_name)
        except Exception:
            return False, "Invalid parameter type: {}".format(param_name)
        try:
            if bool(getattr(param, "IsReadOnly", False)):
                return False, "Parameter is read-only: {}".format(param_name)
        except Exception:
            pass
        return True, ""

>>>>>>> main
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

    def _raise_action_operation(self, operation_key, circuit_ids, options, callback=None):
        if self._operation_gateway.is_busy():
            forms.alert("An operation is already running. Please wait.", title=TITLE)
            return False
        self._set_status("Applying action...")
        self._operation_gateway.raise_operation(
            operation_key=operation_key,
            circuit_ids=list(circuit_ids or []),
            source="pane",
            options=dict(options or {}),
            callback=callback or self._on_operation_complete,
        )
        return True

    def _build_neutral_rows(self, targets):
        rows = []
        doc = self._get_active_doc()
        settings = None
        if doc is not None:
            try:
                settings = settings_manager.load_circuit_settings(doc)
            except Exception:
                settings = None
        lock_map = self._build_writeback_lock_map([getattr(x, "circuit", None) for x in list(targets or [])], doc=doc, settings=settings)
        for item in targets:
            circuit = item.circuit
<<<<<<< HEAD
            btype = (item.branch_type or "").upper()
=======
            btype = self._resolve_action_branch_type(circuit, fallback_type=getattr(item, "branch_type", ""), settings=settings)
>>>>>>> main
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

<<<<<<< HEAD
=======
            if is_enabled:
                ok_param, param_reason = self._validate_include_param(circuit, "CKT_Include Neutral_CED")
                if not ok_param:
                    reason = param_reason
                    is_enabled = False

>>>>>>> main
            lock_row = lock_map.get(_elid_value(circuit.Id))
            if lock_row is not None:
                reason = self._format_writeback_lock_reason(lock_row)
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
<<<<<<< HEAD
=======
            row.target_include = None
>>>>>>> main

            try:
                branch = self._simulate_branch(row.circuit, include_neutral=include_neutral)
                new_qty = int(branch.neutral_wire_quantity or 0)
                new_size = branch.neutral_wire_size or ""
                new_wire = branch.get_wire_size_callout() or ""
            except Exception:
                row.recompute_state()
                continue

<<<<<<< HEAD
            changed = (new_qty != int(row.current_qty or 0)) or (str(new_size or "") != str(row.current_size or "")) or (
                str(new_wire or "") != str(row.current_wire or "")
            )
            if changed:
=======
            qty_changed = new_qty != int(row.current_qty or 0)
            if qty_changed:
>>>>>>> main
                row.new_qty = new_qty
                row.new_size = new_size
                row.new_wire = new_wire or "-"
                row.new_wire_font_style = "Normal"
<<<<<<< HEAD
                row.new_qty_changed = new_qty != int(row.current_qty or 0)
                row.new_size_changed = str(new_size or "") != str(row.current_size or "")
                row.new_wire_changed = str(new_wire or "") != str(row.current_wire or "")
            row.recompute_state()

    def _apply_neutral_rows(self, rows, mode):
        changed_ids = [x.circuit_id for x in rows if x.is_enabled and x.is_checked and x.is_changed]
=======
                row.new_qty_changed = True
                row.new_size_changed = str(new_size or "") != str(row.current_size or "")
                row.new_wire_changed = str(new_wire or "") != str(row.current_wire or "")
                row.target_include = 1 if int(new_qty or 0) > 0 else 0
            row.recompute_state()

    def _apply_neutral_rows(self, rows, mode):
        updates = []
        for row in list(rows or []):
            if not (row.is_enabled and row.is_changed):
                continue
            include_value = getattr(row, "target_include", None)
            if include_value not in (0, 1):
                continue
            updates.append(
                {
                    "circuit_id": int(row.circuit_id or 0),
                    "include": int(include_value),
                }
            )
        changed_ids = [x.get("circuit_id") for x in updates if int(x.get("circuit_id") or 0) > 0]
>>>>>>> main
        if not changed_ids:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False
        return self._raise_action_operation(
            "set_neutral_and_recalculate",
            changed_ids,
<<<<<<< HEAD
            {"mode": mode, "show_output": False},
=======
            {"mode": mode, "updates": updates, "show_output": False},
>>>>>>> main
        )

    def _build_ig_rows(self, targets):
        rows = []
        doc = self._get_active_doc()
        settings = None
        if doc is not None:
            try:
                settings = settings_manager.load_circuit_settings(doc)
            except Exception:
                settings = None
        lock_map = self._build_writeback_lock_map([getattr(x, "circuit", None) for x in list(targets or [])], doc=doc, settings=settings)
        for item in targets:
            circuit = item.circuit
<<<<<<< HEAD
            btype = (item.branch_type or "").upper()
=======
            btype = self._resolve_action_branch_type(circuit, fallback_type=getattr(item, "branch_type", ""), settings=settings)
>>>>>>> main
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

<<<<<<< HEAD
=======
            if is_enabled:
                ok_param, param_reason = self._validate_include_param(circuit, "CKT_Include Isolated Ground_CED")
                if not ok_param:
                    reason = param_reason
                    is_enabled = False

>>>>>>> main
            lock_row = lock_map.get(_elid_value(circuit.Id))
            if lock_row is not None:
                reason = self._format_writeback_lock_reason(lock_row)
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
<<<<<<< HEAD
=======
            row.target_include = None
>>>>>>> main

            try:
                branch = self._simulate_branch(row.circuit, include_ig=include_ig)
                new_qty = int(branch.isolated_ground_wire_quantity or 0)
                new_size = branch.isolated_ground_wire_size or ""
                new_wire = branch.get_wire_size_callout() or ""
            except Exception:
                row.recompute_state()
                continue

<<<<<<< HEAD
            changed = (new_qty != int(row.current_qty or 0)) or (str(new_size or "") != str(row.current_size or "")) or (
                str(new_wire or "") != str(row.current_wire or "")
            )
            if changed:
=======
            qty_changed = new_qty != int(row.current_qty or 0)
            if qty_changed:
>>>>>>> main
                row.new_qty = new_qty
                row.new_size = new_size
                row.new_wire = new_wire or "-"
                row.new_wire_font_style = "Normal"
<<<<<<< HEAD
                row.new_qty_changed = new_qty != int(row.current_qty or 0)
                row.new_size_changed = str(new_size or "") != str(row.current_size or "")
                row.new_wire_changed = str(new_wire or "") != str(row.current_wire or "")
            row.recompute_state()

    def _apply_ig_rows(self, rows, mode):
        changed_ids = [x.circuit_id for x in rows if x.is_enabled and x.is_checked and x.is_changed]
=======
                row.new_qty_changed = True
                row.new_size_changed = str(new_size or "") != str(row.current_size or "")
                row.new_wire_changed = str(new_wire or "") != str(row.current_wire or "")
                row.target_include = 1 if int(new_qty or 0) > 0 else 0
            row.recompute_state()

    def _apply_ig_rows(self, rows, mode):
        updates = []
        for row in list(rows or []):
            if not (row.is_enabled and row.is_changed):
                continue
            include_value = getattr(row, "target_include", None)
            if include_value not in (0, 1):
                continue
            updates.append(
                {
                    "circuit_id": int(row.circuit_id or 0),
                    "include": int(include_value),
                }
            )
        changed_ids = [x.get("circuit_id") for x in updates if int(x.get("circuit_id") or 0) > 0]
>>>>>>> main
        if not changed_ids:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False
        return self._raise_action_operation(
            "set_ig_and_recalculate",
            changed_ids,
<<<<<<< HEAD
            {"mode": mode, "show_output": False},
=======
            {"mode": mode, "updates": updates, "show_output": False},
>>>>>>> main
        )

    def _build_breaker_rows(self, targets):
        rows = []
        doc = self._get_active_doc()
        is_workshared = bool(getattr(doc, "IsWorkshared", False)) if doc is not None else False
        lock_candidate_circuits = []
        for item in list(targets or []):
            btype = (getattr(item, "branch_type", "") or "").upper()
            if btype in BLOCKED_BRANCH_TYPES:
                continue
            if btype not in IG_BREAKER_ALLOWED_TYPES:
                continue
            circuit = getattr(item, "circuit", None)
            if circuit is not None:
                lock_candidate_circuits.append(circuit)
        settings = None
        if doc is not None:
            try:
                settings = settings_manager.load_circuit_settings(doc)
            except Exception:
                settings = None
        lock_map = self._build_writeback_lock_map(lock_candidate_circuits, doc=doc, settings=settings)
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

            lock_row = lock_map.get(_elid_value(circuit.Id)) if is_workshared else None
            if lock_row is not None:
                reason = self._format_writeback_lock_reason(lock_row)
                is_enabled = False

            load_current = _lookup_param_value(circuit, "Circuit Load Current_CED")
            cur_rating = None
            cur_frame = None
            if btype != "SPACE":
                try:
                    cur_rating = circuit.Rating
                except Exception:
                    cur_rating = None
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

    def _preview_breaker_rows(
        self,
        rows,
        set_breaker,
        set_frame,
        allow_15a=False,
        upsize_only=False,
        max_load_percent=80,
    ):
        min_ocp = 15 if allow_15a else 20
        try:
            target_load_pct = float(max_load_percent)
        except Exception:
            target_load_pct = 80.0
        if target_load_pct < 50.0:
            target_load_pct = 50.0
        elif target_load_pct > 100.0:
            target_load_pct = 100.0
        breaker_factor = 100.0 / target_load_pct
        for row in rows:
            load_value = _safe_float(row._load_current_value)
            auto_rating = self._next_ocp_size(
                (load_value * breaker_factor) if load_value is not None else None,
                min_ocp=min_ocp,
            )
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

    def _apply_breaker_rows(
        self,
        rows,
        set_breaker,
        set_frame,
        allow_15a=False,
        upsize_only=False,
        max_load_percent=80,
    ):
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
            {
                "updates": updates,
                "show_output": False,
                "allow_15a": bool(allow_15a),
            },
        )

    def _build_mark_existing_rows(self, targets):
        rows = []
        doc = self._get_active_doc()
        is_workshared = bool(getattr(doc, "IsWorkshared", False)) if doc is not None else False
        settings = None
        if doc is not None:
            try:
                settings = settings_manager.load_circuit_settings(doc)
            except Exception:
                settings = None
        lock_map = self._build_writeback_lock_map([getattr(x, "circuit", None) for x in list(targets or [])], doc=doc, settings=settings)
        for item in targets:
            circuit = item.circuit
            reason = ""
            is_enabled = True
            lock_row = lock_map.get(_elid_value(circuit.Id)) if is_workshared else None
            if lock_row is not None:
                reason = self._format_writeback_lock_reason(lock_row)
                is_enabled = False

            current_notes = _lookup_schedule_notes_text(circuit)
            current_wire = self._conduit_wire_size_string(circuit)
            current_sets = self._param_int(circuit, "CKT_Number of Sets_CED", 1)
            current_conduit_size = self._conduit_size_string(circuit)
            current_wire_size = self._wire_size_string(circuit)
            row = MarkExistingActionRow(
                item,
                is_enabled,
                reason,
                current_notes,
                current_wire,
                current_sets,
                current_conduit_size,
                current_wire_size,
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
            row.action_mode = mode_text if mode_text in ("existing", "new") else "existing"
            row.action_set_notes = bool(set_notes)
            row.preview_mode = mode_text
            row.preview_clear_wire = bool(clear_wire) and not is_new_mode
            row.preview_clear_conduit = bool(clear_conduit) and not is_new_mode
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
                if clear_wire or clear_conduit:
                    if clear_wire and clear_conduit:
                        row.new_wire = "-"
                    elif clear_wire:
                        row.new_wire = row._existing_conduit_only_display()
                    elif clear_conduit:
                        row.new_wire = row._existing_wire_only_display()
                else:
                    row.new_wire = row.current_wire
            row.recompute_state()

    def _apply_mark_existing_rows(self, rows, mode, set_notes, clear_wire, clear_conduit):
        updates = []
        changed_ids = []
        for row in list(rows or []):
            if not (row.is_enabled and row.is_changed):
                continue
            mode_text = str(getattr(row, "action_mode", "existing") or "existing").strip().lower()
            if mode_text not in ("existing", "new"):
                mode_text = "existing"
            update = {
                "circuit_id": int(row.circuit_id or 0),
                "mode": mode_text,
                "set_notes": bool(getattr(row, "action_set_notes", True)),
                "clear_wire": bool(getattr(row, "preview_clear_wire", False)),
                "clear_conduit": bool(getattr(row, "preview_clear_conduit", False)),
            }
            if mode_text == "new":
                update["clear_wire"] = False
                update["clear_conduit"] = False
            cid = int(update.get("circuit_id") or 0)
            if cid <= 0:
                continue
            updates.append(update)
            changed_ids.append(cid)
        if not changed_ids:
            forms.alert("No circuits are marked for modification.", title=TITLE)
            return False
        return self._raise_action_operation(
            "mark_existing_and_recalculate",
            changed_ids,
            {
                "updates": updates,
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

    def action_edit_properties_clicked(self, sender, args):
        self._run_edit_circuit_properties(self._collect_action_targets(), checked_only=True)

    def action_edit_properties_rows_clicked(self, sender, args):
        self._run_edit_circuit_properties(self._collect_selected_row_targets(), checked_only=False)

    def _show_edit_properties_lock_alert(self, locked_rows):
        rows = list(locked_rows or [])
        if not rows:
            return
        owners = []
        for row in rows:
            owner = str((row or {}).get("circuit_owner") or "").strip()
            if owner and owner not in owners:
                owners.append(owner)
        message_lines = [
            "{} circuit(s) were skipped because they are owned by another user.".format(len(rows)),
        ]
        if owners:
            message_lines.append("")
            message_lines.append("Owner(s): {}".format(", ".join(owners)))
        forms.alert("\n".join(message_lines), title="Edit Circuit Properties")

    def _on_edit_circuit_properties_complete(self, status, request, result, error):
        reselect_ids = list(self._edit_properties_reselect_ids or [])
        self._edit_properties_reselect_ids = []
        request_ids = []
        for raw in list(getattr(request, "circuit_ids", []) or []):
            try:
                cid = int(raw or 0)
            except Exception:
                cid = 0
            if cid > 0:
                request_ids.append(cid)

        if status == "error":
            self._set_status("Edit Circuit Properties failed.")
            try:
                self._logger.warning("Edit Circuit Properties apply failed: %s", error)
            except Exception:
                pass
            self._refresh_items_by_circuit_ids(request_ids, reselect_ids=reselect_ids)
            return

        payload = dict(result or {})
        self._update_session_sync_lock_map(payload)
        locked_rows = list(payload.get("locked_rows") or [])
        self._refresh_items_by_circuit_ids(request_ids, reselect_ids=reselect_ids)

        if locked_rows:
            self._show_edit_properties_lock_alert(locked_rows)

        if payload.get("status") == "ok":
            try:
                edited_count = int(payload.get("edited_circuits", 0) or 0)
            except Exception:
                edited_count = 0
            try:
                recalculated = int(payload.get("updated_circuits", 0) or 0)
            except Exception:
                recalculated = 0
            self._set_status("Edited {} circuit(s) | Recalculated {}".format(edited_count, recalculated))
            return

        reason = str(payload.get("reason", "unknown") or "unknown")
        if reason == "no_updates":
            self._set_status("No staged circuit edits to apply.")
            return
        if reason == "no_changes":
            self._set_status("No circuit property changes were applied.")
            return
        self._set_status("Edit Circuit Properties cancelled ({})".format(reason))

    def _run_edit_circuit_properties(self, targets, checked_only):
        targets = list(targets or [])
        if not targets:
            forms.alert(
                "Check one or more circuits first." if checked_only else "Select one or more rows first.",
                title=TITLE,
            )
            return
        if len(targets) > 300:
            choice = forms.alert(
                "{} circuits will be loaded for this action.\n\nContinue?".format(len(targets)),
                title="Large Action Selection",
                options=["Continue", "Cancel"],
            )
            if choice != "Continue":
                return

        doc = self._get_active_doc()
        if doc is None:
            forms.alert("Open a model document first.", title=TITLE)
            return
        try:
            settings = settings_manager.load_circuit_settings(doc)
        except Exception as ex:
            forms.alert("Unable to load circuit settings:\n\n{}".format(ex), title=TITLE)
            return

        xaml_path = os.path.abspath(os.path.join(_THIS_DIR, "CircuitEditPropertiesWindow.xaml"))
        try:
            window = CircuitPropertiesEditorWindow(
                xaml_path=xaml_path,
                targets=targets,
                settings=settings,
                theme_mode=self._theme_mode,
                accent_mode=self._accent_mode,
                resources_root=UI_RESOURCES_ROOT,
            )
            window.ShowDialog()
        except Exception as ex:
            forms.alert("Failed to open Edit Circuit Properties window:\n\n{}".format(ex), title=TITLE)
            return

        if not bool(getattr(window, "apply_requested", False)):
            self._set_status("Edit Circuit Properties cancelled")
            return

        payload = dict(getattr(window, "apply_payload", {}) or {})
        updates = list(payload.get("updates") or [])
        if not updates:
            self._set_status("No circuit property edits were staged.")
            return

        circuit_ids = []
        for row in list(updates or []):
            try:
                cid = int((row or {}).get("circuit_id") or 0)
            except Exception:
                cid = 0
            if cid > 0:
                circuit_ids.append(cid)
        if not circuit_ids:
            self._set_status("No valid staged circuit edits found.")
            return

        self._edit_properties_reselect_ids = self._collect_selected_row_ids()
        self._set_status("Applying circuit property edits...")
        raised = self._raise_action_operation(
            "edit_circuit_properties_and_recalculate",
            circuit_ids,
            {
                "updates": updates,
                "show_output": False,
            },
            callback=self._on_edit_circuit_properties_complete,
        )
        if not raised:
            self._edit_properties_reselect_ids = []

    def action_move_checked_clicked(self, sender, args):
        self._run_move_selected_circuits(self._collect_action_targets(), checked_only=True)

    def action_move_selected_rows_clicked(self, sender, args):
        self._run_move_selected_circuits(self._collect_selected_row_targets(), checked_only=False)

    def _reset_filters(self):
        self._active_type_filters = set(self._default_active_type_filters())
        self._warnings_only = False
        self._warnings_active_only = False
        self._overrides_only = False
        self._syncblocked_only = False
        self._checked_only = False
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_reset_clicked(self, sender, args):
        self._reset_filters()
        try:
            if self._filter_menu is not None:
                self._filter_menu.IsOpen = False
        except Exception:
            pass

    def filter_select_all_clicked(self, sender, args):
        self._active_type_filters = set(self._type_options)
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_clear_all_clicked(self, sender, args):
        self._active_type_filters = set()
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_type_toggled(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
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
        self._sync_filter_menu_state()

    def filter_warnings_toggled(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
        self._set_exclusive_filter("warnings", bool(getattr(sender, "IsChecked", False)))
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_warnings_parent_clicked(self, sender, args):
        # Parent is submenu-only; preserve current checked state and just open submenu.
        if self._suppress_filter_toggle_events:
            return
        try:
            sender.IsChecked = bool(self._warnings_only)
        except Exception:
            pass
        try:
            sender.IsSubmenuOpen = True
        except Exception:
            pass
        self._sync_filter_menu_state()

    def filter_warnings_parent_mouse_enter(self, sender, args):
        # Match native menu behavior + explicit immediate hover-open for this submenu.
        if self._suppress_filter_toggle_events:
            return
        try:
            sender.IsSubmenuOpen = True
        except Exception:
            pass

    def filter_warnings_show_all_clicked(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
        self._warnings_active_only = False
        if not self._warnings_only:
            self._set_exclusive_filter("warnings", True)
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_warnings_active_only_clicked(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
        self._warnings_active_only = True
        if not self._warnings_only:
            self._set_exclusive_filter("warnings", True)
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_sync_blocked_toggled(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
        self._set_exclusive_filter("syncblocked", bool(getattr(sender, "IsChecked", False)))
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_overrides_toggled(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
        self._set_exclusive_filter("overrides", bool(getattr(sender, "IsChecked", False)))
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def filter_checked_toggled(self, sender, args):
        if self._suppress_filter_toggle_events:
            return
        self._set_exclusive_filter("checked", bool(getattr(sender, "IsChecked", False)))
        self._update_filter_button_style()
        self._refresh_list()
        self._sync_filter_menu_state()

    def _set_card_view(self, is_card_view):
        new_card_view = bool(is_card_view)
        self._is_card_view = bool(new_card_view)
        if self._is_card_view:
            self._list.ItemTemplate = self._card_template
        else:
            self._list.ItemTemplate = self._compact_template
        self._update_toggle_button_visual()
        self._sync_browser_options_menu_state()
        self._refresh_list()

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

<<<<<<< HEAD
=======
    def panel_preview_key_down(self, sender, args):
        try:
            key = getattr(args, "Key", None)
        except Exception:
            key = None
        if key != Key.F:
            return
        try:
            modifiers = Keyboard.Modifiers
        except Exception:
            modifiers = getattr(ModifierKeys, "None")
        if (modifiers & ModifierKeys.Control) != ModifierKeys.Control:
            return
        if self._search is None:
            return
        try:
            self._search.Focus()
            self._search.SelectAll()
            args.Handled = True
        except Exception:
            pass

>>>>>>> main
    def list_preview_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        if _find_visual_ancestor(source, ListViewItem) is None:
            self._clear_list_selection()

    def list_mouse_double_click(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        if _find_visual_ancestor(source, Button) is not None:
            return
        row = _find_visual_ancestor(source, ListViewItem)
        if row is None:
            return
        try:
            item = row.DataContext
        except Exception:
            item = None
        if item is None:
            return
        item.is_checked = not bool(getattr(item, "is_checked", False))
        # Row models do not always raise property change notifications for programmatic
        # toggles, so refresh once on explicit double-click interaction.
        try:
            self._list.Items.Refresh()
        except Exception:
            pass
        self.selection_changed(None, None)
        try:
            args.Handled = True
        except Exception:
            pass

    def _ensure_context_row_selection(self, source):
        row = _find_visual_ancestor(source, ListViewItem)
        if row is None:
            return
        try:
            clicked_item = row.DataContext
        except Exception:
            clicked_item = None
        if clicked_item is None:
            return
        try:
            selected_items = list(self._list.SelectedItems)
        except Exception:
            selected_items = []
        if clicked_item in selected_items:
            return
        try:
            self._list.SelectedItems.Clear()
        except Exception:
            pass
        try:
            self._list.SelectedItem = clicked_item
        except Exception:
            return
        self.selection_changed(None, None)

    def list_preview_mouse_right_button_up(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        self._ensure_context_row_selection(source)
        menu = ContextMenu()
        _set_if_resource(self, menu, "Style", "CED.ContextMenu.Base")
        if not self._is_card_view:
            toggle = MenuItem()
            _set_if_resource(self, toggle, "Style", "CED.MenuItem.Base")
            toggle.Header = "Show Circuit Type Badges"
            toggle.IsCheckable = True
            toggle.IsChecked = bool(self._compact_show_type_badges)
            toggle.Click += self.toggle_compact_badges_clicked
            menu.Items.Add(toggle)
        compress_item = MenuItem()
        _set_if_resource(self, compress_item, "Style", "CED.MenuItem.Base")
        compress_item.Header = "Compress Item Width"
        compress_item.IsCheckable = True
        compress_item.IsChecked = bool(self._compress_item_width)
        compress_item.Click += self.toggle_compress_item_width_clicked
        menu.Items.Add(compress_item)

        selected_rows = self._collect_selected_row_targets()
        if menu.Items.Count > 0:
            sep = Separator()
            _set_if_resource(self, sep, "Style", "CED.Separator.Menu")
            menu.Items.Add(sep)

        select_menu = MenuItem()
        _set_if_resource(self, select_menu, "Style", "CED.MenuItem.Base")
        select_menu.Header = "Select in Model"
        select_menu.IsEnabled = bool(selected_rows)
        panel_item = MenuItem()
        _set_if_resource(self, panel_item, "Style", "CED.MenuItem.Base")
        panel_item.Header = "Panel"
        panel_item.IsEnabled = bool(selected_rows)
        panel_item.Click += self.select_rows_equipment_clicked
        circuit_item = MenuItem()
        _set_if_resource(self, circuit_item, "Style", "CED.MenuItem.Base")
        circuit_item.Header = "Circuit"
        circuit_item.IsEnabled = bool(selected_rows)
        circuit_item.Click += self.select_rows_circuits_clicked
        device_item = MenuItem()
        _set_if_resource(self, device_item, "Style", "CED.MenuItem.Base")
        device_item.Header = "Device"
        device_item.IsEnabled = bool(selected_rows)
        device_item.Click += self.select_rows_downstream_clicked
        select_menu.Items.Add(panel_item)
        select_menu.Items.Add(circuit_item)
        select_menu.Items.Add(device_item)
        menu.Items.Add(select_menu)

<<<<<<< HEAD
=======
        show_devices_item = MenuItem()
        _set_if_resource(self, show_devices_item, "Style", "CED.MenuItem.Base")
        show_devices_item.Header = "Show Devices in Model"
        show_devices_item.IsEnabled = bool(selected_rows)
        show_devices_item.Click += self.show_rows_devices_in_model_clicked
        menu.Items.Add(show_devices_item)

        show_panel_item = MenuItem()
        _set_if_resource(self, show_panel_item, "Style", "CED.MenuItem.Base")
        show_panel_item.Header = "Show Panel in Model"
        show_panel_item.IsEnabled = bool(selected_rows)
        show_panel_item.Click += self.show_rows_panel_in_model_clicked
        menu.Items.Add(show_panel_item)

>>>>>>> main
        sep_select_edit = Separator()
        _set_if_resource(self, sep_select_edit, "Style", "CED.Separator.Menu")
        menu.Items.Add(sep_select_edit)

        edit_item = MenuItem()
        _set_if_resource(self, edit_item, "Style", "CED.MenuItem.Base")
        edit_item.Header = "Edit Circuit Properties"
        edit_item.IsEnabled = bool(selected_rows)
        edit_item.Click += self.action_edit_properties_rows_clicked
        menu.Items.Add(edit_item)

        move_item = MenuItem()
        _set_if_resource(self, move_item, "Style", "CED.MenuItem.Base")
        move_item.Header = "Move Selected Circuits"
        move_item.IsEnabled = bool(selected_rows)
        move_item.Click += self.action_move_selected_rows_clicked
        menu.Items.Add(move_item)
        menu.PlacementTarget = self._list
        menu.IsOpen = True
        args.Handled = True

    def toggle_compact_badges_clicked(self, sender, args):
        self._compact_show_type_badges = bool(getattr(sender, "IsChecked", True))
        self._refresh_list()

    def toggle_compress_item_width_clicked(self, sender, args):
        self._compress_item_width = bool(getattr(sender, "IsChecked", False))
        self._sync_browser_options_menu_state()
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

    def item_checkbox_clicked(self, sender, args):
        self._apply_check_state(sender, bool(getattr(sender, "IsChecked", False)))
        self.selection_changed(None, None)

    def select_equipment_clicked(self, sender, args):
        targets = []
        for item in self._target_items():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "panel"))
            except Exception:
                continue
        self._set_revit_selection(targets)

    def select_circuits_clicked(self, sender, args):
        targets = []
        for item in self._target_items():
            targets.extend(collect_circuit_targets(item.circuit, "circuit"))
        self._set_revit_selection(targets)

    def select_downstream_clicked(self, sender, args):
        targets = []
        for item in self._target_items():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "device"))
            except Exception:
                continue
        self._set_revit_selection(targets)

    def select_rows_equipment_clicked(self, sender, args):
        targets = []
        for item in self._collect_selected_row_targets():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "panel"))
            except Exception:
                continue
        self._set_revit_selection(targets)

    def select_rows_circuits_clicked(self, sender, args):
        targets = []
        for item in self._collect_selected_row_targets():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "circuit"))
            except Exception:
                continue
        self._set_revit_selection(targets)

    def select_rows_downstream_clicked(self, sender, args):
        targets = []
        for item in self._collect_selected_row_targets():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "device"))
            except Exception:
                continue
        self._set_revit_selection(targets)

<<<<<<< HEAD
=======
    def show_rows_devices_in_model_clicked(self, sender, args):
        targets = []
        for item in self._collect_selected_row_targets():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "device"))
            except Exception:
                continue
        self._show_and_select_revit_targets(targets)

    def show_rows_panel_in_model_clicked(self, sender, args):
        targets = []
        for item in self._collect_selected_row_targets():
            try:
                targets.extend(collect_circuit_targets(item.circuit, "panel"))
            except Exception:
                continue
        self._show_and_select_revit_targets(targets)

>>>>>>> main
    def clear_revit_selection_clicked(self, sender, args):
        clear_revit_selection(uidoc=revit.uidoc)

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
        circuit_ids = [_elid_value(x.circuit.Id) for x in items if getattr(x, "circuit", None)]
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
        circuit_ids = [_elid_value(x.circuit.Id) for x in items if getattr(x, "circuit", None)]
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
        self._sync_theme_from_config(apply_if_changed=True)
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
            circuit_ids=[_elid_value(item.circuit.Id)],
            source="pane",
            options={"hidden_definition_ids": list(hidden_ids)},
            callback=self._on_alert_visibility_saved,
        )


def ensure_panel_visible():
    try:
<<<<<<< HEAD
        if not forms.is_registered_dockable_panel(CircuitBrowserPanel):
            forms.register_dockable_panel(CircuitBrowserPanel, default_visible=False)
    except Exception:
        pass

    try:
=======
>>>>>>> main
        forms.open_dockable_panel(CircuitBrowserPanel.panel_id)
    except Exception:
        try:
            forms.open_dockable_panel(CircuitBrowserPanel)
        except Exception:
            pass

    panel = CircuitBrowserPanel.get_instance()
    if panel is not None:
        panel.refresh_on_open()

