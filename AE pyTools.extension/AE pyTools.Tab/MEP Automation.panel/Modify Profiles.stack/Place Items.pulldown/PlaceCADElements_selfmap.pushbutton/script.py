# -*- coding: utf-8 -*-
"""Place cad stuff.

Supports two sources:
1. CSV CAD export
2. Lighting fixtures from a selected Revit link model
"""

import codecs
import csv
import math
import os
import re
from collections import defaultdict
from difflib import SequenceMatcher

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    Level,
    Line,
    LocationPoint,
    RevitLinkInstance,
    StorageType,
    Structure,
    Transaction,
    TransactionGroup,
    XYZ,
)
from Autodesk.Revit.Exceptions import InvalidOperationException
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System import TimeSpan
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, Visibility
from System.Windows.Controls import ComboBox, ComboBoxItem, TextBox
from System.Windows.Media import VisualTreeHelper
from System.Windows.Threading import DispatcherTimer
from pyrevit import forms, revit, script

TITLE = "Let There Be Light"
CSV_SOURCE_MODE = "csv"
LINK_SOURCE_MODE = "link"
MAPPING_CONFIG_SECTION = "let_there_be_light_mapping"
DUPLICATE_TOLERANCE_FT = 0.02
CSV_MARKER_START = "LUMINAIRE-"
CSV_MARKER_END = "_Symbol"

CATEGORY_OPTIONS = [
    ("Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
    ("Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
    ("Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
    ("Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),
    ("Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
    ("Sprinklers", BuiltInCategory.OST_Sprinklers),
    ("Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices),
    ("Communication Devices", BuiltInCategory.OST_CommunicationDevices),
    ("Data Devices", BuiltInCategory.OST_DataDevices),
    ("Security Devices", BuiltInCategory.OST_SecurityDevices),
    ("Nurse Call Devices", BuiltInCategory.OST_NurseCallDevices),
]
DEFAULT_CATEGORY = BuiltInCategory.OST_LightingFixtures

LOGGER = script.get_logger()
OUTPUT = script.get_output()
THIS_DIR = os.path.abspath(os.path.dirname(__file__))


# -----------------------------------------------------------------------------
# UIClasses bootstrap
try:
    from UIClasses import pathing as ui_pathing
except Exception:
    import sys

    fallback_lib = os.path.abspath(
        os.path.join(THIS_DIR, "..", "..", "..", "..", "..", "CEDLib.lib")
    )
    if os.path.isdir(fallback_lib) and fallback_lib not in sys.path:
        sys.path.append(fallback_lib)
    from UIClasses import pathing as ui_pathing

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate CEDLib.lib for {}.".format(TITLE), title=TITLE)
    raise SystemExit

from UIClasses import load_theme_state_from_config
from UIClasses.ui_bases import CEDWindowBase, TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK


# -----------------------------------------------------------------------------
# Utility helpers
def _safe_text(value):
    try:
        if value is None:
            return ""
        return str(value).strip()
    except Exception:
        return ""


def _idval(value):
    try:
        return int(value.IntegerValue)
    except Exception:
        pass
    try:
        return int(value.Id.IntegerValue)
    except Exception:
        return -1


def _compact_norm(value):
    return re.sub(r"[^a-z0-9]+", "", _safe_text(value).lower())


def _spaced_norm(value):
    text = _safe_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text


def _tokens(value):
    text = _spaced_norm(value)
    if not text:
        return set()
    return set([token for token in text.split(" ") if token])


def _format_elevation(feet_value):
    try:
        return "{0:.3f} ft".format(float(feet_value))
    except Exception:
        return "0.000 ft"


def _extract_fixture_code(csv_name):
    text = _safe_text(csv_name)
    if not text:
        return ""
    start_index = text.find(CSV_MARKER_START)
    end_index = text.find(CSV_MARKER_END, start_index if start_index >= 0 else 0)
    if start_index == -1 or end_index == -1:
        return ""
    return _safe_text(text[start_index + len(CSV_MARKER_START) : end_index])


def _as_location_point(element):
    try:
        location = getattr(element, "Location", None)
    except Exception:
        location = None
    if isinstance(location, LocationPoint):
        return location
    return None


def _visual_descendants(root, target_type):
    if root is None:
        return []
    found = []
    queue = [root]
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
                found.append(child)
            queue.append(child)
    return found


def _symbol_names(symbol):
    family_name = "Unknown Family"
    type_name = "Unknown Type"
    try:
        family_name = _safe_text(symbol.Family.Name) or family_name
    except Exception:
        pass
    try:
        type_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if type_param:
            type_name = _safe_text(type_param.AsString()) or type_name
    except Exception:
        pass
    return family_name, type_name


def _symbol_type_mark(symbol):
    """Return identity type mark (preferred) fallback to type mark/name-like fields."""
    candidate_names = [
        "Identity Type Mark",
        "Type Mark",
        "Fixture Type_CEDT",
        "CED-E-FIXTURE TYPE",
    ]
    for name in candidate_names:
        try:
            param = symbol.LookupParameter(name)
        except Exception:
            param = None
        if param is None:
            continue
        try:
            value = _safe_text(param.AsString())
        except Exception:
            value = ""
        if value:
            return value
    return ""


def _parameter_value_as_text(parameter, doc=None):
    if parameter is None:
        return ""

    value = ""
    try:
        value = _safe_text(parameter.AsString())
    except Exception:
        value = ""
    if value:
        return value

    try:
        value = _safe_text(parameter.AsValueString())
    except Exception:
        value = ""
    if value:
        return value

    try:
        storage_type = parameter.StorageType
    except Exception:
        storage_type = None

    try:
        if storage_type == StorageType.Integer:
            return str(int(parameter.AsInteger()))
        if storage_type == StorageType.Double:
            return "{0:.8f}".format(float(parameter.AsDouble())).rstrip("0").rstrip(".")
        if storage_type == StorageType.ElementId:
            element_id = parameter.AsElementId()
            id_val = _idval(element_id)
            if id_val <= 0:
                return ""
            if doc is not None:
                try:
                    element = doc.GetElement(element_id)
                except Exception:
                    element = None
                if element is not None:
                    name = _safe_text(getattr(element, "Name", ""))
                    if name:
                        return name
            return str(id_val)
    except Exception:
        pass

    return ""


def _collect_type_parameter_values(symbol, doc=None):
    values = {}
    if symbol is None:
        return values

    try:
        parameters = list(symbol.Parameters or [])
    except Exception:
        parameters = []

    for parameter in parameters:
        if parameter is None:
            continue
        try:
            definition = parameter.Definition
            name = _safe_text(getattr(definition, "Name", ""))
        except Exception:
            name = ""
        if not name:
            continue

        text_value = _parameter_value_as_text(parameter, doc)
        existing = _safe_text(values.get(name, ""))
        if existing and not text_value:
            continue
        values[name] = text_value

    return values


def feet_inch_to_inches(value):
    """Parse feet-inch text like 5'-6 1/4" into inches."""
    try:
        if value is None:
            return None
        text = _safe_text(value)
        if not text:
            return None

        text = text.replace('"', "")
        sign = 1.0
        if text.startswith("-"):
            sign = -1.0
            text = text[1:].strip()

        feet = 0.0
        inches = 0.0

        if "'" in text:
            feet_part, inch_part = text.split("'", 1)
            feet_part = _safe_text(feet_part)
            if feet_part:
                feet = float(feet_part)
            text = _safe_text(inch_part)
        else:
            text = _safe_text(text)

        if text:
            parts = text.split()
            if len(parts) == 1:
                if "/" in parts[0]:
                    num, den = parts[0].split("/")
                    inches = abs(float(num) / float(den))
                else:
                    inches = abs(float(parts[0]))
            elif len(parts) >= 2:
                whole = abs(float(parts[0]))
                num, den = parts[1].split("/")
                inches = whole + (float(num) / float(den))

        return sign * (feet * 12.0 + inches)
    except Exception:
        return None


def _to_float(value, fallback=0.0):
    try:
        return float(value)
    except Exception:
        return float(fallback)


def _transform_rotation(local_rotation, transform):
    base = _to_float(local_rotation, 0.0)
    if transform is None:
        return base
    try:
        local_vec = XYZ(math.cos(base), math.sin(base), 0.0)
        host_vec = transform.OfVector(local_vec)
        return math.atan2(host_vec.Y, host_vec.X)
    except Exception:
        return base


# -----------------------------------------------------------------------------
# Data models
class LightPlacementData(object):
    def __init__(
        self,
        source_id,
        location,
        rotation_radians,
        source_level_key=None,
        source_level_name="",
        source_level_elevation_host=None,
        source_level_offset=None,
        source_unit="feet",
        source_family_name="",
        source_type_name="",
        source_ref="",
    ):
        self.source_id = _safe_text(source_id)
        self.location = location
        self.rotation = _to_float(rotation_radians, 0.0)
        self.source_level_key = _safe_text(source_level_key)
        self.source_level_name = _safe_text(source_level_name)
        self.source_level_elevation_host = source_level_elevation_host
        self.source_level_offset = source_level_offset
        self.source_unit = _safe_text(source_unit).lower() or "feet"
        self.source_family_name = _safe_text(source_family_name)
        self.source_type_name = _safe_text(source_type_name)
        self.source_ref = _safe_text(source_ref)


class FixtureGroupInfo(object):
    def __init__(self, source_id, family_name, type_name, count, source_type_params=None):
        self.source_id = _safe_text(source_id)
        self.family_name = _safe_text(family_name)
        self.type_name = _safe_text(type_name)
        self.count = int(count or 0)
        self.source_type_params = dict(source_type_params or {})


class SourceLevelInfo(object):
    def __init__(self, level_key, level_name, elevation_host, count=0):
        self.level_key = _safe_text(level_key)
        self.level_name = _safe_text(level_name)
        self.elevation_host = _to_float(elevation_host, 0.0)
        self.count = int(count or 0)


class LightSourceResult(object):
    def __init__(self, mode, source_label):
        self.mode = _safe_text(mode)
        self.source_label = _safe_text(source_label)
        self.placements = []
        self.fixture_groups = []
        self.levels = []
        self.warnings = []


class HostFixtureOption(object):
    def __init__(self, symbol, family_name, type_name, label, type_mark, type_params=None):
        self.symbol = symbol
        self.family_name = _safe_text(family_name)
        self.type_name = _safe_text(type_name)
        self.type_mark = _safe_text(type_mark)
        self.label = _safe_text(label)
        self.family_type = "{} : {}".format(self.family_name, self.type_name)
        self.type_params = dict(type_params or {})
        self.type_params_norm = {}
        self.type_param_names_norm = {}
        for name, value in self.type_params.items():
            norm_name = _compact_norm(name)
            if not norm_name:
                continue
            if norm_name not in self.type_param_names_norm:
                self.type_param_names_norm[norm_name] = _safe_text(name)
            existing = _safe_text(self.type_params_norm.get(norm_name, ""))
            incoming = _safe_text(value)
            if existing and not incoming:
                continue
            self.type_params_norm[norm_name] = incoming

        self.family_norm = _compact_norm(self.family_name)
        self.type_norm = _compact_norm(self.type_name)
        self.label_norm = _compact_norm(self.label)
        self.family_type_norm = _compact_norm(self.family_type)
        self.type_mark_norm = _compact_norm(self.type_mark)
        self.search_text_norm = _compact_norm(
            "{} {} {}".format(self.family_name, self.type_name, self.type_mark)
        )
        self.tokens = _tokens("{} {} {}".format(self.family_name, self.type_name, self.type_mark))


class HostLevelOption(object):
    def __init__(self, level, label):
        self.level = level
        self.name = _safe_text(getattr(level, "Name", ""))
        self.elevation = _to_float(getattr(level, "Elevation", 0.0), 0.0)
        self.label = _safe_text(label)
        self.name_norm = _compact_norm(self.name)
        self.label_norm = _compact_norm(self.label)


class LinkInstanceOption(object):
    def __init__(self, link_instance, is_loaded, display_name):
        self.link_instance = link_instance
        self.is_loaded = bool(is_loaded)
        self.display_name = _safe_text(display_name)
        self.link_name = _safe_text(getattr(link_instance, "Name", "")) if link_instance else ""


class FixtureMappingRow(object):
    def __init__(
        self,
        source_id,
        source_label,
        count,
        target_options,
        source_family_name="",
        source_type_name="",
        source_type_params=None,
    ):
        self.source_id = _safe_text(source_id)
        self.source_label = _safe_text(source_label)
        self.count = int(count or 0)
        self.target_options = list(target_options or [])
        self._target_label = ""
        self.on_target_changed = None

        self.source_family_name = _safe_text(source_family_name)
        self.source_type_name = _safe_text(source_type_name)
        self.source_type_params = dict(source_type_params or {})
        self.source_type_params_norm = {}
        self.source_type_param_names_norm = {}
        for name, value in self.source_type_params.items():
            norm_name = _compact_norm(name)
            if not norm_name:
                continue
            if norm_name not in self.source_type_param_names_norm:
                self.source_type_param_names_norm[norm_name] = _safe_text(name)
            existing = _safe_text(self.source_type_params_norm.get(norm_name, ""))
            incoming = _safe_text(value)
            if existing and not incoming:
                continue
            self.source_type_params_norm[norm_name] = incoming

        self.fixture_code = _extract_fixture_code(self.source_label)

        self.source_norm = _compact_norm(self.source_label)
        self.source_tokens = _tokens(self.source_label)
        self.family_norm = _compact_norm(self.source_family_name)
        self.type_norm = _compact_norm(self.source_type_name)
        self.fixture_code_norm = _compact_norm(self.fixture_code)

    @property
    def target_label(self):
        return self._target_label

    @target_label.setter
    def target_label(self, value):
        next_value = _safe_text(value)
        if next_value == self._target_label:
            return
        self._target_label = next_value
        callback = getattr(self, "on_target_changed", None)
        if callable(callback):
            try:
                callback(self)
            except Exception:
                pass


class LevelMappingRow(object):
    def __init__(self, source_level_key, source_level_name, source_elevation, source_count, target_options):
        self.source_level_key = _safe_text(source_level_key)
        self.source_level_name = _safe_text(source_level_name)
        self.source_elevation = _to_float(source_elevation, 0.0)
        self.count = int(source_count or 0)
        self.source_elevation_display = _format_elevation(self.source_elevation)
        self.target_options = list(target_options or [])
        self._target_label = ""
        self.on_target_changed = None

        self.source_name_norm = _compact_norm(self.source_level_name)

    @property
    def target_label(self):
        return self._target_label

    @target_label.setter
    def target_label(self, value):
        next_value = _safe_text(value)
        if next_value == self._target_label:
            return
        self._target_label = next_value
        callback = getattr(self, "on_target_changed", None)
        if callable(callback):
            try:
                callback(self)
            except Exception:
                pass


class PlacementReport(object):
    def __init__(self):
        self.total = 0
        self.placed = 0
        self.skipped_unmapped_fixture = 0
        self.skipped_unmapped_level = 0
        self.skipped_invalid_point = 0
        self.skipped_duplicates = 0
        self.skipped_errors = 0
        self.error_messages = []
        self.fatal_error = ""

    def add_error(self, message):
        text = _safe_text(message)
        if not text:
            return
        if len(self.error_messages) < 20:
            self.error_messages.append(text)
        self.skipped_errors += 1

    def summary_lines(self):
        lines = [
            "Total source fixtures: {}".format(int(self.total)),
            "Placed: {}".format(int(self.placed)),
            "Skipped - unmapped fixture: {}".format(int(self.skipped_unmapped_fixture)),
            "Skipped - unmapped level: {}".format(int(self.skipped_unmapped_level)),
            "Skipped - invalid point: {}".format(int(self.skipped_invalid_point)),
            "Skipped - duplicates: {}".format(int(self.skipped_duplicates)),
            "Skipped - errors: {}".format(int(self.skipped_errors)),
        ]
        if self.fatal_error:
            lines.append("Fatal error: {}".format(self.fatal_error))
        return lines

    def short_alert(self):
        return (
            "Placed: {placed}\n"
            "Skipped (unmapped fixture): {sf}\n"
            "Skipped (unmapped level): {sl}\n"
            "Skipped (invalid point): {sp}\n"
            "Skipped (duplicates): {sd}\n"
            "Skipped (errors): {se}"
        ).format(
            placed=int(self.placed),
            sf=int(self.skipped_unmapped_fixture),
            sl=int(self.skipped_unmapped_level),
            sp=int(self.skipped_invalid_point),
            sd=int(self.skipped_duplicates),
            se=int(self.skipped_errors),
        )


# -----------------------------------------------------------------------------
# Mapping persistence (disabled for now; keep API surface for future re-enable)
class MappingStore(object):
    def __init__(self):
        self._cfg = None

    def _load_json(self, key):
        return {}

    def _save_json(self, key, data):
        return

    def load_fixture_map(self, mode):
        mode_key = _safe_text(mode).lower() or CSV_SOURCE_MODE
        return self._load_json("fixture_map_{}".format(mode_key))

    def save_fixture_map(self, mode, mapping):
        mode_key = _safe_text(mode).lower() or CSV_SOURCE_MODE
        self._save_json("fixture_map_{}".format(mode_key), mapping)

    def load_level_map(self):
        return self._load_json("level_map_link")

    def save_level_map(self, mapping):
        self._save_json("level_map_link", mapping)


# -----------------------------------------------------------------------------
# Data source implementations
class CsvLightSource(object):
    def __init__(self, csv_path):
        self.csv_path = _safe_text(csv_path)

    def collect(self):
        result = LightSourceResult(CSV_SOURCE_MODE, self.csv_path)
        if not self.csv_path:
            result.warnings.append("No CSV path supplied.")
            return result
        if not os.path.exists(self.csv_path):
            result.warnings.append("CSV file does not exist: {}".format(self.csv_path))
            return result

        grouped = {}
        warning_count = 0

        with codecs.open(self.csv_path, "r", encoding="utf-8-sig") as stream:
            reader = csv.DictReader(stream, delimiter=",")
            if reader.fieldnames:
                reader.fieldnames = [h.strip() for h in reader.fieldnames if h is not None]

            for row_index, row in enumerate(reader, 2):
                count_text = _safe_text(row.get("Count"))
                if count_text and count_text != "1":
                    continue

                source_name = _safe_text(row.get("Name"))
                if not source_name:
                    warning_count += 1
                    if warning_count <= 20:
                        result.warnings.append("Row {} has no Name. Skipped.".format(row_index))
                    continue

                x_inches = feet_inch_to_inches(row.get("Position X"))
                y_inches = feet_inch_to_inches(row.get("Position Y"))
                z_inches = feet_inch_to_inches(row.get("Position Z"))
                if x_inches is None or y_inches is None or z_inches is None:
                    warning_count += 1
                    if warning_count <= 20:
                        result.warnings.append(
                            "Row {} has invalid XYZ. Skipped '{}'.".format(row_index, source_name)
                        )
                    continue

                rotation_deg = _to_float(row.get("Rotation"), 0.0)
                rotation_rad = math.radians(rotation_deg)

                data = LightPlacementData(
                    source_id=source_name,
                    location=XYZ(float(x_inches), float(y_inches), float(z_inches)),
                    rotation_radians=rotation_rad,
                    source_level_key="",
                    source_level_name="",
                    source_level_elevation_host=None,
                    source_level_offset=None,
                    source_unit="inches",
                    source_family_name="",
                    source_type_name="",
                    source_ref="csv:{}".format(row_index),
                )
                result.placements.append(data)

                group = grouped.get(source_name)
                if group is None:
                    grouped[source_name] = FixtureGroupInfo(
                        source_id=source_name,
                        family_name="",
                        type_name="",
                        count=1,
                        source_type_params={},
                    )
                else:
                    group.count += 1

        result.fixture_groups = sorted(
            list(grouped.values()),
            key=lambda g: (_safe_text(g.source_id).lower(), int(g.count)),
        )
        return result


class RevitLinkLightSource(object):
    UNASSIGNED_LEVEL_KEY = "__NO_LEVEL__"
    UNASSIGNED_LEVEL_NAME = "<No Level Assigned in Link>"

    def __init__(self, host_doc, link_instance, category=None):
        self.host_doc = host_doc
        self.link_instance = link_instance
        self.category = category if category is not None else DEFAULT_CATEGORY

    def _get_transform(self):
        if self.link_instance is None:
            return None
        try:
            return self.link_instance.GetTotalTransform()
        except Exception:
            pass
        try:
            return self.link_instance.GetTransform()
        except Exception:
            return None

    def _host_point_from_link(self, transform, point):
        if transform is None:
            return point
        try:
            return transform.OfPoint(point)
        except Exception:
            return point

    def _host_level_elevation(self, transform, link_level_elevation):
        point = XYZ(0.0, 0.0, _to_float(link_level_elevation, 0.0))
        if transform is None:
            return point.Z
        try:
            return transform.OfPoint(point).Z
        except Exception:
            return point.Z

    def _level_from_element_id(self, link_doc, element_id):
        if link_doc is None or element_id is None:
            return None
        try:
            if element_id == ElementId.InvalidElementId:
                return None
        except Exception:
            pass
        try:
            int_id = int(element_id.IntegerValue)
            if int_id <= 0:
                return None
        except Exception:
            int_id = -1
            try:
                int_id = int(element_id)
            except Exception:
                int_id = -1
            if int_id <= 0:
                return None
            try:
                element_id = ElementId(int_id)
            except Exception:
                return None

        try:
            candidate = link_doc.GetElement(element_id)
        except Exception:
            candidate = None
        return candidate if isinstance(candidate, Level) else None

    def _level_from_bip(self, element, link_doc, bip_name):
        try:
            bip = getattr(BuiltInParameter, str(bip_name or ""), None)
        except Exception:
            bip = None
        if bip is None:
            return None
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
        if param is None:
            return None
        try:
            return self._level_from_element_id(link_doc, param.AsElementId())
        except Exception:
            return None

    def _level_from_lookup(self, element, link_doc, param_name):
        try:
            param = element.LookupParameter(str(param_name or ""))
        except Exception:
            param = None
        if param is None:
            return None
        try:
            return self._level_from_element_id(link_doc, param.AsElementId())
        except Exception:
            return None

    def _resolve_linked_level(self, element, link_doc):
        # 1) Start with native LevelId.
        try:
            native_level = self._level_from_element_id(link_doc, getattr(element, "LevelId", None))
        except Exception:
            native_level = None
        if native_level is not None:
            return native_level

        # 2) Fallback through common level params used by hosted/face-based families.
        bip_candidates = (
            "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM",
            "SCHEDULE_LEVEL_PARAM",
            "FAMILY_LEVEL_PARAM",
            "INSTANCE_LEVEL_PARAM",
            "INSTANCE_REFERENCE_LEVEL_PARAM",
        )
        for bip_name in bip_candidates:
            level = self._level_from_bip(element, link_doc, bip_name)
            if level is not None:
                return level

        # 3) Final fallback through named parameters.
        lookup_candidates = ("Schedule Level", "Reference Level", "Level")
        for lookup_name in lookup_candidates:
            level = self._level_from_lookup(element, link_doc, lookup_name)
            if level is not None:
                return level

        return None

    def collect(self):
        link_name = _safe_text(getattr(self.link_instance, "Name", ""))
        result = LightSourceResult(LINK_SOURCE_MODE, link_name)

        if self.link_instance is None:
            result.warnings.append("No Revit link instance selected.")
            return result

        try:
            link_doc = self.link_instance.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            result.warnings.append("Selected link is unloaded or unavailable.")
            return result

        transform = self._get_transform()
        grouped = {}
        levels_by_key = {}
        warning_count = 0

        collector = (
            FilteredElementCollector(link_doc)
            .OfCategory(self.category)
            .WhereElementIsNotElementType()
        )

        for element in collector:
            if not isinstance(element, FamilyInstance):
                continue

            loc_point = _as_location_point(element)
            if loc_point is None:
                warning_count += 1
                if warning_count <= 20:
                    result.warnings.append(
                        "Link fixture id {} has no point location. Skipped.".format(_idval(element))
                    )
                continue

            symbol = getattr(element, "Symbol", None)
            if symbol is None:
                warning_count += 1
                if warning_count <= 20:
                    result.warnings.append(
                        "Link fixture id {} has no FamilySymbol. Skipped.".format(_idval(element))
                    )
                continue

            family_name, type_name = _symbol_names(symbol)
            source_id = "{} : {}".format(family_name, type_name)

            host_point = self._host_point_from_link(transform, loc_point.Point)
            host_rotation = _transform_rotation(getattr(loc_point, "Rotation", 0.0), transform)

            level_key = ""
            level_name = ""
            level_elevation_host = None
            level_offset = None
            linked_level = self._resolve_linked_level(element, link_doc)

            if isinstance(linked_level, Level):
                level_key = str(_idval(linked_level.Id))
                level_name = _safe_text(linked_level.Name) or "Level {}".format(level_key)
                level_elevation_host = self._host_level_elevation(transform, linked_level.Elevation)
                level_offset = float(host_point.Z) - float(level_elevation_host)
            else:
                level_key = self.UNASSIGNED_LEVEL_KEY
                level_name = self.UNASSIGNED_LEVEL_NAME
                level_elevation_host = float(host_point.Z)
                level_offset = 0.0

            if level_key not in levels_by_key:
                levels_by_key[level_key] = SourceLevelInfo(
                    level_key=level_key,
                    level_name=level_name,
                    elevation_host=level_elevation_host,
                    count=0,
                )
            levels_by_key[level_key].count += 1

            placement = LightPlacementData(
                source_id=source_id,
                location=host_point,
                rotation_radians=host_rotation,
                source_level_key=level_key,
                source_level_name=level_name,
                source_level_elevation_host=level_elevation_host,
                source_level_offset=level_offset,
                source_unit="feet",
                source_family_name=family_name,
                source_type_name=type_name,
                source_ref="link:{}".format(_idval(element)),
            )
            result.placements.append(placement)

            group = grouped.get(source_id)
            if group is None:
                source_type_params = _collect_type_parameter_values(symbol, link_doc)
                grouped[source_id] = FixtureGroupInfo(
                    source_id=source_id,
                    family_name=family_name,
                    type_name=type_name,
                    count=1,
                    source_type_params=source_type_params,
                )
            else:
                group.count += 1

        result.fixture_groups = sorted(
            list(grouped.values()),
            key=lambda g: (_safe_text(g.family_name).lower(), _safe_text(g.type_name).lower()),
        )
        result.levels = sorted(
            list(levels_by_key.values()),
            key=lambda lvl: (_to_float(lvl.elevation_host, 0.0), _safe_text(lvl.level_name).lower()),
        )
        return result


# -----------------------------------------------------------------------------
# Matching services
class FixtureMatcher(object):
    def __init__(self, host_options):
        self.options = list(host_options or [])
        self.by_family_type = defaultdict(list)
        self.by_type = defaultdict(list)
        self.by_type_mark = defaultdict(list)
        self.by_label = defaultdict(list)
        for option in self.options:
            if option.family_norm or option.type_norm:
                self.by_family_type["{}|{}".format(option.family_norm, option.type_norm)].append(option)
            if option.type_norm:
                self.by_type[option.type_norm].append(option)
            if option.type_mark_norm:
                self.by_type_mark[option.type_mark_norm].append(option)
            if option.label_norm:
                self.by_label[option.label_norm].append(option)

    def _best_similarity(self, row):
        source_norm = row.source_norm
        if not source_norm:
            source_norm = "{}{}".format(row.family_norm, row.type_norm)
        source_tokens = set(row.source_tokens or [])

        best_score = 0.0
        best_option = None
        for option in self.options:
            base_ratio = SequenceMatcher(None, source_norm, option.search_text_norm).ratio()
            token_overlap = 0.0
            if source_tokens and option.tokens:
                overlap = float(len(source_tokens.intersection(option.tokens)))
                token_overlap = overlap * 0.03
            score = base_ratio + token_overlap
            if score > best_score:
                best_score = score
                best_option = option
        if best_option is not None and best_score >= 0.60:
            return best_option.label
        return ""

    def _single_label(self, options):
        if len(options or []) == 1:
            return options[0].label
        return ""

    def match(self, row):
        if row is None:
            return ""

        if row.fixture_code_norm:
            code_match = self._single_label(self.by_type_mark.get(row.fixture_code_norm, []))
            if code_match:
                return code_match

        if row.family_norm or row.type_norm:
            exact_key = "{}|{}".format(row.family_norm, row.type_norm)
            exact_match = self._single_label(self.by_family_type.get(exact_key, []))
            if exact_match:
                return exact_match

        if row.type_norm:
            type_match = self._single_label(self.by_type.get(row.type_norm, []))
            if type_match:
                return type_match

        if row.source_norm:
            exact_label = self._single_label(self.by_label.get(row.source_norm, []))
            if exact_label:
                return exact_label

        return self._best_similarity(row)


class LevelMatcher(object):
    def __init__(self, host_levels):
        self.options = list(host_levels or [])
        self.by_name = defaultdict(list)
        for option in self.options:
            if option.name_norm:
                self.by_name[option.name_norm].append(option)

    def match(self, row):
        if row is None:
            return ""
        name_norm = row.source_name_norm
        if name_norm:
            candidates = list(self.by_name.get(name_norm, []))
            if len(candidates) == 1:
                return candidates[0].label
            if len(candidates) > 1:
                candidates.sort(key=lambda x: abs(float(x.elevation) - float(row.source_elevation)))
                return candidates[0].label

        # Elevation fallback
        if self.options:
            elev_sorted = sorted(
                self.options,
                key=lambda x: abs(float(x.elevation) - float(row.source_elevation)),
            )
            best = elev_sorted[0]
            if abs(float(best.elevation) - float(row.source_elevation)) <= 0.5:
                return best.label

        # Name similarity fallback
        best_score = 0.0
        best_option = None
        source_norm = name_norm
        for option in self.options:
            ratio = SequenceMatcher(None, source_norm, option.name_norm).ratio()
            if ratio > best_score:
                best_score = ratio
                best_option = option
        if best_option is not None and best_score >= 0.72:
            return best_option.label
        return ""


# -----------------------------------------------------------------------------
# Placement engine
class PlacementEngine(object):
    def __init__(self, doc, category=None):
        self.doc = doc
        self.category = category if category is not None else DEFAULT_CATEGORY

    def _activate_symbols(self, symbols):
        changed = False
        for symbol in set([s for s in list(symbols or []) if s is not None]):
            try:
                if not symbol.IsActive:
                    symbol.Activate()
                    changed = True
            except Exception:
                pass
        if changed:
            self.doc.Regenerate()

    def _duplicate_key(self, symbol_id, point):
        tol = float(DUPLICATE_TOLERANCE_FT)
        return (
            int(symbol_id),
            int(round(float(point.X) / tol)),
            int(round(float(point.Y) / tol)),
            int(round(float(point.Z) / tol)),
        )

    def _build_existing_index(self, symbol_ids):
        index = set()
        if not symbol_ids:
            return index

        collector = (
            FilteredElementCollector(self.doc)
            .OfCategory(self.category)
            .WhereElementIsNotElementType()
        )
        for instance in collector:
            if not isinstance(instance, FamilyInstance):
                continue
            try:
                symbol = instance.Symbol
            except Exception:
                symbol = None
            symbol_id = _idval(getattr(symbol, "Id", None))
            if symbol_id <= 0 or symbol_id not in symbol_ids:
                continue

            loc_point = _as_location_point(instance)
            if loc_point is None:
                continue
            index.add(self._duplicate_key(symbol_id, loc_point.Point))
        return index

    def _resolve_target_point(self, data, target_level, csv_divisor):
        if data is None or data.location is None:
            return None

        point = data.location
        if _safe_text(data.source_unit).lower() == "inches":
            divisor = _to_float(csv_divisor, 0.0)
            if abs(divisor) < 1e-9:
                return None
            point = XYZ(point.X / divisor, point.Y / divisor, point.Z / divisor)

        if data.source_level_key and data.source_level_offset is not None and target_level is not None:
            try:
                point = XYZ(point.X, point.Y, float(target_level.Elevation) + float(data.source_level_offset))
            except Exception:
                pass
        return point

    def _create_instance(self, point, symbol, level):
        try:
            return self.doc.Create.NewFamilyInstance(
                point, symbol, level, Structure.StructuralType.NonStructural
            )
        except Exception:
            # Some fixture families can still place without explicit level.
            return self.doc.Create.NewFamilyInstance(point, symbol, Structure.StructuralType.NonStructural)

    def place(
        self,
        placements,
        fixture_symbol_by_source,
        default_level,
        level_by_source,
        csv_divisor,
        skip_duplicates,
    ):
        report = PlacementReport()
        report.total = int(len(list(placements or [])))

        if not placements:
            return report

        symbol_ids = set()
        for symbol in list(fixture_symbol_by_source.values()):
            sid = _idval(getattr(symbol, "Id", None))
            if sid > 0:
                symbol_ids.add(sid)

        existing_index = set()
        if skip_duplicates:
            existing_index = self._build_existing_index(symbol_ids)

        group = TransactionGroup(self.doc, "Let There Be Light")
        transaction = Transaction(self.doc, "Place Fixtures")
        group.Start()
        transaction.Start()
        try:
            self._activate_symbols(list(fixture_symbol_by_source.values()))

            for data in list(placements or []):
                symbol = fixture_symbol_by_source.get(data.source_id)
                if symbol is None:
                    report.skipped_unmapped_fixture += 1
                    continue

                target_level = default_level
                if data.source_level_key and level_by_source:
                    mapped_level = level_by_source.get(data.source_level_key)
                    if mapped_level is not None:
                        target_level = mapped_level

                if target_level is None:
                    report.skipped_unmapped_level += 1
                    continue

                point = self._resolve_target_point(data, target_level, csv_divisor)
                if point is None:
                    report.skipped_invalid_point += 1
                    continue

                symbol_id = _idval(getattr(symbol, "Id", None))
                dup_key = self._duplicate_key(symbol_id, point)
                if skip_duplicates and dup_key in existing_index:
                    report.skipped_duplicates += 1
                    continue

                try:
                    instance = self._create_instance(point, symbol, target_level)
                    if abs(float(data.rotation or 0.0)) > 1e-9:
                        axis = Line.CreateBound(point, point + XYZ(0.0, 0.0, 1.0))
                        ElementTransformUtils.RotateElement(self.doc, instance.Id, axis, float(data.rotation))
                    report.placed += 1
                    if skip_duplicates:
                        existing_index.add(dup_key)
                except InvalidOperationException as ex:
                    report.add_error("Placement failed for '{}': {}".format(data.source_id, ex))
                except Exception as ex:
                    report.add_error("Placement failed for '{}': {}".format(data.source_id, ex))

            transaction.Commit()
            group.Assimilate()
        except Exception as ex:
            report.fatal_error = _safe_text(ex)
            try:
                transaction.RollBack()
            except Exception:
                pass
            try:
                group.RollBack()
            except Exception:
                pass

        return report


# -----------------------------------------------------------------------------
# Revit data collection for host model
def collect_host_fixture_options(doc, category=None):
    if category is None:
        category = DEFAULT_CATEGORY
    symbols = (
        FilteredElementCollector(doc)
        .OfClass(FamilySymbol)
        .OfCategory(category)
        .ToElements()
    )
    records = []
    base_counts = defaultdict(int)
    for symbol in symbols:
        family_name, type_name = _symbol_names(symbol)
        base_label = "{} : {}".format(family_name, type_name)
        base_counts[base_label] += 1
        records.append(
            (
                symbol,
                family_name,
                type_name,
                base_label,
                _symbol_type_mark(symbol),
                _collect_type_parameter_values(symbol, doc),
            )
        )

    options = []
    for symbol, family_name, type_name, base_label, type_mark, type_params in records:
        label = base_label
        if base_counts[base_label] > 1:
            label = "{} [Id {}]".format(base_label, _idval(symbol.Id))
        options.append(
            HostFixtureOption(
                symbol=symbol,
                family_name=family_name,
                type_name=type_name,
                label=label,
                type_mark=type_mark,
                type_params=type_params,
            )
        )

    options.sort(key=lambda x: (x.family_name.lower(), x.type_name.lower(), x.label.lower()))
    by_label = dict([(option.label, option) for option in options])
    return options, by_label


def collect_host_levels(doc):
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    name_counts = defaultdict(int)
    for level in levels:
        name_counts[_safe_text(level.Name)] += 1

    options = []
    for level in levels:
        name = _safe_text(level.Name) or "Unnamed Level"
        elevation = _to_float(level.Elevation, 0.0)
        label = "{} ({})".format(name, _format_elevation(elevation))
        if name_counts[name] > 1:
            label = "{} [Id {}]".format(label, _idval(level.Id))
        options.append(HostLevelOption(level, label))

    options.sort(key=lambda x: (float(x.elevation), x.name.lower()))
    by_label = dict([(option.label, option) for option in options])
    return options, by_label


def collect_link_options(doc):
    collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    options = []
    for link in collector:
        loaded = False
        try:
            loaded = link.GetLinkDocument() is not None
        except Exception:
            loaded = False
        status = "Loaded" if loaded else "Unloaded"
        link_name = _safe_text(getattr(link, "Name", "")) or "Unnamed Link"
        display_name = "{} ({})".format(link_name, status)
        options.append(LinkInstanceOption(link_instance=link, is_loaded=loaded, display_name=display_name))

    options.sort(
        key=lambda x: (
            0 if x.is_loaded else 1,
            x.display_name.lower(),
        )
    )
    return options


# -----------------------------------------------------------------------------
class AutoMatchFieldOption(object):
    def __init__(self, key, label):
        self.key = _safe_text(key)
        self.label = _safe_text(label)


class AutoMatchRow(object):
    def __init__(self, fixture_row):
        self.source_id = _safe_text(fixture_row.source_id)
        self.source_label = _safe_text(fixture_row.source_label)
        self.source_family_name = _safe_text(getattr(fixture_row, "source_family_name", ""))
        self.source_type_name = _safe_text(getattr(fixture_row, "source_type_name", ""))
        self.source_type_params = dict(getattr(fixture_row, "source_type_params", {}) or {})
        self.source_type_params_norm = {}
        self.source_type_param_names_norm = {}
        for name, value in self.source_type_params.items():
            norm_name = _compact_norm(name)
            if not norm_name:
                continue
            if norm_name not in self.source_type_param_names_norm:
                self.source_type_param_names_norm[norm_name] = _safe_text(name)
            existing = _safe_text(self.source_type_params_norm.get(norm_name, ""))
            incoming = _safe_text(value)
            if existing and not incoming:
                continue
            self.source_type_params_norm[norm_name] = incoming

        self.source_fixture_code = _safe_text(getattr(fixture_row, "fixture_code", ""))
        self.target_label = _safe_text(getattr(fixture_row, "target_label", ""))
        self.is_complete = bool(self.target_label)
        self.source_value_display = ""

    def source_param_value_by_norm(self, norm_name):
        norm = _compact_norm(norm_name)
        if not norm:
            return ""
        return _safe_text(self.source_type_params_norm.get(norm, ""))

    def source_field_value(self, field_key):
        key = _safe_text(field_key).lower()
        if key.startswith("param:"):
            return self.source_param_value_by_norm(key.split("param:", 1)[1])
        if key == "source_family_name":
            return self.source_family_name
        if key == "source_type_name":
            return self.source_type_name
        if key == "source_fixture_code":
            return self.source_fixture_code
        if key == "source_id":
            return self.source_id
        return self.source_label


class AutoMatchWindow(CEDWindowBase):
    theme_aware = True
    use_config_theme = True

    def __init__(
        self,
        fixture_rows,
        host_fixture_options,
        source_mode=LINK_SOURCE_MODE,
        theme_mode=None,
        accent_mode=None,
    ):
        self.host_fixture_options = list(host_fixture_options or [])
        self.rows = [AutoMatchRow(row) for row in list(fixture_rows or [])]
        self.source_mode = _safe_text(source_mode).lower() or LINK_SOURCE_MODE
        self.is_csv_mode = self.source_mode == CSV_SOURCE_MODE
        self.result_map = {}
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "AutoMatchWindow.xaml"))
        CEDWindowBase.__init__(
            self,
            xaml_source=xaml_path,
            theme_mode=theme_mode,
            accent_mode=accent_mode,
            theme_aware=True,
            use_config_theme=True,
        )
        self._bind_controls()
        self._wire_events()
        self._initialize()

    def _bind_controls(self):
        self.LinkParamPanel = self.FindName("LinkParamPanel")
        self.SourceParamLabel = self.FindName("SourceParamLabel")
        self.SourceParamCombo = self.FindName("SourceParamCombo")
        self.TargetParamLabel = self.FindName("TargetParamLabel")
        self.LinkTargetParamCombo = self.FindName("LinkTargetParamCombo")
        self.CsvTargetParamCombo = self.FindName("CsvTargetParamCombo")
        self.TargetParamCombo = None
        self.CsvDerivePanel = self.FindName("CsvDerivePanel")
        self.CsvBeforeTextBox = self.FindName("CsvBeforeTextBox")
        self.CsvAfterTextBox = self.FindName("CsvAfterTextBox")
        self.CsvSampleTextBox = self.FindName("CsvSampleTextBox")
        self.LinkRunMatchButton = self.FindName("LinkRunMatchButton")
        self.CsvRunMatchButton = self.FindName("CsvRunMatchButton")
        self.RunMatchButton = None
        self.RunStatusText = self.FindName("RunStatusText")
        self.MatchGrid = self.FindName("MatchGrid")
        self.StatusText = self.FindName("StatusText")
        self.OkButton = self.FindName("OkButton")
        self.CancelButton = self.FindName("CancelButton")

    def _wire_events(self):
        if self.SourceParamCombo is not None:
            self.SourceParamCombo.SelectionChanged += self._source_field_changed
        if self.LinkTargetParamCombo is not None:
            self.LinkTargetParamCombo.SelectionChanged += self._target_field_changed
        if self.CsvTargetParamCombo is not None:
            self.CsvTargetParamCombo.SelectionChanged += self._target_field_changed
        if self.CsvBeforeTextBox is not None:
            self.CsvBeforeTextBox.TextChanged += self._csv_strip_text_changed
        if self.CsvAfterTextBox is not None:
            self.CsvAfterTextBox.TextChanged += self._csv_strip_text_changed
        if self.LinkRunMatchButton is not None:
            self.LinkRunMatchButton.Click += self._run_match_clicked
        if self.CsvRunMatchButton is not None:
            self.CsvRunMatchButton.Click += self._run_match_clicked
        if self.MatchGrid is not None:
            self.MatchGrid.SelectionChanged += self._match_grid_selection_changed
        self.OkButton.Click += self._ok_clicked
        self.CancelButton.Click += self._cancel_clicked

    def _initialize(self):
        source_param_fields = self._build_source_parameter_fields()
        target_param_fields = self._build_target_parameter_fields()

        if self.is_csv_mode:
            self.source_fields = []
            if self.LinkParamPanel is not None:
                self.LinkParamPanel.Visibility = Visibility.Collapsed
            self._set_source_param_controls_visible(False)
            self._set_csv_controls_visible(True)
            self.TargetParamCombo = self.CsvTargetParamCombo
            self.RunMatchButton = self.CsvRunMatchButton
            run_status = "CSV mode: source value derives from Text After and Text Before."
        else:
            self.source_fields = list(source_param_fields)
            if not self.source_fields:
                self.source_fields = [
                    AutoMatchFieldOption("source_label", "Source Label"),
                    AutoMatchFieldOption("source_fixture_code", "Fixture Code"),
                    AutoMatchFieldOption("source_family_name", "Family Name"),
                    AutoMatchFieldOption("source_type_name", "Type Name"),
                    AutoMatchFieldOption("source_id", "Source Id"),
                ]
            if self.LinkParamPanel is not None:
                self.LinkParamPanel.Visibility = Visibility.Visible
            self._set_source_param_controls_visible(True)
            self._set_csv_controls_visible(False)
            self.TargetParamCombo = self.LinkTargetParamCombo
            self.RunMatchButton = self.LinkRunMatchButton
            run_status = "Link mode: source value uses selected source type parameter."

        self.target_fields = list(target_param_fields)
        if not self.target_fields:
            self.target_fields = [
                AutoMatchFieldOption("label", "Full Label"),
                AutoMatchFieldOption("family_type", "Family + Type"),
                AutoMatchFieldOption("type_mark", "Type Mark"),
                AutoMatchFieldOption("family_name", "Family Name"),
                AutoMatchFieldOption("type_name", "Type Name"),
            ]

        if self.SourceParamCombo is not None:
            self.SourceParamCombo.ItemsSource = self.source_fields
            self.SourceParamCombo.DisplayMemberPath = "label"
            if self.source_fields:
                source_index = self._preferred_param_index(self.source_fields)
                self.SourceParamCombo.SelectedIndex = source_index if source_index >= 0 else 0
            else:
                self.SourceParamCombo.SelectedIndex = -1

        if self.TargetParamCombo is not None:
            self.TargetParamCombo.ItemsSource = self.target_fields
            self.TargetParamCombo.DisplayMemberPath = "label"
            if self.target_fields:
                target_index = self._preferred_param_index(self.target_fields)
                self.TargetParamCombo.SelectedIndex = target_index if target_index >= 0 else 0
            else:
                self.TargetParamCombo.SelectedIndex = -1

        if self.CsvBeforeTextBox is not None:
            self.CsvBeforeTextBox.Text = ""
        if self.CsvAfterTextBox is not None:
            self.CsvAfterTextBox.Text = ""

        self.MatchGrid.ItemsSource = self.rows
        if self.RunStatusText is not None:
            self.RunStatusText.Text = run_status
        self._refresh_source_value_display()
        self._refresh_sample_source_text()
        self._update_status("Ready. Choose fields and run match.")

    def _set_source_param_controls_visible(self, visible):
        vis = Visibility.Visible if bool(visible) else Visibility.Collapsed
        if self.SourceParamLabel is not None:
            self.SourceParamLabel.Visibility = vis
        if self.SourceParamCombo is not None:
            self.SourceParamCombo.Visibility = vis

    def _set_csv_controls_visible(self, visible):
        if self.CsvDerivePanel is None:
            return
        self.CsvDerivePanel.Visibility = Visibility.Visible if bool(visible) else Visibility.Collapsed

    def _selected_source_field(self):
        if self.is_csv_mode:
            return ""
        selected = self.SourceParamCombo.SelectedItem
        return _safe_text(getattr(selected, "key", "")) or "source_label"

    def _selected_target_field(self):
        if self.TargetParamCombo is None:
            return "label"
        selected = self.TargetParamCombo.SelectedItem
        return _safe_text(getattr(selected, "key", "")) or "label"

    def _target_field_value(self, option, field_key):
        key = _safe_text(field_key).lower()
        if key.startswith("param:"):
            param_norm = _compact_norm(key.split("param:", 1)[1])
            return _safe_text(getattr(option, "type_params_norm", {}).get(param_norm, ""))
        if key == "family_name":
            return option.family_name
        if key == "type_name":
            return option.type_name
        if key == "type_mark":
            return option.type_mark
        if key == "family_type":
            return option.family_type
        return option.label

    def _derive_csv_source_value(self, source_text):
        value = _safe_text(source_text)
        if not value:
            return ""

        after_text = _safe_text(getattr(self.CsvAfterTextBox, "Text", ""))
        before_text = _safe_text(getattr(self.CsvBeforeTextBox, "Text", ""))

        if after_text:
            idx = value.find(after_text)
            if idx >= 0:
                value = value[idx + len(after_text) :]

        if before_text:
            idx = value.find(before_text)
            if idx >= 0:
                value = value[:idx]

        return _safe_text(value)

    def _source_value_for_row(self, row):
        if not isinstance(row, AutoMatchRow):
            return ""
        if self.is_csv_mode:
            return self._derive_csv_source_value(row.source_label)
        source_field = self._selected_source_field()
        return row.source_field_value(source_field)

    def _refresh_source_value_display(self):
        for row in self.rows:
            row.source_value_display = self._source_value_for_row(row)
        try:
            self.MatchGrid.Items.Refresh()
        except Exception:
            pass
        self._refresh_sample_source_text()

    def _source_field_changed(self, sender, args):
        if self.is_csv_mode:
            return
        self._refresh_source_value_display()

    def _target_field_changed(self, sender, args):
        return

    def _csv_strip_text_changed(self, sender, args):
        if not self.is_csv_mode:
            return
        self._refresh_source_value_display()

    def _selected_match_row(self):
        row = getattr(self.MatchGrid, "SelectedItem", None)
        if isinstance(row, AutoMatchRow):
            return row
        if self.rows:
            return self.rows[0]
        return None

    def _refresh_sample_source_text(self):
        if self.CsvSampleTextBox is None:
            return
        row = self._selected_match_row()
        sample = _safe_text(getattr(row, "source_label", "")) if row is not None else ""
        self.CsvSampleTextBox.Text = sample

    def _match_grid_selection_changed(self, sender, args):
        self._refresh_sample_source_text()

    def _update_status(self, text):
        if self.StatusText is not None:
            self.StatusText.Text = _safe_text(text)

    def _build_source_parameter_fields(self):
        source_names = {}
        for row in self.rows:
            for norm_name, display_name in dict(getattr(row, "source_type_param_names_norm", {}) or {}).items():
                if norm_name and norm_name not in source_names:
                    source_names[norm_name] = _safe_text(display_name)
        norms = sorted(list(source_names.keys()), key=lambda n: (_safe_text(source_names.get(n, "")).lower(), n))
        return [AutoMatchFieldOption("param:{}".format(norm), source_names.get(norm, norm)) for norm in norms]

    def _build_target_parameter_fields(self):
        target_names = {}
        for option in self.host_fixture_options:
            for norm_name, display_name in dict(getattr(option, "type_param_names_norm", {}) or {}).items():
                if norm_name and norm_name not in target_names:
                    target_names[norm_name] = _safe_text(display_name)
        norms = sorted(list(target_names.keys()), key=lambda n: (_safe_text(target_names.get(n, "")).lower(), n))
        return [AutoMatchFieldOption("param:{}".format(norm), target_names.get(norm, norm)) for norm in norms]

    def _preferred_param_index(self, param_fields):
        preferred = ("identitytypemark", "typemark", "fixturetypecedt", "cedefixturetype")
        for index, field in enumerate(list(param_fields or [])):
            key = _safe_text(getattr(field, "key", "")).lower()
            if not key.startswith("param:"):
                continue
            param_norm = _safe_text(key.split("param:", 1)[1])
            if param_norm in preferred:
                return int(index)
        if param_fields:
            return 0
        return -1

    def _run_match_clicked(self, sender, args):
        target_field = self._selected_target_field()
        target_map = defaultdict(list)
        for option in self.host_fixture_options:
            value = _safe_text(self._target_field_value(option, target_field))
            norm = _compact_norm(value)
            if not norm:
                continue
            target_map[norm].append(option.label)

        matched = 0
        unresolved = 0
        ambiguous = 0
        skipped_complete = 0

        for row in self.rows:
            row.source_value_display = self._source_value_for_row(row)
            if bool(row.is_complete):
                skipped_complete += 1
                continue

            source_value = _safe_text(row.source_value_display)
            source_norm = _compact_norm(source_value)
            if not source_norm:
                unresolved += 1
                continue

            candidates = sorted(list(set(target_map.get(source_norm, []))))
            if len(candidates) == 1:
                row.target_label = candidates[0]
                row.is_complete = True
                matched += 1
            elif len(candidates) > 1:
                ambiguous += 1
            else:
                unresolved += 1

        try:
            self.MatchGrid.Items.Refresh()
        except Exception:
            pass
        if self.RunStatusText is not None:
            self.RunStatusText.Text = (
                "Matched: {0} | Ambiguous: {1} | Unresolved: {2} | Locked: {3}".format(
                    int(matched),
                    int(ambiguous),
                    int(unresolved),
                    int(skipped_complete),
                )
            )
        self._update_status("Run complete. Review rows and apply results.")

    def _ok_clicked(self, sender, args):
        self.result_map = {}
        for row in self.rows:
            label = _safe_text(row.target_label)
            if label:
                self.result_map[row.source_id] = label
        self.DialogResult = True
        self.Close()

    def _cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


# -----------------------------------------------------------------------------
# Main UI window / controller
class LightMappingWindow(CEDWindowBase):
    theme_aware = True
    use_config_theme = True
    auto_wire_textboxes = False
    text_select_all_on_click = False
    text_select_all_on_focus = False

    def __init__(self, doc):
        self.doc = doc
        self._mapping_store = MappingStore()
        self._placement_engine = PlacementEngine(doc, DEFAULT_CATEGORY)
        self._suppress_mode_events = False
        self._combo_filter_suspended = False
        self._wired_combo_filter_keys = set()
        self._wired_combo_textbox_keys = set()
        self._current_mode = CSV_SOURCE_MODE
        self._source_result = None
        self._fixture_rows = []
        self._level_rows = []
        self._divisor_tooltip_timer = None

        theme_mode, accent_mode = load_theme_state_from_config(default_theme="light", default_accent="blue")
        xaml_path = os.path.abspath(os.path.join(THIS_DIR, "MappingWindow.XAML"))
        CEDWindowBase.__init__(
            self,
            xaml_source=xaml_path,
            theme_mode=theme_mode,
            accent_mode=accent_mode,
            theme_aware=True,
            use_config_theme=True,
            auto_wire_textboxes=False,
        )

        self._bind_controls()
        self._wire_events()
        self._initialize_host_data()
        self._initialize_ui_defaults()

    # -- setup ---------------------------------------------------------------
    def _bind_controls(self):
        self.CategoryCombo = self.FindName("CategoryCombo")
        self.SourceModeCombo = self.FindName("SourceModeCombo")
        self.ModeWarningText = self.FindName("ModeWarningText")
        self.SubHeaderText = self.FindName("SubHeaderText")
        self.LoadedCountText = self.FindName("LoadedCountText")

        self.CsvOptionsBorder = self.FindName("CsvOptionsBorder")
        self.CsvPathTextBox = self.FindName("CsvPathTextBox")
        self.BrowseCsvButton = self.FindName("BrowseCsvButton")
        self.LoadCsvButton = self.FindName("LoadCsvButton")
        self.CsvLevelCombo = self.FindName("CsvLevelCombo")
        self.DivisorTextBox = self.FindName("DivisorTextBox")

        self.LinkOptionsBorder = self.FindName("LinkOptionsBorder")
        self.LinkCombo = self.FindName("LinkCombo")
        self.LoadLinkButton = self.FindName("LoadLinkButton")
        self.LinkSummaryText = self.FindName("LinkSummaryText")

        self.FixtureGrid = self.FindName("FixtureGrid")
        self.BulkFixtureCombo = self.FindName("BulkFixtureCombo")
        self.ApplyBulkFixtureButton = self.FindName("ApplyBulkFixtureButton")
        self.ClearBulkFixtureButton = self.FindName("ClearBulkFixtureButton")
        self.FixtureMappingStatusText = self.FindName("FixtureMappingStatusText")

        self.LevelMappingSection = self.FindName("LevelMappingSection")
        self.LevelGrid = self.FindName("LevelGrid")

        self.SkipDuplicatesCheck = self.FindName("SkipDuplicatesCheck")
        self.ValidationText = self.FindName("ValidationText")
        self.AutoMapButton = self.FindName("AutoMapButton")
        self.CancelButton = self.FindName("CancelButton")
        self.PlaceButton = self.FindName("PlaceButton")

    def _wire_events(self):
        self.CategoryCombo.SelectionChanged += self._category_changed
        self.SourceModeCombo.SelectionChanged += self._source_mode_changed
        self.BrowseCsvButton.Click += self._browse_csv_clicked
        self.LoadCsvButton.Click += self._load_csv_clicked
        self.LoadLinkButton.Click += self._load_link_clicked
        self.ApplyBulkFixtureButton.Click += self._apply_bulk_fixture_clicked
        self.ClearBulkFixtureButton.Click += self._clear_bulk_fixture_clicked
        self.AutoMapButton.Click += self._auto_map_clicked
        self.PlaceButton.Click += self._place_clicked
        self.CancelButton.Click += self._cancel_clicked
        try:
            self.FixtureGrid.LostKeyboardFocus += self._fixture_grid_lost_keyboard_focus
        except Exception:
            pass
        if self.DivisorTextBox is not None:
            self.DivisorTextBox.MouseEnter += self._divisor_tooltip_mouse_enter
            self.DivisorTextBox.MouseLeave += self._divisor_tooltip_mouse_leave

    def _initialize_host_data(self):
        self.host_fixture_options, self.host_fixture_by_label = collect_host_fixture_options(self.doc, self._selected_category)
        self.host_fixture_labels = [option.label for option in self.host_fixture_options]
        self.host_level_options, self.host_level_by_label = collect_host_levels(self.doc)
        self.host_level_labels = [option.label for option in self.host_level_options]
        self.link_options = collect_link_options(self.doc)

        self._fixture_matcher = FixtureMatcher(self.host_fixture_options)
        self._level_matcher = LevelMatcher(self.host_level_options)

    def _initialize_ui_defaults(self):
        self.CategoryCombo.ItemsSource = [label for label, _cat in CATEGORY_OPTIONS]
        self.CategoryCombo.SelectedIndex = 0

        self.BulkFixtureCombo.ItemsSource = self.host_fixture_labels
        if self.host_fixture_labels:
            self.BulkFixtureCombo.SelectedIndex = 0

        self.CsvLevelCombo.ItemsSource = self.host_level_options
        if self.host_level_options:
            self.CsvLevelCombo.SelectedIndex = 0

        self.LinkCombo.ItemsSource = self.link_options
        self._select_default_link_option()

        self._suppress_mode_events = True
        self._select_mode_in_combo(CSV_SOURCE_MODE)
        self._suppress_mode_events = False
        self._set_mode(CSV_SOURCE_MODE, initializing=True)
        if self.ModeWarningText is not None:
            self.ModeWarningText.Visibility = Visibility.Collapsed
        self._set_divisor_tooltip(expanded=False)

        self._wire_combo_filter(self.BulkFixtureCombo)
        self._wire_combo_textbox_filter(self.BulkFixtureCombo)
        try:
            self.FixtureGrid.PreviewKeyUp += self._datagrid_preview_key_up
        except Exception:
            pass

        self.apply_textbox_interaction_modes(
            textbox_mode_map={"DivisorTextBox": TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK}
        )
        self._clear_loaded_source()

    def _select_default_link_option(self):
        if not self.link_options:
            return
        loaded_indices = [i for i, opt in enumerate(self.link_options) if bool(opt.is_loaded)]
        if loaded_indices:
            self.LinkCombo.SelectedIndex = loaded_indices[0]
        else:
            self.LinkCombo.SelectedIndex = 0

    def _set_divisor_tooltip(self, expanded):
        if self.DivisorTextBox is None:
            return
        if expanded:
            self.DivisorTextBox.ToolTip = (
                "Adjustment factor to convert units to feet.\n"
                "Example: if CSV coordinates are in feet, enter 1.\n"
                "If CSV coordinates are in inches, enter 12."
            )
        else:
            self.DivisorTextBox.ToolTip = "Adjustment factor to convert units to feet."

    def _divisor_tooltip_mouse_enter(self, sender, args):
        self._set_divisor_tooltip(expanded=False)
        try:
            if self._divisor_tooltip_timer is None:
                self._divisor_tooltip_timer = DispatcherTimer()
                self._divisor_tooltip_timer.Interval = TimeSpan.FromMilliseconds(900)
                self._divisor_tooltip_timer.Tick += self._divisor_tooltip_timer_tick
            self._divisor_tooltip_timer.Stop()
            self._divisor_tooltip_timer.Start()
        except Exception:
            pass

    def _divisor_tooltip_mouse_leave(self, sender, args):
        try:
            if self._divisor_tooltip_timer is not None:
                self._divisor_tooltip_timer.Stop()
        except Exception:
            pass
        self._set_divisor_tooltip(expanded=False)

    def _divisor_tooltip_timer_tick(self, sender, args):
        try:
            if self._divisor_tooltip_timer is not None:
                self._divisor_tooltip_timer.Stop()
        except Exception:
            pass
        self._set_divisor_tooltip(expanded=True)

    # -- combo filtering -----------------------------------------------------
    def _window_loaded(self, sender, args):
        return

    def _mapping_grid_loading_row(self, sender, args):
        return

    def _combo_filter_key(self, combo):
        try:
            return int(combo.GetHashCode())
        except Exception:
            return id(combo)

    def _find_ancestor_combo(self, element):
        current = element
        while current is not None:
            if isinstance(current, ComboBox):
                return current
            try:
                current = VisualTreeHelper.GetParent(current)
            except Exception:
                return None
        return None

    def _resolve_fixture_combo_from_event(self, args):
        source = getattr(args, "OriginalSource", None)
        if isinstance(source, ComboBox):
            return source
        if isinstance(source, TextBox):
            combo = self._combo_from_editable_textbox(source)
            if combo is not None:
                return combo
        return self._find_ancestor_combo(source)

    def _validate_fixture_combo_entry(self, combo):
        if combo is None:
            return
        row = getattr(combo, "DataContext", None)
        if not isinstance(row, FixtureMappingRow):
            return

        typed = _safe_text(getattr(combo, "Text", "")) or _safe_text(getattr(combo, "SelectedItem", ""))
        valid_options = set(list(row.target_options or self.host_fixture_labels or []))
        if typed and typed in valid_options:
            if row.target_label != typed:
                row.target_label = typed
            return

        row.target_label = ""
        try:
            combo.SelectedItem = None
        except Exception:
            pass
        try:
            combo.Text = ""
        except Exception:
            pass

    def _fixture_grid_lost_keyboard_focus(self, sender, args):
        combo = self._resolve_fixture_combo_from_event(args)
        if combo is None:
            return
        self._validate_fixture_combo_entry(combo)

    def _is_mapping_combo(self, combo):
        if combo is None:
            return False
        if combo is self.BulkFixtureCombo:
            return True
        data_context = getattr(combo, "DataContext", None)
        if isinstance(data_context, FixtureMappingRow):
            return True
        if isinstance(data_context, LevelMappingRow):
            return True
        return False

    def _datagrid_preview_key_up(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if source is None:
            return
        combo = None
        if isinstance(source, ComboBox):
            combo = source
        elif isinstance(source, TextBox):
            combo = self._combo_from_editable_textbox(source)
            if combo is None:
                combo = self._find_ancestor_combo(source)
        else:
            combo = self._find_ancestor_combo(source)
        if combo is None or not self._is_mapping_combo(combo):
            return
        self._wire_combo_filter(combo)
        self._wire_combo_textbox_filter(combo)
        textbox = self._editable_combo_textbox(combo)
        typed = _safe_text(getattr(textbox, "Text", "")) if textbox is not None else _safe_text(
            getattr(combo, "Text", "")
        )
        self._apply_combo_text_filter(combo, typed)

    def _wire_combo_filters_in_container(self, root):
        for combo in list(_visual_descendants(root, ComboBox) or []):
            self._wire_combo_filter(combo)

    def _wire_combo_filter(self, combo):
        if combo is None:
            return
        key = self._combo_filter_key(combo)
        if key in self._wired_combo_filter_keys:
            return
        self._wired_combo_filter_keys.add(key)
        try:
            combo.KeyUp += self._combo_filter_key_up
            combo.DropDownOpened += self._combo_filter_dropdown_opened
            combo.GotFocus += self._combo_filter_got_focus
            combo.LostFocus += self._combo_filter_lost_focus
            combo.SelectionChanged += self._combo_filter_selection_changed
        except Exception:
            pass

    def _editable_combo_textbox(self, combo):
        if combo is None:
            return None
        try:
            combo.ApplyTemplate()
        except Exception:
            pass
        try:
            template = getattr(combo, "Template", None)
            if template is None:
                raise Exception("Missing template")
            textbox = template.FindName("PART_EditableTextBox", combo)
            if isinstance(textbox, TextBox):
                return textbox
        except Exception:
            pass
        for textbox in list(_visual_descendants(combo, TextBox) or []):
            if isinstance(textbox, TextBox):
                return textbox
        return None

    def _wire_combo_textbox_filter(self, combo):
        textbox = self._editable_combo_textbox(combo)
        if textbox is None:
            return
        key = self._combo_filter_key(textbox)
        if key in self._wired_combo_textbox_keys:
            return
        self._wired_combo_textbox_keys.add(key)
        try:
            textbox.TextChanged += self._combo_textbox_text_changed
            textbox.LostKeyboardFocus += self._combo_textbox_lost_focus
        except Exception:
            pass

    def _base_options_for_combo(self, combo):
        if combo is self.BulkFixtureCombo:
            return list(self.host_fixture_labels or [])

        data_context = getattr(combo, "DataContext", None)
        if isinstance(data_context, FixtureMappingRow):
            return list(data_context.target_options or self.host_fixture_labels or [])
        if isinstance(data_context, LevelMappingRow):
            return list(data_context.target_options or self.host_level_labels or [])

        try:
            return list(combo.ItemsSource or [])
        except Exception:
            return []

    def _apply_combo_text_filter(self, combo, text_value):
        if combo is None:
            return
        selected_label = _safe_text(getattr(combo, "SelectedItem", None))
        typed = _safe_text(text_value)
        typed_lower = typed.lower()

        self._combo_filter_suspended = True
        try:
            if typed_lower:
                try:
                    combo.SelectedItem = None
                except Exception:
                    pass
                combo.Items.Filter = lambda item: typed_lower in _safe_text(item).lower()
            else:
                combo.Items.Filter = None
            try:
                combo.Items.Refresh()
            except Exception:
                pass
            if selected_label and (not typed_lower):
                combo.SelectedItem = selected_label
            if typed_lower and (not bool(combo.IsDropDownOpen)):
                combo.IsDropDownOpen = True
        except Exception:
            pass
        finally:
            self._combo_filter_suspended = False

    def _combo_from_editable_textbox(self, textbox):
        if textbox is None:
            return None
        for combo in list(_visual_descendants(self, ComboBox) or []):
            part = self._editable_combo_textbox(combo)
            if part is textbox:
                return combo
        return None

    def _restore_combo_options(self, combo):
        if combo is None:
            return
        selected_label = _safe_text(getattr(combo, "SelectedItem", None)) or _safe_text(
            getattr(combo, "Text", "")
        )
        self._combo_filter_suspended = True
        try:
            combo.Items.Filter = None
            try:
                combo.Items.Refresh()
            except Exception:
                pass
            if selected_label:
                combo.Text = selected_label
        except Exception:
            pass
        finally:
            self._combo_filter_suspended = False

    def _combo_textbox_bubbled_changed(self, sender, args):
        combo = sender if isinstance(sender, ComboBox) else None
        if combo is None or self._combo_filter_suspended:
            return
        if not self._is_mapping_combo(combo):
            return
        textbox = self._editable_combo_textbox(combo)
        typed = _safe_text(getattr(textbox, "Text", "")) if textbox is not None else _safe_text(
            getattr(combo, "Text", "")
        )
        self._apply_combo_text_filter(combo, typed)

    def _combo_filter_key_up(self, sender, args):
        combo = sender if isinstance(sender, ComboBox) else None
        if combo is None or self._combo_filter_suspended:
            return
        if not self._is_mapping_combo(combo):
            return
        textbox = self._editable_combo_textbox(combo)
        typed = _safe_text(getattr(textbox, "Text", "")) if textbox is not None else _safe_text(
            getattr(combo, "Text", "")
        )
        self._apply_combo_text_filter(combo, typed)

    def _combo_textbox_text_changed(self, sender, args):
        if self._combo_filter_suspended:
            return
        textbox = sender if isinstance(sender, TextBox) else None
        if textbox is None:
            return
        combo = self._combo_from_editable_textbox(textbox)
        if combo is None:
            return
        if not self._is_mapping_combo(combo):
            return
        self._apply_combo_text_filter(combo, getattr(textbox, "Text", ""))

    def _combo_filter_dropdown_opened(self, sender, args):
        combo = sender if isinstance(sender, ComboBox) else None
        if combo is None:
            return
        if not self._is_mapping_combo(combo):
            return
        self._wire_combo_filter(combo)
        self._wire_combo_textbox_filter(combo)
        textbox = self._editable_combo_textbox(combo)
        typed = _safe_text(getattr(textbox, "Text", "")) if textbox is not None else _safe_text(
            getattr(combo, "Text", "")
        )
        self._apply_combo_text_filter(combo, typed)

    def _combo_filter_got_focus(self, sender, args):
        combo = sender if isinstance(sender, ComboBox) else None
        if combo is None:
            return
        if not self._is_mapping_combo(combo):
            return
        self._wire_combo_filter(combo)
        self._wire_combo_textbox_filter(combo)

    def _combo_filter_lost_focus(self, sender, args):
        combo = sender if isinstance(sender, ComboBox) else None
        if combo is None:
            return
        if not self._is_mapping_combo(combo):
            return
        self._restore_combo_options(combo)

    def _combo_textbox_lost_focus(self, sender, args):
        textbox = sender if isinstance(sender, TextBox) else None
        if textbox is None:
            return
        combo = self._combo_from_editable_textbox(textbox)
        if combo is None:
            return
        if not self._is_mapping_combo(combo):
            return
        self._restore_combo_options(combo)

    def _combo_filter_selection_changed(self, sender, args):
        combo = sender if isinstance(sender, ComboBox) else None
        if combo is None or self._combo_filter_suspended:
            return
        if not self._is_mapping_combo(combo):
            return
        selected = _safe_text(getattr(combo, "SelectedItem", None))
        if selected:
            self._combo_filter_suspended = True
            try:
                combo.Items.Filter = None
                try:
                    combo.Items.Refresh()
                except Exception:
                    pass
                combo.Text = selected
                combo.IsDropDownOpen = False
            except Exception:
                pass
            finally:
                self._combo_filter_suspended = False

    # -- mode / UI state -----------------------------------------------------
    def _get_combo_mode(self):
        selected = self.SourceModeCombo.SelectedItem
        if isinstance(selected, ComboBoxItem):
            mode = _safe_text(getattr(selected, "Tag", ""))
        else:
            mode = ""
        mode = mode.lower()
        if mode not in (CSV_SOURCE_MODE, LINK_SOURCE_MODE):
            return CSV_SOURCE_MODE
        return mode

    def _select_mode_in_combo(self, mode):
        target = _safe_text(mode).lower()
        count = int(self.SourceModeCombo.Items.Count)
        for index in range(count):
            item = self.SourceModeCombo.Items[index]
            if not isinstance(item, ComboBoxItem):
                continue
            item_mode = _safe_text(getattr(item, "Tag", "")).lower()
            if item_mode == target:
                self.SourceModeCombo.SelectedIndex = index
                return

    @property
    def _selected_category(self):
        idx = self.CategoryCombo.SelectedIndex
        if idx < 0 or idx >= len(CATEGORY_OPTIONS):
            return DEFAULT_CATEGORY
        return CATEGORY_OPTIONS[idx][1]

    def _category_changed(self, sender, args):
        if self._suppress_mode_events:
            return
        self._refresh_host_fixture_data()

    def _refresh_host_fixture_data(self):
        category = self._selected_category
        self._placement_engine.category = category
        self.host_fixture_options, self.host_fixture_by_label = collect_host_fixture_options(self.doc, category)
        self.host_fixture_labels = [option.label for option in self.host_fixture_options]
        self._fixture_matcher = FixtureMatcher(self.host_fixture_options)
        self.BulkFixtureCombo.ItemsSource = self.host_fixture_labels
        if self.host_fixture_labels:
            self.BulkFixtureCombo.SelectedIndex = 0
        for row in self._fixture_rows:
            row.target_options = list(self.host_fixture_labels)
            row.target_label = ""
        try:
            self.FixtureGrid.Items.Refresh()
        except Exception:
            pass
        self._refresh_status_texts()

    def _source_mode_changed(self, sender, args):
        if self._suppress_mode_events:
            return
        requested_mode = self._get_combo_mode()
        self._set_mode(requested_mode, initializing=False)

    def _has_loaded_work(self):
        return bool(self._source_result is not None and len(self._fixture_rows) > 0)

    def _set_mode(self, mode, initializing=False):
        mode = _safe_text(mode).lower()
        if mode not in (CSV_SOURCE_MODE, LINK_SOURCE_MODE):
            mode = CSV_SOURCE_MODE

        if not initializing and mode != self._current_mode and self._has_loaded_work():
            message = (
                "Switching data source will clear loaded fixtures and mappings.\n\n"
                "Do you want to continue?"
            )
            decision = MessageBox.Show(
                message,
                TITLE,
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
            )
            if decision != MessageBoxResult.Yes:
                self._suppress_mode_events = True
                self._select_mode_in_combo(self._current_mode)
                self._suppress_mode_events = False
                return
            self._clear_loaded_source()

        self._current_mode = mode
        self.CsvOptionsBorder.Visibility = (
            Visibility.Visible if mode == CSV_SOURCE_MODE else Visibility.Collapsed
        )
        self.LinkOptionsBorder.Visibility = (
            Visibility.Visible if mode == LINK_SOURCE_MODE else Visibility.Collapsed
        )
        self.LevelMappingSection.Visibility = (
            Visibility.Visible if mode == LINK_SOURCE_MODE else Visibility.Collapsed
        )

        if mode == CSV_SOURCE_MODE:
            self.SubHeaderText.Text = (
                "CSV mode: map CAD block names to host fixture types, then place on one selected level."
            )
        else:
            self.SubHeaderText.Text = (
                "Link mode: collect linked fixtures, map family/types and levels, then place in host model."
            )
        self._refresh_status_texts()

    def _clear_loaded_source(self):
        self._source_result = None
        self._fixture_rows = []
        self._level_rows = []
        self.FixtureGrid.ItemsSource = self._fixture_rows
        self.LevelGrid.ItemsSource = self._level_rows
        self.LinkSummaryText.Text = "Select a loaded link and click Load Link Fixtures."
        self.LoadedCountText.Text = "No source loaded"
        if self.FixtureMappingStatusText is not None:
            self.FixtureMappingStatusText.Text = "No fixture source loaded."
        self._refresh_status_texts()

    def _on_mapping_target_changed(self, row):
        self._refresh_status_texts()

    def _count_valid_fixture_mapped_rows(self):
        valid_labels = set(self.host_fixture_labels or [])
        count = 0
        for row in list(self._fixture_rows or []):
            label = _safe_text(getattr(row, "target_label", ""))
            if label and label in valid_labels:
                count += 1
        return int(count)

    def _count_valid_level_mapped_rows(self):
        valid_labels = set(self.host_level_labels or [])
        count = 0
        for row in list(self._level_rows or []):
            label = _safe_text(getattr(row, "target_label", ""))
            if label and label in valid_labels:
                count += 1
        return int(count)

    def _refresh_status_texts(self):
        if self.ModeWarningText is not None:
            self.ModeWarningText.Visibility = (
                Visibility.Visible if self._has_loaded_work() else Visibility.Collapsed
            )

        fixture_total = int(len(self._fixture_rows))
        fixture_mapped = self._count_valid_fixture_mapped_rows()
        level_total = int(len(self._level_rows))
        level_mapped = self._count_valid_level_mapped_rows()

        if fixture_total <= 0:
            if self.FixtureMappingStatusText is not None:
                self.FixtureMappingStatusText.Text = "No fixture source loaded."
            self.ValidationText.Text = "Load a source to begin."
            return

        if self.FixtureMappingStatusText is not None:
            self.FixtureMappingStatusText.Text = "Fixture mappings: {}/{} (unmapped rows will be skipped)".format(
                int(fixture_mapped), int(fixture_total)
            )

        if self._current_mode == LINK_SOURCE_MODE:
            self.ValidationText.Text = (
                "Fixture mappings: {0}/{1} (optional) | Level mappings: {2}/{3} (required)".format(
                    int(fixture_mapped),
                    int(fixture_total),
                    int(level_mapped),
                    int(level_total),
                )
            )
        else:
            self.ValidationText.Text = "Fixture mappings: {0}/{1} (optional; unmapped rows are skipped)".format(
                int(fixture_mapped), int(fixture_total)
            )

    # -- source loading ------------------------------------------------------
    def _browse_csv_clicked(self, sender, args):
        csv_path = forms.pick_file(file_ext="csv", title="Select Fixture CSV")
        if not csv_path:
            return
        self.CsvPathTextBox.Text = _safe_text(csv_path)
        self._load_csv_source()

    def _load_csv_clicked(self, sender, args):
        self._load_csv_source()

    def _load_csv_source(self):
        csv_path = _safe_text(self.CsvPathTextBox.Text)
        if not csv_path:
            forms.alert("Select a CSV file first.", title=TITLE)
            return
        if not os.path.exists(csv_path):
            forms.alert("CSV file not found:\n{}".format(csv_path), title=TITLE)
            return

        source = CsvLightSource(csv_path)
        result = source.collect()
        self._apply_source_result(result)
        self._restore_saved_mappings()
        self._apply_auto_match(overwrite_existing=False)

    def _load_link_clicked(self, sender, args):
        self._load_link_source()

    def _load_link_source(self):
        selected = self.LinkCombo.SelectedItem
        if selected is None:
            forms.alert("Select a Revit link instance first.", title=TITLE)
            return
        if not getattr(selected, "is_loaded", False):
            forms.alert("Selected link is unloaded. Load/reload the link and try again.", title=TITLE)
            return

        source = RevitLinkLightSource(self.doc, selected.link_instance, self._selected_category)
        result = source.collect()
        self._apply_source_result(result)
        self._restore_saved_mappings()
        self._apply_auto_match(overwrite_existing=False)

    def _apply_source_result(self, result):
        self._source_result = result
        self._fixture_rows = []
        self._level_rows = []

        if result is None:
            self._clear_loaded_source()
            return

        for group in list(result.fixture_groups or []):
            source_label = group.source_id
            if self._current_mode == LINK_SOURCE_MODE:
                source_label = "{} : {}".format(group.family_name, group.type_name)
            row = FixtureMappingRow(
                source_id=group.source_id,
                source_label=source_label,
                count=group.count,
                target_options=self.host_fixture_labels,
                source_family_name=group.family_name,
                source_type_name=group.type_name,
                source_type_params=getattr(group, "source_type_params", {}),
            )
            row.on_target_changed = self._on_mapping_target_changed
            self._fixture_rows.append(row)

        for level_info in list(result.levels or []):
            level_row = LevelMappingRow(
                source_level_key=level_info.level_key,
                source_level_name=level_info.level_name,
                source_elevation=level_info.elevation_host,
                source_count=level_info.count,
                target_options=self.host_level_labels,
            )
            level_row.on_target_changed = self._on_mapping_target_changed
            self._level_rows.append(level_row)

        self.FixtureGrid.ItemsSource = self._fixture_rows
        self.LevelGrid.ItemsSource = self._level_rows
        self.LevelMappingSection.Visibility = (
            Visibility.Visible if self._current_mode == LINK_SOURCE_MODE else Visibility.Collapsed
        )

        fixture_count = int(len(list(result.placements or [])))
        type_count = int(len(list(result.fixture_groups or [])))
        self.LoadedCountText.Text = "{} fixtures | {} source types".format(fixture_count, type_count)

        if self._current_mode == LINK_SOURCE_MODE:
            self.LinkSummaryText.Text = "Loaded {} fixtures from link; {} source levels found.".format(
                int(fixture_count), int(len(self._level_rows))
            )
            no_level_row = None
            for row in self._level_rows:
                if _safe_text(row.source_level_key) == RevitLinkLightSource.UNASSIGNED_LEVEL_KEY:
                    no_level_row = row
                    break
            if no_level_row is not None and int(no_level_row.count) > 0:
                self.LinkSummaryText.Text = (
                    "{} {} fixture(s) are tagged as {}.".format(
                        self.LinkSummaryText.Text,
                        int(no_level_row.count),
                        RevitLinkLightSource.UNASSIGNED_LEVEL_NAME,
                    )
                )
        else:
            self.LinkSummaryText.Text = "CSV source loaded."

        if result.warnings:
            LOGGER.warning(
                "{}: {} warnings while reading source.".format(
                    TITLE, int(len(list(result.warnings or [])))
                )
            )
            for warning in list(result.warnings or [])[:15]:
                LOGGER.warning(" - %s", warning)

        self._refresh_status_texts()

    # -- mapping actions -----------------------------------------------------
    def _apply_auto_match(self, overwrite_existing):
        if not self._fixture_rows:
            self._refresh_status_texts()
            return

        for row in self._fixture_rows:
            if overwrite_existing or not _safe_text(row.target_label):
                guess = self._fixture_matcher.match(row)
                if guess:
                    row.target_label = guess

        if self._current_mode == LINK_SOURCE_MODE:
            for row in self._level_rows:
                if overwrite_existing or not _safe_text(row.target_label):
                    guess = self._level_matcher.match(row)
                    if guess:
                        row.target_label = guess

        try:
            self.FixtureGrid.Items.Refresh()
        except Exception:
            pass
        try:
            self.LevelGrid.Items.Refresh()
        except Exception:
            pass
        self._refresh_status_texts()

    def _restore_saved_mappings(self):
        # Intentionally disabled: start each tool run with no persisted mappings.
        return

    def _auto_map_clicked(self, sender, args):
        if not self._fixture_rows:
            forms.alert("Load source data first.", title=TITLE)
            return
        dialog = AutoMatchWindow(
            fixture_rows=self._fixture_rows,
            host_fixture_options=self.host_fixture_options,
            source_mode=self._current_mode,
            theme_mode=getattr(self, "_theme_mode", "light"),
            accent_mode=getattr(self, "_accent_mode", "blue"),
        )
        dialog_result = dialog.ShowDialog()
        if not dialog_result:
            return

        applied = 0
        for row in self._fixture_rows:
            new_label = _safe_text(dialog.result_map.get(row.source_id, ""))
            if new_label and row.target_label != new_label:
                row.target_label = new_label
                applied += 1
        try:
            self.FixtureGrid.Items.Refresh()
        except Exception:
            pass
        self._refresh_status_texts()
        if applied > 0:
            self.ValidationText.Text = "{} | Auto-match applied {} mapping updates.".format(
                _safe_text(getattr(self.ValidationText, "Text", "")),
                int(applied),
            )

    def _apply_bulk_fixture_clicked(self, sender, args):
        selected_label = _safe_text(self.BulkFixtureCombo.SelectedItem) or _safe_text(
            getattr(self.BulkFixtureCombo, "Text", "")
        )
        if not selected_label:
            forms.alert("Choose a target family/type in the bulk mapping combo.", title=TITLE)
            return
        if selected_label not in self.host_fixture_by_label:
            forms.alert(
                "Bulk target '{}' is not a loaded host fixture type.".format(selected_label),
                title=TITLE,
            )
            return

        selected_rows = list(getattr(self.FixtureGrid, "SelectedItems", []) or [])
        if not selected_rows:
            forms.alert("Select one or more fixture rows first.", title=TITLE)
            return

        for row in selected_rows:
            if isinstance(row, FixtureMappingRow):
                row.target_label = selected_label
        try:
            self.FixtureGrid.Items.Refresh()
        except Exception:
            pass
        self._refresh_status_texts()

    def _clear_bulk_fixture_clicked(self, sender, args):
        selected_rows = list(getattr(self.FixtureGrid, "SelectedItems", []) or [])
        if not selected_rows:
            forms.alert("Select one or more fixture rows first.", title=TITLE)
            return
        for row in selected_rows:
            if isinstance(row, FixtureMappingRow):
                row.target_label = ""
        try:
            self.FixtureGrid.Items.Refresh()
        except Exception:
            pass
        self._refresh_status_texts()

    # -- placement -----------------------------------------------------------
    def _resolve_csv_divisor(self):
        text = _safe_text(getattr(self.DivisorTextBox, "Text", "12"))
        divisor = _to_float(text, 0.0)
        if abs(divisor) < 1e-9:
            return None
        return divisor

    def _build_fixture_symbol_map(self):
        fixture_map = {}
        unmapped = []
        invalid = []
        for row in self._fixture_rows:
            label = _safe_text(row.target_label)
            if not label:
                unmapped.append(row.source_label)
                continue
            option = self.host_fixture_by_label.get(label)
            if option is None or option.symbol is None:
                invalid.append("{} -> {}".format(row.source_label, label))
                continue
            fixture_map[row.source_id] = option.symbol
        return fixture_map, unmapped, invalid

    def _build_level_map(self):
        if self._current_mode != LINK_SOURCE_MODE:
            return {}, [], []
        level_map = {}
        unmapped = []
        invalid = []
        for row in self._level_rows:
            label = _safe_text(row.target_label)
            if not label:
                unmapped.append(row.source_level_name)
                continue
            option = self.host_level_by_label.get(label)
            if option is None or option.level is None:
                invalid.append("{} -> {}".format(row.source_level_name, label))
                continue
            level_map[row.source_level_key] = option.level
        return level_map, unmapped, invalid

    def _resolve_default_level(self):
        if self._current_mode == LINK_SOURCE_MODE:
            return None
        selected = self.CsvLevelCombo.SelectedItem
        if selected is None:
            return None
        if isinstance(selected, HostLevelOption):
            return selected.level
        return None

    def _validate_ready_to_place(self):
        if self._source_result is None or not self._fixture_rows:
            forms.alert("Load source data first.", title=TITLE)
            return None

        csv_divisor = 12.0
        if self._current_mode == CSV_SOURCE_MODE:
            divisor = self._resolve_csv_divisor()
            if divisor is None:
                forms.alert("Divisor must be a non-zero number.", title=TITLE)
                return None
            csv_divisor = divisor

        fixture_symbol_map, fixture_unmapped, fixture_invalid = self._build_fixture_symbol_map()
        if fixture_invalid:
            sample = "\n".join(["- {}".format(name) for name in fixture_invalid[:12]])
            forms.alert("Invalid fixture mappings found:\n\n{}".format(sample), title=TITLE)
            return None
        if fixture_unmapped:
            sample = "\n".join(["- {}".format(name) for name in fixture_unmapped[:12]])
            remaining = int(len(fixture_unmapped)) - min(12, int(len(fixture_unmapped)))
            if remaining > 0:
                sample = "{}\n- ...and {} more".format(sample, int(remaining))
            message = (
                "Some source fixture types are not mapped.\n\n"
                "{}\n\n"
                "Unmapped fixture rows will be skipped.\n"
                "Do you want to continue?"
            ).format(sample)
            decision = MessageBox.Show(
                message,
                TITLE,
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
            )
            if decision != MessageBoxResult.Yes:
                return None

        level_map = {}
        if self._current_mode == LINK_SOURCE_MODE:
            level_map, level_unmapped, level_invalid = self._build_level_map()
            if level_unmapped:
                sample = "\n".join(["- {}".format(name) for name in level_unmapped[:12]])
                forms.alert(
                    "Unmapped source levels found (map all before placing):\n\n{}".format(sample),
                    title=TITLE,
                )
                return None
            if level_invalid:
                sample = "\n".join(["- {}".format(name) for name in level_invalid[:12]])
                forms.alert("Invalid level mappings found:\n\n{}".format(sample), title=TITLE)
                return None

        default_level = self._resolve_default_level()
        if self._current_mode == CSV_SOURCE_MODE and default_level is None:
            forms.alert("Select a host level for CSV placement.", title=TITLE)
            return None

        return {
            "fixture_symbol_map": fixture_symbol_map,
            "level_map": level_map,
            "default_level": default_level,
            "csv_divisor": csv_divisor,
        }

    def _save_mapping_state(self):
        # Intentionally disabled: no config writes while mapping behavior is iterated.
        return

    def _place_clicked(self, sender, args):
        payload = self._validate_ready_to_place()
        if payload is None:
            return

        self._save_mapping_state()

        skip_duplicates = False
        report = self._placement_engine.place(
            placements=self._source_result.placements,
            fixture_symbol_by_source=payload["fixture_symbol_map"],
            default_level=payload["default_level"],
            level_by_source=payload["level_map"],
            csv_divisor=payload["csv_divisor"],
            skip_duplicates=skip_duplicates,
        )
        self._print_report(report)
        forms.alert(report.short_alert(), title=TITLE)

        if report.fatal_error:
            forms.alert("Placement failed:\n{}".format(report.fatal_error), title=TITLE)

    def _print_report(self, report):
        OUTPUT.print_md("## {} Summary".format(TITLE))
        for line in report.summary_lines():
            OUTPUT.print_md("- {}".format(line))
        if report.error_messages:
            OUTPUT.print_md("### First Placement Errors")
            for message in report.error_messages:
                OUTPUT.print_md("- {}".format(message))

    def _cancel_clicked(self, sender, args):
        self.Close()


# -----------------------------------------------------------------------------
# Entrypoint
def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active Revit document.", title=TITLE)
        return

    window = LightMappingWindow(doc)
    window.ShowDialog()


if __name__ == "__main__":
    main()
