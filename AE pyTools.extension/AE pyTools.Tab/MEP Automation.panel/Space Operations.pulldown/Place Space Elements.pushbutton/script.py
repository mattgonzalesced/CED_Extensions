# -*- coding: utf-8 -*-
"""
Place Space Elements
--------------------
Place saved space profile templates into classified spaces.
"""

import math
import os
import re
import sys
import uuid
from collections import OrderedDict

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    GroupType,
    HostObjectUtils,
    Line,
    LocationPoint,
    RevitLinkInstance,
    ShellLayerType,
    Wall,
    SpatialElementBoundaryLocation,
    SpatialElementBoundaryOptions,
    Transaction,
    XYZ,
)

try:
    from Autodesk.Revit.DB.Structure import StructuralType as RevitStructuralType  # type: ignore
except Exception:
    RevitStructuralType = None
output = script.get_output()
output.close_others()

TITLE = "Place Space Elements"
CLASSIFICATION_STORAGE_ID = "space_operations.classifications.v1"
KEY_TYPE_ELEMENTS = "space_type_elements"
KEY_SPACE_OVERRIDES = "space_overrides"

BUCKETS = [
    "Restrooms",
    "Offices",
    "Sales Floor",
    "Freezers",
    "Coolers",
    "Receiving",
    "Break",
    "Food Prep",
    "Utility",
    "Storage",
    "Other",
]

PLACEMENT_OPTIONS = [
    "Ceiling Corner Furthest from door",
    "One Foot off doorway wall",
    "Center of Furthest wall",
    "Center Ceiling",
    "Center Floor",
    "Center of Room",
    "Ceiling Corner Nearest Door",
]

DEFAULT_PLACEMENT_OPTION = "Center of Room"
CORNER_INSET_FT = 0.5
DOOR_WALL_OFFSET_FT = 1.0
DOOR_WALL_CLEARANCE_FT = 0.08
FLOOR_Z_OFFSET_FT = 0.1
CEILING_Z_OFFSET_FT = 0.1
HINGE_INTERIOR_NUDGE_FT = 0.15


def _resolve_lib_root():
    cursor = os.path.abspath(os.path.dirname(__file__))
    for _ in range(12):
        candidate = os.path.join(cursor, "CEDLib.lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if not parent or parent == cursor:
            break
        cursor = parent
    return None


LIB_ROOT = _resolve_lib_root()
if LIB_ROOT and LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

try:
    from ExtensibleStorage import ExtensibleStorage  # noqa: E402
except Exception:
    ExtensibleStorage = None


def _element_id_value(elem_id, default=""):
    if elem_id is None:
        return default
    for attr in ("IntegerValue", "Value"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return str(int(value))
        except Exception:
            try:
                return str(value)
            except Exception:
                continue
    return default


def _try_int(value, default=None):
    try:
        return int(str(value).strip())
    except Exception:
        return default


def _normalize_text(value):
    return " ".join(str(value or "").strip().lower().split())


def _normalize_bucket(value, default="Other"):
    needle = _normalize_text(value)
    if not needle:
        return default

    for bucket in BUCKETS:
        if needle == _normalize_text(bucket):
            return bucket

    if needle in ("salesfloor", "sales"):
        return "Sales Floor"
    if needle in ("foodprep", "prep"):
        return "Food Prep"

    return default

def _param_text(element, built_in_param):
    if element is None:
        return ""
    try:
        param = element.get_Parameter(built_in_param)
    except Exception:
        param = None
    if not param:
        return ""
    for getter_name in ("AsString", "AsValueString"):
        try:
            getter = getattr(param, getter_name)
            value = getter()
        except Exception:
            value = None
        if value is not None:
            text = str(value).strip()
            if text:
                return text
    return ""


def _space_name(space):
    name = _param_text(space, BuiltInParameter.ROOM_NAME)
    if name:
        return name
    try:
        value = getattr(space, "Name", None)
    except Exception:
        value = None
    return str(value).strip() if value else ""


def _space_number(space):
    return _param_text(space, BuiltInParameter.ROOM_NUMBER)


def _space_label(space):
    number = _space_number(space) or "<No Number>"
    name = _space_name(space) or "<Unnamed Space>"
    return "{} - {}".format(number, name)


def _space_key(space):
    sid = _element_id_value(getattr(space, "Id", None), default="")
    uid = ""
    try:
        uid = str(getattr(space, "UniqueId", "") or "").strip()
    except Exception:
        uid = ""
    return (uid or sid).strip()


def _collect_spaces(doc):
    spaces = []

    try:
        spaces = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_MEPSpaces)
            .WhereElementIsNotElementType()
        )
    except Exception:
        spaces = []

    if not spaces:
        try:
            spaces = list(
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
            )
        except Exception:
            spaces = []

    return spaces


def _build_host_level_lookup(doc):
    lookup = {}
    try:
        levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        levels = []

    for level in levels:
        try:
            name = str(getattr(level, "Name", "") or "").strip()
        except Exception:
            name = ""
        key = _normalize_text(name)
        if key and key not in lookup:
            lookup[key] = level

    return lookup


def _extract_first_number(text):
    value = str(text or "").strip()
    if not value:
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?", value)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def _numeric_host_levels(doc):
    try:
        levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        levels = []

    data = []
    for level in levels:
        try:
            name = str(getattr(level, "Name", "") or "").strip()
        except Exception:
            name = ""
        num = _extract_first_number(name)
        if num is None:
            continue
        try:
            elev = float(getattr(level, "Elevation", 0.0) or 0.0)
        except Exception:
            elev = 0.0
        data.append((level, num, elev))
    return data


def _fallback_host_level(host_doc, source_level=None):
    candidates = _numeric_host_levels(host_doc)
    if not candidates:
        return None

    src_num = None
    src_elev = None
    if source_level is not None:
        try:
            src_name = str(getattr(source_level, "Name", "") or "").strip()
        except Exception:
            src_name = ""
        src_num = _extract_first_number(src_name)
        try:
            src_elev = float(getattr(source_level, "Elevation", None))
        except Exception:
            src_elev = None

    best = None
    best_score = None
    for level, lvl_num, lvl_elev in candidates:
        num_delta = abs(lvl_num - src_num) if src_num is not None else 0.0
        elev_delta = abs(lvl_elev - src_elev) if src_elev is not None else 0.0
        # Prefer numeric-name match first, elevation second.
        score = (num_delta, elev_delta)
        if best is None or score < best_score:
            best = level
            best_score = score

    if best is not None:
        return best
    return candidates[0][0]


def _resolve_host_level(host_doc, source_doc, host_level_lookup, space):
    if host_doc is source_doc:
        try:
            level_id = getattr(space, "LevelId", None)
            if level_id is not None:
                return host_doc.GetElement(level_id)
        except Exception:
            pass

    try:
        level_id = getattr(space, "LevelId", None)
    except Exception:
        level_id = None

    source_level = None
    if level_id is not None:
        try:
            source_level = source_doc.GetElement(level_id)
        except Exception:
            source_level = None

    if source_level is not None:
        key = _normalize_text(getattr(source_level, "Name", "") or "")
        if key and key in host_level_lookup:
            return host_level_lookup.get(key)

    return _fallback_host_level(host_doc, source_level=source_level)


def _make_assignment_key(space_id, unique_id):
    return (unique_id or "").strip() or (space_id or "").strip()


def _lookup_assignment(assignments, space_id, unique_id, space_number=None, space_name=None, allow_space_id_direct=True):
    if not isinstance(assignments, dict):
        return None

    uid = (unique_id or "").strip()
    sid = (space_id or "").strip()

    if uid and uid in assignments and isinstance(assignments.get(uid), dict):
        return assignments.get(uid)
    if allow_space_id_direct and sid and sid in assignments and isinstance(assignments.get(sid), dict):
        return assignments.get(sid)

    values = [value for value in assignments.values() if isinstance(value, dict)]

    for value in values:
        entry_uid = str(value.get("unique_id") or "").strip()
        entry_sid = str(value.get("space_id") or "").strip()
        if uid and entry_uid and uid == entry_uid:
            return value
        if allow_space_id_direct and sid and entry_sid and sid == entry_sid:
            return value

    target_number = _normalize_text(space_number)
    target_name = _normalize_text(space_name)
    if not target_number and not target_name:
        return None

    exact_both = []
    number_matches = []
    name_matches = []

    for value in values:
        value_number = _normalize_text(value.get("space_number"))
        value_name = _normalize_text(value.get("space_name"))

        if target_number and value_number and target_number == value_number:
            number_matches.append(value)
        if target_name and value_name and target_name == value_name:
            name_matches.append(value)
        if target_number and target_name and value_number == target_number and value_name == target_name:
            exact_both.append(value)

    if exact_both:
        return exact_both[0]

    if target_number and not target_name and len(number_matches) == 1:
        return number_matches[0]
    if target_name and not target_number and len(name_matches) == 1:
        return name_matches[0]

    best = None
    best_score = 0
    for value in values:
        value_number = _normalize_text(value.get("space_number"))
        value_name = _normalize_text(value.get("space_name"))

        score = 0
        if target_number and value_number and target_number == value_number:
            score += 2
        if target_name and value_name and target_name == value_name:
            score += 1

        if score > best_score:
            best = value
            best_score = score
            if score >= 3:
                break

    return best


def _collect_classified_spaces(host_doc, source_doc, assignments, source_transform=None, source_label="Host Spaces"):
    allow_space_id_direct = host_doc is source_doc
    rows = []
    host_level_lookup = _build_host_level_lookup(host_doc)

    for space in _collect_spaces(source_doc):
        sid = _element_id_value(getattr(space, "Id", None), default="")
        uid = ""
        try:
            uid = str(getattr(space, "UniqueId", "") or "").strip()
        except Exception:
            uid = ""
        number = _space_number(space)
        name = _space_name(space) or "<Unnamed Space>"

        assignment = _lookup_assignment(assignments, sid, uid, number, name, allow_space_id_direct=allow_space_id_direct)
        bucket = "Other"
        if isinstance(assignment, dict):
            bucket = _normalize_bucket(assignment.get("bucket"), default="Other")

        space_key = _make_assignment_key(sid, uid)
        if isinstance(assignment, dict):
            assignment_key = _make_assignment_key(assignment.get("space_id"), assignment.get("unique_id"))
            if assignment_key and allow_space_id_direct:
                space_key = assignment_key

        rows.append(
            {
                "space": space,
                "space_id": sid,
                "unique_id": uid,
                "space_key": space_key,
                "space_number": number,
                "space_name": name,
                "bucket": bucket,
                "source_transform": source_transform,
                "source_label": source_label,
                "host_level": _resolve_host_level(host_doc, source_doc, host_level_lookup, space),
            }
        )

    rows.sort(key=lambda x: ((x.get("space_number") or "").lower(), (x.get("space_name") or "").lower()))
    return rows



def _transform_points_to_source(host_points, source_transform):
    points = []
    if not host_points:
        return points
    if source_transform is None:
        return list(host_points)

    inverse = None
    try:
        inverse = source_transform.Inverse
    except Exception:
        inverse = None

    if inverse is None:
        return points

    for pt in host_points:
        if pt is None:
            continue
        try:
            points.append(inverse.OfPoint(pt))
        except Exception:
            continue

    return points


def _transform_dir_to_source(host_dir_xy, source_transform):
    if host_dir_xy is None:
        return None

    try:
        hx, hy = _normalize_xy(float(host_dir_xy[0]), float(host_dir_xy[1]))
    except Exception:
        return None

    if abs(hx) <= 1e-9 and abs(hy) <= 1e-9:
        return None

    if source_transform is None:
        return hx, hy

    inverse = None
    try:
        inverse = source_transform.Inverse
    except Exception:
        inverse = None

    if inverse is None:
        return hx, hy

    try:
        p0 = inverse.OfPoint(XYZ(0.0, 0.0, 0.0))
        p1 = inverse.OfPoint(XYZ(hx, hy, 0.0))
        sx, sy = _normalize_xy(p1.X - p0.X, p1.Y - p0.Y)
        if abs(sx) > 1e-9 or abs(sy) > 1e-9:
            return sx, sy
    except Exception:
        pass

    return hx, hy


def _transform_door_rows_to_source(host_rows, source_transform):
    rows = []
    for row in host_rows or []:
        if not isinstance(row, dict):
            continue

        host_point = row.get("point")
        if host_point is None or _is_origin_point(host_point):
            continue

        source_points = _transform_points_to_source([host_point], source_transform)
        if not source_points:
            continue

        source_row = dict(row)
        source_row["point"] = source_points[0]

        for key in ("wall_dir", "hand_dir", "facing_dir", "opening_dir"):
            source_row[key] = _transform_dir_to_source(row.get(key), source_transform)

        rows.append(source_row)

    return rows
def _collect_loaded_links(doc):
    links = []
    try:
        instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
    except Exception:
        instances = []

    for link in instances:
        try:
            link_doc = link.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue

        name = str(getattr(link, "Name", "") or "").strip()
        if not name:
            name = str(getattr(link_doc, "Title", "") or "").strip()
        if not name:
            name = "<Linked Model>"

        links.append({"name": name, "link": link, "doc": link_doc})

    links.sort(key=lambda x: x.get("name", "").lower())
    return links



def _collect_all_space_sources(doc):
    sources = [
        {
            "source_doc": doc,
            "source_transform": None,
            "source_label": "Host Spaces",
        }
    ]

    for entry in _collect_loaded_links(doc):
        link_doc = entry.get("doc")
        if link_doc is None:
            continue

        transform = None
        try:
            transform = entry.get("link").GetTransform()
        except Exception:
            transform = None

        label = "Linked: {}".format(entry.get("name") or "<Linked Model>")
        sources.append(
            {
                "source_doc": link_doc,
                "source_transform": transform,
                "source_label": label,
            }
        )

    return sources
def _choose_space_source(doc):
    links = _collect_loaded_links(doc)

    options = ["Host Spaces"]
    if links:
        options.append("Linked Model Spaces")

    mode = forms.CommandSwitchWindow.show(options, message="Select space source for placement")
    if not mode:
        return None

    if mode == "Host Spaces":
        return {"source_doc": doc, "source_transform": None, "source_label": "Host Spaces"}

    names = [entry.get("name") for entry in links]
    chosen = forms.SelectFromList.show(
        names,
        title="Select Linked Model",
        button_name="Use Link",
        multiselect=False,
    )
    if not chosen:
        return None

    selected_name = chosen[0] if isinstance(chosen, list) else chosen
    selected = None
    for entry in links:
        if entry.get("name") == selected_name:
            selected = entry
            break
    if selected is None:
        return None

    transform = None
    try:
        transform = selected.get("link").GetTransform()
    except Exception:
        transform = None

    return {
        "source_doc": selected.get("doc"),
        "source_transform": transform,
        "source_label": "Linked: {}".format(selected.get("name")),
    }



def _as_dict(value):
    if isinstance(value, dict):
        return value

    keys = getattr(value, "Keys", None)
    if keys is not None:
        try:
            return {str(k): value[k] for k in list(keys)}
        except Exception:
            pass

    return {}


def _as_list(value):
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    if value is None:
        return []
    if isinstance(value, (str, bytes)):
        return []
    try:
        return list(value)
    except Exception:
        return []


def _sanitize_parameter_map(parameters):
    clean = OrderedDict()
    parameters = _as_dict(parameters)
    if not parameters:
        return clean

    for name, data in parameters.items():
        key = str(name or "").strip()
        if not key:
            continue

        data_map = _as_dict(data)
        if data_map:
            storage_type = str(data_map.get("storage_type") or "String")
            value = data_map.get("value")
            read_only = bool(data_map.get("read_only"))
        else:
            storage_type = "String"
            value = data
            read_only = False

        clean[key] = {
            "storage_type": storage_type,
            "value": "" if value is None else str(value),
            "read_only": read_only,
        }

    ordered = OrderedDict()
    for key in sorted(clean.keys(), key=lambda x: x.lower()):
        ordered[key] = clean[key]
    return ordered


def _normalize_entry_kind(raw_kind, entry_id=""):
    kind = str(raw_kind or "").strip().lower().replace("-", "_").replace(" ", "_")

    if kind in ("family_type", "family", "familytype", "symbol", "family_symbol"):
        return "family_type"
    if kind in ("model_group", "modelgroup", "group", "group_type"):
        return "model_group"

    prefix = str(entry_id or "").strip().lower()
    if prefix.startswith("family_type:") or prefix.startswith("family:"):
        return "family_type"
    if prefix.startswith("model_group:") or prefix.startswith("modelgroup:") or prefix.startswith("group:") or prefix.startswith("group_type:"):
        return "model_group"

    return ""


def _sanitize_template_entry(entry):
    entry = _as_dict(entry)
    if not entry:
        return None

    entry_id = str(entry.get("id") or entry.get("entry_id") or "").strip()
    kind = _normalize_entry_kind(entry.get("kind"), entry_id)

    element_type_id = str(
        entry.get("element_type_id")
        or entry.get("type_id")
        or entry.get("symbol_id")
        or entry.get("group_type_id")
        or ""
    ).strip()
    if not element_type_id:
        if ":" in entry_id:
            element_type_id = entry_id.split(":", 1)[-1]
        elif _try_int(entry_id, default=None) is not None:
            element_type_id = entry_id

    name = str(entry.get("name") or entry.get("display_name") or "").strip()
    if not kind:
        kind = _normalize_entry_kind("", entry_id)
        lowered_name = _normalize_text(name)
        if not kind and "model group" in lowered_name:
            kind = "model_group"
        if not kind and "family" in lowered_name:
            kind = "family_type"

    if kind not in ("family_type", "model_group"):
        return None

    if element_type_id and not entry_id:
        entry_id = "{}:{}".format(kind, element_type_id)
    elif element_type_id and ":" not in entry_id and _try_int(entry_id, default=None) is not None:
        entry_id = "{}:{}".format(kind, element_type_id)

    if not entry_id:
        return None

    if not name:
        name = "Family Type" if kind == "family_type" else "Model Group"

    placement_rule = str(entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION).strip()
    if placement_rule not in PLACEMENT_OPTIONS:
        placement_rule = DEFAULT_PLACEMENT_OPTION

    entry_uid = str(entry.get("entry_uid") or "").strip()
    if not entry_uid:
        entry_uid = uuid.uuid4().hex

    return {
        "id": entry_id,
        "entry_uid": entry_uid,
        "kind": kind,
        "element_type_id": element_type_id,
        "name": name,
        "placement_rule": placement_rule,
        "parameters": _sanitize_parameter_map(entry.get("parameters") or {}),
    }


def _sanitize_template_list(raw_list):
    map_payload = _as_dict(raw_list)
    source_entries = None
    if map_payload:
        source_entries = []
        for map_key, map_value in map_payload.items():
            row = _as_dict(map_value)
            if not row:
                continue
            if not str(row.get("id") or "").strip():
                row["id"] = str(map_key or "").strip()
            source_entries.append(row)
    else:
        source_entries = _as_list(raw_list)

    clean = []
    for raw in source_entries or []:
        entry = _sanitize_template_entry(raw)
        if not entry:
            continue
        clean.append(entry)
    return clean


def _sanitize_type_elements(raw_map):
    data = {bucket: [] for bucket in BUCKETS}
    raw_map = _as_dict(raw_map)
    if not raw_map:
        return data

    for raw_bucket, raw_entries in raw_map.items():
        bucket = _normalize_bucket(raw_bucket, default=None)
        if not bucket:
            continue
        data[bucket].extend(_sanitize_template_list(raw_entries))

    return data


def _sanitize_space_overrides(raw_map):
    data = {}
    raw_map = _as_dict(raw_map)
    if not raw_map:
        return data

    for space_key, entries in raw_map.items():
        key = str(space_key or "").strip()
        if not key:
            continue
        sanitized = _sanitize_template_list(entries or [])
        if sanitized:
            data[key] = sanitized
    return data


def _infer_bucket_from_space_row(space_row):
    name = str(space_row.get("space_name") or "")
    number = str(space_row.get("space_number") or "")
    text = _normalize_text("{} {}".format(number, name))
    if not text:
        return "Other"

    token_text = re.sub(r"[^a-z0-9]+", " ", text)
    token_text = " ".join(token_text.split())
    tokens = set(token_text.split())

    def contains_any(parts):
        for part in parts:
            if not part:
                continue
            needle = _normalize_text(part)
            if not needle:
                continue
            if needle in text or needle in token_text:
                return True
        return False

    def token_any(parts):
        for part in parts:
            if str(part or "").strip().lower() in tokens:
                return True
        return False

    if contains_any(["restroom", "bathroom", "toilet", "lavatory", "water closet", "mens", "men s", "womens", "women s", "wc", "rr"]) or token_any(["wc", "rr"]):
        return "Restrooms"
    if contains_any(["office", "admin", "desk"]):
        return "Offices"
    if contains_any(["sales floor", "selling", "sell"]):
        return "Sales Floor"
    if contains_any(["freezer"]):
        return "Freezers"
    if contains_any(["cooler"]):
        return "Coolers"
    if contains_any(["receiving", "dock"]):
        return "Receiving"
    if contains_any(["break", "lounge"]):
        return "Break"
    if contains_any(["prep", "food prep"]):
        return "Food Prep"
    if contains_any(["electrical", "mechanical", "mech", "elec", "utility"]):
        return "Utility"
    if contains_any(["storage", "janitor", "ware", "warehouse"]):
        return "Storage"

    return "Other"


def _resolve_template_bucket(space_row, type_elements):
    saved_bucket = _normalize_bucket(space_row.get("bucket"), default="Other")
    if type_elements.get(saved_bucket):
        return saved_bucket

    inferred_bucket = _infer_bucket_from_space_row(space_row)
    if inferred_bucket in BUCKETS and type_elements.get(inferred_bucket):
        return inferred_bucket

    if type_elements.get("Other"):
        return "Other"

    return saved_bucket


def _effective_entries(space_row, type_elements, space_overrides):
    bucket = _resolve_template_bucket(space_row, type_elements)
    type_entries = type_elements.get(bucket) or []
    override_entries = space_overrides.get(space_row.get("space_key")) or []

    combined = []
    for entry in type_entries:
        sanitized = _sanitize_template_entry(entry)
        if sanitized:
            combined.append(sanitized)
    for entry in override_entries:
        sanitized = _sanitize_template_entry(entry)
        if sanitized:
            combined.append(sanitized)

    return combined, bucket

def _distance_xy(a, b):
    if a is None or b is None:
        return float("inf")
    dx = float(a.X) - float(b.X)
    dy = float(a.Y) - float(b.Y)
    return math.sqrt(dx * dx + dy * dy)


def _format_xyz(point):
    if point is None:
        return "<none>"
    try:
        return "X={:.3f}, Y={:.3f}, Z={:.3f}".format(float(point.X), float(point.Y), float(point.Z))
    except Exception:
        return "<invalid>"


def _is_origin_point(point, tol=1e-3):
    if point is None:
        return False
    try:
        return abs(float(point.X)) <= tol and abs(float(point.Y)) <= tol and abs(float(point.Z)) <= tol
    except Exception:
        return False


def _clean_origin_points(points):
    data = [pt for pt in (points or []) if pt is not None]
    if not data:
        return []
    return [pt for pt in data if not _is_origin_point(pt)]


def _clean_origin_rows(rows):
    items = [row for row in (rows or []) if isinstance(row, dict) and row.get("point") is not None]
    if not items:
        return []
    return [row for row in items if not _is_origin_point(row.get("point"))]


def _element_location_point(element, fallback_point=None):
    if element is None:
        return fallback_point

    try:
        loc = getattr(element, "Location", None)
    except Exception:
        loc = None

    if isinstance(loc, LocationPoint):
        try:
            return loc.Point
        except Exception:
            pass

    try:
        point = getattr(loc, "Point", None)
        if point is not None:
            return point
    except Exception:
        pass

    return fallback_point


def _normalize_xy(dx, dy):
    length = math.sqrt(dx * dx + dy * dy)
    if length <= 1e-9:
        return 0.0, 0.0
    return dx / length, dy / length


def _space_center(space):
    loc = getattr(space, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point

    try:
        pt = getattr(loc, "Point", None)
    except Exception:
        pt = None
    if pt is not None:
        return pt

    try:
        bbox = space.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None

    try:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
    except Exception:
        return None

def _space_center_robust(space):
    center = _space_center(space)
    if center is not None and not _is_origin_point(center):
        return center

    corners, segment_mids = _boundary_geometry(space)
    if corners:
        robust = _average_xy(corners, corners[0])
        if robust is not None:
            return robust
    if segment_mids:
        robust = _average_xy(segment_mids, segment_mids[0])
        if robust is not None:
            return robust

    return center


def _door_rows_from_points(door_points, width_ft=3.0):
    rows = []
    for pt in (door_points or []):
        if pt is None or _is_origin_point(pt):
            continue
        rows.append({
            "point": pt,
            "width": float(width_ft),
            "wall_dir": None,
        })
    return rows

def _space_floor_ceil_center_z(space, center_pt):
    floor_z = center_pt.Z if center_pt else 0.0
    ceil_z = floor_z
    center_z = floor_z

    try:
        bbox = space.get_BoundingBox(None)
    except Exception:
        bbox = None

    if bbox is not None:
        try:
            floor_z = float(bbox.Min.Z)
            ceil_z = float(bbox.Max.Z)
            center_z = (floor_z + ceil_z) / 2.0
        except Exception:
            pass

    return floor_z, ceil_z, center_z


def _boundary_geometry(space):
    corners = []
    segment_mids = []

    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None

    if not loops:
        return corners, segment_mids

    for loop in loops:
        for segment in loop:
            try:
                curve = segment.GetCurve()
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
            except Exception:
                continue

            corners.append(p0)
            corners.append(p1)
            segment_mids.append(
                XYZ(
                    (p0.X + p1.X) / 2.0,
                    (p0.Y + p1.Y) / 2.0,
                    (p0.Z + p1.Z) / 2.0,
                )
            )

    unique = []
    for pt in corners:
        if all(_distance_xy(pt, existing) > 0.05 for existing in unique):
            unique.append(pt)

    return unique, segment_mids


def _average_xy(points, fallback_pt):
    if not points:
        return fallback_pt

    sx = 0.0
    sy = 0.0
    count = 0
    for pt in points:
        sx += float(pt.X)
        sy += float(pt.Y)
        count += 1

    if count <= 0:
        return fallback_pt

    z = fallback_pt.Z if fallback_pt else 0.0
    return XYZ(sx / float(count), sy / float(count), z)


def _pick_nearest(points, ref_pt):
    if not points:
        return None
    return min(points, key=lambda pt: _distance_xy(pt, ref_pt))


def _pick_farthest(points, ref_pt):
    if not points:
        return None
    return max(points, key=lambda pt: _distance_xy(pt, ref_pt))


def _collect_door_points(doc):
    points = []
    try:
        doors = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Doors)
            .WhereElementIsNotElementType()
        )
    except Exception:
        doors = []

    for door in doors:
        loc = getattr(door, "Location", None)
        if isinstance(loc, LocationPoint):
            pt = getattr(loc, "Point", None)
            if pt is not None:
                points.append(pt)
                continue

        try:
            bbox = door.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue

        try:
            pt = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
            points.append(pt)
        except Exception:
            pass

    return points



def _read_param_double(param):
    if param is None:
        return None
    try:
        return float(param.AsDouble())
    except Exception:
        return None


def _door_width_ft(door):
    if door is None:
        return 3.0

    width = None
    try:
        width = _read_param_double(door.get_Parameter(BuiltInParameter.DOOR_WIDTH))
    except Exception:
        width = None

    if width is None:
        try:
            symbol = getattr(door, "Symbol", None)
            width = _read_param_double(symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH)) if symbol is not None else None
        except Exception:
            width = None

    if width is None or width <= 0:
        return 3.0
    return float(width)


def _wall_direction_xy_from_wall(wall):
    if wall is None:
        return None
    try:
        loc = getattr(wall, "Location", None)
        curve = getattr(loc, "Curve", None)
        p0 = curve.GetEndPoint(0) if curve is not None else None
        p1 = curve.GetEndPoint(1) if curve is not None else None
    except Exception:
        p0 = None
        p1 = None

    if p0 is None or p1 is None:
        return None

    nx, ny = _normalize_xy(p1.X - p0.X, p1.Y - p0.Y)
    if abs(nx) <= 1e-9 and abs(ny) <= 1e-9:
        return None
    return nx, ny


def _door_orientation_xy(door, attr_name):
    if door is None:
        return None
    try:
        vec = getattr(door, attr_name, None)
    except Exception:
        vec = None
    if vec is None:
        return None

    try:
        nx, ny = _normalize_xy(float(vec.X), float(vec.Y))
    except Exception:
        return None

    if abs(nx) <= 1e-9 and abs(ny) <= 1e-9:
        return None
    return nx, ny


def _collect_door_rows(doc):
    rows = []
    try:
        doors = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Doors)
            .WhereElementIsNotElementType()
        )
    except Exception:
        doors = []

    for door in doors:
        point = None
        try:
            loc = getattr(door, "Location", None)
            if isinstance(loc, LocationPoint):
                point = loc.Point
            elif getattr(loc, "Point", None) is not None:
                point = loc.Point
        except Exception:
            point = None

        if point is None:
            try:
                bbox = door.get_BoundingBox(None)
                if bbox is not None:
                    point = XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0,
                    )
            except Exception:
                point = None

        if point is None:
            continue

        wall = None
        try:
            wall = getattr(door, "Host", None)
        except Exception:
            wall = None

        facing_dir = _door_orientation_xy(door, "FacingOrientation")
        hand_dir = _door_orientation_xy(door, "HandOrientation")
        wall_dir = _wall_direction_xy_from_wall(wall)

        opening_dir = hand_dir or wall_dir
        if opening_dir is None and facing_dir is not None:
            fx, fy = facing_dir
            opening_dir = (-fy, fx)

        rows.append(
            {
                "point": point,
                "width": _door_width_ft(door),
                "wall_id": _element_id_value(getattr(wall, "Id", None), default=""),
                "wall_dir": wall_dir,
                "hand_dir": hand_dir,
                "facing_dir": facing_dir,
                "opening_dir": opening_dir,
            }
        )

    return rows
def _nearest_door_point(ref_pt, door_points):
    if ref_pt is None or not door_points:
        return None
    return min(door_points, key=lambda pt: _distance_xy(pt, ref_pt))



def _nearest_door_row(ref_pt, door_rows):
    rows = [row for row in (door_rows or []) if isinstance(row, dict) and row.get("point") is not None]
    if not rows:
        return None
    if ref_pt is None:
        return rows[0]
    return min(rows, key=lambda row: _distance_xy(ref_pt, row.get("point")))


def _door_row_host_wall(doc, space, door_row):
    if doc is None or space is None or not isinstance(door_row, dict):
        return None

    try:
        space_doc = getattr(space, "Document", None)
    except Exception:
        space_doc = None
    if space_doc is not doc:
        return None

    wall_id = _try_int(door_row.get("wall_id"), default=None)
    if wall_id is None:
        return None

    try:
        wall = doc.GetElement(ElementId(wall_id))
    except Exception:
        wall = None
    if wall is None:
        return None
    if not _is_main_model_element(wall):
        return None
    return wall


def _door_row_source_wall(space, door_row):
    if space is None or not isinstance(door_row, dict):
        return None

    wall_id = _try_int(door_row.get("wall_id"), default=None)
    if wall_id is None:
        return None

    try:
        source_doc = getattr(space, "Document", None)
    except Exception:
        source_doc = None
    if source_doc is None:
        return None

    try:
        wall = source_doc.GetElement(ElementId(wall_id))
    except Exception:
        wall = None
    if wall is None:
        return None

    is_wall = isinstance(wall, Wall)
    if not is_wall:
        try:
            is_wall = "Wall" in str(wall.GetType().Name or "")
        except Exception:
            is_wall = False
    if not is_wall:
        return None

    return wall


def _door_row_project_point_to_wall(space, door_row, point):
    if point is None:
        return None
    wall = _door_row_source_wall(space, door_row)
    if wall is None:
        return point
    projected = _wall_curve_point_near(wall, point)
    if projected is None:
        return point
    return XYZ(projected.X, projected.Y, point.Z)
def _boundary_tangent_near_point(space, ref_pt):
    if space is None or ref_pt is None:
        return None

    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None

    if not loops:
        return None

    best_dir = None
    best_dist = float("inf")
    for loop in loops:
        for segment in loop:
            try:
                curve = segment.GetCurve()
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
            except Exception:
                continue

            nx, ny = _normalize_xy(p1.X - p0.X, p1.Y - p0.Y)
            if abs(nx) <= 1e-9 and abs(ny) <= 1e-9:
                continue

            mid = XYZ((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, (p0.Z + p1.Z) / 2.0)
            dist = _distance_xy(mid, ref_pt)
            if dist < best_dist:
                best_dist = dist
                best_dir = (nx, ny)

    return best_dir



def _boundary_segment_rows(space):
    rows = []
    if space is None:
        return rows

    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None

    if not loops:
        return rows

    for loop in loops:
        for segment in loop:
            try:
                curve = segment.GetCurve()
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
            except Exception:
                continue

            sx, sy = _normalize_xy(p1.X - p0.X, p1.Y - p0.Y)
            if abs(sx) <= 1e-9 and abs(sy) <= 1e-9:
                continue

            mid = XYZ((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, (p0.Z + p1.Z) / 2.0)
            rows.append(
                {
                    "curve": curve,
                    "p0": p0,
                    "p1": p1,
                    "dir": (sx, sy),
                    "mid": mid,
                }
            )

    return rows

def _one_foot_off_hinge_point(space, door_rows, fallback_center=None, return_row=False):
    rows = list(door_rows or [])
    if not rows:
        if return_row:
            return None, None
        return None

    segment_rows = _boundary_segment_rows(space)
    if not segment_rows:
        if return_row:
            return None, None
        return None

    def _is_inside_probe(pt):
        if pt is None:
            return False
        if _point_in_space(space, pt):
            return True
        for dx, dy in ((0.2, 0.0), (-0.2, 0.0), (0.0, 0.2), (0.0, -0.2)):
            probe = XYZ(pt.X + dx, pt.Y + dy, pt.Z)
            if _point_in_space(space, probe):
                return True
        return False

    def _project_to_curve(curve, ref_pt):
        if curve is None or ref_pt is None:
            return None
        try:
            ir = curve.Project(ref_pt)
            xyz = getattr(ir, "XYZPoint", None) if ir is not None else None
            if xyz is not None:
                return xyz
        except Exception:
            pass
        return None

    def _resolve_wall_dir(row, door_pt):
        wall_dir = row.get("wall_dir") or row.get("opening_dir") or row.get("hand_dir")
        if wall_dir is None:
            facing_dir = row.get("facing_dir")
            if facing_dir is not None:
                fx, fy = facing_dir
                wall_dir = (-fy, fx)
        if wall_dir is None:
            wall_dir = _boundary_tangent_near_point(space, door_pt)
        if wall_dir is None:
            return None

        wx, wy = wall_dir
        wx, wy = _normalize_xy(wx, wy)
        if abs(wx) <= 1e-9 and abs(wy) <= 1e-9:
            return None
        return wx, wy

    def _boundary_distance(pt):
        if pt is None:
            return float("inf")
        dist = float("inf")
        for seg in segment_rows:
            anchor = _project_to_curve(seg.get("curve"), pt)
            if anchor is None:
                continue
            d = _distance_xy(anchor, pt)
            if d < dist:
                dist = d
        return dist

    def _find_door_wall_segment(door_pt, wx, wy):
        best = None
        best_score = float("inf")
        for seg in segment_rows:
            sx, sy = seg.get("dir") or (0.0, 0.0)
            parallel = abs((sx * wx) + (sy * wy))
            if parallel < 0.85:
                continue

            anchor = _project_to_curve(seg.get("curve"), door_pt)
            if anchor is None:
                continue

            dist = _distance_xy(anchor, door_pt)
            score = (dist * 10.0) + ((1.0 - parallel) * 5.0)
            if score < best_score:
                best_score = score
                best = (seg, anchor)
        return best

    def _side_lengths(seg, anchor, ux, uy):
        p0 = seg.get("p0")
        p1 = seg.get("p1")
        if p0 is None or p1 is None or anchor is None:
            return 0.0, 0.0

        t0 = ((p0.X - anchor.X) * ux) + ((p0.Y - anchor.Y) * uy)
        t1 = ((p1.X - anchor.X) * ux) + ((p1.Y - anchor.Y) * uy)

        pos_len = max(0.0, t0, t1)
        neg_len = max(0.0, -t0, -t1)
        return pos_len, neg_len

    center_ref = fallback_center or _space_center_robust(space)

    candidate_rows = []
    for row in rows:
        door_pt = row.get("point")
        if door_pt is None or _is_origin_point(door_pt):
            continue

        inside = _is_inside_probe(door_pt)
        boundary_dist = _boundary_distance(door_pt)
        width = max(1.0, float(row.get("width") or 3.0))
        near_boundary = boundary_dist <= max(1.25, min(width, 4.0))

        if not inside and not near_boundary:
            continue

        center_dist = _distance_xy(door_pt, center_ref) if center_ref is not None else boundary_dist
        candidate_rows.append((row, inside, boundary_dist, center_dist))

    if not candidate_rows:
        for row in rows:
            door_pt = row.get("point")
            if door_pt is None or _is_origin_point(door_pt):
                continue
            boundary_dist = _boundary_distance(door_pt)
            center_dist = _distance_xy(door_pt, center_ref) if center_ref is not None else boundary_dist
            candidate_rows.append((row, False, boundary_dist, center_dist))

    candidate_rows.sort(key=lambda r: (0 if r[1] else 1, r[2], r[3]))

    best = None
    best_row = None
    best_score = float("inf")

    for row_rank, row_info in enumerate(candidate_rows[:12]):
        row, _inside, _bnd, _cd = row_info
        door_pt = row.get("point")
        if door_pt is None or _is_origin_point(door_pt):
            continue

        wall_dir = _resolve_wall_dir(row, door_pt)
        if wall_dir is None:
            continue
        wx, wy = wall_dir

        seg_match = _find_door_wall_segment(door_pt, wx, wy)
        if seg_match is None:
            continue

        seg, anchor = seg_match
        sx, sy = seg.get("dir") or (0.0, 0.0)
        sx, sy = _normalize_xy(sx, sy)
        if abs(sx) <= 1e-9 and abs(sy) <= 1e-9:
            continue

        # Align the segment axis with the door wall axis.
        if ((sx * wx) + (sy * wy)) < 0.0:
            ux, uy = -sx, -sy
        else:
            ux, uy = sx, sy

        half_width = max(0.5, min(float(row.get("width") or 3.0) / 2.0, 4.0))
        pos_len, neg_len = _side_lengths(seg, anchor, ux, uy)

        # Final rule: place on the same door wall on the longer side of the opening.
        side_sign = 1.0 if pos_len >= neg_len else -1.0
        side_x, side_y = (ux * side_sign, uy * side_sign)

        opening_side = XYZ(
            anchor.X + side_x * half_width,
            anchor.Y + side_y * half_width,
            anchor.Z,
        )

        target_on_wall = XYZ(
            opening_side.X + side_x * DOOR_WALL_OFFSET_FT,
            opening_side.Y + side_y * DOOR_WALL_OFFSET_FT,
            opening_side.Z,
        )

        n1 = (-uy, ux)
        n2 = (uy, -ux)
        probe1 = XYZ(
            target_on_wall.X + n1[0] * DOOR_WALL_CLEARANCE_FT,
            target_on_wall.Y + n1[1] * DOOR_WALL_CLEARANCE_FT,
            target_on_wall.Z,
        )
        probe2 = XYZ(
            target_on_wall.X + n2[0] * DOOR_WALL_CLEARANCE_FT,
            target_on_wall.Y + n2[1] * DOOR_WALL_CLEARANCE_FT,
            target_on_wall.Z,
        )

        if _is_inside_probe(probe1) and not _is_inside_probe(probe2):
            nx, ny = n1
        elif _is_inside_probe(probe2) and not _is_inside_probe(probe1):
            nx, ny = n2
        else:
            center = center_ref or _space_center_robust(space)
            if center is not None:
                c1 = (n1[0] * (center.X - target_on_wall.X)) + (n1[1] * (center.Y - target_on_wall.Y))
                c2 = (n2[0] * (center.X - target_on_wall.X)) + (n2[1] * (center.Y - target_on_wall.Y))
                nx, ny = n1 if c1 >= c2 else n2
            else:
                nx, ny = n1

        candidates = [
            XYZ(
                target_on_wall.X + nx * DOOR_WALL_CLEARANCE_FT,
                target_on_wall.Y + ny * DOOR_WALL_CLEARANCE_FT,
                target_on_wall.Z,
            ),
            XYZ(
                target_on_wall.X + nx * HINGE_INTERIOR_NUDGE_FT,
                target_on_wall.Y + ny * HINGE_INTERIOR_NUDGE_FT,
                target_on_wall.Z,
            ),
        ]

        for c_index, candidate in enumerate(candidates):
            inside = _is_inside_probe(candidate)
            score = 0.0
            if not inside:
                score += 100.0

            score += abs(_distance_xy(candidate, opening_side) - DOOR_WALL_OFFSET_FT) * 10.0
            score += float(c_index)
            score += float(row_rank) * 200.0

            if score < best_score:
                best_score = score
                best = candidate
                best_row = row

    if best is not None and not _is_origin_point(best):
        if return_row:
            return best, best_row
        return best

    if return_row:
        return None, None
    return None
def _point_in_space(space, point):
    if space is None or point is None:
        return False

    methods = ["IsPointInSpace", "IsPointInRoom"]
    probes = [point]
    # Door location is often on the boundary; test slight XY offsets too.
    for dx, dy in ((0.25, 0.0), (-0.25, 0.0), (0.0, 0.25), (0.0, -0.25)):
        probes.append(XYZ(point.X + dx, point.Y + dy, point.Z))

    for method_name in methods:
        checker = getattr(space, method_name, None)
        if checker is None:
            continue
        for probe in probes:
            try:
                if bool(checker(probe)):
                    return True
            except Exception:
                continue

    return False


def _door_point_for_space(space, door_points, fallback_point=None):
    points = list(door_points or [])
    if not points:
        return None

    contained = []
    for pt in points:
        if _point_in_space(space, pt):
            contained.append(pt)

    if contained:
        ref = fallback_point or _space_center(space) or contained[0]
        return _nearest_door_point(ref, contained) or contained[0]

    ref = fallback_point or _space_center(space)
    if ref is not None:
        return _nearest_door_point(ref, points)

    return points[0]


def _apply_corner_inset(corner_pt, center_xy):
    if corner_pt is None:
        return None
    if center_xy is None:
        return corner_pt

    nx, ny = _normalize_xy(center_xy.X - corner_pt.X, center_xy.Y - corner_pt.Y)
    return XYZ(corner_pt.X + nx * CORNER_INSET_FT, corner_pt.Y + ny * CORNER_INSET_FT, corner_pt.Z)

def _compute_target_xy(rule, center_xy, corners, segment_mids, door_pt):
    target = center_xy

    if rule == "Center of Furthest wall":
        if segment_mids:
            if door_pt is not None:
                target = _pick_farthest(segment_mids, door_pt)
            else:
                # This mode is defined relative to a doorway.
                # If no door is available, keep center as fallback.
                target = center_xy

    elif rule == "One Foot off doorway wall":
        anchor = None
        if segment_mids:
            if door_pt is not None:
                anchor = _pick_nearest(segment_mids, door_pt)
            else:
                anchor = _pick_nearest(segment_mids, center_xy)

        if anchor is not None:
            tx = anchor.X
            ty = anchor.Y
            if center_xy is not None:
                nx, ny = _normalize_xy(center_xy.X - anchor.X, center_xy.Y - anchor.Y)
                if abs(nx) <= 1e-9 and abs(ny) <= 1e-9 and door_pt is not None:
                    nx, ny = _normalize_xy(anchor.X - door_pt.X, anchor.Y - door_pt.Y)
                tx = anchor.X + nx * DOOR_WALL_OFFSET_FT
                ty = anchor.Y + ny * DOOR_WALL_OFFSET_FT
            target = XYZ(tx, ty, anchor.Z)

    elif rule == "Ceiling Corner Furthest from door":
        if corners:
            corner = _pick_farthest(corners, door_pt) if door_pt is not None else _pick_farthest(corners, center_xy)
            inset = _apply_corner_inset(corner, center_xy)
            if inset is not None:
                target = inset

    elif rule == "Ceiling Corner Nearest Door":
        if corners:
            corner = _pick_nearest(corners, door_pt) if door_pt is not None else _pick_nearest(corners, center_xy)
            inset = _apply_corner_inset(corner, center_xy)
            if inset is not None:
                target = inset

    return target or center_xy


def _compute_placement_point(space, placement_rule, door_points):
    rule = str(placement_rule or DEFAULT_PLACEMENT_OPTION).strip()
    if rule not in PLACEMENT_OPTIONS:
        rule = DEFAULT_PLACEMENT_OPTION

    center_pt = _space_center(space)
    corners, segment_mids = _boundary_geometry(space)

    # Some linked/phase-filtered spaces do not expose Location/BBox reliably.
    # Fall back to boundary-derived center when available.
    if center_pt is None:
        if corners:
            center_pt = _average_xy(corners, corners[0])
        elif segment_mids:
            center_pt = _average_xy(segment_mids, segment_mids[0])

    door_pt = _door_point_for_space(space, door_points, fallback_point=center_pt)

    # Last-resort center fallback: use nearest doorway point.
    if center_pt is None and door_pt is not None:
        center_pt = door_pt

    if center_pt is None:
        return None

    floor_z, ceil_z, center_z = _space_floor_ceil_center_z(space, center_pt)
    center_xy = _average_xy(corners, center_pt)

    xy = _compute_target_xy(rule, center_xy, corners, segment_mids, door_pt)

    # If boundary segments are unavailable, still support door-based placement intent.
    if xy is None and rule == "One Foot off doorway wall" and door_pt is not None:
        nx, ny = _normalize_xy(center_xy.X - door_pt.X, center_xy.Y - door_pt.Y)
        if abs(nx) <= 1e-9 and abs(ny) <= 1e-9:
            nx, ny = 1.0, 0.0
        xy = XYZ(
            door_pt.X + nx * DOOR_WALL_OFFSET_FT,
            door_pt.Y + ny * DOOR_WALL_OFFSET_FT,
            door_pt.Z,
        )

    if xy is None:
        xy = center_pt

    if rule == "Center Ceiling" or rule.startswith("Ceiling Corner"):
        z = ceil_z - CEILING_Z_OFFSET_FT
    elif rule in ("Center Floor", "One Foot off doorway wall", "Center of Furthest wall"):
        z = floor_z + FLOOR_Z_OFFSET_FT
    elif rule == "Center of Room":
        z = center_pt.Z
    else:
        z = center_z

    return XYZ(xy.X, xy.Y, z)

def _to_host_point(source_point, source_transform):
    if source_point is None:
        return None
    if source_transform is None:
        return source_point
    try:
        return source_transform.OfPoint(source_point)
    except Exception:
        return source_point




def _to_host_direction(source_dir_xy, source_transform):
    if source_dir_xy is None:
        return None

    try:
        sx, sy = _normalize_xy(float(source_dir_xy[0]), float(source_dir_xy[1]))
    except Exception:
        return None

    if abs(sx) <= 1e-9 and abs(sy) <= 1e-9:
        return None

    if source_transform is None:
        return sx, sy

    try:
        p0 = source_transform.OfPoint(XYZ(0.0, 0.0, 0.0))
        p1 = source_transform.OfPoint(XYZ(sx, sy, 0.0))
        hx, hy = _normalize_xy(p1.X - p0.X, p1.Y - p0.Y)
        if abs(hx) > 1e-9 or abs(hy) > 1e-9:
            return hx, hy
    except Exception:
        pass

    return sx, sy


def _door_parallel_dir_host(door_row, source_transform):
    if not isinstance(door_row, dict):
        return None

    parallel = door_row.get("wall_dir") or door_row.get("opening_dir") or door_row.get("hand_dir")
    if parallel is None:
        facing = door_row.get("facing_dir")
        if facing is not None:
            fx, fy = _normalize_xy(facing[0], facing[1])
            if abs(fx) > 1e-9 or abs(fy) > 1e-9:
                parallel = (-fy, fx)

    return _to_host_direction(parallel, source_transform)


def _rotate_instance_facing_perpendicular_to_door(instance, door_row, source_transform, center_host=None):
    if instance is None or not isinstance(door_row, dict):
        return False

    loc_point = _element_location_point(instance, fallback_point=None)
    if loc_point is None:
        return False

    parallel = _door_parallel_dir_host(door_row, source_transform)
    if parallel is None:
        return False
    ux, uy = parallel
    desired_a = (ux, uy)
    desired_b = (-ux, -uy)

    current_x = None
    current_y = None
    try:
        facing = getattr(instance, "FacingOrientation", None)
        if facing is not None:
            fx, fy = _normalize_xy(float(facing.X), float(facing.Y))
            if abs(fx) > 1e-9 or abs(fy) > 1e-9:
                current_x, current_y = fx, fy
    except Exception:
        current_x = None
        current_y = None

    if current_x is None or current_y is None:
        try:
            hand = getattr(instance, "HandOrientation", None)
            if hand is not None:
                hx, hy = _normalize_xy(float(hand.X), float(hand.Y))
                if abs(hx) > 1e-9 or abs(hy) > 1e-9:
                    current_x, current_y = hx, hy
        except Exception:
            current_x = None
            current_y = None

    if current_x is None or current_y is None:
        return False
    # Pick the perpendicular direction that requires the least rotation.
    dot_a = (current_x * desired_a[0]) + (current_y * desired_a[1])
    dot_b = (current_x * desired_b[0]) + (current_y * desired_b[1])
    if dot_b > dot_a:
        desired_x, desired_y = desired_b
    else:
        desired_x, desired_y = desired_a

    dot = max(-1.0, min(1.0, (current_x * desired_x) + (current_y * desired_y)))
    cross_z = (current_x * desired_y) - (current_y * desired_x)
    angle = math.atan2(cross_z, dot)

    if abs(angle) <= 1e-6:
        return True

    try:
        loc = getattr(instance, "Location", None)
        if not isinstance(loc, LocationPoint):
            return False
        axis = Line.CreateBound(
            XYZ(loc_point.X, loc_point.Y, loc_point.Z),
            XYZ(loc_point.X, loc_point.Y, loc_point.Z + 10.0),
        )
        return bool(loc.Rotate(axis, angle))
    except Exception:
        return False



def _get_family_symbol(doc, element_type_id, name_hint=None):
    int_id = _try_int(element_type_id, default=None)
    if int_id is not None:
        try:
            elem = doc.GetElement(ElementId(int_id))
        except Exception:
            elem = None
        if isinstance(elem, FamilySymbol):
            return elem

    # Fallback by name for imported configs or stale ids.
    hint = _normalize_text(name_hint)
    if not hint:
        return None

    try:
        symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).WhereElementIsElementType())
    except Exception:
        symbols = []

    for symbol in symbols:
        candidates = []
        try:
            fam = getattr(symbol, "Family", None)
            fam_name = str(getattr(fam, "Name", "") or "").strip()
            type_name = str(getattr(symbol, "Name", "") or "").strip()
            if fam_name and type_name:
                candidates.append("{} : {}".format(fam_name, type_name))
            if type_name:
                candidates.append(type_name)
        except Exception:
            pass

        for bip in (BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM, BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME):
            try:
                p = symbol.get_Parameter(bip)
                if p:
                    v = p.AsString() or p.AsValueString() or ""
                    if v:
                        candidates.append(str(v).strip())
            except Exception:
                pass

        for cand in candidates:
            if _normalize_text(cand) == hint:
                return symbol

    return None


def _get_group_type(doc, element_type_id, name_hint=None):
    int_id = _try_int(element_type_id, default=None)
    elem = None
    if int_id is not None:
        try:
            elem = doc.GetElement(ElementId(int_id))
        except Exception:
            elem = None

    if isinstance(elem, GroupType):
        return elem

    # Some Revit/IronPython combos can return a CLR type that fails direct isinstance checks.
    if elem is not None:
        try:
            clr_name = str(elem.GetType().Name or "")
        except Exception:
            clr_name = ""
        if "GroupType" in clr_name:
            return elem

    # Fallback: resolve by name for imported/cross-project configs where ids differ.
    raw_hint = str(name_hint or "").strip()
    if ":" in raw_hint and raw_hint.lower().startswith("model group"):
        raw_hint = raw_hint.split(":", 1)[1].strip()

    hint = _normalize_text(raw_hint)
    if not hint:
        return None

    try:
        model_group_cat = ElementId(BuiltInCategory.OST_IOSModelGroups).IntegerValue
    except Exception:
        model_group_cat = None

    try:
        group_types = list(FilteredElementCollector(doc).OfClass(GroupType).WhereElementIsElementType())
    except Exception:
        group_types = []

    for group_type in group_types:
        if model_group_cat is not None:
            try:
                cat = getattr(group_type, "Category", None)
                cat_id = getattr(getattr(cat, "Id", None), "IntegerValue", None)
                if cat_id != model_group_cat:
                    continue
            except Exception:
                continue

        candidates = []
        try:
            candidates.append(str(getattr(group_type, "Name", "") or "").strip())
        except Exception:
            pass

        for bip in (BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME):
            try:
                p = group_type.get_Parameter(bip)
                if p:
                    v = p.AsString() or p.AsValueString() or ""
                    candidates.append(str(v).strip())
            except Exception:
                pass

        for cand in candidates:
            if _normalize_text(cand) == hint:
                return group_type

    return None

def _space_level(doc, space):
    try:
        level_id = getattr(space, "LevelId", None)
        if level_id is None:
            return None
        return doc.GetElement(level_id)
    except Exception:
        return None


def _level_elevation(level):
    if level is None:
        return None
    try:
        elevation = getattr(level, "Elevation", None)
        if elevation is None:
            return None
        return float(elevation)
    except Exception:
        return None


def _is_numeric_level_name(level):
    if level is None:
        return False
    try:
        name = str(getattr(level, "Name", "") or "").strip()
    except Exception:
        name = ""
    if not name:
        return False
    return re.search(r"\d", name) is not None


def _is_main_model_element(element):
    if element is None:
        return False

    try:
        design_option = getattr(element, "DesignOption", None)
    except Exception:
        design_option = None

    # Main model elements are not in a design option.
    if design_option is None:
        return True

    try:
        option_id = getattr(design_option, "Id", None)
        value = getattr(option_id, "IntegerValue", None)
        if value is None:
            return False
        return int(value) < 0
    except Exception:
        return False


def _wall_curve_point_near(wall, ref_point):
    if wall is None or ref_point is None:
        return None

    try:
        loc = getattr(wall, "Location", None)
        curve = getattr(loc, "Curve", None)
    except Exception:
        curve = None

    if curve is None:
        return None

    try:
        ir = curve.Project(ref_point)
        if ir is not None:
            xyz = getattr(ir, "XYZPoint", None)
            if xyz is not None:
                return xyz
    except Exception:
        pass

    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        return None


def _space_boundary_walls(doc, space):
    walls = []
    seen = set()

    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None

    if not loops:
        return walls

    for loop in loops:
        for segment in loop:
            elem_id = None
            try:
                elem_id = segment.ElementId
            except Exception:
                elem_id = None
            if elem_id is None:
                continue

            sid = _element_id_value(elem_id, default="")
            if not sid or sid in seen:
                continue

            try:
                elem = doc.GetElement(elem_id)
            except Exception:
                elem = None
            if elem is None:
                continue

            is_wall = isinstance(elem, Wall)
            if not is_wall:
                try:
                    is_wall = "Wall" in str(elem.GetType().Name or "")
                except Exception:
                    is_wall = False
            if not is_wall:
                continue

            if not _is_main_model_element(elem):
                continue

            seen.add(sid)
            walls.append(elem)

    return walls


def _nearest_space_wall(doc, space, ref_point):
    best_wall = None
    best_point = None
    best_dist = float("inf")

    for wall in _space_boundary_walls(doc, space):
        probe = _wall_curve_point_near(wall, ref_point) or ref_point
        dist = _distance_xy(probe, ref_point)
        if dist < best_dist:
            best_dist = dist
            best_wall = wall
            best_point = probe

    return best_wall, best_point



def _nearest_host_wall(doc, ref_point):
    if doc is None or ref_point is None:
        return None, None

    try:
        walls = list(FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType())
    except Exception:
        walls = []

    best_wall = None
    best_point = None
    best_dist = float("inf")

    for wall in walls:
        if not _is_main_model_element(wall):
            continue

        probe = _wall_curve_point_near(wall, ref_point) or ref_point
        dist = _distance_xy(probe, ref_point)
        if dist < best_dist:
            best_dist = dist
            best_wall = wall
            best_point = probe

    return best_wall, best_point


def _linked_wall_face_reference(link_instance, wall, source_point):
    if link_instance is None or wall is None:
        return None

    side_refs = []
    for side in (ShellLayerType.Interior, ShellLayerType.Exterior):
        try:
            refs = HostObjectUtils.GetSideFaces(wall, side)
        except Exception:
            refs = None
        if refs:
            for ref in refs:
                side_refs.append(ref)

    if not side_refs:
        return None

    best_ref = None
    best_dist = float("inf")

    for face_ref in side_refs:
        dist = float("inf")
        try:
            face = wall.GetGeometryObjectFromReference(face_ref)
        except Exception:
            face = None

        if face is not None and source_point is not None:
            try:
                ir = face.Project(source_point)
            except Exception:
                ir = None

            if ir is not None:
                xyz = getattr(ir, "XYZPoint", None)
                if xyz is not None:
                    dist = _distance_xy(xyz, source_point)
                else:
                    try:
                        dist = float(getattr(ir, "Distance", None))
                    except Exception:
                        dist = float("inf")

        if dist < best_dist:
            best_dist = dist
            best_ref = face_ref

    if best_ref is None:
        best_ref = side_refs[0]

    try:
        return best_ref.CreateLinkReference(link_instance)
    except Exception:
        return None



def _exc_text(exc):
    if exc is None:
        return "<unknown>"
    try:
        text = str(exc)
    except Exception:
        text = ""
    text = " ".join(str(text or "").split())
    if not text:
        try:
            text = str(exc.__class__.__name__)
        except Exception:
            text = "<unknown>"
    if len(text) > 220:
        text = text[:217] + "..."
    return text
def _linked_wall_reference_direction(link_transform, wall):
    try:
        loc = getattr(wall, "Location", None)
        curve = getattr(loc, "Curve", None)
        p0 = curve.GetEndPoint(0) if curve is not None else None
        p1 = curve.GetEndPoint(1) if curve is not None else None
    except Exception:
        p0 = None
        p1 = None

    if p0 is not None and p1 is not None and link_transform is not None:
        try:
            hp0 = link_transform.OfPoint(p0)
            hp1 = link_transform.OfPoint(p1)
            nx, ny = _normalize_xy(hp1.X - hp0.X, hp1.Y - hp0.Y)
            if abs(nx) > 1e-9 or abs(ny) > 1e-9:
                return XYZ(nx, ny, 0.0)
        except Exception:
            pass

    return XYZ(0.0, 0.0, 1.0)


def _place_family_instance_on_linked_wall_face(
    doc,
    symbol,
    point,
    preferred_link_doc=None,
    preferred_wall_id=None,
    preferred_host_point=None,
    max_candidates=24,
    return_debug=False,
):
    if doc is None or symbol is None or point is None:
        if return_debug:
            return None, "Missing doc/symbol/point for linked-wall placement"
        return None

    candidate_rows = []
    errors = []
    wall_id_target = _try_int(preferred_wall_id, default=None)
    seed_host_point = preferred_host_point or point
    placement_type = str(getattr(symbol, "FamilyPlacementType", "<unknown>") or "<unknown>")

    try:
        link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance).WhereElementIsNotElementType())
    except Exception as exc:
        link_instances = []
        errors.append("Link collection failed: {}".format(_exc_text(exc)))

    for link_instance in link_instances:
        if not _is_main_model_element(link_instance):
            continue

        try:
            link_doc = link_instance.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue

        try:
            link_transform = link_instance.GetTotalTransform()
        except Exception:
            link_transform = None
        if link_transform is None:
            try:
                link_transform = link_instance.GetTransform()
            except Exception:
                link_transform = None
        if link_transform is None:
            continue

        try:
            inverse = link_transform.Inverse
            source_ref = inverse.OfPoint(seed_host_point)
        except Exception as exc:
            errors.append("Link transform inverse failed: {}".format(_exc_text(exc)))
            continue

        try:
            link_walls = list(FilteredElementCollector(link_doc).OfClass(Wall).WhereElementIsNotElementType())
        except Exception as exc:
            link_walls = []
            errors.append("Linked wall collection failed: {}".format(_exc_text(exc)))

        for link_wall in link_walls:
            if not _is_main_model_element(link_wall):
                continue

            source_probe = _wall_curve_point_near(link_wall, source_ref) or source_ref
            try:
                host_probe = link_transform.OfPoint(source_probe)
            except Exception:
                host_probe = None

            wall_id = _try_int(_element_id_value(getattr(link_wall, "Id", None), default=""), default=None)
            preferred_doc_hit = preferred_link_doc is not None and link_doc is preferred_link_doc
            preferred_wall_hit = preferred_doc_hit and wall_id_target is not None and wall_id == wall_id_target
            dist = _distance_xy(host_probe or point, seed_host_point)
            candidate_rows.append(
                {
                    "dist": dist,
                    "preferred_doc_hit": preferred_doc_hit,
                    "preferred_wall_hit": preferred_wall_hit,
                    "wall_id": wall_id,
                    "link_instance": link_instance,
                    "link_transform": link_transform,
                    "wall": link_wall,
                    "source_probe": source_probe,
                    "host_probe": host_probe,
                }
            )

    if not candidate_rows:
        reason = "No linked wall candidates resolved near target point"
        if errors:
            reason += "; " + errors[0]
        reason += " (family placement type: {})".format(placement_type)
        if return_debug:
            return None, reason
        return None

    preferred_hits = any(bool(row.get("preferred_wall_hit")) for row in candidate_rows)

    candidate_rows = sorted(
        candidate_rows,
        key=lambda row: (
            0 if row.get("preferred_wall_hit") else 1,
            0 if row.get("preferred_doc_hit") else 1,
            row.get("dist", float("inf")),
        ),
    )

    for row in candidate_rows[:max_candidates]:
        link_ref = _linked_wall_face_reference(
            row.get("link_instance"),
            row.get("wall"),
            row.get("source_probe"),
        )
        if link_ref is None:
            errors.append("Linked wall face reference unavailable for wall id {}".format(row.get("wall_id") or "?"))
            continue

        reference_dir = _linked_wall_reference_direction(row.get("link_transform"), row.get("wall"))
        host_probe = row.get("host_probe") or seed_host_point

        for candidate in (host_probe, seed_host_point, point):
            try:
                return (doc.Create.NewFamilyInstance(link_ref, candidate, reference_dir, symbol), None) if return_debug else doc.Create.NewFamilyInstance(link_ref, candidate, reference_dir, symbol)
            except Exception as exc:
                errors.append("Linked-face create failed: {}".format(_exc_text(exc)))
                try:
                    return (doc.Create.NewFamilyInstance(link_ref, candidate, XYZ(0.0, 0.0, 1.0), symbol), None) if return_debug else doc.Create.NewFamilyInstance(link_ref, candidate, XYZ(0.0, 0.0, 1.0), symbol)
                except Exception as exc2:
                    errors.append("Linked-face create (vertical dir) failed: {}".format(_exc_text(exc2)))

    reason = "Linked-wall face host failed"
    if preferred_link_doc is not None and wall_id_target is not None and not preferred_hits:
        reason += "; target linked wall id {} not found in loaded links".format(wall_id_target)
    if errors:
        reason += "; " + errors[0]
    reason += " (family placement type: {})".format(placement_type)

    if return_debug:
        return None, reason
    return None


def _place_family_instance_on_wall(
    doc,
    symbol,
    point,
    level,
    space,
    preferred_wall=None,
    preferred_link_doc=None,
    preferred_link_wall_id=None,
    preferred_link_host_point=None,
    return_debug=False,
):
    if symbol is None or point is None or space is None:
        if return_debug:
            return None, "Missing symbol/point/space for wall-hosted placement"
        return None

    if not getattr(symbol, "IsActive", True):
        try:
            symbol.Activate()
            doc.Regenerate()
        except Exception:
            pass

    errors = []
    placement_type = str(getattr(symbol, "FamilyPlacementType", "<unknown>") or "<unknown>")
    wall = None
    wall_point = None

    if preferred_wall is not None and _is_main_model_element(preferred_wall):
        wall = preferred_wall
        wall_point = _wall_curve_point_near(wall, point) or point

    space_doc = None
    try:
        space_doc = getattr(space, "Document", None)
    except Exception:
        space_doc = None

    # Only use space-boundary walls when the space is from host doc.
    if wall is None and space_doc is doc:
        wall, wall_point = _nearest_space_wall(doc, space, point)

    if wall is None:
        wall, wall_point = _nearest_host_wall(doc, point)

    linked_reason = None
    if wall is None:
        if preferred_link_doc is None and space_doc is not None and space_doc is not doc:
            preferred_link_doc = space_doc
        linked_result = _place_family_instance_on_linked_wall_face(
            doc,
            symbol,
            point,
            preferred_link_doc=preferred_link_doc,
            preferred_wall_id=preferred_link_wall_id,
            preferred_host_point=preferred_link_host_point,
            return_debug=True,
        )
        linked_elem, linked_reason = linked_result if isinstance(linked_result, tuple) else (linked_result, None)
        if linked_elem is not None:
            if return_debug:
                return linked_elem, "Placed on linked wall face"
            return linked_elem

        reason = "No valid host wall resolved"
        if linked_reason:
            reason += "; " + linked_reason
        reason += " (family placement type: {})".format(placement_type)
        if return_debug:
            return None, reason
        return None

    candidates = []
    if wall_point is not None:
        candidates.append(wall_point)
    candidates.append(point)

    for candidate in candidates:
        if candidate is None:
            continue

        if level is not None and RevitStructuralType is not None:
            try:
                created = doc.Create.NewFamilyInstance(candidate, symbol, wall, level, RevitStructuralType.NonStructural)
                return (created, "Placed on host wall") if return_debug else created
            except Exception as exc:
                errors.append("Host-wall create (level+struct) failed: {}".format(_exc_text(exc)))

        if RevitStructuralType is not None:
            try:
                created = doc.Create.NewFamilyInstance(candidate, symbol, wall, RevitStructuralType.NonStructural)
                return (created, "Placed on host wall") if return_debug else created
            except Exception as exc:
                errors.append("Host-wall create (struct) failed: {}".format(_exc_text(exc)))

        if level is not None:
            try:
                created = doc.Create.NewFamilyInstance(candidate, symbol, wall, level)
                return (created, "Placed on host wall") if return_debug else created
            except Exception as exc:
                errors.append("Host-wall create (level) failed: {}".format(_exc_text(exc)))

        try:
            created = doc.Create.NewFamilyInstance(candidate, symbol, wall)
            return (created, "Placed on host wall") if return_debug else created
        except Exception as exc:
            errors.append("Host-wall create failed: {}".format(_exc_text(exc)))

    # Host-wall creation failed; attempt linked wall face fallback as secondary path.
    if preferred_link_doc is None and space_doc is not None and space_doc is not doc:
        preferred_link_doc = space_doc

    linked_result = _place_family_instance_on_linked_wall_face(
        doc,
        symbol,
        point,
        preferred_link_doc=preferred_link_doc,
        preferred_wall_id=preferred_link_wall_id,
        preferred_host_point=preferred_link_host_point,
        return_debug=True,
    )
    linked_elem, linked_reason = linked_result if isinstance(linked_result, tuple) else (linked_result, None)
    if linked_elem is not None:
        if return_debug:
            return linked_elem, "Placed on linked wall face after host-wall failure"
        return linked_elem

    reason = "Host-wall placement failed"
    if errors:
        reason += "; " + errors[0]
    if linked_reason:
        reason += " | " + linked_reason
    reason += " (family placement type: {})".format(placement_type)

    if return_debug:
        return None, reason
    return None


def _place_family_instance(doc, symbol, point, level, space=None):
    if symbol is None or point is None:
        return None

    if not getattr(symbol, "IsActive", True):
        try:
            symbol.Activate()
            doc.Regenerate()
        except Exception:
            pass

    if level is not None and RevitStructuralType is not None:
        try:
            return doc.Create.NewFamilyInstance(point, symbol, level, RevitStructuralType.NonStructural)
        except Exception:
            pass

    if RevitStructuralType is not None:
        try:
            return doc.Create.NewFamilyInstance(point, symbol, RevitStructuralType.NonStructural)
        except Exception:
            pass

    if level is not None:
        try:
            return doc.Create.NewFamilyInstance(point, symbol, level)
        except Exception:
            pass

    try:
        return doc.Create.NewFamilyInstance(point, symbol)
    except Exception:
        pass

    # Retry for wall-hosted families (e.g., many light switches).
    return _place_family_instance_on_wall(doc, symbol, point, level, space)

def _try_set_instance_elevation(instance, target_z):
    if instance is None or target_z is None:
        return

    loc = getattr(instance, "Location", None)
    if isinstance(loc, LocationPoint):
        try:
            pt = loc.Point
            loc.Point = XYZ(pt.X, pt.Y, target_z)
            return
        except Exception:
            pass


def _set_parameter_value(param, storage_type_name, value_text):
    if param is None:
        return False

    try:
        if param.IsReadOnly:
            return False
    except Exception:
        return False

    storage = str(storage_type_name or "")
    text = "" if value_text is None else str(value_text)

    try:
        if "String" in storage:
            return bool(param.Set(text))

        if "Integer" in storage:
            try:
                return bool(param.Set(int(float(text))))
            except Exception:
                return bool(param.SetValueString(text))

        if "Double" in storage:
            try:
                return bool(param.Set(float(text)))
            except Exception:
                return bool(param.SetValueString(text))

        if "ElementId" in storage:
            int_id = _try_int(text, default=None)
            if int_id is None:
                return False
            return bool(param.Set(ElementId(int_id)))

        try:
            return bool(param.SetValueString(text))
        except Exception:
            return bool(param.Set(text))
    except Exception:
        return False


def _apply_parameter_overrides(element, parameters):
    result = {
        "set": 0,
        "missing": 0,
        "readonly": 0,
        "failed": 0,
    }

    for param_name, data in (parameters or {}).items():
        key = str(param_name or "").strip()
        if not key:
            continue

        try:
            param = element.LookupParameter(key)
        except Exception:
            param = None

        if not param:
            result["missing"] += 1
            continue

        try:
            if param.IsReadOnly:
                result["readonly"] += 1
                continue
        except Exception:
            result["readonly"] += 1
            continue

        storage_type = "String"
        value = ""
        if isinstance(data, dict):
            storage_type = str(data.get("storage_type") or "String")
            value = data.get("value")
        else:
            value = data

        if _set_parameter_value(param, storage_type, value):
            result["set"] += 1
        else:
            result["failed"] += 1

    return result



class _ProfileSelectionOption(forms.TemplateListItem):
    @property
    def name(self):
        data = self.item or {}
        kind = "Family Type" if str(data.get("kind") or "") == "family_type" else "Model Group"
        count = int(data.get("request_count") or 0)
        plural = "" if count == 1 else "s"
        base = "[{}] {} [{}] - {} request{}".format(
            data.get("bucket") or "Other",
            data.get("name") or data.get("id") or "<Profile>",
            kind,
            count,
            plural,
        )

        space_labels = list(data.get("space_labels") or [])
        if not space_labels:
            return base

        lines = [base, "    Spaces:"]
        for label in space_labels:
            lines.append("      - {}".format(label))
        return "\n".join(lines)


def _space_request_label(space_row):
    number = str(space_row.get("space_number") or "<No Number>").strip()
    name = str(space_row.get("space_name") or "<Unnamed Space>").strip()
    return "{} - {}".format(number, name)
def _request_profile_key(space_row, entry):
    bucket = _normalize_bucket(space_row.get("bucket"), default="Other")
    entry_uid = str(entry.get("entry_uid") or "").strip()
    if entry_uid:
        return "{}|{}".format(bucket, entry_uid)

    entry_id = str(entry.get("id") or "").strip()
    if not entry_id:
        return ""
    return "{}|{}".format(bucket, entry_id)


def _build_profile_selection_options(requests):
    stats = OrderedDict()
    for space_row, entry in requests or []:
        key = _request_profile_key(space_row, entry)
        if not key:
            continue

        if key not in stats:
            stats[key] = {
                "key": key,
                "bucket": space_row.get("bucket") or "Other",
                "id": entry.get("id") or "",
                "entry_uid": entry.get("entry_uid") or "",
                "name": entry.get("name") or entry.get("id") or "<Profile>",
                "kind": _normalize_entry_kind(entry.get("kind"), entry.get("id")) or "family_type",
                "request_count": 0,
                "space_labels": [],
            }

        stats[key]["request_count"] += 1

        label = _space_request_label(space_row)
        if label and label not in stats[key]["space_labels"]:
            stats[key]["space_labels"].append(label)

    ordered = sorted(
        stats.values(),
        key=lambda d: (
            str(d.get("bucket") or "").lower(),
            str(d.get("name") or "").lower(),
            str(d.get("id") or "").lower(),
        ),
    )

    return [_ProfileSelectionOption(item, checked=True) for item in ordered]
def _prompt_profile_selection(requests):
    options = _build_profile_selection_options(requests)
    if not options:
        return None

    selected = forms.SelectFromList.show(
        options,
        title="Select Space Profiles To Place",
        button_name="Place Selected",
        multiselect=True,
        return_all=True,
    )
    if selected is None:
        return None

    chosen_keys = set()
    for option in selected:
        try:
            checked = bool(option)
        except Exception:
            checked = False
        if not checked:
            continue
        key = str((option.item or {}).get("key") or "").strip()
        if key:
            chosen_keys.add(key)

    return chosen_keys


def _request_rows(spaces, type_elements, space_overrides):
    rows = []
    for space_row in spaces:
        entries, resolved_bucket = _effective_entries(space_row, type_elements, space_overrides)
        request_space_row = space_row

        original_bucket = _normalize_bucket(space_row.get("bucket"), default="Other")
        if resolved_bucket != original_bucket:
            request_space_row = dict(space_row)
            request_space_row["bucket"] = resolved_bucket
            request_space_row["resolved_from_saved_bucket"] = original_bucket

        for entry in entries:
            rows.append((request_space_row, entry))
    return rows

def _bucket_counts(spaces):
    counts = OrderedDict((bucket, 0) for bucket in BUCKETS)
    for row in spaces:
        bucket = _normalize_bucket(row.get("bucket"), default="Other")
        counts[bucket] += 1
    return counts



def _bucket_counts_for_source(spaces, source_label):
    filtered = [row for row in (spaces or []) if str(row.get("source_label") or "") == str(source_label or "")]
    return _bucket_counts(filtered)
def _template_counts_by_bucket(type_elements):
    counts = OrderedDict((bucket, 0) for bucket in BUCKETS)
    for bucket in BUCKETS:
        counts[bucket] = len(type_elements.get(bucket) or [])
    return counts


def _template_kind_counts(type_elements, space_overrides):
    counts = OrderedDict([("family_type", 0), ("model_group", 0)])

    for bucket in BUCKETS:
        for entry in type_elements.get(bucket) or []:
            kind = _normalize_entry_kind((entry or {}).get("kind"), (entry or {}).get("id"))
            if kind in counts:
                counts[kind] += 1

    for entries in space_overrides.values():
        for entry in entries or []:
            kind = _normalize_entry_kind((entry or {}).get("kind"), (entry or {}).get("id"))
            if kind in counts:
                counts[kind] += 1

    return counts


def _run_placement(doc, rows, door_points_by_source, door_rows_by_source):
    placed = 0
    failures = []
    placed_rows = []
    param_totals = {
        "set": 0,
        "missing": 0,
        "readonly": 0,
        "failed": 0,
    }

    tx = Transaction(doc, TITLE)
    tx.Start()
    try:
        for space_row, entry in rows:
            source_space = space_row.get("space")
            rule = entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION

            source_label = str(space_row.get("source_label") or "Host Spaces")
            source_door_points = []
            source_door_rows = []
            if isinstance(door_points_by_source, dict):
                source_door_points = door_points_by_source.get(source_label) or []
            else:
                source_door_points = door_points_by_source or []

            if isinstance(door_rows_by_source, dict):
                source_door_rows = door_rows_by_source.get(source_label) or []

            if not source_door_points and isinstance(door_points_by_source, dict):
                host_candidates = door_points_by_source.get("__host::" + source_label) or door_points_by_source.get("__all_host__") or []
                if host_candidates:
                    source_door_points = _transform_points_to_source(host_candidates, space_row.get("source_transform"))
                    source_door_points = _clean_origin_points(source_door_points)

            if not source_door_rows and source_door_points:
                source_door_rows = _door_rows_from_points(source_door_points)

            source_point = None
            source_door_row = None
            point = None

            if rule == "One Foot off doorway wall":
                if not source_door_rows and source_door_points:
                    source_door_rows = _door_rows_from_points(source_door_points)

                center_source = _space_center_robust(source_space)
                source_point, source_door_row = _one_foot_off_hinge_point(
                    source_space,
                    source_door_rows,
                    fallback_center=center_source,
                    return_row=True,
                )

                if source_door_row is None and source_point is not None and source_door_rows:
                    source_door_row = _nearest_door_row(source_point, source_door_rows)

                if source_point is None or _is_origin_point(source_point):
                    failures.append((space_row, entry, "No valid doorway-wall placement point could be resolved", source_point))
                    continue

                point = _to_host_point(source_point, space_row.get("source_transform"))
                if point is None or _is_origin_point(point):
                    failures.append((space_row, entry, "Doorway-wall point could not be transformed to valid host coordinates", point or source_point))
                    continue

            else:
                source_point = _compute_placement_point(source_space, rule, source_door_points)

                if source_point is None:
                    fallback_door = _door_point_for_space(source_space, source_door_points, fallback_point=None)
                    if fallback_door is None and source_door_points:
                        fallback_door = source_door_points[0]
                    if fallback_door is not None:
                        source_point = fallback_door
                    else:
                        source_point = _space_center_robust(source_space)

                if source_point is None and isinstance(door_points_by_source, dict):
                    host_candidates = door_points_by_source.get("__host::" + source_label) or door_points_by_source.get("__all_host__") or []
                    if host_candidates:
                        center_source = _space_center_robust(source_space)
                        host_ref = _to_host_point(center_source, space_row.get("source_transform")) if center_source is not None else None
                        point = _nearest_door_point(host_ref, host_candidates) if host_ref is not None else host_candidates[0]

                if source_point is None and point is None:
                    failures.append((space_row, entry, "Could not calculate placement point", None))
                    continue

                if point is None:
                    point = _to_host_point(source_point, space_row.get("source_transform"))
                if point is None:
                    failures.append((space_row, entry, "Could not transform placement point to host", source_point))
                    continue

                # Guard against bad API coordinates resolving to model origin.
                if _is_origin_point(point):
                    fallback_point = None
                    center_source = _space_center_robust(source_space)
                    center_host = _to_host_point(center_source, space_row.get("source_transform")) if center_source is not None else None
                    if center_host is not None and not _is_origin_point(center_host):
                        fallback_point = center_host
                    else:
                        host_candidates = []
                        if isinstance(door_points_by_source, dict):
                            host_candidates = door_points_by_source.get("__host::" + source_label) or door_points_by_source.get("__all_host__") or []
                        if host_candidates:
                            if center_host is not None:
                                fallback_point = _nearest_door_point(center_host, host_candidates)
                            else:
                                fallback_point = host_candidates[0]

                    if fallback_point is not None and not _is_origin_point(fallback_point):
                        point = fallback_point
                    else:
                        failures.append((space_row, entry, "Computed placement point resolved to origin", point))
                        continue

            kind = _normalize_entry_kind(entry.get("kind"), entry.get("id"))
            element = None

            if kind == "family_type":
                symbol = _get_family_symbol(doc, entry.get("element_type_id"), entry.get("name"))
                if symbol is None:
                    failures.append((space_row, entry, "Family type not found by id/name", point))
                    continue
                level = space_row.get("host_level")
                if not _is_numeric_level_name(level):
                    level_name = str(getattr(level, "Name", "<None>") or "<None>")
                    failures.append((space_row, entry, "Host level is non-numeric '{}'".format(level_name), point))
                    continue

                if rule == "One Foot off doorway wall":
                    placement_type = str(getattr(symbol, "FamilyPlacementType", "<unknown>") or "<unknown>")
                    element = _place_family_instance(doc, symbol, point, level, None)
                    wall_reason = None
                else:
                    element = _place_family_instance(doc, symbol, point, level, None)
                    wall_reason = None
                if element is None:
                    if rule == "One Foot off doorway wall":
                        reason = "Family instance could not be placed near doorway wall"
                        reason = "{} (family placement type: {})".format(reason, placement_type)
                        failures.append((space_row, entry, reason, point))
                    else:
                        failures.append((space_row, entry, "Family instance creation failed (invalid point or unsupported host requirements)", point))
                    continue

                if rule == "One Foot off doorway wall":
                    center_source = _space_center_robust(source_space)
                    center_host = _to_host_point(center_source, space_row.get("source_transform")) if center_source is not None else None
                    _rotate_instance_facing_perpendicular_to_door(element, source_door_row, space_row.get("source_transform"), center_host=center_host)

                if rule not in ("Center of Room", "One Foot off doorway wall"):
                    _try_set_instance_elevation(element, point.Z)

            elif kind == "model_group":
                group_type = _get_group_type(doc, entry.get("element_type_id"), entry.get("name"))
                if group_type is None:
                    failures.append((space_row, entry, "Model group type not found by id", point))
                    continue

                level = space_row.get("host_level")
                if not _is_numeric_level_name(level):
                    level_name = str(getattr(level, "Name", "<None>") or "<None>")
                    failures.append((space_row, entry, "Host level is non-numeric '{}'".format(level_name), point))
                    continue

                level_z = _level_elevation(level)
                attempted_points = []
                if level_z is not None:
                    attempted_points.append(XYZ(point.X, point.Y, level_z))
                attempted_points.append(point)

                element = None
                place_errors = []
                for group_point in attempted_points:
                    try:
                        element = doc.Create.PlaceGroup(group_point, group_type)
                        break
                    except Exception as exc:
                        place_errors.append(str(exc))

                if element is None:
                    reason = "Model group placement failed"
                    if place_errors:
                        reason = "{}: {}".format(reason, place_errors[0])
                    failures.append((space_row, entry, reason, point))
                    continue

            else:
                failures.append((space_row, entry, "Unsupported template kind '{}'".format(kind), point))
                continue

            stats = _apply_parameter_overrides(element, entry.get("parameters") or {})
            for key in param_totals.keys():
                param_totals[key] += int(stats.get(key, 0) or 0)

            placed += 1
            actual_point = _element_location_point(element, fallback_point=None)
            placed_rows.append(
                {
                    "space_number": space_row.get("space_number") or "<No Number>",
                    "space_name": space_row.get("space_name") or "<Unnamed>",
                    "entry_name": entry.get("name") or entry.get("id") or "<Entry>",
                    "bucket": space_row.get("bucket") or "Other",
                    "element_id": _element_id_value(getattr(element, "Id", None), default="<unknown>"),
                    "target_point": point,
                    "actual_point": actual_point,
                }
            )

        tx.Commit()
    except Exception:
        tx.RollBack()
        raise

    return placed, failures, param_totals, placed_rows


def _summary_lines(space_rows, request_rows, placed_count, failures, param_totals, door_count, source_label, placed_rows=None, door_stats=None):
    counts = _bucket_counts(space_rows)
    lines = [
        "Placed space profile elements.",
        "Storage ID: {}".format(CLASSIFICATION_STORAGE_ID),
        "Source: {}".format(source_label),
        "",
        "Classified spaces processed: {}".format(len(space_rows)),
        "Placement requests: {}".format(len(request_rows)),
        "Successfully placed: {}".format(placed_count),
        "Failed placements: {}".format(len(failures)),
        "Detected doors in source model: {}".format(door_count),
        "",
        "Door detection breakdown:",
        " - Host doors: {}".format((door_stats or {}).get("host", 0)),
        " - Linked doors: {}".format((door_stats or {}).get("linked", 0)),
        " - Linked models with doors: {}".format((door_stats or {}).get("linked_models_with_doors", 0)),
        "",
        "Parameter overrides:",
        " - Set: {}".format(param_totals.get("set", 0)),
        " - Missing on placed element: {}".format(param_totals.get("missing", 0)),
        " - Read-only: {}".format(param_totals.get("readonly", 0)),
        " - Failed to set: {}".format(param_totals.get("failed", 0)),
        "",
        "Space buckets with counts:",
    ]

    for bucket in BUCKETS:
        count = counts.get(bucket, 0)
        if count <= 0:
            continue
        lines.append(" - {}: {}".format(bucket, count))

    if placed_rows:
        lines.append("")
        lines.append("Placed coordinates (first 50):")
        for row in placed_rows[:50]:
            label = "{} - {}".format(row.get("space_number") or "<No Number>", row.get("space_name") or "<Unnamed>")
            entry_name = row.get("entry_name") or "<Entry>"
            target_xyz = _format_xyz(row.get("target_point"))
            actual_xyz = _format_xyz(row.get("actual_point"))
            element_id = row.get("element_id") or "<unknown>"
            lines.append(" - {} | {} | id {} | target {} | actual {}".format(label, entry_name, element_id, target_xyz, actual_xyz))
        if len(placed_rows) > 50:
            lines.append(" - ... ({} more)".format(len(placed_rows) - 50))

    if failures:
        lines.append("")
        lines.append("Placement failures (first 20):")
        for failure in failures[:20]:
            row = failure[0]
            entry = failure[1]
            reason = failure[2]
            target_point = failure[3] if len(failure) > 3 else None
            label = "{} - {}".format(row.get("space_number") or "<No Number>", row.get("space_name") or "<Unnamed>")
            entry_name = entry.get("name") or entry.get("id") or "<Entry>"
            if target_point is not None:
                lines.append(" - {} | {} | {} | target {}".format(label, entry_name, reason, _format_xyz(target_point)))
            else:
                lines.append(" - {} | {} | {}".format(label, entry_name, reason))
        if len(failures) > 20:
            lines.append(" - ... ({} more)".format(len(failures) - 20))

    return lines


def _load_latest_space_payload(doc):
    if ExtensibleStorage is None:
        return {}

    # Defensive cache clear so each run re-resolves schema/data-storage for this document.
    try:
        schema_cache = getattr(ExtensibleStorage, "_schema_cache", None)
        if isinstance(schema_cache, dict):
            schema_cache.clear()
    except Exception:
        pass

    payload = None
    try:
        payload = ExtensibleStorage.get_project_data(doc, CLASSIFICATION_STORAGE_ID, default=None)
    except Exception:
        payload = None

    payload_map = _as_dict(payload)
    if payload_map:
        return payload_map

    # Fallback: read root storage directly, then project_data map.
    try:
        root = ExtensibleStorage._read_storage(doc)
    except Exception:
        root = None

    root_map = _as_dict(root)
    meta_map = _as_dict(root_map.get("meta"))
    project_data_map = _as_dict(meta_map.get("project_data"))
    return _as_dict(project_data_map.get(CLASSIFICATION_STORAGE_ID))


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    if ExtensibleStorage is None:
        forms.alert("Failed to load ExtensibleStorage library from CEDLib.lib.", title=TITLE)
        return

    payload = _load_latest_space_payload(doc)
    if not isinstance(payload, dict):
        forms.alert(
            "No saved space data found.\n\nRun Classify Spaces and save first.",
            title=TITLE,
        )
        return

    assignments = payload.get("space_assignments")
    if not isinstance(assignments, dict) or not assignments:
        forms.alert(
            "No space assignments found in saved data.\n\nRun Classify Spaces and save first.",
            title=TITLE,
        )
        return

    profile_saved_utc = str(payload.get("space_profile_saved_utc") or "").strip()

    type_elements = _sanitize_type_elements(payload.get(KEY_TYPE_ELEMENTS) or {})
    space_overrides = _sanitize_space_overrides(payload.get(KEY_SPACE_OVERRIDES) or {})

    sources = _collect_all_space_sources(doc)
    spaces = []
    door_points_by_source = {}
    door_rows_by_source = {}
    source_transform_by_label = {}

    for source in sources:
        source_doc = source.get("source_doc")
        if source_doc is None:
            continue

        source_transform = source.get("source_transform")
        source_row_label = source.get("source_label") or "Host Spaces"
        source_transform_by_label[source_row_label] = source_transform

        source_spaces = _collect_classified_spaces(
            doc,
            source_doc,
            assignments,
            source_transform=source_transform,
            source_label=source_row_label,
        )
        if source_spaces:
            spaces.extend(source_spaces)

        if source_row_label not in door_points_by_source:
            door_points_by_source[source_row_label] = _collect_door_points(source_doc)
        if source_row_label not in door_rows_by_source:
            door_rows_by_source[source_row_label] = _collect_door_rows(source_doc)

        door_points_by_source[source_row_label] = _clean_origin_points(door_points_by_source.get(source_row_label) or [])
        door_rows_by_source[source_row_label] = _clean_origin_rows(door_rows_by_source.get(source_row_label) or [])
    # Backfill linked-door points from host coordinates when link docs do not expose doors.
    host_door_points = list(door_points_by_source.get("Host Spaces") or [])
    if host_door_points:
        for source in sources:
            source_label_key = source.get("source_label") or "Host Spaces"
            if source_label_key == "Host Spaces":
                continue

            existing = door_points_by_source.get(source_label_key) or []
            if existing:
                continue

            source_transform = source.get("source_transform")
            transformed = _transform_points_to_source(host_door_points, source_transform)
            if transformed:
                door_points_by_source[source_label_key] = _clean_origin_points(transformed)

    # Backfill linked doorway rows from transformed host rows (preserve orientation metadata)
    # and finally from doorway points when row metadata is unavailable.
    host_door_rows = _clean_origin_rows(door_rows_by_source.get("Host Spaces") or [])
    for source in sources:
        source_label_key = source.get("source_label") or "Host Spaces"
        if source_label_key == "Host Spaces":
            continue

        source_rows = _clean_origin_rows(door_rows_by_source.get(source_label_key) or [])
        if source_rows:
            door_rows_by_source[source_label_key] = source_rows
            continue

        source_transform = source.get("source_transform")
        if host_door_rows:
            transformed_rows = _clean_origin_rows(_transform_door_rows_to_source(host_door_rows, source_transform))
            if transformed_rows:
                door_rows_by_source[source_label_key] = transformed_rows
                continue

        source_points = _clean_origin_points(door_points_by_source.get(source_label_key) or [])
        if source_points:
            door_rows_by_source[source_label_key] = _door_rows_from_points(source_points)

    # Build host-coordinate door maps per source label for robust fallback placement.
    all_host_door_points = []
    for source_label_key, source_points in list(door_points_by_source.items()):
        if not source_points:
            continue

        if source_label_key == "Host Spaces":
            host_points = list(source_points)
        else:
            host_points = []
            source_transform = source_transform_by_label.get(source_label_key)
            for source_pt in source_points:
                host_pt = _to_host_point(source_pt, source_transform)
                if host_pt is not None:
                    host_points.append(host_pt)

        if not host_points:
            continue

        door_points_by_source["__host::" + source_label_key] = host_points
        all_host_door_points.extend(host_points)

    if all_host_door_points:
        door_points_by_source["__all_host__"] = all_host_door_points
    if not spaces:
        forms.alert("No spaces were found in host or loaded linked models.", title=TITLE)
        return

    linked_count = max(0, len(sources) - 1)
    if linked_count > 0:
        source_label = "Host + {} linked model{}".format(linked_count, "" if linked_count == 1 else "s")
    else:
        source_label = "Host Spaces"

    requests = _request_rows(spaces, type_elements, space_overrides)
    if not requests:
        space_counts = _bucket_counts(spaces)
        host_space_counts = _bucket_counts_for_source(spaces, "Host Spaces")
        template_bucket_counts = _template_counts_by_bucket(type_elements)
        override_total = sum(len(entries or []) for entries in space_overrides.values())
        template_kinds = _template_kind_counts(type_elements, space_overrides)

        lines = [
            "No placement requests were resolved for the current spaces.",
            "",
            "Saved template totals:",
            " - Profiles saved UTC: {}".format(profile_saved_utc or "<unknown>"),
            " - Type templates: {}".format(sum(template_bucket_counts.values())),
            " - Space overrides: {}".format(override_total),
            " - Family Type templates: {}".format(template_kinds.get("family_type", 0)),
            " - Model Group templates: {}".format(template_kinds.get("model_group", 0)),
            "",
            "Host space counts by bucket (Manage Space Profiles scope):",
        ]
        for bucket in BUCKETS:
            lines.append(" - {}: {}".format(bucket, host_space_counts.get(bucket, 0)))

        lines.append("")
        lines.append("Saved type templates by bucket:")
        for bucket in BUCKETS:
            lines.append(" - {}: {}".format(bucket, template_bucket_counts.get(bucket, 0)))

        lines.extend(
            [
                "",
                "If this is unexpected, open Manage Space Profiles and confirm templates are saved under the same bucket names as the classified spaces.",
            ]
        )

        forms.alert("\n".join(lines), title=TITLE)
        return

    selected_profile_keys = _prompt_profile_selection(requests)
    if selected_profile_keys is None:
        return
    if not selected_profile_keys:
        forms.alert("No space profiles were selected.", title=TITLE)
        return

    requests = [
        (space_row, entry)
        for space_row, entry in requests
        if _request_profile_key(space_row, entry) in selected_profile_keys
    ]
    if not requests:
        forms.alert("No placement requests matched the selected space profiles.", title=TITLE)
        return

    counts = _bucket_counts(spaces)
    host_counts = _bucket_counts_for_source(spaces, "Host Spaces")
    preview = [
        "This will place selected space profiles into classified spaces.",
        "",
        "Source: {}".format(source_label),
        "Spaces: {}".format(len(spaces)),
        "Selected profiles: {}".format(len(selected_profile_keys)),
        "Placement requests: {}".format(len(requests)),
        "Profiles saved UTC: {}".format(profile_saved_utc or "<unknown>"),
        "",
        "Host Buckets (matches Manage Space Profiles):",
    ]
    for bucket in BUCKETS:
        count = host_counts.get(bucket, 0)
        if count <= 0:
            continue
        preview.append(" - {}: {}".format(bucket, count))

    preview.append("")
    preview.append("Continue?")

    proceed = forms.alert("\n".join(preview), title=TITLE, yes=True, no=True)
    if not proceed:
        return

    host_door_count = len(door_points_by_source.get("Host Spaces") or [])
    linked_door_count = sum(
        len(points or [])
        for key, points in door_points_by_source.items()
        if (not str(key).startswith("__")) and str(key) != "Host Spaces"
    )
    linked_models_with_doors = sum(
        1
        for key, points in door_points_by_source.items()
        if (not str(key).startswith("__")) and str(key) != "Host Spaces" and len(points or []) > 0
    )
    total_door_count = host_door_count + linked_door_count

    door_stats = {
        "host": host_door_count,
        "linked": linked_door_count,
        "linked_models_with_doors": linked_models_with_doors,
    }

    try:
        placed_count, failures, param_totals, placed_rows = _run_placement(doc, requests, door_points_by_source, door_rows_by_source)
    except Exception as exc:
        forms.alert("Placement failed:\n\n{}".format(exc), title=TITLE)
        return

    forms.alert(
        "\n".join(_summary_lines(spaces, requests, placed_count, failures, param_totals, total_door_count, source_label, placed_rows=placed_rows, door_stats=door_stats)),
        title=TITLE,
    )


if __name__ == "__main__":
    main()
























































































































