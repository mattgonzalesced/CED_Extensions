# -*- coding: utf-8 -*-
"""Place selected profiles in linked rooms using the active YAML store."""

import imp
import os
import sys

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Transaction,
    FilteredElementCollector,
    FamilySymbol,
    FamilyInstance,
    ElementId,
    ElementTransformUtils,
    Level,
    LocationPoint,
    RevitLinkInstance,
    SpatialElementBoundaryOptions,
    XYZ,
)

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Place by Room"
LOG = script.get_logger()

try:
    basestring
except NameError:
    basestring = str


def _build_repository(data):
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def _find_level_one(doc):
    if doc is None:
        return None
    targets = {"level1", "level01", "lvl1", "l1"}
    try:
        levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        levels = []
    for level in levels:
        name = getattr(level, "Name", None)
        key = "".join(ch for ch in str(name).lower() if ch.isalnum())
        if key in targets:
            return level
    return None


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _build_level_lookup(doc):
    lookup = {}
    try:
        levels = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        levels = []
    for level in levels:
        name = getattr(level, "Name", None)
        key = _normalize_name(name)
        if key and key not in lookup:
            lookup[key] = level
    return lookup


def _room_center(room):
    if room is None:
        return None
    loc = getattr(room, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point
    try:
        bbox = room.get_BoundingBox(None)
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


def _room_extents(room):
    if room is None:
        return None
    points = []
    try:
        options = SpatialElementBoundaryOptions()
        loops = room.GetBoundarySegments(options)
    except Exception:
        loops = None
    if loops:
        for loop in loops:
            for segment in loop:
                try:
                    curve = segment.GetCurve()
                except Exception:
                    curve = None
                if curve is None:
                    continue
                try:
                    points.append(curve.GetEndPoint(0))
                    points.append(curve.GetEndPoint(1))
                except Exception:
                    pass
    if not points:
        try:
            bbox = room.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            return None
        try:
            points = [bbox.Min, bbox.Max]
        except Exception:
            return None
    min_x = min(point.X for point in points)
    max_x = max(point.X for point in points)
    min_y = min(point.Y for point in points)
    max_y = max(point.Y for point in points)
    return {
        "min_x": min_x,
        "max_x": max_x,
        "min_y": min_y,
        "max_y": max_y,
    }


def _room_boundary_points(room):
    if room is None:
        return []
    points = []
    try:
        options = SpatialElementBoundaryOptions()
        loops = room.GetBoundarySegments(options)
    except Exception:
        loops = None
    if loops:
        for loop in loops:
            for segment in loop:
                try:
                    curve = segment.GetCurve()
                except Exception:
                    curve = None
                if curve is None:
                    continue
                try:
                    points.append(curve.GetEndPoint(0))
                    points.append(curve.GetEndPoint(1))
                except Exception:
                    pass
    if points:
        return points
    try:
        bbox = room.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return []
    try:
        return [bbox.Min, bbox.Max]
    except Exception:
        return []


def _room_boundary_segments(room):
    if room is None:
        return []
    segments = []
    try:
        options = SpatialElementBoundaryOptions()
        loops = room.GetBoundarySegments(options)
    except Exception:
        loops = None
    if loops:
        for loop in loops:
            for segment in loop:
                try:
                    curve = segment.GetCurve()
                except Exception:
                    curve = None
                if curve is None:
                    continue
                try:
                    start = curve.GetEndPoint(0)
                    end = curve.GetEndPoint(1)
                except Exception:
                    continue
                segments.append((start, end))
    return segments


def _room_extents_host(room, transform):
    points = _room_boundary_points(room)
    if not points:
        return None
    transformed = []
    if transform:
        for point in points:
            try:
                transformed.append(transform.OfPoint(point))
            except Exception:
                transformed.append(point)
    else:
        transformed = points
    min_x = min(point.X for point in transformed)
    max_x = max(point.X for point in transformed)
    min_y = min(point.Y for point in transformed)
    max_y = max(point.Y for point in transformed)
    return {
        "min_x": min_x,
        "max_x": max_x,
        "min_y": min_y,
        "max_y": max_y,
    }


def _axis_for_rotation(rotation_deg):
    if rotation_deg is None:
        return None
    try:
        angle = float(rotation_deg) % 360.0
    except Exception:
        return None
    # Treat ~90/270 as Y-axis spacing, everything else as X-axis.
    if abs(angle - 90.0) <= 15.0 or abs(angle - 270.0) <= 15.0:
        return "y"
    return "x"


def _rotation_is_vertical(rotation_deg):
    if rotation_deg is None:
        return False
    try:
        angle = float(rotation_deg) % 360.0
    except Exception:
        return False
    return abs(angle - 90.0) <= 15.0 or abs(angle - 270.0) <= 15.0


def _axis_for_symbol(symbol, fallback_axis):
    if symbol is None:
        return fallback_axis
    try:
        bbox = symbol.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return fallback_axis
    try:
        span_x = abs(bbox.Max.X - bbox.Min.X)
        span_y = abs(bbox.Max.Y - bbox.Min.Y)
    except Exception:
        return fallback_axis
    if span_x <= 1e-6 or span_y <= 1e-6:
        return fallback_axis
    return "x" if span_x >= span_y else "y"


def _room_layout_points(room, count, axis_hint=None, element_width=None, min_offset=None):
    if room is None or count <= 0:
        return []
    center = _room_center(room)
    if center is None:
        return []
    extents = _room_extents(room)
    if not extents:
        return [center]
    span_x = extents["max_x"] - extents["min_x"]
    span_y = extents["max_y"] - extents["min_y"]
    if span_x <= 1e-6 or span_y <= 1e-6:
        return [center]
    if axis_hint == "x":
        span = span_x
        base = extents["min_x"]
        fixed = (extents["min_y"] + extents["max_y"]) / 2.0
        axis = "x"
    elif axis_hint == "y":
        span = span_y
        base = extents["min_y"]
        fixed = (extents["min_x"] + extents["max_x"]) / 2.0
        axis = "y"
    elif span_x >= span_y:
        span = span_x
        base = extents["min_x"]
        fixed = (extents["min_y"] + extents["max_y"]) / 2.0
        axis = "x"
    else:
        span = span_y
        base = extents["min_y"]
        fixed = (extents["min_x"] + extents["max_x"]) / 2.0
        axis = "y"
    width = 0.0 if element_width is None else float(element_width)
    offset = 0.0 if min_offset is None else float(min_offset)
    if width < 0.0:
        width = 0.0
    spacing = (span - (count * width)) / float(count + 1)
    if spacing < 0.0:
        spacing = 0.0
    points = []
    for idx in range(count):
        position = base + spacing - offset + (width + spacing) * idx
        if axis == "x":
            points.append(XYZ(position, fixed, center.Z))
        else:
            points.append(XYZ(fixed, position, center.Z))
    return points


def _room_name(room):
    try:
        name = getattr(room, "Name", None)
    except Exception:
        name = None
    if name:
        return name
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
    except Exception:
        param = None
    if param:
        try:
            return param.AsString() or ""
        except Exception:
            return ""
    return ""




def _room_number(room):
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
    except Exception:
        param = None
    if param:
        try:
            return param.AsString() or ""
        except Exception:
            return ""
    return ""


def _build_row(name, point, rotation_deg, level_id=None):
    row = {
        "Name": name,
        "Count": "1",
        "Position X": str(point.X * 12.0),
        "Position Y": str(point.Y * 12.0),
        "Position Z": str(point.Z * 12.0),
        "Rotation": str(rotation_deg or 0.0),
    }
    if level_id is not None:
        row["LevelId"] = str(level_id)
    return row


def _place_requests(doc, repo, selection_map, rows, default_level=None):
    if not selection_map or not rows:
        return {"placed": 0}
    engine = PlaceElementsEngine(
        doc,
        repo,
        default_level=default_level,
        allow_tags=False,
        transaction_name="Place by Room",
        apply_recorded_level=False,
    )
    return engine.place_from_csv(rows, selection_map)


def _select_link_instance(doc):
    try:
        links = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
    except Exception:
        links = []
    if not links:
        return None
    options = []
    option_map = {}
    for link in links:
        try:
            link_doc = link.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        name = getattr(link, "Name", None) or getattr(link_doc, "Title", None) or "<Linked Model>"
        if name in option_map:
            continue
        option_map[name] = link
        options.append(name)
    if not options:
        return None
    options.sort(key=lambda value: value.lower())
    selection = forms.SelectFromList.show(
        options,
        title="Select Linked Model",
        button_name="Select",
        multiselect=False,
    )
    if not selection:
        return None
    chosen = selection[0] if isinstance(selection, list) else selection
    return option_map.get(chosen)


def _select_profiles(repo):
    profiles = list(repo.cad_names() or [])
    if not profiles:
        return []
    profiles.sort(key=lambda value: value.lower())
    selection = forms.SelectFromList.show(
        profiles,
        title="Select Profiles",
        button_name="Place",
        multiselect=True,
    )
    if not selection:
        return []
    if isinstance(selection, basestring):
        return [selection]
    return list(selection)


def _room_label(room):
    number = _room_number(room)
    name = _room_name(room)
    label = "{} - {}".format(number, name).strip(" -")
    return label or "<Room>"


def _load_counts_ui():
    module_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "PlaceByRoomUI.py"))
    if not os.path.exists(module_path):
        forms.alert("Counts UI not found at:\n{}".format(module_path), title=TITLE)
        return None
    try:
        return imp.load_source("ced_place_by_room_ui", module_path)
    except Exception as exc:
        forms.alert("Failed to load counts UI:\n{}\n\n{}".format(module_path, exc), title=TITLE)
        return None


def _prompt_counts_ui(room_entries, profiles, profile_labels):
    module = _load_counts_ui()
    if module is None:
        return None
    xaml_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "PlaceByRoomWindow.xaml"))
    if not os.path.exists(xaml_path):
        forms.alert("Counts XAML not found at:\n{}".format(xaml_path), title=TITLE)
        return None
    window = module.PlaceByRoomWindow(xaml_path, room_entries, profiles, profile_labels)
    result = window.show_dialog()
    if not result:
        return None
    return getattr(window, "counts", None)


def _build_symbol_map(doc):
    symbol_map = {}
    try:
        symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements())
    except Exception:
        symbols = []
    for sym in symbols:
        try:
            family = getattr(sym, "Family", None)
            fam_name = getattr(family, "Name", None) if family else None
            if not fam_name:
                continue
            type_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = type_param.AsString() if type_param else None
            if not type_name and hasattr(sym, "Name"):
                type_name = sym.Name
            if not type_name:
                continue
            label = u"{} : {}".format(fam_name, type_name)
            symbol_map[label] = sym
        except Exception:
            continue
    return symbol_map


def _symbol_axis_metrics(symbol, axis_hint):
    if symbol is None or axis_hint not in ("x", "y"):
        return None, None
    try:
        bbox = symbol.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None, None
    try:
        min_x = bbox.Min.X
        max_x = bbox.Max.X
        min_y = bbox.Min.Y
        max_y = bbox.Max.Y
        width_x = abs(max_x - min_x)
        width_y = abs(max_y - min_y)
    except Exception:
        return None, None
    if axis_hint == "x":
        return width_x, min_x
    return width_y, min_y


def _collect_symbol_instance_ids(doc, symbol_ids):
    ids_map = {sid: set() for sid in symbol_ids}
    if not ids_map:
        return ids_map
    try:
        instances = list(FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType())
    except Exception:
        instances = []
    for inst in instances:
        try:
            symbol = inst.Symbol
            symbol_id = symbol.Id.IntegerValue if symbol is not None else None
        except Exception:
            symbol_id = None
        if symbol_id is None:
            continue
        if symbol_id in ids_map:
            try:
                ids_map[symbol_id].add(inst.Id.IntegerValue)
            except Exception:
                continue
    return ids_map


def _instance_point(instance):
    if instance is None:
        return None
    try:
        loc = instance.Location
    except Exception:
        loc = None
    if isinstance(loc, LocationPoint):
        try:
            return loc.Point
        except Exception:
            return None
    if loc is not None:
        try:
            curve = loc.Curve
        except Exception:
            curve = None
        if curve is not None:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                pass
    try:
        bbox = instance.get_BoundingBox(None)
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


def _instance_axis_bounds(instance, axis_hint):
    if instance is None or axis_hint not in ("x", "y"):
        return None, None
    try:
        bbox = instance.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None, None
    if axis_hint == "x":
        return bbox.Min.X, bbox.Max.X
    return bbox.Min.Y, bbox.Max.Y


def _instance_axis_from_bbox(instance, fallback_axis):
    if instance is None:
        return fallback_axis
    try:
        bbox = instance.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return fallback_axis
    try:
        span_x = abs(bbox.Max.X - bbox.Min.X)
        span_y = abs(bbox.Max.Y - bbox.Min.Y)
    except Exception:
        return fallback_axis
    if span_x <= 1e-6 or span_y <= 1e-6:
        return fallback_axis
    return "x" if span_x >= span_y else "y"


def _axis_from_extents(extents):
    span_x = extents["max_x"] - extents["min_x"]
    span_y = extents["max_y"] - extents["min_y"]
    return "x" if span_x >= span_y else "y"


def _median(values):
    if not values:
        return None
    ordered = sorted(values)
    count = len(ordered)
    mid = count // 2
    if count % 2:
        return ordered[mid]
    return (ordered[mid - 1] + ordered[mid]) / 2.0


def _axis_span_from_segments(segments, axis_hint, coord, tol=1e-6):
    if not segments or axis_hint not in ("x", "y") or coord is None:
        return None
    values = []
    for p1, p2 in segments:
        if axis_hint == "x":
            y0 = p1.Y
            y1 = p2.Y
            if abs(y1 - y0) <= tol:
                if abs(y0 - coord) <= tol:
                    values.append(p1.X)
                    values.append(p2.X)
                continue
            if (y0 <= coord <= y1) or (y1 <= coord <= y0):
                t = (coord - y0) / (y1 - y0)
                values.append(p1.X + t * (p2.X - p1.X))
        else:
            x0 = p1.X
            x1 = p2.X
            if abs(x1 - x0) <= tol:
                if abs(x0 - coord) <= tol:
                    values.append(p1.Y)
                    values.append(p2.Y)
                continue
            if (x0 <= coord <= x1) or (x1 <= coord <= x0):
                t = (coord - x0) / (x1 - x0)
                values.append(p1.Y + t * (p2.Y - p1.Y))
    if len(values) < 2:
        return None
    return min(values), max(values)


def _axis_center_coord(instances, axis_hint):
    if not instances or axis_hint not in ("x", "y"):
        return None
    coords = []
    for inst in instances:
        pt = _instance_point(inst)
        if pt is None:
            continue
        coords.append(pt.Y if axis_hint == "x" else pt.X)
    return _median(coords)


def _nudge_instances(doc, groups, room_extents_map, room_boundary_map, axis_map):
    moved = 0
    if not groups:
        return moved
    t = Transaction(doc, "Nudge Place by Room")
    t.Start()
    try:
        for (room_id, symbol_id), instances in groups.items():
            extents = room_extents_map.get(room_id)
            if not extents:
                continue
            axis_hint = axis_map.get(symbol_id) or _axis_from_extents(extents)
            axis_locked = symbol_id in axis_map
            axis_hint_group = axis_hint
            if not axis_locked and instances:
                axis_hint_group = _instance_axis_from_bbox(instances[0], axis_hint_group)
            axis_min = None
            axis_max = None
            segments = room_boundary_map.get(room_id) if room_boundary_map else None
            if segments:
                coord = _axis_center_coord(instances, axis_hint_group)
                span = _axis_span_from_segments(segments, axis_hint_group, coord)
                if span:
                    axis_min, axis_max = span
            if axis_min is None or axis_max is None:
                if axis_hint_group == "x":
                    axis_min = extents["min_x"]
                    axis_max = extents["max_x"]
                else:
                    axis_min = extents["min_y"]
                    axis_max = extents["max_y"]
            span = axis_max - axis_min
            if span <= 1e-6:
                continue
            items = []
            for inst in instances:
                min_val, max_val = _instance_axis_bounds(inst, axis_hint_group)
                if min_val is None or max_val is None:
                    continue
                width = max_val - min_val
                if width <= 1e-6:
                    continue
                items.append({"inst": inst, "min": min_val, "width": width})
            if not items:
                continue
            items.sort(key=lambda item: item["min"])
            total_width = sum(item["width"] for item in items)
            gap = (span - total_width) / float(len(items) + 1)
            if gap < 0.0:
                gap = 0.0
            target_min = axis_min + gap
            for item in items:
                delta = target_min - item["min"]
                if abs(delta) > 1e-6:
                    move_vec = XYZ(delta, 0.0, 0.0) if axis_hint_group == "x" else XYZ(0.0, delta, 0.0)
                    try:
                        ElementTransformUtils.MoveElement(doc, item["inst"].Id, move_vec)
                        moved += 1
                    except Exception:
                        pass
                target_min += item["width"] + gap
    except Exception:
        t.RollBack()
        raise
    else:
        t.Commit()
    return moved


def _select_rooms(link_doc):
    try:
        rooms = list(FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())
    except Exception:
        rooms = []
    display_map = {}
    options = []
    for room in rooms:
        try:
            area = getattr(room, "Area", 0.0) or 0.0
        except Exception:
            area = 0.0
        if area <= 0.0:
            continue
        name = _room_name(room)
        number = _room_number(room)
        if not name and not number:
            continue
        label = _room_label(room)
        if label in display_map:
            continue
        display_map[label] = room
        options.append(label)
    if not options:
        return []
    options.sort(key=lambda value: value.lower())
    selection = forms.SelectFromList.show(
        options,
        title="Select Rooms",
        button_name="Place",
        multiselect=True,
    )
    if not selection:
        return []
    if isinstance(selection, basestring):
        selection = [selection]
    return [display_map[label] for label in selection if label in display_map]


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)
    repo = _build_repository(data)
    symbol_map = _build_symbol_map(doc)

    profiles = _select_profiles(repo)
    if not profiles:
        return

    repo_names = set(repo.cad_names() or [])
    profile_labels = {}
    missing_profiles = set()
    missing_labels = set()
    for profile_name in profiles:
        if profile_name not in repo_names:
            missing_profiles.add(profile_name)
            continue
        labels = repo.labels_for_cad(profile_name)
        if not labels:
            missing_labels.add(profile_name)
            continue
        profile_labels[profile_name] = labels

    profiles = [name for name in profiles if name in profile_labels]
    if not profiles:
        summary = [
            "No valid profiles were selected.",
        ]
        if missing_profiles:
            summary.append("")
            summary.append("Missing profiles:")
            for name in sorted(missing_profiles):
                summary.append(" - {}".format(name))
        if missing_labels:
            summary.append("")
            summary.append("Profiles missing linked types:")
            for name in sorted(missing_labels):
                summary.append(" - {}".format(name))
        forms.alert("\n".join(summary), title=TITLE)
        return

    link = _select_link_instance(doc)
    if link is None:
        forms.alert("No linked model with rooms found.", title=TITLE)
        return

    try:
        link_doc = link.GetLinkDocument()
    except Exception:
        link_doc = None
    if link_doc is None:
        forms.alert("Linked model is not loaded.", title=TITLE)
        return

    rooms = _select_rooms(link_doc)
    if not rooms:
        forms.alert("No rooms selected from linked model.", title=TITLE)
        return

    room_entries = []
    for room in rooms:
        room_id = getattr(room.Id, "IntegerValue", None)
        room_entries.append({"id": room_id, "label": _room_label(room)})

    room_counts = _prompt_counts_ui(room_entries, profiles, profile_labels)
    if not room_counts:
        return

    level_lookup = _build_level_lookup(doc)
    default_level = _find_level_one(doc)
    default_level_id = None
    if default_level is not None:
        try:
            default_level_id = default_level.Id.IntegerValue
        except Exception:
            default_level_id = None

    rows = []
    selection_map = {}
    deduped = 0
    seen = set()
    symbol_axis_map = {}
    target_symbol_ids = set()

    try:
        transform = link.GetTransform()
    except Exception:
        transform = None

    row_name_map = {}
    for profile_name in profiles:
        labels = profile_labels.get(profile_name) or []
        for label in labels:
            row_name = "{}:{}".format(profile_name, label)
            selection_map[row_name] = [label]
            row_name_map[(profile_name, label)] = row_name
            try:
                linked_def = repo.definition_for_label(profile_name, label)
            except Exception:
                linked_def = None
            if linked_def is not None:
                placement = linked_def.get_placement()
                rotation = placement.get_rotation_degrees() if placement else None
                axis_hint = _axis_for_rotation(rotation)
                axis_locked = _rotation_is_vertical(rotation)
                family_name = linked_def.get_family()
                type_name = linked_def.get_type()
                if family_name and type_name:
                    symbol_key = u"{} : {}".format(family_name, type_name)
                    symbol = symbol_map.get(symbol_key)
                    if symbol is not None and not axis_locked:
                        axis_hint = _axis_for_symbol(symbol, axis_hint)
                    try:
                        symbol_id = symbol.Id.IntegerValue if symbol is not None else None
                    except Exception:
                        symbol_id = None
                    if symbol_id is not None:
                        target_symbol_ids.add(symbol_id)
                        if symbol_id not in symbol_axis_map and axis_hint:
                            symbol_axis_map[symbol_id] = axis_hint

    pre_existing = _collect_symbol_instance_ids(doc, target_symbol_ids)
    room_extents_map = {}
    room_boundary_map = {}
    for room in rooms:
        room_id = getattr(room.Id, "IntegerValue", None)
        room_extents_map[room_id] = _room_extents_host(room, transform)
        segments = _room_boundary_segments(room)
        if segments and transform:
            transformed = []
            for start, end in segments:
                try:
                    start = transform.OfPoint(start)
                except Exception:
                    pass
                try:
                    end = transform.OfPoint(end)
                except Exception:
                    pass
                transformed.append((start, end))
            segments = transformed
        room_boundary_map[room_id] = segments

    for profile_name in profiles:
        labels = profile_labels.get(profile_name) or []
        for room in rooms:
            room_id = getattr(room.Id, "IntegerValue", None)
            room_payload = room_counts.get(room_id) or {}
            profile_mult = (room_payload.get("profile") or {}).get(profile_name, 1)
            type_counts = (room_payload.get("types") or {}).get(profile_name, {})

            level_id = default_level_id
            try:
                link_level_id = getattr(room, "LevelId", None)
                if link_level_id and link_doc:
                    link_level = link_doc.GetElement(link_level_id)
                    level_name = getattr(link_level, "Name", None)
                    level_key = _normalize_name(level_name)
                    host_level = level_lookup.get(level_key)
                    if host_level is not None:
                        level_id = host_level.Id.IntegerValue
            except Exception:
                pass

            for label in labels:
                axis_hint = None
                element_width = None
                min_offset = None
                try:
                    linked_def = repo.definition_for_label(profile_name, label)
                except Exception:
                    linked_def = None
                if linked_def is not None:
                    placement = linked_def.get_placement()
                    rotation = placement.get_rotation_degrees() if placement else None
                    axis_hint = _axis_for_rotation(rotation)
                    axis_locked = _rotation_is_vertical(rotation)
                    family_name = linked_def.get_family()
                    type_name = linked_def.get_type()
                    if family_name and type_name:
                        symbol_key = u"{} : {}".format(family_name, type_name)
                        symbol = symbol_map.get(symbol_key)
                        if not axis_locked:
                            axis_hint = _axis_for_symbol(symbol, axis_hint)
                        element_width, min_offset = _symbol_axis_metrics(symbol, axis_hint)
                type_count = type_counts.get(label, 1)
                total_count = profile_mult * type_count
                if total_count <= 0:
                    continue
                points = _room_layout_points(
                    room,
                    total_count,
                    axis_hint=axis_hint,
                    element_width=element_width,
                    min_offset=min_offset,
                )
                if not points:
                    continue
                if transform:
                    transformed = []
                    for point in points:
                        try:
                            transformed.append(transform.OfPoint(point))
                        except Exception:
                            transformed.append(point)
                    points = transformed

                row_name = row_name_map.get((profile_name, label)) or "{}:{}".format(profile_name, label)
                for idx, point in enumerate(points):
                    key = (room_id, profile_name, label, idx)
                    if key in seen:
                        deduped += 1
                        continue
                    seen.add(key)
                    rows.append(_build_row(row_name, point, 0.0, level_id))

    if not rows:
        summary = [
            "No placements were generated.",
            "Check selected rooms and profiles.",
        ]
        if missing_profiles:
            summary.append("")
            summary.append("Missing profiles:")
            for name in sorted(missing_profiles):
                summary.append(" - {}".format(name))
        if missing_labels:
            summary.append("")
            summary.append("Profiles missing linked types:")
            for name in sorted(missing_labels):
                summary.append(" - {}".format(name))
        forms.alert("\n".join(summary), title=TITLE)
        return

    results = _place_requests(doc, repo, selection_map, rows, default_level=default_level)
    placed = results.get("placed", 0)

    post_existing = _collect_symbol_instance_ids(doc, target_symbol_ids)
    new_ids_map = {}
    for symbol_id in target_symbol_ids:
        before = pre_existing.get(symbol_id, set())
        after = post_existing.get(symbol_id, set())
        new_ids = after - before
        if new_ids:
            new_ids_map[symbol_id] = new_ids

    groups = {}
    tol = 0.1
    for symbol_id, new_ids in new_ids_map.items():
        for elem_id in new_ids:
            inst = None
            try:
                inst = doc.GetElement(ElementId(int(elem_id)))
            except Exception:
                inst = None
            if inst is None:
                continue
            pt = _instance_point(inst)
            if pt is None:
                continue
            assigned = False
            for room_id, extents in room_extents_map.items():
                if not extents:
                    continue
                if (
                    pt.X >= extents["min_x"] - tol
                    and pt.X <= extents["max_x"] + tol
                    and pt.Y >= extents["min_y"] - tol
                    and pt.Y <= extents["max_y"] + tol
                ):
                    groups.setdefault((room_id, symbol_id), []).append(inst)
                    assigned = True
                    break
            if not assigned:
                continue

    nudged = _nudge_instances(doc, groups, room_extents_map, room_boundary_map, symbol_axis_map)

    summary = [
        "Processed {} placement(s).".format(len(rows)),
        "Placed {} element(s).".format(placed),
    ]
    if nudged:
        summary.append("Nudged {} element(s) for equal edge gaps.".format(nudged))
    if deduped:
        summary.append("Skipped {} duplicate placement(s).".format(deduped))
    if missing_profiles:
        summary.append("")
        summary.append("Missing profiles:")
        sample = sorted(missing_profiles)[:6]
        summary.extend(" - {}".format(name) for name in sample)
        if len(missing_profiles) > len(sample):
            summary.append("   (+{} more)".format(len(missing_profiles) - len(sample)))
    if missing_labels:
        summary.append("")
        summary.append("Profiles missing linked types:")
        sample = sorted(missing_labels)[:6]
        summary.extend(" - {}".format(name) for name in sample)
        if len(missing_labels) > len(sample):
            summary.append("   (+{} more)".format(len(missing_labels) - len(sample)))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
