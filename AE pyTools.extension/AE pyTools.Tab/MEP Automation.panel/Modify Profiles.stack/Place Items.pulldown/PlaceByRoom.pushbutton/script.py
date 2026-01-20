# -*- coding: utf-8 -*-
"""Place selected profiles in linked rooms using the active YAML store."""

import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    Level,
    LocationPoint,
    RevitLinkInstance,
    XYZ,
)

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Place by Room"

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


def _room_label(room):
    number = _room_number(room)
    name = _room_name(room)
    label = "{} - {}".format(number, name).strip(" -")
    return label or "<Room>"


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


def _rotation_for_label(repo, profile_name, label):
    try:
        linked_def = repo.definition_for_label(profile_name, label)
    except Exception:
        linked_def = None
    if linked_def is None:
        return 0.0
    try:
        placement = linked_def.get_placement()
    except Exception:
        placement = None
    if placement is None:
        return 0.0
    try:
        rotation = placement.get_rotation_degrees()
    except Exception:
        rotation = None
    try:
        return float(rotation or 0.0)
    except Exception:
        return 0.0


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

    repo = _build_repository(data)

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

    level_lookup = _build_level_lookup(doc)
    default_level = _find_level_one(doc)
    default_level_id = None
    if default_level is not None:
        try:
            default_level_id = default_level.Id.IntegerValue
        except Exception:
            default_level_id = None

    try:
        transform = link.GetTransform()
    except Exception:
        transform = None

    room_centers = {}
    room_level_map = {}
    skipped_rooms = []
    for room in rooms:
        room_id = getattr(room.Id, "IntegerValue", None)
        center = _room_center(room)
        if center is None:
            skipped_rooms.append(_room_label(room))
            continue
        if transform:
            try:
                center = transform.OfPoint(center)
            except Exception:
                pass
        room_centers[room_id] = center

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
        room_level_map[room_id] = level_id

    rows = []
    selection_map = {}
    for profile_name in profiles:
        labels = profile_labels.get(profile_name) or []
        for label in labels:
            row_name = "{}:{}".format(profile_name, label)
            selection_map[row_name] = [label]
            rotation = _rotation_for_label(repo, profile_name, label)
            for room in rooms:
                room_id = getattr(room.Id, "IntegerValue", None)
                point = room_centers.get(room_id)
                if point is None:
                    continue
                level_id = room_level_map.get(room_id)
                rows.append(_build_row(row_name, point, rotation, level_id))

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

    summary = [
        "Processed {} placement(s).".format(len(rows)),
        "Placed {} element(s).".format(placed),
    ]
    if skipped_rooms:
        summary.append("Skipped {} room(s) with no center point.".format(len(skipped_rooms)))
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
