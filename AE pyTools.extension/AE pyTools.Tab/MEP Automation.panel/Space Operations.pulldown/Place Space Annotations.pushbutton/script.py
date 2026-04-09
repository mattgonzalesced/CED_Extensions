# -*- coding: utf-8 -*-
"""
Place Space Annotations
-----------------------
Place saved tag/keynote annotations for resolved space-profile elements.
"""

import imp
import os
import re
import uuid
from collections import OrderedDict

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    ElementId,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    Group,
    IndependentTag,
    Reference,
    TagMode,
    TagOrientation,
    Transaction,
    XYZ,
)

try:
    from Autodesk.Revit.DB.Structure import StructuralType as RevitStructuralType  # type: ignore
except Exception:
    RevitStructuralType = None

output = script.get_output()
output.close_others()

TITLE = "Place Space Annotations"
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

ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
SPACE_BASED_LINKER_PATTERN = re.compile(r"SPACE\s*BASED\s*,?\s*ID\s*NUMBER\s*=\s*(\d+)", re.IGNORECASE)



def _load_helper():
    pulldown_dir = os.path.dirname(os.path.dirname(__file__))
    helper_path = os.path.join(pulldown_dir, "Place Space Elements.pushbutton", "script.py")
    if not os.path.exists(helper_path):
        return None, "Could not find Place Space Elements helper script."
    try:
        helper = imp.load_source("place_space_elements_helper_" + uuid.uuid4().hex, helper_path)
    except Exception as exc:
        return None, "Failed to load Place Space Elements helper:\n\n{}".format(exc)
    return helper, ""


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


def _as_float(value, default=0.0):
    try:
        return float(str(value).strip())
    except Exception:
        return float(default)



def _read_element_linker_text(element):
    if element is None:
        return ""
    for param_name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = element.LookupParameter(param_name)
        except Exception:
            param = None
        if not param:
            continue
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


def _read_space_based_duplicate_id(element):
    text = _read_element_linker_text(element)
    if not text:
        return None
    match = SPACE_BASED_LINKER_PATTERN.search(text)
    if not match:
        return None
    try:
        return int(match.group(1))
    except Exception:
        return None

def _sanitize_offset(offset):
    offset = offset if isinstance(offset, dict) else {}
    return {
        "x": _as_float(offset.get("x"), 0.0),
        "y": _as_float(offset.get("y"), 0.0),
        "z": _as_float(offset.get("z"), 0.0),
    }


def _sanitize_parameter_map(parameters):
    out = OrderedDict()
    if not isinstance(parameters, dict):
        return out
    for key, data in parameters.items():
        name = str(key or "").strip()
        if not name:
            continue
        if isinstance(data, dict):
            storage_type = str(data.get("storage_type") or "String")
            value = "" if data.get("value") is None else str(data.get("value"))
            read_only = bool(data.get("read_only"))
        else:
            storage_type = "String"
            value = "" if data is None else str(data)
            read_only = False
        out[name] = {"storage_type": storage_type, "value": value, "read_only": read_only}
    return out


def _sanitize_annotation(annotation):
    if not isinstance(annotation, dict):
        return None
    kind = str(annotation.get("kind") or "").strip().lower()
    if kind not in ("tag", "keynote"):
        return None

    symbol_id = str(annotation.get("symbol_id") or annotation.get("element_type_id") or "").strip()
    if kind == "keynote" and not symbol_id:
        return None

    return {
        "kind": kind,
        "name": str(annotation.get("name") or ("Tag" if kind == "tag" else "GA_Keynote Symbol_CED")).strip(),
        "symbol_id": symbol_id,
        "offset": _sanitize_offset(annotation.get("offset") or {}),
        "parameters": _sanitize_parameter_map(annotation.get("parameters") or {}),
    }


def _sanitize_annotation_list(raw_list):
    annotations = []
    for raw in raw_list or []:
        ann = _sanitize_annotation(raw)
        if ann:
            annotations.append(ann)
    return annotations


def _sanitize_entry_with_annotations(helper, raw_entry):
    if not isinstance(raw_entry, dict):
        return None
    base = helper._sanitize_template_entry(raw_entry)
    if not isinstance(base, dict):
        return None
    base["annotations"] = _sanitize_annotation_list(raw_entry.get("annotations") or [])
    return base


def _sanitize_entry_list(helper, raw_entries):
    rows = []
    if isinstance(raw_entries, dict):
        source = []
        for map_key, map_value in raw_entries.items():
            row = map_value if isinstance(map_value, dict) else {}
            if row and not str(row.get("id") or "").strip():
                row = dict(row)
                row["id"] = str(map_key or "").strip()
            source.append(row)
    else:
        source = list(raw_entries or [])

    for raw in source:
        entry = _sanitize_entry_with_annotations(helper, raw if isinstance(raw, dict) else {})
        if entry:
            rows.append(entry)
    return rows


def _sanitize_type_elements(helper, raw_map):
    data = {bucket: [] for bucket in BUCKETS}
    raw_map = raw_map if isinstance(raw_map, dict) else {}
    for raw_bucket, raw_entries in raw_map.items():
        bucket = helper._normalize_bucket(raw_bucket, default=None)
        if not bucket:
            continue
        data[bucket].extend(_sanitize_entry_list(helper, raw_entries))
    return data


def _sanitize_space_overrides(helper, raw_map):
    data = {}
    raw_map = raw_map if isinstance(raw_map, dict) else {}
    for space_key, raw_entries in raw_map.items():
        key = str(space_key or "").strip()
        if not key:
            continue
        entries = _sanitize_entry_list(helper, raw_entries)
        if entries:
            data[key] = entries
    return data


def _resolve_template_bucket(helper, space_row, type_elements):
    saved_bucket = helper._normalize_bucket(space_row.get("bucket"), default="Other")
    if type_elements.get(saved_bucket):
        return saved_bucket

    inferred = helper._infer_bucket_from_space_row(space_row)
    if inferred in BUCKETS and type_elements.get(inferred):
        return inferred

    if type_elements.get("Other"):
        return "Other"

    return saved_bucket


def _request_rows(helper, spaces, type_elements, space_overrides):
    rows = []

    for space_row in spaces:
        bucket = _resolve_template_bucket(helper, space_row, type_elements)
        type_entries = type_elements.get(bucket) or []
        override_entries = space_overrides.get(space_row.get("space_key")) or []

        req_space = space_row
        original_bucket = helper._normalize_bucket(space_row.get("bucket"), default="Other")
        if bucket != original_bucket:
            req_space = dict(space_row)
            req_space["bucket"] = bucket
            req_space["resolved_from_saved_bucket"] = original_bucket

        for entry in (type_entries + override_entries):
            annotations = _sanitize_annotation_list(entry.get("annotations") or [])
            if annotations:
                e = dict(entry)
                e["annotations"] = annotations
                e["profile_duplicate_id"] = _try_int(entry.get("profile_duplicate_id") or entry.get("duplicate_id"), default=None)
                e["_request_uid"] = uuid.uuid4().hex
                rows.append((req_space, e))
    return rows
def _collect_space_context(helper, doc, assignments):
    sources = helper._collect_all_space_sources(doc)
    spaces = []
    door_points_by_source = {}
    door_rows_by_source = {}
    source_transform_by_label = {}

    for source in sources:
        source_doc = source.get("source_doc")
        if source_doc is None:
            continue

        source_transform = source.get("source_transform")
        source_label = source.get("source_label") or "Host Spaces"
        source_transform_by_label[source_label] = source_transform

        source_spaces = helper._collect_classified_spaces(
            doc,
            source_doc,
            assignments,
            source_transform=source_transform,
            source_label=source_label,
        )
        if source_spaces:
            spaces.extend(source_spaces)

        if source_label not in door_points_by_source:
            door_points_by_source[source_label] = helper._collect_door_points(source_doc)
        if source_label not in door_rows_by_source:
            door_rows_by_source[source_label] = helper._collect_door_rows(source_doc)

        door_points_by_source[source_label] = helper._clean_origin_points(door_points_by_source.get(source_label) or [])
        door_rows_by_source[source_label] = helper._clean_origin_rows(door_rows_by_source.get(source_label) or [])

    host_door_points = list(door_points_by_source.get("Host Spaces") or [])
    if host_door_points:
        for source in sources:
            source_label = source.get("source_label") or "Host Spaces"
            if source_label == "Host Spaces":
                continue
            existing = door_points_by_source.get(source_label) or []
            if existing:
                continue
            transformed = helper._transform_points_to_source(host_door_points, source.get("source_transform"))
            if transformed:
                door_points_by_source[source_label] = helper._clean_origin_points(transformed)

    host_door_rows = helper._clean_origin_rows(door_rows_by_source.get("Host Spaces") or [])
    for source in sources:
        source_label = source.get("source_label") or "Host Spaces"
        if source_label == "Host Spaces":
            continue

        source_rows = helper._clean_origin_rows(door_rows_by_source.get(source_label) or [])
        if source_rows:
            door_rows_by_source[source_label] = source_rows
            continue

        source_transform = source.get("source_transform")
        if host_door_rows:
            transformed_rows = helper._clean_origin_rows(helper._transform_door_rows_to_source(host_door_rows, source_transform))
            if transformed_rows:
                door_rows_by_source[source_label] = transformed_rows
                continue

        source_points = helper._clean_origin_points(door_points_by_source.get(source_label) or [])
        if source_points:
            door_rows_by_source[source_label] = helper._door_rows_from_points(source_points)

    all_host_door_points = []
    for source_label, source_points in list(door_points_by_source.items()):
        if not source_points:
            continue

        if source_label == "Host Spaces":
            host_points = list(source_points)
        else:
            host_points = []
            source_transform = source_transform_by_label.get(source_label)
            for source_pt in source_points:
                host_pt = helper._to_host_point(source_pt, source_transform)
                if host_pt is not None:
                    host_points.append(host_pt)

        if not host_points:
            continue

        door_points_by_source["__host::" + source_label] = host_points
        all_host_door_points.extend(host_points)

    if all_host_door_points:
        door_points_by_source["__all_host__"] = all_host_door_points

    return spaces, sources, door_points_by_source, door_rows_by_source


def _compute_target_point(helper, space_row, entry, door_points_by_source, door_rows_by_source):
    source_space = space_row.get("space")
    source_label = str(space_row.get("source_label") or "Host Spaces")
    source_transform = space_row.get("source_transform")
    rule = str(entry.get("placement_rule") or helper.DEFAULT_PLACEMENT_OPTION).strip()

    source_door_points = door_points_by_source.get(source_label) or []
    source_door_rows = door_rows_by_source.get(source_label) or []

    if not source_door_points:
        host_candidates = door_points_by_source.get("__host::" + source_label) or door_points_by_source.get("__all_host__") or []
        if host_candidates:
            source_door_points = helper._clean_origin_points(helper._transform_points_to_source(host_candidates, source_transform))

    if not source_door_rows and source_door_points:
        source_door_rows = helper._door_rows_from_points(source_door_points)

    if rule == "One Foot off doorway wall":
        center_source = helper._space_center_robust(source_space)
        source_point, _ = helper._one_foot_off_hinge_point(source_space, source_door_rows, fallback_center=center_source, return_row=True)
        if source_point is None or helper._is_origin_point(source_point):
            return None, "No valid doorway-wall placement point could be resolved"

        host_point = helper._to_host_point(source_point, source_transform)
        if host_point is None or helper._is_origin_point(host_point):
            return None, "Doorway-wall point could not be transformed to host coordinates"
        return host_point, ""

    source_point = helper._compute_placement_point(source_space, rule, source_door_points)
    if source_point is None:
        source_point = helper._space_center_robust(source_space)

    host_point = helper._to_host_point(source_point, source_transform) if source_point is not None else None
    if host_point is None:
        return None, "Could not calculate placement point"
    if helper._is_origin_point(host_point):
        return None, "Computed placement point resolved to origin"
    return host_point, ""



def _request_row_key(space_row, entry):
    source_label = str(space_row.get("source_label") or "Host Spaces")
    space_key = str(space_row.get("space_key") or space_row.get("unique_id") or space_row.get("space_id") or "").strip()
    entry_key = str(entry.get("_request_uid") or entry.get("entry_uid") or entry.get("id") or "").strip()
    return "{}|{}|{}".format(source_label, space_key, entry_key)


def _host_point_in_space_row(helper, space_row, host_point):
    if host_point is None:
        return False

    source_space = space_row.get("space")
    if source_space is None:
        return False

    source_transform = space_row.get("source_transform")
    if source_transform is None:
        source_point = host_point
    else:
        try:
            source_point = source_transform.Inverse.OfPoint(host_point)
        except Exception:
            source_point = None
    if source_point is None:
        return False

    try:
        return bool(helper._point_in_space(source_space, source_point))
    except Exception:
        return False


def _entry_target_host_point(helper, space_row, entry, door_points_by_source, door_rows_by_source):
    target_point, _reason = _compute_target_point(helper, space_row, entry, door_points_by_source, door_rows_by_source)
    if target_point is not None:
        return target_point

    source_center = helper._space_center_robust(space_row.get("space"))
    if source_center is None:
        return None
    return helper._to_host_point(source_center, space_row.get("source_transform"))


def _assign_entries_to_candidates(helper, entry_rows, candidate_rows, target_points):
    assignment = {}
    if not candidate_rows:
        return assignment

    remaining_entries = list(range(len(entry_rows)))
    remaining_candidates = list(range(len(candidate_rows)))

    while remaining_entries and remaining_candidates:
        best = None
        for ei in remaining_entries:
            target = target_points.get(ei)
            for ci in remaining_candidates:
                cand_point = candidate_rows[ci][1]
                dist = helper._distance_xy(cand_point, target) if target is not None else float("inf")
                if best is None or dist < best[0]:
                    best = (dist, ei, ci)

        if best is None:
            break

        _dist, ei, ci = best
        assignment[ei] = candidate_rows[ci]
        remaining_entries.remove(ei)
        remaining_candidates.remove(ci)

    return assignment


def _build_entry_assignment_map(
        helper,
        doc,
        all_requests,
        door_points_by_source,
        door_rows_by_source):
    grouped = OrderedDict()

    for idx, pair in enumerate(all_requests or []):
        space_row, entry = pair
        kind = str(entry.get("kind") or "").strip().lower()
        type_id = str(entry.get("element_type_id") or "").strip()
        if kind not in ("family_type", "model_group") or not type_id:
            continue

        source_label = str(space_row.get("source_label") or "Host Spaces")
        space_key = str(space_row.get("space_key") or space_row.get("unique_id") or space_row.get("space_id") or "").strip()
        group_key = "{}|{}|{}|{}".format(source_label, space_key, kind, type_id)
        grouped.setdefault(group_key, []).append((idx, space_row, entry))

    assignment_map = {}
    candidate_cache = {}

    for _group_key, rows in grouped.items():
        if not rows:
            continue

        probe_entry = rows[0][2]
        all_candidates = _candidate_rows_for_entry(helper, doc, probe_entry, candidate_cache)
        if not all_candidates:
            for _idx, space_row, entry in rows:
                assignment_map[_request_row_key(space_row, entry)] = {
                    "element": None,
                    "point": None,
                    "target": _entry_target_host_point(helper, space_row, entry, door_points_by_source, door_rows_by_source),
                }
            continue

        group_space_row = rows[0][1]
        in_space_candidates = []
        for cand in all_candidates:
            cand_point = cand[1]
            if _host_point_in_space_row(helper, group_space_row, cand_point):
                in_space_candidates.append(cand)
        if in_space_candidates and len(in_space_candidates) >= len(rows):
            candidate_rows = in_space_candidates
        else:
            candidate_rows = all_candidates

        entry_rows = []
        target_points = {}
        for local_i, (_idx, space_row, entry) in enumerate(rows):
            entry_rows.append((space_row, entry))
            target_points[local_i] = _entry_target_host_point(helper, space_row, entry, door_points_by_source, door_rows_by_source)

        local_assignment = {}
        used_candidate_indexes = set()

        # First pass: exact match by SPACE BASED duplicate id from Element Linker.
        for local_i, (_idx, _space_row, entry) in enumerate(rows):
            expected_dup = _try_int((entry or {}).get("profile_duplicate_id"), default=None)
            if expected_dup is None:
                continue

            target_pt = target_points.get(local_i)
            best_ci = None
            best_dist = None
            for ci, cand in enumerate(candidate_rows):
                if ci in used_candidate_indexes:
                    continue

                cand_dup = _try_int(cand[2], default=None) if len(cand) > 2 else None
                if cand_dup != expected_dup:
                    continue

                cand_point = cand[1]
                dist = helper._distance_xy(cand_point, target_pt) if target_pt is not None else 0.0
                if best_ci is None or dist < best_dist:
                    best_ci = ci
                    best_dist = dist

            if best_ci is not None:
                local_assignment[local_i] = candidate_rows[best_ci]
                used_candidate_indexes.add(best_ci)

        # Second pass: nearest one-to-one for rows without a duplicate-id match.
        remaining_entry_rows = []
        remaining_target_points = {}
        remaining_index_map = {}
        for local_i, wrapped in enumerate(rows):
            if local_i in local_assignment:
                continue
            sub_i = len(remaining_entry_rows)
            remaining_entry_rows.append((wrapped[1], wrapped[2]))
            remaining_target_points[sub_i] = target_points.get(local_i)
            remaining_index_map[sub_i] = local_i

        remaining_candidate_rows = [cand for ci, cand in enumerate(candidate_rows) if ci not in used_candidate_indexes]
        if remaining_entry_rows and remaining_candidate_rows:
            secondary_assignment = _assign_entries_to_candidates(
                helper,
                remaining_entry_rows,
                remaining_candidate_rows,
                remaining_target_points,
            )
            for sub_i, cand in secondary_assignment.items():
                local_i = remaining_index_map.get(sub_i)
                if local_i is not None:
                    local_assignment[local_i] = cand

        for local_i, (_idx, space_row, entry) in enumerate(rows):
            req_key = _request_row_key(space_row, entry)
            assigned = local_assignment.get(local_i)
            if assigned is None:
                expected_dup = _try_int((entry or {}).get("profile_duplicate_id"), default=None)
                seen_ids = []
                for cand in candidate_rows:
                    cand_dup = _try_int(cand[2], default=None) if len(cand) > 2 else None
                    if cand_dup is not None and cand_dup not in seen_ids:
                        seen_ids.append(cand_dup)
                id_note = "Expected SPACE BASED ID {}. Candidate IDs in scope: {}".format(
                    expected_dup if expected_dup is not None else "<none>",
                    seen_ids if seen_ids else "<none>",
                )
                assignment_map[req_key] = {
                    "element": None,
                    "point": None,
                    "target": target_points.get(local_i),
                    "reason": "No unique placed element match was found for this profile entry in the space. " + id_note,
                }
            else:
                cand_elem = assigned[0]
                cand_point = assigned[1]
                assignment_map[req_key] = {
                    "element": cand_elem,
                    "point": cand_point,
                    "target": target_points.get(local_i),
                }

    return assignment_map
def _candidate_rows_for_entry(helper, doc, entry, cache):
    kind = str(entry.get("kind") or "").strip().lower()
    type_id = str(entry.get("element_type_id") or "").strip()
    key = "{}|{}".format(kind, type_id)
    if key in cache:
        return cache.get(key) or []

    rows = []
    if kind == "family_type":
        try:
            elements = list(FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType())
        except Exception:
            elements = []
        for element in elements:
            symbol = getattr(element, "Symbol", None)
            symbol_id = _element_id_value(getattr(symbol, "Id", None), default="")
            if symbol_id != type_id:
                continue
            point = helper._element_location_point(element, fallback_point=None)
            if point is None:
                continue
            rows.append((element, point, _read_space_based_duplicate_id(element)))

    elif kind == "model_group":
        try:
            elements = list(FilteredElementCollector(doc).OfClass(Group).WhereElementIsNotElementType())
        except Exception:
            elements = []
        for element in elements:
            group_type = getattr(element, "GroupType", None)
            group_type_id = _element_id_value(getattr(group_type, "Id", None), default="")
            if group_type_id != type_id:
                continue
            point = helper._element_location_point(element, fallback_point=None)
            if point is None:
                continue
            rows.append((element, point, _read_space_based_duplicate_id(element)))

    cache[key] = rows
    return rows


def _pick_target_element(helper, space_row, entry, target_point, cache):
    rows = _candidate_rows_for_entry(helper, revit.doc, entry, cache)
    if not rows:
        return None, None, "No matching placed element found for template type"

    best = None
    best_score = None
    source_space = space_row.get("space")
    source_transform = space_row.get("source_transform")

    for row in rows:
        element = row[0]
        point = row[1]
        score = helper._distance_xy(point, target_point) if target_point is not None else 0.0
        if source_space is not None:
            if source_transform is None:
                source_point = point
            else:
                try:
                    source_point = source_transform.Inverse.OfPoint(point)
                except Exception:
                    source_point = None
            inside = helper._point_in_space(source_space, source_point) if source_point is not None else False
            if not inside:
                score += 1000000.0

        if best is None or score < best_score:
            best = (element, point)
            best_score = score

    if best is None:
        return None, None, "Could not resolve target element for annotation"
    return best[0], best[1], ""


def _point_with_offset(base_point, offset):
    if base_point is None:
        return None
    offset = _sanitize_offset(offset or {})
    return XYZ(base_point.X + offset["x"], base_point.Y + offset["y"], base_point.Z + offset["z"])


def _ensure_symbol_active(symbol):
    if symbol is None:
        return
    try:
        if not bool(symbol.IsActive):
            symbol.Activate()
    except Exception:
        pass


def _get_symbol_by_id(doc, symbol_id):
    sid = _try_int(symbol_id, default=None)
    if sid is None:
        return None
    try:
        elem = doc.GetElement(ElementId(sid))
    except Exception:
        elem = None
    return elem if isinstance(elem, FamilySymbol) else None


def _place_tag(doc, view, target_element, symbol_id, point):
    if target_element is None:
        return None, "Missing target element for tag"
    try:
        ref = Reference(target_element)
    except Exception as exc:
        return None, "Could not reference target element for tag: {}".format(exc)

    first_error = ""
    for mode in (TagMode.TM_ADDBY_CATEGORY, TagMode.TM_ADDBY_MULTICATEGORY):
        try:
            tag = IndependentTag.Create(doc, view.Id, ref, False, mode, TagOrientation.Horizontal, point)
            if tag is None:
                continue
            type_int = _try_int(symbol_id, default=None)
            if type_int is not None:
                try:
                    tag.ChangeTypeId(ElementId(type_int))
                except Exception:
                    pass
            return tag, ""
        except Exception as exc:
            if not first_error:
                first_error = str(exc)

    return None, first_error or "Tag creation failed"


def _target_level(doc, target_element, view):
    if target_element is not None:
        try:
            level_id = getattr(target_element, "LevelId", None)
            if level_id is not None and int(level_id.IntegerValue) > 0:
                level = doc.GetElement(level_id)
                if level is not None:
                    return level
        except Exception:
            pass
    try:
        level = getattr(view, "GenLevel", None)
        if level is not None:
            return level
    except Exception:
        pass
    return None


def _place_keynote(doc, view, target_element, annotation, point):
    symbol = _get_symbol_by_id(doc, annotation.get("symbol_id"))
    if symbol is None:
        return None, "Keynote symbol id '{}' was not found".format(annotation.get("symbol_id") or "")

    _ensure_symbol_active(symbol)

    errors = []
    try:
        elem = doc.Create.NewFamilyInstance(point, symbol, view)
        if elem is not None:
            return elem, ""
    except Exception as exc:
        errors.append(str(exc))

    level = _target_level(doc, target_element, view)
    if level is not None and RevitStructuralType is not None:
        try:
            elem = doc.Create.NewFamilyInstance(point, symbol, level, RevitStructuralType.NonStructural)
            if elem is not None:
                return elem, ""
        except Exception as exc:
            errors.append(str(exc))

    return None, errors[0] if errors else "Keynote placement failed"


def _format_xyz(point):
    if point is None:
        return "<none>"
    try:
        return "X={:.3f}, Y={:.3f}, Z={:.3f}".format(float(point.X), float(point.Y), float(point.Z))
    except Exception:
        return "<invalid>"


def _summary_lines(spaces, requests, selected_profile_count, placed_rows, failures, source_label):
    counts = OrderedDict((bucket, 0) for bucket in BUCKETS)
    for row in spaces:
        bucket = str(row.get("bucket") or "Other")
        if bucket not in counts:
            bucket = "Other"
        counts[bucket] += 1

    lines = [
        "Placed space annotations.",
        "Storage ID: {}".format(CLASSIFICATION_STORAGE_ID),
        "Source: {}".format(source_label),
        "",
        "Classified spaces processed: {}".format(len(spaces)),
        "Resolved profile requests with annotations: {}".format(len(requests)),
        "Selected profile templates: {}".format(selected_profile_count),
        "Annotations successfully placed: {}".format(len(placed_rows)),
        "Annotation placement failures: {}".format(len(failures)),
        "",
        "Space buckets with counts:",
    ]
    for bucket in BUCKETS:
        if counts.get(bucket, 0) > 0:
            lines.append(" - {}: {}".format(bucket, counts[bucket]))

    if placed_rows:
        lines.append("")
        lines.append("Placed annotations (first 60):")
        for row in placed_rows[:60]:
            lines.append(
                " - {} - {} | {} | {} | id {} | on element {} | {}".format(
                    row.get("space_number") or "<No Number>",
                    row.get("space_name") or "<Unnamed>",
                    row.get("entry_name") or "<Entry>",
                    row.get("annotation_name") or "Annotation",
                    row.get("annotation_id") or "<unknown>",
                    row.get("target_element_id") or "<unknown>",
                    row.get("point") or "<none>",
                )
            )
        if len(placed_rows) > 60:
            lines.append(" - ... ({} more)".format(len(placed_rows) - 60))

    if failures:
        lines.append("")
        lines.append("Placement failures (first 30):")
        for row in failures[:30]:
            space_row = row.get("space_row") or {}
            entry = row.get("entry") or {}
            annotation = row.get("annotation") or {}
            lines.append(
                " - {} - {} | {} | {} | {}".format(
                    space_row.get("space_number") or "<No Number>",
                    space_row.get("space_name") or "<Unnamed>",
                    entry.get("name") or entry.get("id") or "<Entry>",
                    annotation.get("name") or annotation.get("kind") or "annotation",
                    row.get("reason") or "Unknown error",
                )
            )
        if len(failures) > 30:
            lines.append(" - ... ({} more)".format(len(failures) - 30))

    return lines


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    view = getattr(doc, "ActiveView", None)
    if view is None:
        forms.alert("No active view detected.", title=TITLE)
        return
    try:
        if bool(view.IsTemplate):
            forms.alert("Active view is a template. Open a model view first.", title=TITLE)
            return
    except Exception:
        pass

    helper, helper_error = _load_helper()
    if helper is None:
        forms.alert(helper_error or "Failed to load helper script.", title=TITLE)
        return

    if getattr(helper, "ExtensibleStorage", None) is None:
        forms.alert("Failed to load ExtensibleStorage library from CEDLib.lib.", title=TITLE)
        return

    payload = helper._load_latest_space_payload(doc)
    if not isinstance(payload, dict):
        forms.alert("No saved space data found.\n\nRun Classify Spaces and save first.", title=TITLE)
        return

    assignments = payload.get("space_assignments")
    if not isinstance(assignments, dict) or not assignments:
        forms.alert("No space assignments found in saved data.\n\nRun Classify Spaces and save first.", title=TITLE)
        return

    type_elements = _sanitize_type_elements(helper, payload.get(KEY_TYPE_ELEMENTS) or {})
    space_overrides = _sanitize_space_overrides(helper, payload.get(KEY_SPACE_OVERRIDES) or {})

    spaces, sources, door_points_by_source, door_rows_by_source = _collect_space_context(helper, doc, assignments)
    if not spaces:
        forms.alert("No spaces were found in host or loaded linked models.", title=TITLE)
        return

    all_profile_requests = helper._request_rows(spaces, type_elements, space_overrides)
    if not all_profile_requests:
        forms.alert("No profile entries were resolved for current spaces.", title=TITLE)
        return

    requests = _request_rows(helper, spaces, type_elements, space_overrides)
    if not requests:
        forms.alert("No profile entries with saved annotations were resolved for current spaces.", title=TITLE)
        return

    selected_profile_keys = helper._prompt_profile_selection(requests)
    if selected_profile_keys is None:
        return
    if not selected_profile_keys:
        forms.alert("No space profiles were selected.", title=TITLE)
        return

    selected_requests = [
        (space_row, entry)
        for space_row, entry in requests
        if helper._request_profile_key(space_row, entry) in selected_profile_keys
    ]
    if not selected_requests:
        forms.alert("No annotation requests matched the selected profiles.", title=TITLE)
        return

    linked_count = max(0, len(sources) - 1)
    source_label = "Host + {} linked model{}".format(linked_count, "" if linked_count == 1 else "s") if linked_count > 0 else "Host Spaces"

    if not forms.alert(
        "\n".join(
            [
                "This will place saved annotations for selected profile entries.",
                "",
                "Source: {}".format(source_label),
                "Classified spaces: {}".format(len(spaces)),
                "Selected profile templates: {}".format(len(selected_profile_keys)),
                "Resolved annotation requests: {}".format(len(selected_requests)),
                "",
                "Continue?",
            ]
        ),
        title=TITLE,
        yes=True,
        no=True,
    ):
        return

    assignment_map = _build_entry_assignment_map(
        helper,
        doc,
        selected_requests,
        door_points_by_source,
        door_rows_by_source,
    )
    candidate_cache = {}
    placed_rows = []
    failures = []

    tx = Transaction(doc, TITLE)
    tx.Start()
    try:
        for space_row, entry in selected_requests:
            req_key = _request_row_key(space_row, entry)
            has_assignment = req_key in assignment_map
            assigned = assignment_map.get(req_key) or {}

            target_point = assigned.get("target")
            point_reason = ""
            if target_point is None:
                target_point, point_reason = _compute_target_point(helper, space_row, entry, door_points_by_source, door_rows_by_source)
                if target_point is None:
                    source_center = helper._space_center_robust(space_row.get("space"))
                    target_point = helper._to_host_point(source_center, space_row.get("source_transform")) if source_center is not None else None

            target_element = assigned.get("element")
            target_element_point = assigned.get("point")
            target_reason = str(assigned.get("reason") or "")
            if (target_element is None or target_element_point is None) and (not has_assignment):
                target_element, target_element_point, target_reason = _pick_target_element(helper, space_row, entry, target_point, candidate_cache)

            if target_element is None:
                failures.append(
                    {
                        "space_row": space_row,
                        "entry": entry,
                        "annotation": {},
                        "reason": target_reason or point_reason or "No target element could be resolved",
                    }
                )
                continue

            for annotation in _sanitize_annotation_list(entry.get("annotations") or []):
                ann_point = _point_with_offset(target_element_point, annotation.get("offset") or {})
                if ann_point is None:
                    failures.append(
                        {
                            "space_row": space_row,
                            "entry": entry,
                            "annotation": annotation,
                            "reason": "Could not calculate annotation point",
                        }
                    )
                    continue

                kind = str(annotation.get("kind") or "").strip().lower()
                created = None
                reason = ""
                if kind == "tag":
                    created, reason = _place_tag(doc, view, target_element, annotation.get("symbol_id"), ann_point)
                elif kind == "keynote":
                    created, reason = _place_keynote(doc, view, target_element, annotation, ann_point)
                    if created is not None:
                        try:
                            helper._apply_parameter_overrides(created, annotation.get("parameters") or {})
                        except Exception:
                            pass
                else:
                    reason = "Unsupported annotation kind '{}'".format(kind or "<blank>")

                if created is None:
                    failures.append(
                        {
                            "space_row": space_row,
                            "entry": entry,
                            "annotation": annotation,
                            "reason": reason or "Annotation placement returned no element",
                        }
                    )
                    continue

                placed_rows.append(
                    {
                        "space_number": space_row.get("space_number") or "<No Number>",
                        "space_name": space_row.get("space_name") or "<Unnamed>",
                        "entry_name": entry.get("name") or entry.get("id") or "<Entry>",
                        "annotation_name": annotation.get("name") or annotation.get("kind") or "annotation",
                        "annotation_id": _element_id_value(getattr(created, "Id", None), default="<unknown>"),
                        "target_element_id": _element_id_value(getattr(target_element, "Id", None), default="<unknown>"),
                        "point": _format_xyz(ann_point),
                    }
                )

        tx.Commit()
    except Exception:
        tx.RollBack()
        raise

    forms.alert(
        "\n".join(_summary_lines(spaces, selected_requests, len(selected_profile_keys), placed_rows, failures, source_label)),
        title=TITLE,
    )


if __name__ == "__main__":
    main()






