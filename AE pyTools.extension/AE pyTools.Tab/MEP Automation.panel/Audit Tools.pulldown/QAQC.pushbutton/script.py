# -*- coding: utf-8 -*-
"""
QA/QC Relationship Audit
-----------------------
Sync-independent audit for YAML profile coverage and Element Linker integrity.
"""

import imp
import math
import os
import re
import sys

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    RevitLinkInstance,
    Transaction,
    XYZ,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Collections.Generic import List

output = script.get_output()
output.close_others()

TITLE = "QA/QC Relationship Audit"
LOG = script.get_logger()
_MODELLESS_WINDOW = None
_FOLLOW_PARENT_MODULE = None
_OPTIMIZE_MODULE = None
_ADJUST_HANDLER = None
_ADJUST_EXTERNAL_EVENT = None

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402

LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
PARENT_ID_KEYS = ("Parent ElementId", "Parent Element ID")
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"

INLINE_LINKER_PATTERN = re.compile(
    r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|"
    r"Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
    r"Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation)\s*:\s*",
    re.IGNORECASE,
)

FAR_FROM_PARENT_THRESHOLD_FT = 10.0


def _element_id_value(elem_id, default=None):
    if elem_id is None:
        return default
    for attr in ("Value", "IntegerValue"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return int(value)
        except Exception:
            try:
                return value
            except Exception:
                continue
    return default


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _normalize_name_ignoring_default_suffix(value):
    normalized = _normalize_name(value)
    if not normalized:
        return ""
    # Treat these suffixes as equivalent to no suffix for Tab 5 host-name drift checks:
    # ": Default", ": Default 2", ": DefaultType"
    cleaned = re.sub(
        r"\s*:\s*(?:default(?:\s*\d+)?|defaulttype|default\s*type)$",
        "",
        normalized,
        flags=re.IGNORECASE,
    ).strip()

    # Treat "Family : Family" and "Family :" as equivalent to just "Family".
    if ":" in cleaned:
        left, right = cleaned.split(":", 1)
        family = (left or "").strip()
        type_name = (right or "").strip()
        if family:
            if not type_name:
                return family
            if type_name == family:
                return family
    return cleaned


def _try_int(value):
    if value in (None, ""):
        return None
    try:
        return int(value)
    except Exception:
        try:
            return int(float(value))
        except Exception:
            return None


def _parse_xyz(value):
    if not value:
        return None
    parts = [part.strip() for part in str(value).split(",")]
    if len(parts) != 3:
        return None
    try:
        return XYZ(float(parts[0]), float(parts[1]), float(parts[2]))
    except Exception:
        return None


def _get_symbol(elem):
    symbol = getattr(elem, "Symbol", None)
    if symbol is not None:
        return symbol
    try:
        type_id = elem.GetTypeId()
    except Exception:
        type_id = None
    if not type_id:
        return None
    try:
        return elem.Document.GetElement(type_id)
    except Exception:
        return None


def _name_variants(elem):
    names = set()
    if elem is None:
        return names
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            names.add(raw_name)
    except Exception:
        pass

    if isinstance(elem, FamilyInstance):
        symbol = _get_symbol(elem)
        family = getattr(symbol, "Family", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if family_name and type_name:
            names.add(u"{} : {}".format(family_name, type_name))
            names.add(u"{} : {}".format(type_name, family_name))
        if family_name:
            names.add(family_name)
        if type_name:
            names.add(type_name)
    elif isinstance(elem, Group):
        group_type = getattr(elem, "GroupType", None)
        gname = getattr(group_type, "Name", None) if group_type else None
        if gname:
            names.add(gname)

    return {_normalize_name(name) for name in names if _normalize_name(name)}


def _element_label(elem):
    if elem is None:
        return "<missing>"
    label = None
    if isinstance(elem, FamilyInstance):
        symbol = _get_symbol(elem)
        family = getattr(symbol, "Family", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if family_name and type_name:
            label = "{} : {}".format(family_name, type_name)
    if not label:
        try:
            label = getattr(elem, "Name", None)
        except Exception:
            label = None
    if not label:
        label = "<element>"
    try:
        elem_id = _element_id_value(elem.Id)
    except Exception:
        elem_id = None
    if elem_id is not None:
        return "{} (Id:{})".format(label, elem_id)
    return str(label)


def _family_type_label(elem):
    if elem is None:
        return ""
    if isinstance(elem, FamilyInstance):
        symbol = _get_symbol(elem)
        family = getattr(symbol, "Family", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if family_name and type_name:
            return u"{} : {}".format(family_name, type_name)
    try:
        name = getattr(elem, "Name", None)
    except Exception:
        name = None
    return str(name or "")


def _get_element_point(elem):
    if elem is None:
        return None
    loc = getattr(elem, "Location", None)
    if loc is not None:
        point = getattr(loc, "Point", None)
        if point is not None:
            return point
        curve = getattr(loc, "Curve", None)
        if curve is not None:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                try:
                    return curve.GetEndPoint(0)
                except Exception:
                    pass
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is not None:
        try:
            return (bbox.Min + bbox.Max) * 0.5
        except Exception:
            return None
    return None


def _xy_distance(point_a, point_b):
    if point_a is None or point_b is None:
        return None
    try:
        dx = float(point_a.X) - float(point_b.X)
        dy = float(point_a.Y) - float(point_b.Y)
    except Exception:
        return None
    return math.sqrt(dx * dx + dy * dy)


def _transform_point(transform, point):
    if transform is None or point is None:
        return point
    try:
        return transform.OfPoint(point)
    except Exception:
        return point


def _doc_key(doc):
    if doc is None:
        return None
    try:
        return doc.PathName or doc.Title
    except Exception:
        return None


def _get_link_transform(link_inst):
    if link_inst is None:
        return None
    try:
        return link_inst.GetTotalTransform()
    except Exception:
        try:
            return link_inst.GetTransform()
        except Exception:
            return None


def _combine_transform(parent_transform, child_transform):
    if parent_transform is None:
        return child_transform
    if child_transform is None:
        return parent_transform
    try:
        return parent_transform.Multiply(child_transform)
    except Exception:
        return None


def _walk_link_documents(doc, parent_transform, doc_chain):
    if doc is None:
        return
    key = _doc_key(doc)
    if key and key in doc_chain:
        return
    next_chain = set(doc_chain or set())
    if key:
        next_chain.add(key)
    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        transform = _combine_transform(parent_transform, _get_link_transform(link_inst))
        yield link_doc, transform
        for nested in _walk_link_documents(link_doc, transform, next_chain):
            yield nested


def _iter_link_documents(doc):
    for link_doc, transform in _walk_link_documents(doc, None, set()):
        yield link_doc, transform


def _collect_family_and_group_instances(doc):
    items = []
    seen = set()
    for cls in (FamilyInstance, Group):
        try:
            collector = FilteredElementCollector(doc).OfClass(cls).WhereElementIsNotElementType()
        except Exception:
            continue
        for elem in collector:
            try:
                elem_id = _element_id_value(elem.Id)
            except Exception:
                elem_id = None
            if elem_id is None:
                continue
            marker = (cls.__name__, elem_id)
            if marker in seen:
                continue
            seen.add(marker)
            items.append(elem)
    return items


def _get_linker_text(elem):
    if elem is None:
        return ""
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        text = None
        try:
            text = param.AsString()
        except Exception:
            text = None
        if not text:
            try:
                text = param.AsValueString()
            except Exception:
                text = None
        if text and str(text).strip():
            return str(text)
    return ""


def _parse_linker_payload(payload_text):
    if not payload_text:
        return {}
    text = str(payload_text)
    entries = {}
    if "\n" in text:
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, remainder = line.partition(":")
            entries[key.strip()] = remainder.strip()
    else:
        matches = list(INLINE_LINKER_PATTERN.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key.strip()] = value.strip(" ,")

    parent_element_id = None
    for key in PARENT_ID_KEYS:
        if key in entries:
            parent_element_id = _try_int(entries.get(key))
            if parent_element_id is not None:
                break

    return {
        "led_id": (entries.get("Linked Element Definition ID") or "").strip(),
        "set_id": (entries.get("Set Definition ID") or "").strip(),
        "host_name": (entries.get("Host Name") or "").strip(),
        "parent_element_id": parent_element_id,
        "location": _parse_xyz(entries.get("Location XYZ (ft)")),
        "parent_location": _parse_xyz(entries.get("Parent_location")),
    }


def _build_yaml_maps(data):
    profiles = []
    set_to_profile = {}
    led_to_profile = {}
    norm_to_profiles = {}
    group_to_profiles = {}
    profile_to_group = {}

    eq_defs = data.get("equipment_definitions") or []
    for eq in eq_defs:
        if not isinstance(eq, dict):
            continue
        profile_name = (eq.get("name") or eq.get("id") or "").strip()
        if not profile_name:
            continue
        if profile_name not in profiles:
            profiles.append(profile_name)
        norm = _normalize_name(profile_name)
        if norm:
            norm_to_profiles.setdefault(norm, []).append(profile_name)

        group_key = (eq.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not group_key:
            group_key = (eq.get("id") or profile_name).strip()
        profile_to_group[profile_name] = group_key
        group_to_profiles.setdefault(group_key, []).append(profile_name)

        for linked_set in eq.get("linked_sets") or []:
            if not isinstance(linked_set, dict):
                continue
            set_id = (linked_set.get("id") or "").strip()
            if set_id:
                set_to_profile[set_id] = profile_name
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_id = (led.get("id") or "").strip()
                if led_id:
                    led_to_profile[led_id] = profile_name

    profiles.sort(key=lambda value: value.lower())
    for key in norm_to_profiles:
        norm_to_profiles[key] = sorted(set(norm_to_profiles[key]), key=lambda value: value.lower())
    return {
        "profiles": profiles,
        "norm_to_profiles": norm_to_profiles,
        "set_to_profile": set_to_profile,
        "led_to_profile": led_to_profile,
        "group_to_profiles": group_to_profiles,
        "profile_to_group": profile_to_group,
    }


def _resolve_profile_name(payload, yaml_maps):
    if not isinstance(payload, dict):
        return None

    host_name = (payload.get("host_name") or "").strip()
    if host_name:
        host_norm = _normalize_name(host_name)
        host_matches = yaml_maps["norm_to_profiles"].get(host_norm) or []
        if host_matches:
            return host_matches[0]

    set_id = (payload.get("set_id") or "").strip()
    if set_id and set_id in yaml_maps["set_to_profile"]:
        return yaml_maps["set_to_profile"][set_id]

    led_id = (payload.get("led_id") or "").strip()
    if led_id and led_id in yaml_maps["led_to_profile"]:
        return yaml_maps["led_to_profile"][led_id]

    return None


def _collect_parent_candidates(doc, yaml_maps):
    norm_to_profiles = yaml_maps["norm_to_profiles"]
    candidates = []
    by_parent_id = {}
    host_parent_elements = {}
    profile_to_candidates = {}

    def _add_candidate(elem, point, is_linked):
        variants = _name_variants(elem)
        matched_profiles = set()
        for variant in variants:
            for profile_name in norm_to_profiles.get(variant) or []:
                matched_profiles.add(profile_name)
        # Keep linker-derived profile separate so only type-change tabs use it.
        # Tab 2 should remain driven by name-based parent matching.
        linker_profile = None
        linker_text = _get_linker_text(elem)
        if linker_text:
            linker_payload = _parse_linker_payload(linker_text)
            linker_profile = _resolve_profile_name(linker_payload, yaml_maps)
        parent_id = _element_id_value(getattr(elem, "Id", None))
        if parent_id is None:
            return
        candidate = {
            "parent_id": parent_id,
            "display_label": _element_label(elem),
            "is_linked": bool(is_linked),
            "point": point,
            "name_variants": sorted(variants, key=lambda value: value.lower()),
            "matched_profiles": sorted(matched_profiles, key=lambda value: value.lower()),
            "linker_profile": linker_profile,
        }
        candidates.append(candidate)
        by_parent_id.setdefault(parent_id, []).append(candidate)
        if not is_linked and parent_id not in host_parent_elements:
            host_parent_elements[parent_id] = elem
        for profile_name in candidate["matched_profiles"]:
            profile_to_candidates.setdefault(profile_name, []).append(candidate)

    for elem in _collect_family_and_group_instances(doc):
        _add_candidate(elem, _get_element_point(elem), is_linked=False)

    for link_doc, transform in _iter_link_documents(doc):
        for elem in _collect_family_and_group_instances(link_doc):
            point = _transform_point(transform, _get_element_point(elem))
            _add_candidate(elem, point, is_linked=True)

    return {
        "all": candidates,
        "by_parent_id": by_parent_id,
        "host_parent_elements": host_parent_elements,
        "profile_to_candidates": profile_to_candidates,
    }


def _collect_placed_instances(doc, yaml_maps):
    records = []
    for elem in _collect_family_and_group_instances(doc):
        payload_text = _get_linker_text(elem)
        if not payload_text:
            continue
        payload = _parse_linker_payload(payload_text)
        if not payload:
            continue
        if not (payload.get("set_id") or payload.get("led_id") or payload.get("host_name")):
            continue
        profile_name = _resolve_profile_name(payload, yaml_maps)
        child_id = _element_id_value(getattr(elem, "Id", None))
        child_point = _get_element_point(elem) or payload.get("location") or payload.get("parent_location")
        records.append({
            "child_id": child_id,
            "child_label": _element_label(elem),
            "child_point": child_point,
            "profile_name": profile_name,
            "set_id": payload.get("set_id"),
            "led_id": payload.get("led_id"),
            "host_name": payload.get("host_name"),
            "parent_element_id": payload.get("parent_element_id"),
            "payload_location": payload.get("location"),
            "payload_parent_location": payload.get("parent_location"),
        })
    return records


def _candidate_text(candidate):
    if not candidate:
        return ""
    scope = "Linked" if candidate.get("is_linked") else "Host"
    return "{} [{}]".format(candidate.get("display_label") or "<parent>", scope)


def _pick_candidate(candidates):
    if not candidates:
        return None
    host_candidates = [item for item in candidates if not item.get("is_linked")]
    if host_candidates:
        return host_candidates[0]
    return candidates[0]


def _build_row(
    profile,
    description,
    parent_text="",
    child_text="",
    child_id=None,
    parent_id=None,
    snap_point=None,
    adjust_enabled=False,
):
    return {
        "profile": profile or "",
        "description": description or "",
        "parent_text": parent_text or "",
        "child_text": child_text or "",
        "child_id": child_id,
        "parent_id": parent_id,
        "snap_point": snap_point,
        "adjust_enabled": bool(adjust_enabled),
    }


def _build_issue_tabs(doc, data):
    yaml_maps = _build_yaml_maps(data)
    parent_data = _collect_parent_candidates(doc, yaml_maps)
    placed_records = _collect_placed_instances(doc, yaml_maps)

    profiles = yaml_maps["profiles"]
    profile_to_candidates = parent_data["profile_to_candidates"]
    by_parent_id = parent_data["by_parent_id"]
    host_parent_elements = parent_data["host_parent_elements"]

    group_to_profiles = yaml_maps["group_to_profiles"]
    profile_to_group = yaml_maps["profile_to_group"]

    group_to_candidate_ids = {}
    for group_key, members in group_to_profiles.items():
        ids = set()
        for member in members or []:
            for candidate in profile_to_candidates.get(member) or []:
                parent_id = candidate.get("parent_id")
                if parent_id is not None:
                    ids.add(parent_id)
        group_to_candidate_ids[group_key] = ids

    placed_by_profile_parent = {}
    placed_by_group_parent = {}
    for record in placed_records:
        profile_name = record.get("profile_name")
        parent_id = record.get("parent_element_id")
        if profile_name and parent_id is not None:
            placed_by_profile_parent.setdefault(profile_name, set()).add(parent_id)
            group_key = profile_to_group.get(profile_name) or profile_name
            placed_by_group_parent.setdefault(group_key, set()).add(parent_id)

    tab1_rows = []
    tab2_rows = []
    tab3_rows = []
    tab4_rows = []
    tab5_rows = []
    tab6_rows = []

    for profile_name in profiles:
        group_key = profile_to_group.get(profile_name) or profile_name
        if not group_to_candidate_ids.get(group_key):
            tab1_rows.append(
                _build_row(
                    profile=profile_name,
                    description="Profile exists in active YAML/extensible storage, but no matching parent elements were found.",
                )
            )

    seen_missing_child = set()
    for profile_name in profiles:
        group_key = profile_to_group.get(profile_name) or profile_name
        candidate_ids = group_to_candidate_ids.get(group_key) or set()
        if not candidate_ids:
            continue
        placed_parent_ids = placed_by_group_parent.get(group_key) or set()
        if candidate_ids and candidate_ids.issubset(placed_parent_ids):
            continue
        for candidate in profile_to_candidates.get(profile_name) or []:
            parent_id = candidate.get("parent_id")
            if parent_id is None or parent_id in placed_parent_ids:
                continue
            unique_key = (profile_name, parent_id, bool(candidate.get("is_linked")))
            if unique_key in seen_missing_child:
                continue
            seen_missing_child.add(unique_key)
            selectable_parent_id = parent_id if parent_id in host_parent_elements else None
            tab2_rows.append(
                _build_row(
                    profile=profile_name,
                    description="Matching parent found, but no placed child instances are currently tracked for this profile-parent.",
                    parent_text=_candidate_text(candidate),
                    parent_id=selectable_parent_id,
                    snap_point=candidate.get("point"),
                )
            )

    for record in placed_records:
        child_id = record.get("child_id")
        profile_name = record.get("profile_name") or (record.get("host_name") or "<unknown profile>")
        parent_id = record.get("parent_element_id")
        if parent_id is None:
            continue
        if parent_id in by_parent_id:
            continue
        tab3_rows.append(
            _build_row(
                profile=profile_name,
                description="Original parent ID from Element Linker no longer exists in host or linked models.",
                parent_text="Missing Parent Id: {}".format(parent_id),
                child_text=record.get("child_label") or "",
                child_id=child_id,
                snap_point=record.get("child_point") or record.get("payload_location"),
            )
        )

    for record in placed_records:
        child_id = record.get("child_id")
        profile_name = record.get("profile_name")
        parent_id = record.get("parent_element_id")
        if not profile_name or parent_id is None:
            continue
        candidates = by_parent_id.get(parent_id) or []
        if not candidates:
            continue
        current_profiles = set()
        current_parent_name_variants = set()
        for candidate in candidates:
            for variant in candidate.get("name_variants") or []:
                if variant:
                    current_parent_name_variants.add(variant)
            for current in candidate.get("matched_profiles") or []:
                current_profiles.add(current)
            linker_profile = (candidate.get("linker_profile") or "").strip()
            if linker_profile:
                current_profiles.add(linker_profile)
        if profile_name in current_profiles:
            continue
        chosen = _pick_candidate(candidates)
        parent_text = _candidate_text(chosen)
        selectable_parent_id = parent_id if parent_id in host_parent_elements else None
        snap_point = (chosen or {}).get("point") or record.get("child_point") or record.get("payload_location")
        child_text = record.get("child_label") or ""
        stored_host_norm = _normalize_name(record.get("host_name"))
        stored_host_no_default = _normalize_name_ignoring_default_suffix(record.get("host_name"))
        current_parent_name_variants_no_default = {
            _normalize_name_ignoring_default_suffix(name)
            for name in current_parent_name_variants
            if name
        }
        parent_name_changed = bool(
            stored_host_norm
            and stored_host_norm not in current_parent_name_variants
            and stored_host_no_default not in current_parent_name_variants_no_default
        )
        if current_profiles:
            new_profile = sorted(current_profiles, key=lambda value: value.lower())[0]
            tab4_rows.append(
                _build_row(
                    profile=profile_name,
                    description="Original parent now matches profile '{}' instead of '{}'.".format(
                        new_profile, profile_name
                    ),
                    parent_text=parent_text,
                    child_text=child_text,
                    child_id=child_id,
                    parent_id=selectable_parent_id,
                    snap_point=snap_point,
                )
            )
        else:
            if not parent_name_changed:
                continue
            tab5_rows.append(
                _build_row(
                    profile=profile_name,
                    description=(
                        "Child Element Linker Host Name no longer matches current parent name, "
                        "and no YAML profile exists for the current parent."
                    ),
                    parent_text=parent_text or "Parent Id: {}".format(parent_id),
                    child_text=child_text,
                    child_id=child_id,
                    parent_id=selectable_parent_id,
                    snap_point=snap_point,
                )
            )

    for record in placed_records:
        child_id = record.get("child_id")
        parent_id = record.get("parent_element_id")
        if child_id is None or parent_id is None:
            continue
        candidates = by_parent_id.get(parent_id) or []
        chosen = _pick_candidate(candidates)
        if chosen is None:
            continue

        child_point = record.get("child_point") or record.get("payload_location")
        current_parent_point = chosen.get("point")
        stored_child_point = record.get("payload_location")
        stored_parent_point = record.get("payload_parent_location")
        if child_point is None or current_parent_point is None:
            continue
        if stored_child_point is None or stored_parent_point is None:
            continue

        actual_xy = _xy_distance(child_point, current_parent_point)
        stored_xy = _xy_distance(stored_child_point, stored_parent_point)
        if actual_xy is None or stored_xy is None:
            continue
        if actual_xy <= FAR_FROM_PARENT_THRESHOLD_FT:
            continue
        if stored_xy > FAR_FROM_PARENT_THRESHOLD_FT:
            continue

        profile_name = record.get("profile_name") or (record.get("host_name") or "<unknown profile>")
        selectable_parent_id = parent_id if parent_id in host_parent_elements else None
        tab6_rows.append(
            _build_row(
                profile=profile_name,
                description=(
                    "Current child-parent XY distance is {:.2f} ft, but stored XY offset is {:.2f} ft (<= {:.0f} ft)."
                ).format(actual_xy, stored_xy, FAR_FROM_PARENT_THRESHOLD_FT),
                parent_text=_candidate_text(chosen) or "Parent Id: {}".format(parent_id),
                child_text=record.get("child_label") or "",
                child_id=child_id,
                parent_id=selectable_parent_id,
                snap_point=child_point,
                adjust_enabled=True,
            )
        )

    tabs = {
        "tab1": tab1_rows,
        "tab2": tab2_rows,
        "tab3": tab3_rows,
        "tab4": tab4_rows,
        "tab5": tab5_rows,
        "tab6": tab6_rows,
    }
    meta = {
        "total_profiles": len(profiles),
        "total_parents_scanned": len(parent_data["all"]),
        "total_children_tracked": len(placed_records),
    }
    return tabs, meta


def _select_element(elem_id):
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return False
    if elem_id in (None, ""):
        return False
    try:
        target = doc.GetElement(ElementId(int(elem_id)))
    except Exception:
        target = None
    if target is None:
        return False
    ids = List[ElementId]()
    ids.Add(target.Id)
    try:
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        return False
    try:
        uidoc.ShowElements(ids)
    except Exception:
        pass
    return True


def _zoom_to_point(point, radius_feet=12.0):
    if point is None:
        return False
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return False
    active_view = getattr(doc, "ActiveView", None)
    if active_view is None:
        return False
    ui_view = None
    try:
        for candidate in uidoc.GetOpenUIViews():
            if candidate.ViewId == active_view.Id:
                ui_view = candidate
                break
    except Exception:
        ui_view = None
    if ui_view is None:
        return False
    try:
        radius = float(radius_feet)
    except Exception:
        radius = 12.0
    min_pt = XYZ(point.X - radius, point.Y - radius, point.Z - radius)
    max_pt = XYZ(point.X + radius, point.Y + radius, point.Z + radius)
    try:
        ui_view.ZoomAndCenterRectangle(min_pt, max_pt)
        return True
    except Exception:
        return False


def _load_window_module():
    module_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "QAQCReportWindow.py"))
    if not os.path.exists(module_path):
        return None
    try:
        return imp.load_source("ced_qaqc_report_window", module_path)
    except Exception as exc:
        LOG.warning("Failed to load QAQC report window module: %s", exc)
        return None


class _AdjustExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.child_id = None

    def GetName(self):
        return "QAQC Adjust External Event"

    def request(self, child_id):
        self.child_id = child_id

    def Execute(self, uiapp):
        child_id = self.child_id
        self.child_id = None
        uidoc = getattr(uiapp, "ActiveUIDocument", None)
        doc = getattr(uidoc, "Document", None) if uidoc else None
        ok, message = _adjust_element(doc, child_id)
        try:
            forms.alert(message, title="{} - Adjust".format(TITLE))
        except Exception:
            LOG.warning("[QAQC Adjust] %s", message)
        if ok and uidoc is not None and child_id not in (None, ""):
            try:
                ids = List[ElementId]()
                ids.Add(ElementId(int(child_id)))
                uidoc.Selection.SetElementIds(ids)
                uidoc.ShowElements(ids)
            except Exception:
                pass


def _ensure_adjust_external_event():
    global _ADJUST_HANDLER, _ADJUST_EXTERNAL_EVENT
    if _ADJUST_HANDLER is not None and _ADJUST_EXTERNAL_EVENT is not None:
        return True
    try:
        _ADJUST_HANDLER = _AdjustExternalEventHandler()
        _ADJUST_EXTERNAL_EVENT = ExternalEvent.Create(_ADJUST_HANDLER)
        return True
    except Exception as exc:
        _ADJUST_HANDLER = None
        _ADJUST_EXTERNAL_EVENT = None
        LOG.warning("Failed to create QAQC adjust external event: %s", exc)
        return False


def _load_follow_parent_module():
    global _FOLLOW_PARENT_MODULE
    if _FOLLOW_PARENT_MODULE is not None:
        return _FOLLOW_PARENT_MODULE
    module_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "Follow Parent.pushbutton", "script.py")
    )
    if not os.path.exists(module_path):
        return None
    try:
        _FOLLOW_PARENT_MODULE = imp.load_source("ced_follow_parent_runtime", module_path)
    except Exception as exc:
        LOG.warning("Failed to load Follow Parent module: %s", exc)
        _FOLLOW_PARENT_MODULE = None
    return _FOLLOW_PARENT_MODULE


def _load_optimize_module():
    global _OPTIMIZE_MODULE
    if _OPTIMIZE_MODULE is not None:
        return _OPTIMIZE_MODULE
    module_path = os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "..",
            "..",
            "Misc Operations.pulldown",
            "Optimize.pushbutton",
            "script.py",
        )
    )
    if not os.path.exists(module_path):
        return None
    try:
        _OPTIMIZE_MODULE = imp.load_source("ced_optimize_runtime", module_path)
    except Exception as exc:
        LOG.warning("Failed to load Optimize module: %s", exc)
        _OPTIMIZE_MODULE = None
    return _OPTIMIZE_MODULE


def _get_optimize_mode_for_element(elem):
    label = (_family_type_label(elem) or "").lower()
    if "wall" in label:
        return "Wall"
    if "drop" in label or "ceiling" in label:
        return "Ceiling"
    if "floor" in label:
        return "Floor"
    return None


def _run_follow_parent_for_element(doc, elem, follow_module):
    payload_text = follow_module._get_linker_text(elem)
    if not payload_text:
        return False, "Element_Linker payload is blank."

    payload = follow_module._parse_linker_payload(payload_text)
    parent_id = (payload or {}).get("parent_element_id")
    if parent_id is None:
        return False, "Parent ElementId was not found in Element_Linker."

    parent_candidates = follow_module._collect_parent_candidates_for_ids(doc, {parent_id})
    candidates = parent_candidates.get(parent_id) or []
    parent_choice = follow_module._choose_parent_candidate(payload, candidates)
    if parent_choice is None:
        return False, "Parent element could not be resolved."

    child_point_old = payload.get("location") or follow_module._get_point(elem)
    parent_point_old = payload.get("parent_location") or parent_choice.get("point")
    parent_point_new = parent_choice.get("point")
    if child_point_old is None or parent_point_old is None or parent_point_new is None:
        return False, "Required location data is missing."

    parent_rot_old = payload.get("parent_rotation_deg")
    if parent_rot_old is None:
        parent_rot_old = parent_choice.get("rotation_deg") or 0.0

    child_rot_old = payload.get("rotation_deg")
    if child_rot_old is None:
        child_rot_old = follow_module._get_rotation_degrees(elem)

    local_offset_old = follow_module._rotate_xy(child_point_old - parent_point_old, -parent_rot_old)
    rotation_offset_old = follow_module._normalize_angle(child_rot_old - parent_rot_old)

    parent_rot_new = parent_choice.get("rotation_deg") or parent_rot_old
    target_point = parent_point_new + follow_module._rotate_xy(local_offset_old, parent_rot_new)
    target_rot = parent_rot_new + rotation_offset_old

    moved, rotated = follow_module._move_and_rotate_child(doc, elem, target_point, target_rot)
    payload_text_new = follow_module._build_linker_payload(
        payload,
        elem,
        target_point,
        target_rot,
        parent_point_new,
        parent_rot_new,
        parent_id,
    )
    payload_updated = follow_module._set_linker_text(elem, payload_text_new)
    return True, "follow moved={}, rotated={}, payload_updated={}".format(
        bool(moved), bool(rotated), bool(payload_updated)
    )


def _run_optimize_for_element(doc, elem, optimize_module, mode):
    if not mode:
        return False, "no optimize mode matched"
    try:
        ok = bool(optimize_module._apply_optimization(doc, elem, mode, "Lower Left"))
    except Exception as exc:
        return False, "optimize {} error: {}".format(mode, exc)
    if ok:
        return True, "optimize {} succeeded".format(mode)
    return False, "optimize {} did not move element".format(mode)


def _adjust_element(doc, child_id):
    if doc is None:
        return False, "No active document available for adjust."
    if child_id in (None, ""):
        return False, "Child element id is missing."
    try:
        elem = doc.GetElement(ElementId(int(child_id)))
    except Exception:
        elem = None
    if elem is None:
        return False, "Child element no longer exists."

    follow_module = _load_follow_parent_module()
    if follow_module is None:
        return False, "Follow Parent module could not be loaded."
    optimize_module = _load_optimize_module()
    if optimize_module is None:
        return False, "Optimize module could not be loaded."

    mode = _get_optimize_mode_for_element(elem)
    txn = Transaction(doc, "QAQC Adjust Element {}".format(int(child_id)))
    try:
        txn.Start()
        follow_ok, follow_msg = _run_follow_parent_for_element(doc, elem, follow_module)
        optimize_ok, optimize_msg = _run_optimize_for_element(doc, elem, optimize_module, mode)
        txn.Commit()
    except Exception as exc:
        try:
            txn.RollBack()
        except Exception:
            pass
        return False, "Adjust failed: {}".format(exc)

    lines = [
        "Child Id {}.".format(int(child_id)),
        "Follow Parent: {}".format(follow_msg),
        "Optimize: {}".format(optimize_msg),
    ]
    return bool(follow_ok or optimize_ok), "\n".join(lines)


def main():
    global _MODELLESS_WINDOW, _ADJUST_HANDLER, _ADJUST_EXTERNAL_EVENT
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    if getattr(doc, "IsFamilyDocument", False):
        forms.alert("Open a project document before running QA/QC.", title=TITLE)
        return
    if not _ensure_adjust_external_event():
        forms.alert(
            "Could not initialize modeless adjust event. Adjust actions will be unavailable.",
            title=TITLE,
        )

    try:
        data_path, data = load_active_yaml_data(doc)
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    except Exception as exc:
        forms.alert("Failed to load active YAML from extensible storage:\n\n{}".format(exc), title=TITLE)
        return

    tabs, meta = _build_issue_tabs(doc, data)
    actionable_issues = (
        len(tabs.get("tab2") or [])
        + len(tabs.get("tab3") or [])
        + len(tabs.get("tab4") or [])
        + len(tabs.get("tab5") or [])
        + len(tabs.get("tab6") or [])
    )
    yaml_label = get_yaml_display_name(data_path)

    summary_text = (
        "Source: {} | Profiles: {} | Parent candidates scanned: {} | "
        "Tracked placed children: {} | Actionable issues: {}"
    ).format(
        yaml_label,
        meta.get("total_profiles", 0),
        meta.get("total_parents_scanned", 0),
        meta.get("total_children_tracked", 0),
        actionable_issues,
    )

    ui_module = _load_window_module()
    xaml_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "QAQCReportWindow.xaml"))
    if ui_module is None or not os.path.exists(xaml_path):
        lines = [
            summary_text,
            "",
            "1) YAML profile has no matching parents: {}".format(len(tabs.get("tab1") or [])),
            "2) Matching parent found but no children placed: {}".format(len(tabs.get("tab2") or [])),
            "3) Placed profile parent ID no longer exists: {}".format(len(tabs.get("tab3") or [])),
            "4) Parent changed type and matching profile exists: {}".format(len(tabs.get("tab4") or [])),
            "5) Parent changed type and matching profile missing: {}".format(len(tabs.get("tab5") or [])),
            "6) Far from Parent: {}".format(len(tabs.get("tab6") or [])),
        ]
        forms.alert("\n".join(lines), title=TITLE)
        return

    def _on_select_child(row):
        if not _select_element((row or {}).get("child_id")):
            forms.alert("Could not select child element.", title=TITLE)

    def _on_select_parent(row):
        if not _select_element((row or {}).get("parent_id")):
            forms.alert("Could not select parent element.\n(Linked parents are not directly selectable.)", title=TITLE)

    def _on_snap(row):
        point = (row or {}).get("snap_point")
        if point is None and (row or {}).get("child_id") is not None:
            try:
                elem = doc.GetElement(ElementId(int(row.get("child_id"))))
            except Exception:
                elem = None
            point = _get_element_point(elem)
        if not _zoom_to_point(point):
            forms.alert("Could not snap to location for this row.", title=TITLE)

    def _on_adjust(row):
        child_id = (row or {}).get("child_id")
        if child_id in (None, ""):
            forms.alert("No child element id found for this row.", title="{} - Adjust".format(TITLE))
            return
        if _ADJUST_HANDLER is None or _ADJUST_EXTERNAL_EVENT is None:
            ok, message = _adjust_element(doc, child_id)
            forms.alert(message, title="{} - Adjust".format(TITLE))
            if ok:
                _select_element(child_id)
            return
        try:
            _ADJUST_HANDLER.request(child_id)
            _ADJUST_EXTERNAL_EVENT.Raise()
        except Exception as exc:
            forms.alert("Failed to queue adjust request:\n\n{}".format(exc), title="{} - Adjust".format(TITLE))

    window = ui_module.QAQCReportWindow(
        xaml_path=xaml_path,
        tab_rows=tabs,
        summary_text=summary_text,
        select_child_callback=_on_select_child,
        select_parent_callback=_on_select_parent,
        snap_callback=_on_snap,
        adjust_callback=_on_adjust,
    )
    # Keep a module-level reference so the modeless window stays alive.
    if _MODELLESS_WINDOW is not None:
        try:
            if bool(_MODELLESS_WINDOW.IsVisible):
                _MODELLESS_WINDOW.Close()
        except Exception:
            pass
        _MODELLESS_WINDOW = None

    def _on_closed(sender, args):
        global _MODELLESS_WINDOW
        try:
            if sender == _MODELLESS_WINDOW:
                _MODELLESS_WINDOW = None
        except Exception:
            _MODELLESS_WINDOW = None

    _MODELLESS_WINDOW = window
    try:
        window.Closed += _on_closed
    except Exception:
        pass
    try:
        window.Show()
    except Exception:
        window.show()


if __name__ == "__main__":
    main()
