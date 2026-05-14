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
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    Reference,
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
_PLACE_SINGLE_PROFILE_MODULE = None
_ADJUST_HANDLER = None
_ADJUST_EXTERNAL_EVENT = None
_FIX_ID_HANDLER = None
_FIX_ID_EXTERNAL_EVENT = None
_PLACE_PROFILE_HANDLER = None
_PLACE_PROFILE_EXTERNAL_EVENT = None

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402

LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
PARENT_ID_KEYS = ("Parent ElementId", "Parent Element ID")
LINKER_ELEMENT_ID_KEYS = ("ElementId", "Element ID", "Element Id")
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"

INLINE_LINKER_PATTERN = re.compile(
    r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|"
    r"Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
    r"Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation)\s*:\s*",
    re.IGNORECASE,
)

FAR_FROM_PARENT_THRESHOLD_FT = 10.0


# A profile is treated as a TRACKED equipment profile (eligible for
# Tab 2) when its name starts with either ``HEB`` or exactly three
# digits, followed by an obvious separator (``_``, ``-``, whitespace,
# or ``:``). Anything else (e.g. "Business Center AHU", "Walk-in
# Cooler") is an independent / annotation-only profile and is skipped
# by Tab 2 — those profiles don't anchor to a parent the way the
# numbered HEB equipment families do. ``allow_parentless`` in the YAML
# is unreliable here (stamped ``true`` on every entry) so the name
# convention is the source of truth.
_TRACKED_PROFILE_PREFIX_RE = re.compile(
    r"^(HEB[_\-\s]|\d{3}[_\-\s:])",
    re.IGNORECASE,
)


def _is_tracked_equipment_profile_name(profile_name):
    if not profile_name:
        return False
    return bool(_TRACKED_PROFILE_PREFIX_RE.match(str(profile_name).strip()))


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


def _category_id_value(category):
    if category is None:
        return None
    try:
        cat_id = category.Id
    except Exception:
        cat_id = None
    return _element_id_value(cat_id)


def _is_fixture_element(elem):
    if elem is None:
        return False
    try:
        category = getattr(elem, "Category", None)
    except Exception:
        category = None

    cat_id_val = _category_id_value(category)
    fixture_ids = set()
    for bic in (
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_PlumbingFixtures,
    ):
        try:
            fixture_ids.add(int(bic))
        except Exception:
            continue
    if cat_id_val in fixture_ids:
        return True

    try:
        cat_name = getattr(category, "Name", None)
    except Exception:
        cat_name = None
    if cat_name and "fixture" in str(cat_name).strip().lower():
        return True
    return False


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


def _walk_link_documents(doc, parent_transform, doc_chain, top_link_inst=None):
    """Yield ``(link_doc, transform, top_link_inst)`` for every linked
    doc reachable from ``doc``.

    ``top_link_inst`` is the host-doc-resident ``RevitLinkInstance``
    that the linked content lives inside. For top-level links it's
    the link instance itself; for nested links it's propagated from
    the outer call so callers always get a reference they can resolve
    against the host doc (e.g. to build a host-coord
    ``Reference.CreateLinkReference``). Nested links can't be deep-
    selected through the host UI, but the top-level instance at
    least lets ``Snap`` highlight SOMETHING relevant.
    """
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
        outer = top_link_inst if top_link_inst is not None else link_inst
        yield link_doc, transform, outer
        for nested in _walk_link_documents(link_doc, transform, next_chain, outer):
            yield nested


def _iter_link_documents(doc):
    for link_doc, transform, link_inst in _walk_link_documents(doc, None, set()):
        yield link_doc, transform, link_inst


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
    param = _find_linker_parameter(elem)
    if not param:
        return ""
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


def _find_linker_parameter(elem):
    if elem is None:
        return None
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if param:
            return param
    return None


def _set_linker_text(elem, text):
    if elem is None:
        return False
    param = _find_linker_parameter(elem)
    if not param:
        return False
    try:
        return bool(param.Set(str(text or "")))
    except Exception:
        return False


def _replace_linker_element_id(payload_text, target_element_id):
    text = str(payload_text or "")
    if not text.strip():
        return text, False
    target_text = str(int(target_element_id))

    multiline_pattern = re.compile(r"^(\s*Element(?:\s+)?Id\s*:\s*).*$", re.IGNORECASE)
    lines = text.splitlines(True)
    if lines:
        replaced = False
        new_lines = []
        for raw_line in lines:
            line_ending = ""
            if raw_line.endswith("\r\n"):
                body = raw_line[:-2]
                line_ending = "\r\n"
            elif raw_line.endswith("\n"):
                body = raw_line[:-1]
                line_ending = "\n"
            else:
                body = raw_line
            match = multiline_pattern.match(body)
            if match:
                body = "{}{}".format(match.group(1), target_text)
                replaced = True
            new_lines.append(body + line_ending)
        if replaced:
            return "".join(new_lines), True

    inline_pattern = re.compile(r"(Element(?:\s+)?Id\s*:\s*)([^,\n\r]+)", re.IGNORECASE)
    updated_text, count = inline_pattern.subn(r"\g<1>{}".format(target_text), text, count=1)
    if count > 0:
        return updated_text, True
    return text, False


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

    entries_lower = {}
    for key, value in entries.items():
        norm_key = str(key or "").strip().lower()
        if norm_key and norm_key not in entries_lower:
            entries_lower[norm_key] = value

    parent_element_id = None
    for key in PARENT_ID_KEYS:
        value = entries.get(key)
        if value is None:
            value = entries_lower.get(str(key).strip().lower())
        if value is not None:
            parent_element_id = _try_int(value)
            if parent_element_id is not None:
                break

    linker_element_id = None
    for key in LINKER_ELEMENT_ID_KEYS:
        value = entries.get(key)
        if value is None:
            value = entries_lower.get(str(key).strip().lower())
        if value is not None:
            linker_element_id = _try_int(value)
            if linker_element_id is not None:
                break

    return {
        "led_id": (entries.get("Linked Element Definition ID") or "").strip(),
        "set_id": (entries.get("Set Definition ID") or "").strip(),
        "host_name": (entries.get("Host Name") or "").strip(),
        "parent_element_id": parent_element_id,
        "linker_element_id": linker_element_id,
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
    # Profiles rolled up under another via legacy ``ced_truth_source_id``
    # or via being listed in another profile's ``merged_aliases``. Used
    # by Tab 1 to skip them.
    merged_member_profiles = set()
    # Profiles whose name doesn't start with the tracked-equipment
    # prefix (``HEB...`` or three-digit ``NNN...``). These are
    # annotation / independent profiles and shouldn't be flagged by
    # Tab 2 for missing children — see _is_tracked_equipment_profile_name.
    independent_profiles = set()
    profile_id_to_name = {}

    def _register_alias(profile_name, alias_value):
        # Adds alias_value (and, if it's a "Family : Type" pair, the
        # family-half) into ``norm_to_profiles`` under the master
        # profile. Used for ``parent_filter.family_name_pattern`` and
        # every entry in ``merged_aliases`` so parents whose Revit
        # family name differs from the profile's display name still
        # match. Independent profiles (per the name prefix) still get
        # their aliases registered so other tabs (e.g. Tab 4 "parent
        # type changed") see them as valid matches — only Tab 2 skips
        # them.
        if not alias_value:
            return
        alias_text = str(alias_value).strip()
        if not alias_text:
            return
        candidates = {alias_text}
        if " : " in alias_text:
            family_half = alias_text.split(" : ", 1)[0].strip()
            if family_half:
                candidates.add(family_half)
        for candidate in candidates:
            alias_norm = _normalize_name(candidate)
            if alias_norm:
                norm_to_profiles.setdefault(alias_norm, []).append(profile_name)

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

        # Register parent_filter.family_name_pattern (if present) so
        # profiles whose name differs from the parent's Revit family
        # name still match against that family. Needed for the Ishida
        # scales and similar profiles whose display name carries a
        # ``: Default`` suffix but whose ``family_name_pattern`` is the
        # clean family name in the model.
        parent_filter = eq.get("parent_filter")
        if isinstance(parent_filter, dict):
            _register_alias(profile_name, parent_filter.get("family_name_pattern"))

        # Register merged_aliases entries so absorbed family names still
        # resolve to the master profile. Also track member names so
        # Tab 1 doesn't false-positive them as "no matching parent".
        merged_aliases = eq.get("merged_aliases") or []
        if isinstance(merged_aliases, (list, tuple)):
            for alias_entry in merged_aliases:
                _register_alias(profile_name, alias_entry)
                if not alias_entry:
                    continue
                alias_text = str(alias_entry).strip()
                if not alias_text:
                    continue
                member_name = alias_text
                if " : " in alias_text:
                    family_half = alias_text.split(" : ", 1)[0].strip()
                    if family_half:
                        member_name = family_half
                if member_name and member_name != profile_name:
                    merged_member_profiles.add(member_name)

        # Classify by name prefix. Tracked equipment profiles start with
        # ``HEB...`` or ``NNN...`` (3 digits) followed by a separator;
        # everything else is independent and excluded from Tab 2.
        if not _is_tracked_equipment_profile_name(profile_name):
            independent_profiles.add(profile_name)

        profile_id = (eq.get("id") or "").strip()
        if profile_id:
            profile_id_to_name[profile_id] = profile_name

        group_key = (eq.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not group_key:
            group_key = (eq.get("id") or profile_name).strip()
        else:
            # Legacy merge model: a non-empty ced_truth_source_id that
            # differs from this profile's own id means this entry is
            # rolled up under another profile — flag it as a merged
            # member so Tab 1 doesn't false-positive it.
            own_id = (eq.get("id") or "").strip()
            if own_id and group_key != own_id:
                merged_member_profiles.add(profile_name)
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
        "merged_member_profiles": merged_member_profiles,
        "independent_profiles": independent_profiles,
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

    def _add_candidate(elem, point, is_linked, link_inst=None):
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
        # ``link_instance_id`` is the host-doc-resident RevitLinkInstance
        # that this linked element lives inside — needed so Snap can
        # build a host-coord Reference and highlight the linked element.
        # Host candidates leave it None.
        link_inst_id = None
        if link_inst is not None:
            link_inst_id = _element_id_value(getattr(link_inst, "Id", None))
        candidate = {
            "parent_id": parent_id,
            "display_label": _element_label(elem),
            "is_linked": bool(is_linked),
            "point": point,
            "name_variants": sorted(variants, key=lambda value: value.lower()),
            "matched_profiles": sorted(matched_profiles, key=lambda value: value.lower()),
            "linker_profile": linker_profile,
            "link_instance_id": link_inst_id,
        }
        candidates.append(candidate)
        by_parent_id.setdefault(parent_id, []).append(candidate)
        if not is_linked and parent_id not in host_parent_elements:
            host_parent_elements[parent_id] = elem
        for profile_name in candidate["matched_profiles"]:
            profile_to_candidates.setdefault(profile_name, []).append(candidate)

    for elem in _collect_family_and_group_instances(doc):
        _add_candidate(elem, _get_element_point(elem), is_linked=False)

    for link_doc, transform, link_inst in _iter_link_documents(doc):
        for elem in _collect_family_and_group_instances(link_doc):
            point = _transform_point(transform, _get_element_point(elem))
            _add_candidate(elem, point, is_linked=True, link_inst=link_inst)

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
            "is_fixture": _is_fixture_element(elem),
            "profile_name": profile_name,
            "set_id": payload.get("set_id"),
            "led_id": payload.get("led_id"),
            "host_name": payload.get("host_name"),
            "parent_element_id": payload.get("parent_element_id"),
            "linker_element_id": payload.get("linker_element_id"),
            "payload_text": payload_text,
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
    fix_id_enabled=False,
    link_instance_id=None,
    linked_element_id=None,
    snap_select_id=None,
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
        "fix_id_enabled": bool(fix_id_enabled),
        # When the row's primary snap target is a linked element these
        # together let Snap build a host-coord Reference and highlight
        # the linked element. Host-only snap targets leave both None.
        "link_instance_id": link_instance_id,
        "linked_element_id": linked_element_id,
        # The host-doc element id that Snap should highlight (e.g. the
        # parent for type-change tabs, the child for far-from-parent
        # tabs). Only used when link_instance_id is not set. Falls back
        # to child_id then parent_id in ``_on_snap`` if not specified.
        "snap_select_id": snap_select_id,
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
    # Parent-level coverage set: every parent_element_id that has at
    # least one placed child, regardless of which profile/group that
    # child's led_id / set_id / host_name resolved to. Used by Tab 2
    # to skip parents that already host SOMETHING — catches cases
    # where the child's profile resolves to a different group than
    # the parent's family matched (merged profile, re-keyed YAML,
    # cross-profile placement).
    parents_with_any_placed_child = set()
    for record in placed_records:
        profile_name = record.get("profile_name")
        parent_id = record.get("parent_element_id")
        if parent_id is not None:
            parents_with_any_placed_child.add(parent_id)
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
    tab7_rows = []

    merged_member_profiles = yaml_maps.get("merged_member_profiles") or set()
    for profile_name in profiles:
        # Skip profiles that are rolled up under another profile (via
        # ced_truth_source_id or merged_aliases) — those are unused on
        # purpose and would otherwise produce false "no matching parent"
        # rows.
        if profile_name in merged_member_profiles:
            continue
        group_key = profile_to_group.get(profile_name) or profile_name
        if not group_to_candidate_ids.get(group_key):
            tab1_rows.append(
                _build_row(
                    profile=profile_name,
                    description="Profile exists in active YAML/extensible storage, but no matching parent elements were found.",
                )
            )

    # Tab 2 — "matching parent found, no placed children". Deduped by
    # parent: each linked element appears at most once, with matching
    # profile names listed. Independent profiles (name doesn't start
    # with ``HEB...`` or three digits) are excluded — they're
    # annotation-only profiles, not tracked equipment.
    independent_profiles = yaml_maps.get("independent_profiles") or set()
    # (parent_id, is_linked) -> {"candidate": ..., "profiles": [names...]}
    tab2_by_parent = {}
    for profile_name in profiles:
        if profile_name in independent_profiles:
            continue
        if profile_name in merged_member_profiles:
            continue
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
            # Cross-profile coverage: if THIS parent already has any
            # placed child (resolved via any profile/group), treat it
            # as covered. Stops Tab 2 from false-flagging a parent
            # whose hosted child's led_id maps to a sibling/merged
            # profile in a different group.
            if parent_id in parents_with_any_placed_child:
                continue
            key = (parent_id, bool(candidate.get("is_linked")))
            slot = tab2_by_parent.get(key)
            if slot is None:
                slot = {"candidate": candidate, "profiles": []}
                tab2_by_parent[key] = slot
            if profile_name not in slot["profiles"]:
                slot["profiles"].append(profile_name)

    for (parent_id, _is_linked), info in tab2_by_parent.items():
        candidate = info["candidate"]
        profile_list = info["profiles"]
        if not profile_list:
            continue
        # Display: actual profile names — comma-separated when there's
        # more than one. No "N profiles" placeholder.
        profile_display = ", ".join(profile_list)
        if len(profile_list) == 1:
            description = (
                "Matching parent found, but no placed children. "
                "Matching profile: {}."
            ).format(profile_list[0])
        else:
            description = (
                "Matching parent found, but no placed children. "
                "Matching profiles: {}."
            ).format(profile_display)
        selectable_parent_id = parent_id if parent_id in host_parent_elements else None
        # Snap should highlight the parent: linked element via Reference
        # if the candidate lives in a link, else the host parent id.
        link_inst_id = (
            candidate.get("link_instance_id") if candidate.get("is_linked") else None
        )
        linked_elem_id = parent_id if candidate.get("is_linked") else None
        tab2_rows.append(
            _build_row(
                profile=profile_display,
                description=description,
                parent_text=_candidate_text(candidate),
                parent_id=selectable_parent_id,
                snap_point=candidate.get("point"),
                link_instance_id=link_inst_id,
                linked_element_id=linked_elem_id,
                snap_select_id=selectable_parent_id,
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
                snap_select_id=child_id,
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
        # Snap target on Tab 4 / Tab 5 is the parent (we zoom to its
        # location). Highlight the linked parent if applicable, else
        # the host parent id.
        chosen_link_inst_id = (
            (chosen or {}).get("link_instance_id") if (chosen or {}).get("is_linked") else None
        )
        chosen_linked_elem_id = parent_id if (chosen or {}).get("is_linked") else None
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
                    link_instance_id=chosen_link_inst_id,
                    linked_element_id=chosen_linked_elem_id,
                    snap_select_id=selectable_parent_id,
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
                    link_instance_id=chosen_link_inst_id,
                    linked_element_id=chosen_linked_elem_id,
                    snap_select_id=selectable_parent_id,
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
        # NOTE: do NOT skip when stored points are missing — actual_xy
        # alone is the gate. Stored-offset diagnostics are best-effort
        # below.

        actual_xy = _xy_distance(child_point, current_parent_point)
        if stored_child_point is not None and stored_parent_point is not None:
            stored_xy = _xy_distance(stored_child_point, stored_parent_point)
        else:
            stored_xy = None
        # Reliability fix: flag whenever the child currently sits more
        # than the threshold from its parent. The old predicate
        # additionally required stored_xy <= threshold so only "drifted
        # from a known-good offset" cases fired — that silently missed
        # children placed with legitimately-long stored offsets that
        # now sit even farther away. User spec: ANY current XY > 10 ft.
        if actual_xy is None:
            continue
        if actual_xy <= FAR_FROM_PARENT_THRESHOLD_FT:
            continue

        if stored_xy is None:
            description = (
                "Current child-parent XY distance is {:.2f} ft (threshold > {:.0f} ft). Stored XY offset is unavailable."
            ).format(actual_xy, FAR_FROM_PARENT_THRESHOLD_FT)
        else:
            description = (
                "Current child-parent XY distance is {:.2f} ft (threshold > {:.0f} ft). Stored XY offset is {:.2f} ft."
            ).format(actual_xy, FAR_FROM_PARENT_THRESHOLD_FT, stored_xy)

        profile_name = record.get("profile_name") or (record.get("host_name") or "<unknown profile>")
        selectable_parent_id = parent_id if parent_id in host_parent_elements else None
        tab6_rows.append(
            _build_row(
                profile=profile_name,
                description=description,
                parent_text=_candidate_text(chosen) or "Parent Id: {}".format(parent_id),
                child_text=record.get("child_label") or "",
                child_id=child_id,
                parent_id=selectable_parent_id,
                snap_point=child_point,
                adjust_enabled=True,
                snap_select_id=child_id,
            )
        )

    # Tab 7a — child's own element_id discrepancy.
    # Dropped the fixture-only restriction so this fires across every
    # element with an Element_Linker, not just fixtures. The user's
    # spec calls for reliable ID-drift detection regardless of category.
    for record in placed_records:
        child_id = record.get("child_id")
        linker_element_id = record.get("linker_element_id")
        if child_id is None or linker_element_id is None:
            continue
        if int(linker_element_id) == int(child_id):
            continue

        profile_name = record.get("profile_name") or (record.get("host_name") or "<unknown profile>")
        parent_id = record.get("parent_element_id")
        selectable_parent_id = parent_id if parent_id in host_parent_elements else None
        parent_text = ""
        if parent_id is not None:
            parent_candidates = by_parent_id.get(parent_id) or []
            chosen_parent = _pick_candidate(parent_candidates)
            parent_text = _candidate_text(chosen_parent) or "Parent Id: {}".format(parent_id)

        tab7_rows.append(
            _build_row(
                profile=profile_name,
                description=(
                    "Element_Linker ElementId is {}, but the actual element Id is {}."
                ).format(int(linker_element_id), int(child_id)),
                parent_text=parent_text,
                child_text=record.get("child_label") or "",
                child_id=child_id,
                parent_id=selectable_parent_id,
                snap_point=record.get("child_point") or record.get("payload_location"),
                fix_id_enabled=True,
                snap_select_id=child_id,
            )
        )

    # Tab 7b — parent-id discrepancy.
    # The stored ``parent_element_id`` resolves to a real element in
    # the host or linked docs, but THAT element's family/type doesn't
    # match the child's stored ``host_name``. Indicates the linker's
    # parent_element_id is pointing at the wrong element — most often
    # a copy/paste of the parent that left the child's linker
    # referencing the original id. Diagnostic only for now (no
    # automated fix wired up).
    for record in placed_records:
        child_id = record.get("child_id")
        parent_id = record.get("parent_element_id")
        if parent_id is None:
            continue
        parent_candidates = by_parent_id.get(parent_id) or []
        if not parent_candidates:
            # parent_element_id doesn't resolve at all — that's Tab 3's
            # job, not ours.
            continue
        stored_host = record.get("host_name")
        stored_host_no_default = _normalize_name_ignoring_default_suffix(stored_host)
        if not stored_host_no_default:
            continue
        chosen_parent = _pick_candidate(parent_candidates)
        parent_name_variants_no_default = set()
        actual_family_display = ""
        for candidate in parent_candidates:
            for variant in candidate.get("name_variants") or []:
                if variant:
                    parent_name_variants_no_default.add(
                        _normalize_name_ignoring_default_suffix(variant)
                    )
                    if not actual_family_display:
                        actual_family_display = variant
        if not parent_name_variants_no_default:
            continue
        if stored_host_no_default in parent_name_variants_no_default:
            continue

        profile_name = record.get("profile_name") or (stored_host or "<unknown profile>")
        selectable_parent_id = parent_id if parent_id in host_parent_elements else None
        parent_text = _candidate_text(chosen_parent) or "Parent Id: {}".format(parent_id)
        tab7_rows.append(
            _build_row(
                profile=profile_name,
                description=(
                    "Element_Linker parent_element_id is {}, but that element's family "
                    "('{}') doesn't match stored host_name ('{}')."
                ).format(int(parent_id), actual_family_display or "<unknown>", stored_host or ""),
                parent_text=parent_text,
                child_text=record.get("child_label") or "",
                child_id=child_id,
                parent_id=selectable_parent_id,
                snap_point=record.get("child_point") or record.get("payload_location"),
                fix_id_enabled=False,
                snap_select_id=child_id,
            )
        )

    tabs = {
        "tab1": tab1_rows,
        "tab2": tab2_rows,
        "tab3": tab3_rows,
        "tab4": tab4_rows,
        "tab5": tab5_rows,
        "tab6": tab6_rows,
        "tab7": tab7_rows,
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


def _select_linked_element(link_inst_id, linked_elem_id):
    """Set the host doc's selection to the given linked element by
    building a host-coord ``Reference`` via
    ``Reference.CreateLinkReference``. Returns True on success.

    Used by Snap to highlight linked parents (instead of just zooming
    to them) so the user can see the specific linked element rather
    than scanning the linked model's overall footprint.
    """
    if link_inst_id in (None, "") or linked_elem_id in (None, ""):
        return False
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return False
    try:
        link_inst = doc.GetElement(ElementId(int(link_inst_id)))
    except Exception:
        return False
    if link_inst is None:
        return False
    try:
        link_doc = link_inst.GetLinkDocument()
    except Exception:
        link_doc = None
    if link_doc is None:
        return False
    try:
        linked_elem = link_doc.GetElement(ElementId(int(linked_elem_id)))
    except Exception:
        linked_elem = None
    if linked_elem is None:
        return False
    try:
        ref = Reference(linked_elem)
    except Exception:
        return False
    try:
        host_ref = ref.CreateLinkReference(link_inst)
    except Exception:
        host_ref = None
    if host_ref is None:
        return False
    refs = List[Reference]()
    refs.Add(host_ref)
    try:
        uidoc.Selection.SetReferences(refs)
        return True
    except Exception:
        return False


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


class _FixIdExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.child_id = None

    def GetName(self):
        return "QAQC Fix ElementId External Event"

    def request(self, child_id):
        self.child_id = child_id

    def Execute(self, uiapp):
        child_id = self.child_id
        self.child_id = None
        uidoc = getattr(uiapp, "ActiveUIDocument", None)
        doc = getattr(uidoc, "Document", None) if uidoc else None
        ok, message = _fix_element_linker_element_id(doc, child_id)
        try:
            forms.alert(message, title="{} - Fix ID".format(TITLE))
        except Exception:
            LOG.warning("[QAQC Fix ID] %s", message)
        if ok and uidoc is not None and child_id not in (None, ""):
            try:
                ids = List[ElementId]()
                ids.Add(ElementId(int(child_id)))
                uidoc.Selection.SetElementIds(ids)
                uidoc.ShowElements(ids)
            except Exception:
                pass


def _ensure_fix_id_external_event():
    global _FIX_ID_HANDLER, _FIX_ID_EXTERNAL_EVENT
    if _FIX_ID_HANDLER is not None and _FIX_ID_EXTERNAL_EVENT is not None:
        return True
    try:
        _FIX_ID_HANDLER = _FixIdExternalEventHandler()
        _FIX_ID_EXTERNAL_EVENT = ExternalEvent.Create(_FIX_ID_HANDLER)
        return True
    except Exception as exc:
        _FIX_ID_HANDLER = None
        _FIX_ID_EXTERNAL_EVENT = None
        LOG.warning("Failed to create QAQC fix-id external event: %s", exc)
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


def _load_place_single_profile_module():
    """Lazy-load the Place Single Profile pushbutton's panel module so
    we can borrow its helpers (_cleaned_profiles_from_raw,
    _build_repository_from_profiles, _gather_child_requests) for the
    Place button on Tab 2 rows. We DO NOT instantiate the dockable
    panel — only the placement engine wiring is used.
    """
    global _PLACE_SINGLE_PROFILE_MODULE
    if _PLACE_SINGLE_PROFILE_MODULE is not None:
        return _PLACE_SINGLE_PROFILE_MODULE
    module_path = os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "..",
            "..",
            "Place Single Profile.pushbutton",
            "PlaceSingleProfilePanel.py",
        )
    )
    if not os.path.exists(module_path):
        return None
    try:
        _PLACE_SINGLE_PROFILE_MODULE = imp.load_source(
            "ced_place_single_profile_runtime", module_path,
        )
    except Exception as exc:
        LOG.warning("Failed to load PlaceSingleProfilePanel module: %s", exc)
        _PLACE_SINGLE_PROFILE_MODULE = None
    return _PLACE_SINGLE_PROFILE_MODULE


def _resolve_parent_for_placement(doc, row):
    """Resolve a Tab 2 row's parent into ``(point, rotation)`` in host
    coordinates.

    Tab 2 rows store ``parent_id`` only for host parents (linked
    parents leave it None) and ``snap_point`` always — already
    transformed into host coords for linked candidates. Returns
    ``(point, rotation_rad)`` or ``(None, 0.0)`` if neither is usable.

    Rotation is read from the host parent's FamilyInstance.Location.
    For linked parents we don't have a cheap way to recover the
    rotation without re-walking the link transforms, so we fall back
    to 0.0 — children land at the offsets the YAML specifies, just
    not rotated to follow a tilted linked parent. "Follow Parent" on
    the placed child fixes that secondarily if needed.
    """
    if not isinstance(row, dict):
        return None, 0.0
    rotation_rad = 0.0
    parent_id = row.get("parent_id")
    point = None
    if parent_id not in (None, ""):
        try:
            elem = doc.GetElement(ElementId(int(parent_id)))
        except Exception:
            elem = None
        if elem is not None:
            point = _get_element_point(elem)
            try:
                loc = elem.Location
                rot = getattr(loc, "Rotation", None)
                if rot is not None:
                    rotation_rad = float(rot)
            except Exception:
                rotation_rad = 0.0
    if point is None:
        point = row.get("snap_point")
    return point, rotation_rad


def _do_place_profile_for_row(doc, raw_data, row):
    """Direct placement-engine call for a Tab 2 row.

    Resolves the row's first matching profile + parent location and
    calls ``PlaceElementsEngine.place_from_csv`` with the parent + its
    LED child requests. No dockable pane involved — runs inside the
    external event so it has a Revit API context for the transaction.

    Returns ``(ok: bool, message: str)``.
    """
    if doc is None:
        return False, "No active document available for placement."
    if not isinstance(row, dict):
        return False, "Invalid row payload."

    # Tab 2 may list multiple matching profiles comma-separated. Place
    # the FIRST one — the user can re-click for the others, or refine
    # with the filter and re-trigger.
    profile_field = (row.get("profile") or "").strip()
    if not profile_field:
        return False, "Row has no profile name."
    cad_choice = profile_field.split(",", 1)[0].strip()
    if not cad_choice:
        return False, "Could not parse a profile name from row."

    parent_point, parent_rotation = _resolve_parent_for_placement(doc, row)
    if parent_point is None:
        return False, "Could not resolve a parent location for placement."

    helper = _load_place_single_profile_module()
    if helper is None:
        return False, "PlaceSingleProfilePanel module not found."

    try:
        from LogicClasses.placement_engine import PlaceElementsEngine
        from LogicClasses.linked_equipment import find_equipment_by_name
    except Exception as exc:
        return False, "Placement engine unavailable: {}".format(exc)

    if not raw_data:
        return False, "Active YAML data is not available for placement."

    try:
        cleaned = helper._cleaned_profiles_from_raw(raw_data)
        repo = helper._build_repository_from_profiles(cleaned)
    except Exception as exc:
        return False, "Failed to build profile repository: {}".format(exc)

    try:
        labels = repo.labels_for_cad(cad_choice)
    except Exception:
        labels = None
    if not labels:
        return False, "Profile '{}' has no linked types to place.".format(cad_choice)

    parent_def = find_equipment_by_name(raw_data, cad_choice)
    if not parent_def:
        return False, "Profile '{}' not found in active YAML.".format(cad_choice)

    # Build the CSV-style rows the engine expects. The first row is the
    # parent's location; subsequent rows are LED children produced by
    # _gather_child_requests at parent + offset.
    selection_map = {cad_choice: labels}
    csv_rows = [{
        "Name": cad_choice,
        "Count": "1",
        "Position X": str(parent_point.X * 12.0),
        "Position Y": str(parent_point.Y * 12.0),
        "Position Z": str(parent_point.Z * 12.0),
        "Rotation": str(parent_rotation or 0.0),
    }]
    try:
        child_requests = helper._gather_child_requests(
            parent_def, parent_point, parent_rotation or 0.0, repo, raw_data,
        )
    except Exception as exc:
        return False, "Failed to gather LED requests: {}".format(exc)
    for request in child_requests or []:
        name = request.get("name")
        req_labels = request.get("labels")
        point = request.get("target_point")
        rotation = request.get("rotation")
        if not name or not req_labels or point is None:
            continue
        selection_map[name] = req_labels
        csv_rows.append({
            "Name": name,
            "Count": "1",
            "Position X": str(point.X * 12.0),
            "Position Y": str(point.Y * 12.0),
            "Position Z": str(point.Z * 12.0),
            "Rotation": str(rotation or 0.0),
        })

    try:
        engine = PlaceElementsEngine(
            doc, repo, allow_tags=False,
            transaction_name="QAQC Place Profile ({})".format(cad_choice),
        )
        results = engine.place_from_csv(csv_rows, selection_map)
    except Exception as exc:
        return False, "Placement engine error: {}".format(exc)

    placed = (results or {}).get("placed", 0)
    return True, "Placed {} element(s) for profile '{}'.".format(placed, cad_choice)


class _PlaceProfileExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self._payload = None

    def GetName(self):  # noqa: N802
        return "QAQC Place Profile External Event"

    def request(self, raw_data, row):
        self._payload = {"raw_data": raw_data, "row": row}

    def Execute(self, uiapp):  # noqa: N802
        payload = self._payload
        self._payload = None
        if not payload:
            return
        uidoc = getattr(uiapp, "ActiveUIDocument", None)
        doc = getattr(uidoc, "Document", None) if uidoc else None
        ok, message = _do_place_profile_for_row(
            doc, payload.get("raw_data"), payload.get("row"),
        )
        try:
            forms.alert(message, title="{} - Place Profile".format(TITLE))
        except Exception:
            LOG.warning("[QAQC Place Profile] %s", message)
        if ok and uidoc is not None:
            # Selecting the parent helps the user see where placement
            # landed (the children are tagged with Element_Linker but
            # there may be several of them — keep the focus on the
            # known parent reference).
            row = payload.get("row") or {}
            parent_id = row.get("parent_id")
            if parent_id not in (None, ""):
                try:
                    ids = List[ElementId]()
                    ids.Add(ElementId(int(parent_id)))
                    uidoc.Selection.SetElementIds(ids)
                    uidoc.ShowElements(ids)
                except Exception:
                    pass


def _ensure_place_profile_external_event():
    global _PLACE_PROFILE_HANDLER, _PLACE_PROFILE_EXTERNAL_EVENT
    if (
        _PLACE_PROFILE_HANDLER is not None
        and _PLACE_PROFILE_EXTERNAL_EVENT is not None
    ):
        return True
    try:
        _PLACE_PROFILE_HANDLER = _PlaceProfileExternalEventHandler()
        _PLACE_PROFILE_EXTERNAL_EVENT = ExternalEvent.Create(_PLACE_PROFILE_HANDLER)
        return True
    except Exception as exc:
        _PLACE_PROFILE_HANDLER = None
        _PLACE_PROFILE_EXTERNAL_EVENT = None
        LOG.warning("Failed to create QAQC place-profile external event: %s", exc)
        return False


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
    try:
        target_point = follow_module._preserve_child_z(elem, target_point)
    except Exception:
        current_point = follow_module._get_point(elem)
        if current_point is None:
            current_point = _get_element_point(elem)
        if current_point is not None:
            target_point = XYZ(target_point.X, target_point.Y, current_point.Z)
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


def _fix_element_linker_element_id(doc, child_id):
    if doc is None:
        return False, "No active document available for ElementId fix."
    if child_id in (None, ""):
        return False, "Child element id is missing."
    try:
        child_id_int = int(child_id)
    except Exception:
        return False, "Child element id is invalid: {}".format(child_id)

    try:
        elem = doc.GetElement(ElementId(child_id_int))
    except Exception:
        elem = None
    if elem is None:
        return False, "Child element no longer exists."

    payload_text = _get_linker_text(elem)
    if not payload_text:
        return False, "Element_Linker payload is blank."
    payload = _parse_linker_payload(payload_text)
    linker_element_id = payload.get("linker_element_id")
    if linker_element_id is None:
        return False, "ElementId was not found in Element_Linker payload."
    if int(linker_element_id) == child_id_int:
        return True, "Element_Linker ElementId already matches actual element Id {}.".format(child_id_int)

    updated_text, replaced = _replace_linker_element_id(payload_text, child_id_int)
    if not replaced:
        return False, "Could not update ElementId in Element_Linker payload."

    txn = Transaction(doc, "QAQC Fix Element_Linker ElementId {}".format(child_id_int))
    try:
        txn.Start()
        set_ok = _set_linker_text(elem, updated_text)
        if not set_ok:
            raise Exception("Failed to write Element_Linker parameter.")
        txn.Commit()
    except Exception as exc:
        try:
            txn.RollBack()
        except Exception:
            pass
        return False, "ElementId fix failed: {}".format(exc)

    return True, "Updated Element_Linker ElementId from {} to {}.".format(int(linker_element_id), child_id_int)


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
    global _MODELLESS_WINDOW, _ADJUST_HANDLER, _ADJUST_EXTERNAL_EVENT, _FIX_ID_HANDLER, _FIX_ID_EXTERNAL_EVENT
    global _PLACE_PROFILE_HANDLER, _PLACE_PROFILE_EXTERNAL_EVENT
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
    if not _ensure_fix_id_external_event():
        forms.alert(
            "Could not initialize modeless ElementId-fix event. Fix ID actions will be unavailable.",
            title=TITLE,
        )
    if not _ensure_place_profile_external_event():
        forms.alert(
            "Could not initialize modeless place-profile event. Place actions will fall back to inline placement.",
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
        + len(tabs.get("tab7") or [])
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
            "7) ID Discrepancies: {}".format(len(tabs.get("tab7") or [])),
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
        # Highlight first, zoom second. Priority order:
        #   1. Linked element (Tab 2/4/5 with a linked parent) — built
        #      via Reference.CreateLinkReference so the specific linked
        #      element gets selected, not the whole RevitLinkInstance.
        #   2. snap_select_id — the row's explicit host-doc target
        #      (parent_id for type-change tabs, child_id for far-from-
        #      parent / ID-discrepancy tabs).
        #   3. Fall back to child_id, then parent_id, so older rows
        #      that didn't populate snap_select_id still highlight
        #      something sensible.
        if row:
            link_inst_id = row.get("link_instance_id")
            linked_elem_id = row.get("linked_element_id")
            selected = False
            if link_inst_id is not None and linked_elem_id is not None:
                selected = _select_linked_element(link_inst_id, linked_elem_id)
            if not selected:
                host_id = row.get("snap_select_id")
                if host_id in (None, ""):
                    host_id = row.get("child_id")
                if host_id in (None, ""):
                    host_id = row.get("parent_id")
                if host_id not in (None, ""):
                    _select_element(host_id)

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

    def _on_fix_id(row):
        child_id = (row or {}).get("child_id")
        if child_id in (None, ""):
            forms.alert("No child element id found for this row.", title="{} - Fix ID".format(TITLE))
            return
        if _FIX_ID_HANDLER is None or _FIX_ID_EXTERNAL_EVENT is None:
            ok, message = _fix_element_linker_element_id(doc, child_id)
            forms.alert(message, title="{} - Fix ID".format(TITLE))
            if ok:
                _select_element(child_id)
            return
        try:
            _FIX_ID_HANDLER.request(child_id)
            _FIX_ID_EXTERNAL_EVENT.Raise()
        except Exception as exc:
            forms.alert("Failed to queue ElementId fix request:\n\n{}".format(exc), title="{} - Fix ID".format(TITLE))

    def _on_place(row):
        # Tab 2 "Place" button — runs the placement engine directly
        # against the row's parent (no dockable pane). Wrapped in the
        # external event so the engine has Revit's API context for its
        # transaction. Falls back to a direct call if the event can't
        # be created (e.g., running outside a UI context).
        if not isinstance(row, dict):
            return
        if not (row.get("profile") or "").strip():
            forms.alert("Row has no profile name to place.",
                        title="{} - Place Profile".format(TITLE))
            return
        if _PLACE_PROFILE_HANDLER is None or _PLACE_PROFILE_EXTERNAL_EVENT is None:
            ok, message = _do_place_profile_for_row(doc, data, row)
            forms.alert(message, title="{} - Place Profile".format(TITLE))
            return
        try:
            _PLACE_PROFILE_HANDLER.request(data, row)
            _PLACE_PROFILE_EXTERNAL_EVENT.Raise()
        except Exception as exc:
            forms.alert(
                "Failed to queue placement request:\n\n{}".format(exc),
                title="{} - Place Profile".format(TITLE),
            )

    window = ui_module.QAQCReportWindow(
        xaml_path=xaml_path,
        tab_rows=tabs,
        summary_text=summary_text,
        select_child_callback=_on_select_child,
        select_parent_callback=_on_select_parent,
        snap_callback=_on_snap,
        adjust_callback=_on_adjust,
        fix_id_callback=_on_fix_id,
        place_callback=_on_place,
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
