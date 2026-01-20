# -*- coding: utf-8 -*-
"""
Audit Linked Model
------------------
Scan linked documents for element names that do not have profiles in the active
YAML, but look like close name matches to existing profiles.
"""

import copy
import os
import re
import sys
from collections import Counter
from difflib import SequenceMatcher

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilyInstance,
    Group,
    RevitLinkInstance,
)

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Audit Linked Model"
LOG = script.get_logger()

DIRECTION_TOKENS = {
    "left", "right", "lh", "rh", "l", "r", "lhs", "rhs", "left-hand", "right-hand",
}
VERSION_TOKEN_RE = re.compile(r"^v\\d+$", re.IGNORECASE)
SEPARATORS_RE = re.compile(r"[_/\\\\-]+")
NON_ALNUM_RE = re.compile(r"[^a-zA-Z0-9 ]+")
MIN_PARTIAL_SCORE = 80.0
MIN_COMMON_TOKENS = 2

try:
    basestring
except NameError:
    basestring = str


class _MissingChoice(object):
    def __init__(self, label, data, checked=True):
        self.label = label
        self.data = data
        self.checked = checked

    def __str__(self):
        return self.label


def _normalize_full_name(value):
    if not value:
        return ""
    text = SEPARATORS_RE.sub(" ", str(value))
    text = NON_ALNUM_RE.sub(" ", text)
    return " ".join(text.lower().split())


def _tokenize_name(value):
    normalized = _normalize_full_name(value)
    return [token for token in normalized.split() if token]


def _strip_trailing_id_tokens(tokens):
    while tokens:
        tail = tokens[-1]
        if tail.isdigit() or VERSION_TOKEN_RE.match(tail):
            tokens = tokens[:-1]
            continue
        break
    return tokens


def _strip_numeric_prefix(tokens):
    while tokens and tokens[0].isdigit() and len(tokens[0]) <= 3:
        tokens = tokens[1:]
    return tokens


def _strip_direction_tokens(tokens):
    return [token for token in tokens if token not in DIRECTION_TOKENS]


def _base_key(value):
    if not value:
        return ""
    tokens = _tokenize_name(value)
    tokens = _strip_numeric_prefix(tokens)
    tokens = _strip_trailing_id_tokens(tokens)
    tokens = _strip_direction_tokens(tokens)
    base = " ".join(tokens)
    return base or _normalize_full_name(value)


def token_set_ratio(a, b):
    a_tokens = set((a or "").lower().split())
    b_tokens = set((b or "").lower().split())
    if not a_tokens or not b_tokens:
        return 0.0
    common = " ".join(sorted(a_tokens & b_tokens))
    a_diff = " ".join(sorted(a_tokens - b_tokens))
    b_diff = " ".join(sorted(b_tokens - a_tokens))
    return max(
        SequenceMatcher(None, common, (common + " " + a_diff).strip()).ratio(),
        SequenceMatcher(None, common, (common + " " + b_diff).strip()).ratio()
    ) * 100.0


def _head_key(value):
    if not value:
        return ""
    head = str(value).split(":", 1)[0].strip()
    return _normalize_full_name(head)


def _match_tokens(value):
    if not value:
        return []
    tokens = _tokenize_name(value)
    tokens = _strip_numeric_prefix(tokens)
    tokens = _strip_trailing_id_tokens(tokens)
    tokens = _strip_direction_tokens(tokens)
    return tokens


def _token_overlap_count(left, right):
    left_tokens = set(_match_tokens(left))
    right_tokens = set(_match_tokens(right))
    if not left_tokens or not right_tokens:
        return 0
    return len(left_tokens & right_tokens)


def _best_match_name(target, candidates):
    target_norm = _normalize_full_name(target)
    best = None
    best_score = -1.0
    for name in candidates:
        score = token_set_ratio(target_norm, _normalize_full_name(name))
        if score > best_score:
            best_score = score
            best = name
    return best, best_score


def _build_label(elem):
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            return u"{} : {}".format(fam_name, type_name)
        if type_name:
            return type_name
        if fam_name:
            return fam_name
    try:
        name = getattr(elem, "Name", None)
        if name:
            return name
    except Exception:
        pass
    return ""


def _split_label(label):
    cleaned = (label or "").strip()
    if not cleaned:
        return "", ""
    if ":" in cleaned:
        fam_part, type_part = cleaned.split(":", 1)
        return fam_part.strip(), type_part.strip()
    return cleaned, ""


def _iter_link_docs(doc):
    if doc is None:
        return
    seen = set()

    def _walk(source_doc):
        if source_doc is None:
            return
        link_instances = []
        try:
            link_instances = FilteredElementCollector(source_doc).OfClass(RevitLinkInstance)
        except Exception:
            link_instances = []
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                continue
            try:
                key = link_doc.GetHashCode()
            except Exception:
                key = id(link_doc)
            if key in seen:
                continue
            seen.add(key)
            yield link_doc
            for nested in _walk(link_doc):
                yield nested

    for linked_doc in _walk(doc):
        yield linked_doc


def _collect_linked_names(doc):
    counts = Counter()
    for link_doc in _iter_link_docs(doc):
        try:
            fam_collector = FilteredElementCollector(link_doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
        except Exception:
            fam_collector = []
        for elem in fam_collector:
            label = _build_label(elem)
            if label:
                counts[label] += 1
        try:
            group_collector = FilteredElementCollector(link_doc).OfClass(Group).WhereElementIsNotElementType()
        except Exception:
            group_collector = []
        for elem in group_collector:
            label = _build_label(elem)
            if label:
                counts[label] += 1
    return counts


def _collect_profile_names(data):
    names = set()
    for eq in data.get("equipment_definitions") or []:
        raw = eq.get("name") or eq.get("id")
        if raw:
            names.add(str(raw).strip())
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                label = led.get("label") or led.get("name") or ""
                if label:
                    names.add(str(label).strip())
    return names


def _find_equipment_def_by_name(data, name):
    if not name:
        return None
    target = str(name).strip()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        eq_name = (eq.get("name") or eq.get("id") or "").strip()
        if eq_name == target:
            return eq
    return None


def _find_equipment_def_by_label(data, label):
    if not label:
        return None
    target = str(label).strip()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_label = (led.get("label") or led.get("name") or "").strip()
                if led_label == target:
                    return eq
    return None


def _next_eq_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        eq_id = (eq.get("id") or "").strip()
        if eq_id.upper().startswith("EQ-"):
            try:
                num = int(eq_id.split("-")[-1])
            except Exception:
                continue
            if num > max_id:
                max_id = num
    return max_id + 1


def _next_set_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            set_id = (linked_set.get("id") or "").strip()
            if not set_id.upper().startswith("SET-"):
                continue
            try:
                num = int(set_id.split("-")[-1])
            except Exception:
                continue
            if num > max_id:
                max_id = num
    return max_id + 1


def _rewrite_linker_payload(params, old_set_id, new_set_id, old_led_id, new_led_id):
    if not isinstance(params, dict):
        return
    for key in ("Element_Linker Parameter", "Element_Linker"):
        value = params.get(key)
        if not isinstance(value, basestring):
            continue
        updated = value
        if old_set_id:
            updated = updated.replace(old_set_id, new_set_id)
        if old_led_id:
            updated = updated.replace(old_led_id, new_led_id)
        if updated != value:
            params[key] = updated


def _clone_equipment_def(source_eq, new_name, new_eq_id, next_set_num):
    eq_copy = copy.deepcopy(source_eq)
    eq_copy["id"] = new_eq_id
    eq_copy["name"] = new_name
    truth_id = source_eq.get("ced_truth_source_id") or source_eq.get("id") or source_eq.get("name") or new_eq_id
    truth_name = source_eq.get("ced_truth_source_name") or source_eq.get("name") or new_name
    eq_copy["ced_truth_source_id"] = truth_id
    eq_copy["ced_truth_source_name"] = truth_name

    parent_filter = eq_copy.get("parent_filter")
    if isinstance(parent_filter, dict):
        family_name, type_name = _split_label(new_name)
        if family_name:
            parent_filter["family_name_pattern"] = family_name
        if ":" in (new_name or ""):
            parent_filter["type_name_pattern"] = type_name or "*"

    linked_sets = eq_copy.get("linked_sets") or []
    if not linked_sets:
        new_set_id = "SET-{:03d}".format(next_set_num)
        next_set_num += 1
        eq_copy["linked_sets"] = [{
            "id": new_set_id,
            "name": "{} Types".format(new_name),
            "linked_element_definitions": [],
        }]
        return eq_copy, next_set_num

    for idx, linked_set in enumerate(linked_sets):
        if not isinstance(linked_set, dict):
            continue
        old_set_id = linked_set.get("id")
        new_set_id = "SET-{:03d}".format(next_set_num)
        next_set_num += 1
        linked_set["id"] = new_set_id
        if idx == 0:
            linked_set["name"] = "{} Types".format(new_name)
        led_list = linked_set.get("linked_element_definitions") or []
        counter = 0
        for led in led_list:
            if not isinstance(led, dict):
                continue
            old_led_id = led.get("id")
            if led.get("is_parent_anchor"):
                new_led_id = "{}-LED-000".format(new_set_id)
            else:
                counter += 1
                new_led_id = "{}-LED-{:03d}".format(new_set_id, counter)
            led["id"] = new_led_id
            _rewrite_linker_payload(led.get("parameters"), old_set_id, new_set_id, old_led_id, new_led_id)
    return eq_copy, next_set_num


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active Revit document.", title=TITLE)
        return
    try:
        _, data = load_active_yaml_data()
    except Exception as exc:
        forms.alert(str(exc), title=TITLE)
        return

    profile_names = _collect_profile_names(data)
    if not profile_names:
        forms.alert("No profiles found in the active YAML.", title=TITLE)
        return

    profile_norms = {_normalize_full_name(name) for name in profile_names if name}
    profile_base_map = {}
    profile_default_heads = set()
    for name in profile_names:
        base = _base_key(name)
        if not base:
            continue
        profile_base_map.setdefault(base, []).append(name)
        head, type_name = _split_label(name)
        type_norm = _normalize_full_name(type_name)
        if type_norm == "default":
            head_norm = _normalize_full_name(head)
            if head_norm:
                profile_default_heads.add(head_norm)

    linked_counts = _collect_linked_names(doc)
    if not linked_counts:
        forms.alert("No linked elements found in any linked documents.", title=TITLE)
        return

    missing = []
    for linked_name, count in linked_counts.items():
        norm = _normalize_full_name(linked_name)
        if norm in profile_norms:
            continue
        if ":" not in linked_name and norm in profile_default_heads:
            continue
        base = _base_key(linked_name)
        if not base:
            continue
        candidates = profile_base_map.get(base) or []
        if candidates:
            best, _ = _best_match_name(linked_name, candidates)
            if best:
                missing.append((linked_name, best, count))
            continue
        best, score = _best_match_name(linked_name, profile_names)
        if not best:
            continue
        if score < MIN_PARTIAL_SCORE:
            continue
        if _token_overlap_count(linked_name, best) < MIN_COMMON_TOKENS:
            continue
        missing.append((linked_name, best, count))

    if not missing:
        forms.alert("No close-name profile gaps found in linked documents.", title=TITLE)
        return

    missing.sort(key=lambda row: row[0].lower())
    items = []
    for linked_name, best, count in missing:
        if count > 1:
            label = "{} (x{}) -> {}".format(linked_name, count, best)
        else:
            label = "{} -> {}".format(linked_name, best)
        items.append(_MissingChoice(label, (linked_name, best, count), checked=True))

    selected = forms.SelectFromList.show(
        items,
        title=TITLE,
        multiselect=True,
        button_name="Create Selected",
        width=900,
        height=600,
    )
    if not selected:
        return
    missing = [item.data for item in selected]

    equipment_defs = data.get("equipment_definitions") or []
    existing_norms = {
        _normalize_full_name(eq.get("name") or eq.get("id") or "")
        for eq in equipment_defs
        if isinstance(eq, dict)
    }
    next_eq_num = _next_eq_number(data)
    next_set_num = _next_set_number(data)
    created = []
    skipped_existing = []
    skipped_unresolved = []
    for linked_name, best, _ in missing:
        new_norm = _normalize_full_name(linked_name)
        if new_norm in existing_norms:
            skipped_existing.append(linked_name)
            continue
        source_eq = _find_equipment_def_by_name(data, best) or _find_equipment_def_by_label(data, best)
        if not source_eq:
            skipped_unresolved.append((linked_name, best))
            continue
        new_eq_id = "EQ-{:03d}".format(next_eq_num)
        next_eq_num += 1
        eq_copy, next_set_num = _clone_equipment_def(source_eq, linked_name, new_eq_id, next_set_num)
        equipment_defs.append(eq_copy)
        existing_norms.add(new_norm)
        created.append((linked_name, best))

    if not created:
        forms.alert("No new profiles were created.", title=TITLE)
        return

    try:
        save_active_yaml_data(
            None,
            data,
            "Audit Linked Model",
            "Created {} profile(s) from audit results".format(len(created)),
        )
    except Exception as exc:
        forms.alert("Failed to save updates:\n\n{}".format(exc), title=TITLE)
        return

    summary = [
        "Created {} profile(s).".format(len(created)),
    ]
    summary.extend(" - {} -> {}".format(name, best) for name, best in created[:20])
    if len(created) > 20:
        summary.append(" (+{} more)".format(len(created) - 20))
    if skipped_existing:
        summary.append("")
        summary.append("Skipped existing:")
        summary.extend(" - {}".format(name) for name in skipped_existing[:10])
        if len(skipped_existing) > 10:
            summary.append(" (+{} more)".format(len(skipped_existing) - 10))
    if skipped_unresolved:
        summary.append("")
        summary.append("Skipped (no source profile found):")
        summary.extend(" - {} -> {}".format(name, best) for name, best in skipped_unresolved[:10])
        if len(skipped_unresolved) > 10:
            summary.append(" (+{} more)".format(len(skipped_unresolved) - 10))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
