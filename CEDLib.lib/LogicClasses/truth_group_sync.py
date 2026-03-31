# -*- coding: utf-8 -*-
"""
Shared truth-group synchronization helpers.

The source profile in a truth group is authoritative. If a non-source member was
edited since the previous snapshot, that member's payload is promoted to the
source, then propagated to the rest of the group.
"""

import copy
import json

try:
    from collections.abc import Mapping
except ImportError:
    from collections import Mapping

TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"

_IDENTITY_KEYS = set(("id", "name", TRUTH_SOURCE_ID_KEY, TRUTH_SOURCE_NAME_KEY))


def _entry_id(entry):
    if not isinstance(entry, Mapping):
        return ""
    value = entry.get("id")
    if value is None:
        return ""
    return str(value).strip()


def _entry_name(entry):
    if not isinstance(entry, Mapping):
        return ""
    value = entry.get("name") or entry.get("id")
    if value is None:
        return ""
    return str(value).strip()


def _entry_truth_source_id(entry):
    if not isinstance(entry, Mapping):
        return ""
    value = entry.get(TRUTH_SOURCE_ID_KEY)
    if value is None:
        return ""
    return str(value).strip()


def _entry_payload(entry):
    payload = {}
    if not isinstance(entry, Mapping):
        return payload
    for key, value in entry.items():
        if key in _IDENTITY_KEYS:
            continue
        payload[key] = copy.deepcopy(value)
    return payload


def _replace_entry_payload(entry, payload):
    if not isinstance(entry, dict):
        return False
    before = _entry_payload(entry)
    if before == payload:
        return False
    keep_id = entry.get("id")
    keep_name = entry.get("name")
    entry.clear()
    if keep_id not in (None, ""):
        entry["id"] = keep_id
    if keep_name not in (None, ""):
        entry["name"] = keep_name
    for key, value in payload.items():
        entry[key] = copy.deepcopy(value)
    return True


def _payload_signature(payload):
    try:
        return json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
    except Exception:
        return repr(payload)


def _build_previous_lookup(previous_defs):
    by_id = {}
    by_name = {}
    for entry in previous_defs or []:
        if not isinstance(entry, Mapping):
            continue
        eq_id = _entry_id(entry)
        eq_name = _entry_name(entry)
        if eq_id and eq_id not in by_id:
            by_id[eq_id] = entry
        if eq_name and eq_name not in by_name:
            by_name[eq_name] = entry
    return by_id, by_name


def _find_previous_entry(entry, previous_by_id, previous_by_name):
    eq_id = _entry_id(entry)
    if eq_id and eq_id in previous_by_id:
        return previous_by_id.get(eq_id)
    eq_name = _entry_name(entry)
    if eq_name and eq_name in previous_by_name:
        return previous_by_name.get(eq_name)
    return None


def _build_groups(equipment_defs):
    groups = {}
    for idx, entry in enumerate(equipment_defs or []):
        if not isinstance(entry, Mapping):
            continue
        source_id = _entry_truth_source_id(entry)
        if not source_id:
            source_id = _entry_id(entry) or _entry_name(entry) or ("group-{}".format(idx + 1))
        groups.setdefault(source_id, []).append(idx)
    return groups


def _pick_source_index(equipment_defs, member_indices, source_id):
    source_lower = (source_id or "").strip().lower()
    for idx in member_indices:
        eq_id = _entry_id(equipment_defs[idx]).lower()
        if source_lower and eq_id and eq_id == source_lower:
            return idx
    for idx in member_indices:
        entry = equipment_defs[idx]
        eq_id = _entry_id(entry)
        truth_id = _entry_truth_source_id(entry)
        if eq_id and truth_id and eq_id.lower() == truth_id.lower():
            return idx
    return member_indices[0]


def _set_truth_metadata(entry, source_id, source_name):
    if not isinstance(entry, dict):
        return False
    before_id = _entry_truth_source_id(entry)
    before_name = (entry.get(TRUTH_SOURCE_NAME_KEY) or "").strip()
    after_id = (source_id or "").strip()
    after_name = (source_name or "").strip() or after_id
    if after_id:
        entry[TRUTH_SOURCE_ID_KEY] = after_id
    if after_name:
        entry[TRUTH_SOURCE_NAME_KEY] = after_name
    return before_id != after_id or before_name != after_name


def synchronize_truth_groups(data, previous_data=None):
    """
    Mutate profile data in place to enforce truth-group consistency.

    Returns a report dictionary.
    """
    report = {
        "changed": False,
        "total_groups": 0,
        "total_profiles": 0,
        "groups_with_drift": 0,
        "groups_promoted_from_member": 0,
        "groups_with_conflicting_member_changes": 0,
        "profiles_payload_synced": 0,
        "profiles_metadata_repaired": 0,
    }

    if not isinstance(data, Mapping):
        return report
    equipment_defs = data.get("equipment_definitions")
    if not isinstance(equipment_defs, list):
        return report

    report["total_profiles"] = sum(1 for entry in equipment_defs if isinstance(entry, Mapping))

    previous_defs = []
    if isinstance(previous_data, Mapping):
        maybe_defs = previous_data.get("equipment_definitions")
        if isinstance(maybe_defs, list):
            previous_defs = maybe_defs
    previous_by_id, previous_by_name = _build_previous_lookup(previous_defs)

    groups = _build_groups(equipment_defs)
    report["total_groups"] = len(groups)

    for declared_source_id, member_indices in groups.items():
        if not member_indices:
            continue
        source_idx = _pick_source_index(equipment_defs, member_indices, declared_source_id)
        source_entry = equipment_defs[source_idx]
        source_payload = _entry_payload(source_entry)

        signatures = set()
        for idx in member_indices:
            signatures.add(_payload_signature(_entry_payload(equipment_defs[idx])))
        if len(signatures) > 1:
            report["groups_with_drift"] += 1

        source_prev = _find_previous_entry(source_entry, previous_by_id, previous_by_name)
        source_changed = False
        if source_prev is not None:
            source_changed = _entry_payload(source_prev) != source_payload

        changed_members = []
        for idx in member_indices:
            if idx == source_idx:
                continue
            entry = equipment_defs[idx]
            prev_entry = _find_previous_entry(entry, previous_by_id, previous_by_name)
            if prev_entry is None:
                continue
            payload = _entry_payload(entry)
            if payload != _entry_payload(prev_entry):
                changed_members.append((idx, payload))

        authoritative_payload = source_payload
        if not source_changed and changed_members:
            if len(changed_members) == 1:
                authoritative_payload = changed_members[0][1]
                report["groups_promoted_from_member"] += 1
            else:
                unique_payloads = {}
                for idx, payload in changed_members:
                    sig = _payload_signature(payload)
                    if sig not in unique_payloads:
                        unique_payloads[sig] = (idx, payload)
                if len(unique_payloads) == 1:
                    authoritative_payload = list(unique_payloads.values())[0][1]
                    report["groups_promoted_from_member"] += 1
                else:
                    chosen = changed_members[0]
                    authoritative_payload = chosen[1]
                    report["groups_promoted_from_member"] += 1
                    report["groups_with_conflicting_member_changes"] += 1

        if _replace_entry_payload(source_entry, authoritative_payload):
            report["profiles_payload_synced"] += 1
            report["changed"] = True

        source_id = _entry_id(source_entry) or _entry_truth_source_id(source_entry) or _entry_name(source_entry)
        source_name = _entry_name(source_entry) or source_id

        if _set_truth_metadata(source_entry, source_id, source_name):
            report["profiles_metadata_repaired"] += 1
            report["changed"] = True

        for idx in member_indices:
            if idx == source_idx:
                continue
            entry = equipment_defs[idx]
            if _replace_entry_payload(entry, authoritative_payload):
                report["profiles_payload_synced"] += 1
                report["changed"] = True
            if _set_truth_metadata(entry, source_id, source_name):
                report["profiles_metadata_repaired"] += 1
                report["changed"] = True

    return report


def normalize_truth_groups(data):
    """One-shot normalization helper for existing YAML payloads."""
    return synchronize_truth_groups(data, previous_data=None)


__all__ = [
    "TRUTH_SOURCE_ID_KEY",
    "TRUTH_SOURCE_NAME_KEY",
    "normalize_truth_groups",
    "synchronize_truth_groups",
]
