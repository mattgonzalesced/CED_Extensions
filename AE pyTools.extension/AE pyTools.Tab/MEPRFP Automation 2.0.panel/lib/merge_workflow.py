# -*- coding: utf-8 -*-
"""
Merge / unmerge engine.

A *merge* declares one profile (the ``source``) as the truth and copies
its structural content (parent_filter, linked_sets, equipment_properties,
flags) into one or more *target* profiles. Targets keep their own ``id``
and ``name`` and are tagged with ``ced_truth_source_id`` /
``ced_truth_source_name`` so audits can trace lineage.

Linked-set / LED / annotation IDs are *renumbered* in the target so two
different profiles never share the same ``SET-NNN-LED-NNN`` IDs. Without
this, Element_Linker references on placed children become ambiguous.

The pure-Python entry points here have no Revit-API dependency; the
pushbutton scripts wrap them with WPF dialogs.
"""

import copy
import io
import os
import re

import truth_groups


class MergeError(Exception):
    pass


# ---------------------------------------------------------------------
# Eligibility
# ---------------------------------------------------------------------

def is_source_for_others(profile_data, profile):
    """True if any other profile in the store points at ``profile`` as its truth source."""
    pid = profile.get("id") if isinstance(profile, dict) else None
    if not pid:
        return False
    for p in profile_data.get("equipment_definitions") or []:
        if not isinstance(p, dict) or p is profile:
            continue
        if truth_groups.truth_source_id(p) == pid:
            return True
    return False


def can_be_target(profile_data, source_profile, candidate_profile):
    """Returns ``(ok: bool, reason: str)``.

    Forbids:
        * candidate == source
        * candidate is already a member of any group
        * candidate is itself a source for other merged profiles
    """
    if candidate_profile is source_profile:
        return False, "Cannot merge a profile into itself"
    if not isinstance(candidate_profile, dict):
        return False, "Invalid candidate"
    if truth_groups.is_group_member(candidate_profile):
        sid = truth_groups.truth_source_id(candidate_profile)
        return False, "Already a member of group {}".format(sid or "?")
    if is_source_for_others(profile_data, candidate_profile):
        return False, "Already a source for other merged profiles — unmerge its members first"
    return True, ""


def can_be_source(profile_data, candidate_profile):
    """A profile can be a source unless it's itself a member of another group.

    (Members of groups can't lead their own groups — that's a chain we
    refuse to build.)
    """
    if not isinstance(candidate_profile, dict):
        return False, "Invalid candidate"
    if truth_groups.is_group_member(candidate_profile):
        sid = truth_groups.truth_source_id(candidate_profile)
        return False, "Profile is itself a member of group {} — unmerge it first".format(sid or "?")
    return True, ""


def eligible_targets(profile_data, source_profile):
    """All profiles that can currently be merged into ``source_profile``."""
    out = []
    for p in profile_data.get("equipment_definitions") or []:
        ok, _ = can_be_target(profile_data, source_profile, p)
        if ok:
            out.append(p)
    return out


def eligible_sources(profile_data):
    out = []
    for p in profile_data.get("equipment_definitions") or []:
        ok, _ = can_be_source(profile_data, p)
        if ok:
            out.append(p)
    return out


# ---------------------------------------------------------------------
# Renumbering helpers (testable offline)
# ---------------------------------------------------------------------

_SUFFIX_RE = re.compile(r"(\d+)$")


def _max_numeric_suffix(strings, prefix):
    best = 0
    for s in strings:
        if not isinstance(s, str) or not s.startswith(prefix):
            continue
        rest = s[len(prefix):]
        try:
            n = int(rest)
        except ValueError:
            continue
        if n > best:
            best = n
    return best


def _collect_set_ids(profile_data):
    out = set()
    for p in profile_data.get("equipment_definitions") or []:
        if not isinstance(p, dict):
            continue
        for s in p.get("linked_sets") or []:
            if isinstance(s, dict) and isinstance(s.get("id"), str):
                out.add(s["id"])
    return out


def renumber_linked_sets(linked_sets, existing_set_ids):
    """Replace every SET / LED / ANN id in ``linked_sets`` (mutates) with
    fresh ids that don't collide with ``existing_set_ids``.

    ``existing_set_ids`` is a *mutable* set that this function adds to
    so callers chaining multiple renumber passes share state.
    """
    next_n = _max_numeric_suffix(existing_set_ids, "SET-") + 1
    for set_dict in linked_sets:
        if not isinstance(set_dict, dict):
            continue
        old_set_id = set_dict.get("id") or ""
        new_set_id = "SET-{:03d}".format(next_n)
        next_n += 1
        existing_set_ids.add(new_set_id)
        set_dict["id"] = new_set_id
        for led in set_dict.get("linked_element_definitions") or []:
            if not isinstance(led, dict):
                continue
            old_led_id = led.get("id") or ""
            led["id"] = _swap_id_prefix(old_led_id, old_set_id, new_set_id)
            new_led_id = led["id"]
            for ann in led.get("annotations") or []:
                if not isinstance(ann, dict):
                    continue
                old_ann_id = ann.get("id") or ""
                ann["id"] = _swap_id_prefix(old_ann_id, old_led_id, new_led_id)
    return linked_sets


def _swap_id_prefix(old_id, old_prefix, new_prefix):
    """If ``old_id`` starts with ``old_prefix + '-'``, replace that prefix
    with ``new_prefix + '-'``. Otherwise return ``old_id`` unchanged
    (defensive fallback for malformed ids — better than crashing)."""
    if not old_id:
        return old_id
    needle = old_prefix + "-"
    if old_id.startswith(needle):
        return new_prefix + "-" + old_id[len(needle):]
    return old_id


# ---------------------------------------------------------------------
# Merge / unmerge
# ---------------------------------------------------------------------

# Fields copied from source -> target on merge. ``id`` and ``name`` are
# explicitly preserved on the target. ``ced_truth_source_*`` is set
# explicitly after the copy.
_COPY_FIELDS = (
    "schema_version",
    "parent_filter",
    "linked_sets",
    "equipment_properties",
    "allow_parentless",
    "allow_unmatched_parents",
    "prompt_on_parent_mismatch",
)


def merge_into(profile_data, source_profile, target_profile):
    """Apply one merge. Mutates ``profile_data`` in place. Raises ``MergeError``
    on eligibility failure."""
    ok, reason = can_be_source(profile_data, source_profile)
    if not ok:
        raise MergeError(reason)
    ok, reason = can_be_target(profile_data, source_profile, target_profile)
    if not ok:
        raise MergeError(reason)

    existing_set_ids = _collect_set_ids(profile_data)
    # Remove the target's OWN set ids from the "existing" set so the
    # renumbered ids of the copy don't collide with the target's about-
    # to-be-replaced sets.
    for s in target_profile.get("linked_sets") or []:
        if isinstance(s, dict) and isinstance(s.get("id"), str):
            existing_set_ids.discard(s["id"])

    new_sets = renumber_linked_sets(
        copy.deepcopy(source_profile.get("linked_sets") or []),
        existing_set_ids,
    )

    for field in _COPY_FIELDS:
        if field == "linked_sets":
            target_profile[field] = new_sets
        elif field in source_profile:
            target_profile[field] = copy.deepcopy(source_profile[field])

    truth_groups.set_truth_source(
        target_profile,
        source_profile.get("id"),
        source_profile.get("name"),
    )


def merge_many(profile_data, source_profile, target_profiles):
    """Apply merges to many targets. Returns ``(succeeded, failed)`` lists.

    Each ``failed`` entry is ``(target_profile, reason_str)``.
    """
    succeeded, failed = [], []
    for t in target_profiles:
        try:
            merge_into(profile_data, source_profile, t)
            succeeded.append(t)
        except MergeError as exc:
            failed.append((t, str(exc)))
    return succeeded, failed


def unmerge(profile_data, member_profile):
    """Detach ``member_profile`` from its truth-source group."""
    if not isinstance(member_profile, dict):
        raise MergeError("Invalid profile")
    if not truth_groups.is_group_member(member_profile):
        raise MergeError("Profile is not currently a group member")
    truth_groups.clear_truth_source(member_profile)


# ---------------------------------------------------------------------
# Bulk merge from CSV
# ---------------------------------------------------------------------

_BULK_HEADER_SOURCE = ("source", "source_id", "source_name")
_BULK_HEADER_TARGET = ("target", "target_id", "target_name")


def _resolve_header(headers, candidates):
    lower = [(h or "").strip().lower() for h in headers]
    for key in candidates:
        if key in lower:
            return lower.index(key)
    return None


def _profile_lookup(profile_data):
    by_id, by_name = {}, {}
    for p in profile_data.get("equipment_definitions") or []:
        if not isinstance(p, dict):
            continue
        if p.get("id"):
            by_id[p["id"]] = p
        if p.get("name"):
            by_name[p["name"]] = p
    return by_id, by_name


class BulkRowResult(object):
    """One row's outcome from bulk-merge CSV processing."""

    def __init__(self, row_number, ok, message, source_label="", target_label=""):
        self.row_number = row_number
        self.ok = ok
        self.message = message
        self.source_label = source_label
        self.target_label = target_label


def bulk_merge_from_csv(profile_data, csv_path):
    """Read ``csv_path`` and apply each row as a merge.

    The CSV must have a row of headers and at least two columns: a
    source column (``source`` / ``source_id`` / ``source_name``) and a
    target column (``target`` / ``target_id`` / ``target_name``).
    Each value is matched against profile ids first, then names.

    Returns ``[BulkRowResult, ...]``. Mutations to ``profile_data`` are
    cumulative — successful rows persist even if later rows fail.
    """
    if not csv_path or not os.path.isfile(csv_path):
        raise MergeError("CSV not found: {}".format(csv_path))
    import csv as _csv
    rows = []
    with io.open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = _csv.reader(f)
        for raw in reader:
            rows.append(raw)
    if not rows:
        raise MergeError("CSV is empty")
    headers = rows[0]
    src_idx = _resolve_header(headers, _BULK_HEADER_SOURCE)
    tgt_idx = _resolve_header(headers, _BULK_HEADER_TARGET)
    if src_idx is None or tgt_idx is None:
        raise MergeError(
            "CSV must have 'source' / 'source_id' / 'source_name' AND "
            "'target' / 'target_id' / 'target_name' columns. "
            "Headers seen: {}".format(headers)
        )

    by_id, by_name = _profile_lookup(profile_data)
    out = []
    for ridx, raw in enumerate(rows[1:], start=2):
        if not raw:
            continue
        src_key = (raw[src_idx] or "").strip() if src_idx < len(raw) else ""
        tgt_key = (raw[tgt_idx] or "").strip() if tgt_idx < len(raw) else ""
        if not src_key or not tgt_key:
            continue
        source = by_id.get(src_key) or by_name.get(src_key)
        target = by_id.get(tgt_key) or by_name.get(tgt_key)
        if source is None:
            out.append(BulkRowResult(
                ridx, False,
                "Source {!r} not found in active store".format(src_key),
                source_label=src_key, target_label=tgt_key,
            ))
            continue
        if target is None:
            out.append(BulkRowResult(
                ridx, False,
                "Target {!r} not found in active store".format(tgt_key),
                source_label=src_key, target_label=tgt_key,
            ))
            continue
        try:
            merge_into(profile_data, source, target)
            out.append(BulkRowResult(
                ridx, True,
                "merged",
                source_label=source.get("name") or source.get("id") or "?",
                target_label=target.get("name") or target.get("id") or "?",
            ))
        except MergeError as exc:
            out.append(BulkRowResult(
                ridx, False, str(exc),
                source_label=source.get("name") or source.get("id") or "?",
                target_label=target.get("name") or target.get("id") or "?",
            ))
    return out
