# -*- coding: utf-8 -*-
"""
Alias-based merge engine.

The 2.0 merge model treats a "merge" as a pure name-resolution alias.
A *source* profile carries a list under ``merged_aliases`` — alternate
strings (typically ``Family : Type`` labels) that should also resolve
to this source during placement matching. No structural data is ever
copied between profiles.

Legacy data captured before this rewrite used a deep-copy model with
``ced_truth_source_id`` / ``ced_truth_source_name`` on each member.
``has_legacy_members`` + ``migrate_legacy_members`` convert that data
into the new alias model on first run.
"""

import io
import os

import truth_groups


MERGED_ALIASES_KEY = "merged_aliases"


class MergeError(Exception):
    pass


# ---------------------------------------------------------------------
# Alias list management
# ---------------------------------------------------------------------

def _normalise_alias(value):
    if value is None:
        return ""
    return str(value).strip()


def _alias_match_key(value):
    """Case-insensitive, trimmed key for de-duplication."""
    return _normalise_alias(value).lower()


def aliases(profile):
    """Return the alias list on ``profile`` (creating it if missing)."""
    if not isinstance(profile, dict):
        return []
    raw = profile.get(MERGED_ALIASES_KEY)
    if not isinstance(raw, list):
        raw = []
        profile[MERGED_ALIASES_KEY] = raw
    return raw


def add_alias(profile, alias):
    """Append ``alias`` to the source profile's list (case-insensitively
    deduped). Returns True if the alias was actually added, False if it
    was already present or empty."""
    if not isinstance(profile, dict):
        raise MergeError("Invalid profile")
    clean = _normalise_alias(alias)
    if not clean:
        return False
    existing = aliases(profile)
    existing_keys = {_alias_match_key(a) for a in existing}
    if _alias_match_key(clean) in existing_keys:
        return False
    existing.append(clean)
    return True


def add_aliases(profile, alias_iterable):
    """Bulk-add. Returns ``(added_count, skipped_duplicates)``."""
    added = 0
    skipped = 0
    for a in alias_iterable or ():
        if add_alias(profile, a):
            added += 1
        else:
            skipped += 1
    return added, skipped


def remove_alias(profile, alias):
    """Remove ``alias`` (case-insensitive match). Returns True if removed."""
    if not isinstance(profile, dict):
        return False
    target_key = _alias_match_key(alias)
    existing = aliases(profile)
    for i, current in enumerate(existing):
        if _alias_match_key(current) == target_key:
            del existing[i]
            return True
    return False


def has_aliases(profile):
    return bool(aliases(profile))


def find_alias_owner(profile_data, alias):
    """Return the profile that lists ``alias`` in its merged_aliases, or None."""
    target_key = _alias_match_key(alias)
    if not target_key:
        return None
    for p in profile_data.get("equipment_definitions") or []:
        if not isinstance(p, dict):
            continue
        for a in p.get(MERGED_ALIASES_KEY) or ():
            if _alias_match_key(a) == target_key:
                return p
    return None


def all_alias_entries(profile_data):
    """Flat enumeration ``[(source_profile, alias_string), ...]``."""
    out = []
    for p in profile_data.get("equipment_definitions") or []:
        if not isinstance(p, dict):
            continue
        for a in p.get(MERGED_ALIASES_KEY) or ():
            out.append((p, a))
    return out


# ---------------------------------------------------------------------
# Legacy migration
# ---------------------------------------------------------------------

def has_legacy_members(profile_data):
    """True if any profile still has the legacy ``ced_truth_source_id``."""
    for p in profile_data.get("equipment_definitions") or []:
        if not isinstance(p, dict):
            continue
        if truth_groups.is_group_member(p):
            return True
    return False


def collect_legacy_members(profile_data):
    """Return list of profiles with ``ced_truth_source_id`` set."""
    return [
        p for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict) and truth_groups.is_group_member(p)
    ]


class LegacyMigrationReport(object):
    def __init__(self):
        self.aliases_added = 0
        self.members_cleared = 0
        self.unresolved_members = []  # members whose source no longer exists


def migrate_legacy_members(profile_data):
    """Convert every legacy member into an alias on its source.

    For each profile with ``ced_truth_source_id``:
      * append the member's ``name`` to the source's ``merged_aliases``
        (deduped)
      * clear the member's ``ced_truth_source_id`` / ``ced_truth_source_name``

    The member profiles themselves are NOT deleted — that's a separate
    optional confirm in the calling UI.

    Returns a ``LegacyMigrationReport``.
    """
    report = LegacyMigrationReport()
    by_id = {
        p.get("id"): p
        for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict) and p.get("id")
    }
    for member in collect_legacy_members(profile_data):
        sid = truth_groups.truth_source_id(member)
        source = by_id.get(sid)
        if source is None:
            report.unresolved_members.append(member)
            # Clear the dangling tag anyway.
            truth_groups.clear_truth_source(member)
            report.members_cleared += 1
            continue
        member_name = (member.get("name") or "").strip()
        if member_name and add_alias(source, member_name):
            report.aliases_added += 1
        truth_groups.clear_truth_source(member)
        report.members_cleared += 1
    return report


def delete_profiles_by_id(profile_data, profile_ids):
    """Remove every profile whose ``id`` is in ``profile_ids``.
    Returns the count actually removed."""
    if not profile_ids:
        return 0
    target_ids = set(profile_ids)
    defs = profile_data.get("equipment_definitions") or []
    keep = []
    removed = 0
    for p in defs:
        if isinstance(p, dict) and p.get("id") in target_ids:
            removed += 1
            continue
        keep.append(p)
    profile_data["equipment_definitions"] = keep
    return removed


# ---------------------------------------------------------------------
# Bulk add via CSV
# ---------------------------------------------------------------------

_BULK_HEADER_SOURCE = ("source", "source_id", "source_name")
_BULK_HEADER_TARGET = ("target", "target_id", "target_name", "alias")


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
    def __init__(self, row_number, ok, message,
                 source_label="", target_label=""):
        self.row_number = row_number
        self.ok = ok
        self.message = message
        self.source_label = source_label
        self.target_label = target_label


def bulk_add_aliases_from_csv(profile_data, csv_path):
    """Read ``csv_path`` and for each row, append ``target`` as an alias
    on ``source``.

    CSV must have headers: a source column (``source`` / ``source_id`` /
    ``source_name``) and a target column (``target`` / ``target_id`` /
    ``target_name`` / ``alias``).

    The source value is matched against profile ids first, then names.
    The target value is added verbatim as an alias string — it is NOT
    matched against existing profiles, so external CAD-only names work.

    Returns ``[BulkRowResult, ...]``.
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
            "'target' / 'target_id' / 'target_name' / 'alias' columns. "
            "Headers seen: {}".format(headers)
        )

    by_id, by_name = _profile_lookup(profile_data)
    out = []
    for ridx, raw in enumerate(rows[1:], start=2):
        if not raw:
            continue
        src_key = (raw[src_idx] or "").strip() if src_idx < len(raw) else ""
        tgt_value = (raw[tgt_idx] or "").strip() if tgt_idx < len(raw) else ""
        if not src_key or not tgt_value:
            continue
        source = by_id.get(src_key) or by_name.get(src_key)
        if source is None:
            out.append(BulkRowResult(
                ridx, False,
                "Source {!r} not found in active store".format(src_key),
                source_label=src_key, target_label=tgt_value,
            ))
            continue
        added = add_alias(source, tgt_value)
        if added:
            out.append(BulkRowResult(
                ridx, True,
                "alias added",
                source_label=source.get("name") or source.get("id") or "?",
                target_label=tgt_value,
            ))
        else:
            out.append(BulkRowResult(
                ridx, False,
                "Alias already present (deduped) or empty",
                source_label=source.get("name") or source.get("id") or "?",
                target_label=tgt_value,
            ))
    return out
