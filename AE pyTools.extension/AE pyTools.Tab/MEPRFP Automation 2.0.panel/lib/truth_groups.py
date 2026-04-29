# -*- coding: utf-8 -*-
"""
Helpers for the truth-source / merge-lineage metadata.

When a profile is "merged" into another, the destination definition gets
``ced_truth_source_id`` and ``ced_truth_source_name`` fields pointing at
the source. These functions read and mutate that metadata; detection of
*drift* between a member and its source is a stage 1+ concern and lives
elsewhere.
"""

TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"


def truth_source_id(profile):
    if not isinstance(profile, dict):
        return None
    value = profile.get(TRUTH_SOURCE_ID_KEY)
    return value or None


def truth_source_name(profile):
    if not isinstance(profile, dict):
        return None
    value = profile.get(TRUTH_SOURCE_NAME_KEY)
    return value or None


def is_group_member(profile):
    """True if the profile carries a truth-source pointer."""
    return truth_source_id(profile) is not None


def set_truth_source(profile, source_id, source_name):
    """Tag a profile as a member of a merge group."""
    if not isinstance(profile, dict):
        raise TypeError("profile must be a dict")
    if not source_id:
        raise ValueError("source_id is required")
    profile[TRUTH_SOURCE_ID_KEY] = source_id
    if source_name is not None:
        profile[TRUTH_SOURCE_NAME_KEY] = source_name


def clear_truth_source(profile):
    """Detach a profile from its merge group (the unmerge operation)."""
    if not isinstance(profile, dict):
        raise TypeError("profile must be a dict")
    profile.pop(TRUTH_SOURCE_ID_KEY, None)
    profile.pop(TRUTH_SOURCE_NAME_KEY, None)


def find_group_source(profiles, source_id):
    """Return the profile whose ``id`` matches ``source_id``, or None."""
    if not source_id:
        return None
    for p in profiles or ():
        if isinstance(p, dict) and p.get("id") == source_id:
            return p
    return None


def find_group_members(profiles, source_id):
    """Return all profiles whose ``ced_truth_source_id`` matches ``source_id``.

    The source profile itself is *not* included.
    """
    if not source_id:
        return []
    out = []
    for p in profiles or ():
        if not isinstance(p, dict):
            continue
        if p.get(TRUTH_SOURCE_ID_KEY) == source_id and p.get("id") != source_id:
            out.append(p)
    return out


def group_members_by_source(profiles):
    """Index ``{source_id: [member_profile, ...]}`` over the input list."""
    out = {}
    for p in profiles or ():
        if not isinstance(p, dict):
            continue
        sid = p.get(TRUTH_SOURCE_ID_KEY)
        if not sid:
            continue
        if p.get("id") == sid:
            continue  # the source itself, skip
        out.setdefault(sid, []).append(p)
    return out
