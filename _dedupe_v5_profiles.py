# -*- coding: utf-8 -*-
"""
One-off dedupe: HEB_profiles_V5.yaml -> HEB_profiles_V5_deduped.yaml.

Group profiles by **family + type + LED-label set**: two profiles only
merge if they share their ``parent_filter.family_name_pattern``,
``parent_filter.type_name_pattern``, AND the exact set of LED labels
under their linked_sets. Profiles that share a family but have
different fixture children (Manager's vs. Customer Service workstation
sharing the same parent family but capturing different LEDs each)
stay separate — that variation is meaningful, not duplication.

Within each group with >1 profile:

  * Pick the master: most LEDs first; alphabetically lowest profile id
    on tie.
  * Add every non-master profile's name to the master's
    ``merged_aliases`` list (case-insensitively deduped).
  * Drop the non-master profiles from ``equipment_definitions``.

Source file is read-only; the consolidated result is written to a new
file alongside it. Re-import that file in Revit to refresh the active
store.
"""

from __future__ import print_function

import io
import os
import re
import sys

import yaml


# Trailing space + digits at end of a profile name. Captures the
# Revit auto-increment pattern (``Default``, ``Default 2``,
# ``Default 3``) without touching embedded digits like ``_RH_1`` or
# ``_1304819`` (those have no leading whitespace).
_TRAILING_INC_RE = re.compile(r"\s+\d+\s*$")


SRC = r"c:\CED_Extensions\HEB_profiles_V5.yaml"
DST = r"c:\CED_Extensions\HEB_profiles_V5_deduped.yaml"


def _led_labels(profile):
    """Sorted tuple of all LED labels in the profile (case-folded)."""
    out = set()
    for s in profile.get("linked_sets") or []:
        if not isinstance(s, dict):
            continue
        for led in s.get("linked_element_definitions") or []:
            if isinstance(led, dict):
                lbl = (led.get("label") or "").strip().lower()
                if lbl:
                    out.add(lbl)
    return tuple(sorted(out))


def _normalize_name(name):
    """Lowercase, trim, and strip a trailing space-prefixed integer
    (Revit's auto-increment pattern) so ``Default`` and ``Default 2``
    collapse to the same key but ``..._LH_1304819`` (no leading space)
    is left untouched."""
    n = (name or "").strip().lower()
    n = _TRAILING_INC_RE.sub("", n)
    return n.strip()


def _key(profile):
    """Group key: (normalized name, family, type, sorted LED-label set).

    Profile name is the strongest discriminator: if the user gave two
    profiles materially different names (e.g. ``..._RH_...`` vs
    ``..._LH_...``), they stay separate. Trailing ``2`` / ``3``
    auto-increments are stripped so re-captures of the same equipment
    (which Revit auto-numbers) still merge.
    """
    pf = profile.get("parent_filter") or {}
    fam = (pf.get("family_name_pattern") or "").strip().lower()
    typ = (pf.get("type_name_pattern") or "").strip().lower()
    name = _normalize_name(profile.get("name"))
    return (name, fam, typ, _led_labels(profile))


def _led_count(profile):
    n = 0
    for s in profile.get("linked_sets") or []:
        if isinstance(s, dict):
            n += len(s.get("linked_element_definitions") or [])
    return n


def _alias_match_key(value):
    return str(value or "").strip().lower()


def _add_alias(master, alias):
    """Append ``alias`` to ``master.merged_aliases`` if not already
    present (case-insensitive). Returns True if added."""
    clean = (alias or "").strip()
    if not clean:
        return False
    existing = master.get("merged_aliases")
    if not isinstance(existing, list):
        existing = []
        master["merged_aliases"] = existing
    seen = {_alias_match_key(a) for a in existing}
    if _alias_match_key(clean) in seen:
        return False
    existing.append(clean)
    return True


def main():
    print("Reading {} ...".format(SRC))
    with io.open(SRC, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    if not isinstance(data, dict):
        print("ERROR: top-level YAML is not a mapping", file=sys.stderr)
        return 1

    profiles = list(data.get("equipment_definitions") or [])
    print("Profiles before dedup: {}".format(len(profiles)))

    groups = {}
    order = []
    for p in profiles:
        if not isinstance(p, dict):
            continue
        k = _key(p)
        if k not in groups:
            groups[k] = []
            order.append(k)
        groups[k].append(p)

    kept = []
    aliases_added = 0
    profiles_dropped = 0
    duplicate_groups = 0

    print("\nGroup summary (only groups with duplicates shown):")
    for k in order:
        group = groups[k]
        if len(group) == 1:
            kept.append(group[0])
            continue
        duplicate_groups += 1

        sorted_group = sorted(
            group,
            key=lambda p: (
                -_led_count(p),
                p.get("id") or "",
            ),
        )
        master = sorted_group[0]
        duplicates = sorted_group[1:]

        for dup in duplicates:
            name = (dup.get("name") or "").strip()
            if name and _add_alias(master, name):
                aliases_added += 1
            for old_alias in dup.get("merged_aliases") or []:
                if isinstance(old_alias, str) and _add_alias(master, old_alias):
                    aliases_added += 1
        profiles_dropped += len(duplicates)
        kept.append(master)

        name_key, fam, typ, _labels = k
        print(
            "  [{}]  master={}  drops={}  master_LEDs={}".format(
                name_key or "(empty)",
                master.get("id") or "?",
                len(duplicates),
                _led_count(master),
            )
        )

    data["equipment_definitions"] = kept

    print("\n----- Result -----")
    print("Profiles after dedup:   {}".format(len(kept)))
    print("Duplicate groups:       {}".format(duplicate_groups))
    print("Profiles dropped:       {}".format(profiles_dropped))
    print("Aliases added:          {}".format(aliases_added))

    print("\nWriting {} ...".format(DST))
    with io.open(DST, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            data,
            f,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
            width=4096,
        )
    print("Done. Bytes written: {}".format(os.path.getsize(DST)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
