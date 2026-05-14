# -*- coding: utf-8 -*-
"""
QAQC — health-check the relationship between the YAML store and the
elements actually placed in the model.

Nine categories::

    A   Orphan profile            Profile in YAML with zero placed
                                  instances AND no matching parent
                                  found anywhere in the host or linked
                                  docs — the profile is irrelevant to
                                  this project. (The "matching parent
                                  exists but no children placed" case
                                  is intentionally NOT flagged here;
                                  Cat G surfaces it per-parent.) Skips
                                  profiles that are merged into another
                                  (legacy ced_truth_source_id members
                                  or alias-target profiles).
    B   Missing parent            Placed child references a
                                  parent_element_id that no longer
                                  exists in the host or linked docs
                                  AND no name-matched parent exists
                                  either — child is genuinely orphan.
    C   Far from parent           Placed child sits more than
                                  Tolerances.FAR_FROM_PARENT_FT from
                                  the pose its stored offset implies.
    D   Child ID discrepancy      Element_Linker.element_id does not
                                  match the live element's Id.Value.
                                  Usually a sign of a copy/paste.
    D2  Parent ID discrepancy     Element_Linker.parent_element_id is
                                  dead, but a parent whose family
                                  matches Element_Linker.host_name DOES
                                  exist in the model. The parent was
                                  likely copy/pasted or had its id
                                  changed by a link reload. Auto-fix
                                  refreshes parent_element_id when
                                  exactly one name-match is found.
    E   Parent type change        Parent's family/type doesn't match
        (no profile)              the profile's parent_filter AND isn't
                                  in the profile's merged_aliases AND
                                  no OTHER profile in the YAML matches
                                  the new family — the new family is
                                  unknown to the profile store.
    E2  Parent type change        Same as E, but a DIFFERENT profile in
        (other profile)           the YAML does match the new family —
                                  the child should probably be
                                  reassigned to that profile.
    F   Host-name mismatch        Parent's current family name does
                                  not match Element_Linker.host_name,
                                  ignoring trailing `` : Default`` /
                                  `` : Default <n>`` decoration.
    G   Missing children          A parent (host or linked) matches a
                                  profile's parent_filter / aliases,
                                  the profile has at least one LED, but
                                  no placed child carries that profile's
                                  led_id back to this parent. The
                                  profile is configured but never
                                  fired against this parent.

Each finding carries a ``fix_kind`` describing what (if any) automated
fix can be applied. The window dispatches on that, calling
``execute_fix()`` inside a Revit transaction.
"""

import math
import re

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    LocationPoint,
    RevitLinkInstance,
)

import element_linker as _el
import element_linker_io as _el_io
import follow_parent_workflow as _fp
import geometry
import links
import placement as _placement
import truth_groups


# ---------------------------------------------------------------------
# Category constants
# ---------------------------------------------------------------------

CAT_A = "A"
CAT_B = "B"
CAT_C = "C"
CAT_D = "D"
CAT_D2 = "D2"
CAT_E = "E"
CAT_E2 = "E2"
CAT_F = "F"
CAT_G = "G"

# Display order: D / D2 and E / E2 sit adjacent so the window groups
# the parent-id and type-change variants visually.
CAT_ALL = (CAT_A, CAT_B, CAT_C, CAT_D, CAT_D2, CAT_E, CAT_E2, CAT_F, CAT_G)

CAT_LABELS = {
    CAT_A:  "A   Orphan profile (no instances + no matching parent)",
    CAT_B:  "B   Missing parent (no id or name match)",
    CAT_C:  "C   Far from parent",
    CAT_D:  "D   Child ID discrepancy",
    CAT_D2: "D2  Parent ID discrepancy (id stale, name-match exists)",
    CAT_E:  "E   Parent type change (no profile matches new family)",
    CAT_E2: "E2  Parent type change (another profile matches new family)",
    CAT_F:  "F   Host-name mismatch",
    CAT_G:  "G   Missing children",
}


# Trailing ``" : Default"`` or ``" : Default 2"`` etc. — Revit's
# default-type decoration that should be ignored when comparing host_name
# to a parent's family name.
_DEFAULT_DECORATION_RE = re.compile(
    r"\s*:\s*Default(?:\s+\d+)?\s*$", re.IGNORECASE,
)


def _strip_default_decoration(name):
    if not name:
        return ""
    return _DEFAULT_DECORATION_RE.sub("", str(name).strip())


def _profile_known_family_names(profile):
    """Lower-cased set of family names the profile recognises — its own
    ``parent_filter.family_name_pattern`` plus the family-half of every
    ``merged_aliases`` entry. Used by the cat-E check so a known alias
    isn't reported as a parent type change.
    """
    if not isinstance(profile, dict):
        return set()
    out = set()
    pf = profile.get("parent_filter") or {}
    fam = (pf.get("family_name_pattern") or "").strip().lower()
    if fam:
        out.add(fam)
    for alias in profile.get("merged_aliases") or []:
        if not isinstance(alias, str):
            continue
        text = alias.strip()
        if " : " in text:
            text = text.split(" : ", 1)[0]
        text = text.strip().lower()
        if text:
            out.add(text)
    return out


def _alias_strip_type(alias_text):
    """Lower-case the alias and drop any trailing ``" : Type"``
    decoration so a stored alias like ``"FAM : Default"`` matches
    a profile name / family of just ``"FAM"``."""
    if not isinstance(alias_text, str):
        return ""
    text = alias_text.strip().lower()
    if " : " in text:
        text = text.split(" : ", 1)[0].strip()
    return text


def _profile_merge_group_family_names(profile, profile_data):
    """All family names recognised by the merge group containing
    ``profile`` — own + every sibling that shares a merge link.

    Cat E / E2 use this instead of ``_profile_known_family_names`` so a
    placement against ANY profile in the same merge group passes the
    "known family" check, regardless of which side of the merge owns
    the alias list. The bare ``_profile_known_family_names`` only
    knows about aliases the profile itself carries, which works when
    the profile is the merge master (its aliases enumerate the
    absorbed siblings) but fails when the profile is a member (the
    aliases live on the master, not here).

    Two profiles are in the same merge group when any of:
      A. One lists the other (by name / id / family, with optional
         ``" : Type"`` decoration on the alias) in its
         ``merged_aliases`` list — checked in BOTH directions so we
         catch the asymmetric "master holds the aliases" case as
         well as the inverse.
      B. Legacy ``ced_truth_source_id`` chain — same truth id on
         both sides, ``other`` points at ``profile`` as its source,
         or ``profile`` points at ``other`` as its source. All three
         arms are checked unconditionally so the source-side profile
         (which has no ``ced_truth_source_id`` itself) still finds
         its members.
    """
    if not isinstance(profile, dict):
        return set()
    own_name = (profile.get("name") or "").strip().lower()
    own_id = (profile.get("id") or "").strip()
    own_id_lower = own_id.lower()
    own_family = _profile_family_name(profile).strip().lower()
    own_keys = set(filter(None, [own_name, own_id_lower, own_family]))
    own_truth_id = (
        truth_groups.truth_source_id(profile)
        if truth_groups.is_group_member(profile) else None
    )

    out = _profile_known_family_names(profile)
    # Include this profile's own NAME (family-portion only) as a
    # known family alias. Merge groups in this project commonly use
    # the profile NAME to record alternate Revit family names that
    # share a single ``parent_filter.family_name_pattern`` — e.g. a
    # source profile named ``"000_Stinger_Cart_1032114 : Default"``
    # has members named ``"000_Stinger_Cart_1"`` / ``"_2"``, all
    # expecting parent_filter family ``"000_Stinger_Cart_1032114"``.
    # A placement against family ``"_1"`` or ``"_2"`` is intended to
    # be valid; adding the profile name's family-portion makes the
    # check honour that.
    own_name_fam = own_name.split(" : ", 1)[0].strip() if own_name else ""
    if own_name_fam:
        out.add(own_name_fam)

    # Precompute this profile's own alias key-set so the reverse
    # direction of Direction A doesn't re-strip on every iteration.
    own_alias_keys = set()
    for alias in profile.get("merged_aliases") or []:
        key = _alias_strip_type(alias)
        if key:
            own_alias_keys.add(key)

    for other in profile_data.get("equipment_definitions") or []:
        if not isinstance(other, dict) or other is profile:
            continue

        sibling = False
        other_id = (other.get("id") or "").strip()
        other_name = (other.get("name") or "").strip().lower()
        other_family = _profile_family_name(other).strip().lower()
        other_keys = set(filter(None, [
            other_name, other_id.lower(), other_family,
        ]))
        other_truth_id = (
            truth_groups.truth_source_id(other)
            if truth_groups.is_group_member(other) else None
        )

        # Direction A1: ``other`` lists this profile in its aliases.
        if own_keys:
            for alias in other.get("merged_aliases") or []:
                key = _alias_strip_type(alias)
                if key and key in own_keys:
                    sibling = True
                    break

        # Direction A2: ``profile`` lists ``other`` in its aliases —
        # the inverse of A1, covers the case where the alias list
        # lives on the side currently being checked.
        if not sibling and own_alias_keys and other_keys:
            if own_alias_keys & other_keys:
                sibling = True

        # Direction B: legacy ``ced_truth_source_id`` chain. Checked
        # unconditionally — works for both source-side and member-
        # side profiles, both directions of the pointer.
        if not sibling:
            if own_truth_id and other_truth_id == own_truth_id:
                sibling = True
            elif own_truth_id and other_id == own_truth_id:
                sibling = True
            elif own_id and other_truth_id == own_id:
                sibling = True

        if sibling:
            out |= _profile_known_family_names(other)
            # Sibling's NAME (family-portion) is also a valid family
            # alias for the merge group — same rationale as the
            # own-name handling above. Critical for groups where the
            # MEMBER profiles' NAMES enumerate the actual Revit
            # family aliases (Stinger_Cart_1 / _2 etc.) while every
            # member shares one parent_filter pointing at the
            # canonical family name.
            other_name_fam = (
                other_name.split(" : ", 1)[0].strip()
                if other_name else ""
            )
            if other_name_fam:
                out.add(other_name_fam)

    return out


def _is_merged_member(profile, profile_data):
    """True if the profile is conceptually folded into another profile.

    Two cases:
      * Legacy: carries ``ced_truth_source_id`` (legacy merge model).
      * Current: this profile's ``name`` appears as an entry in any
        other profile's ``merged_aliases`` list.

    Used by the cat-A check so genuinely-merged profiles aren't flagged
    as orphans just because their data lives under a master.
    """
    if not isinstance(profile, dict):
        return False
    if truth_groups.is_group_member(profile):
        # ced_truth_source_id pointing at a different profile id.
        sid = truth_groups.truth_source_id(profile)
        if sid and sid != (profile.get("id") or ""):
            return True
    name = (profile.get("name") or "").strip().lower()
    if not name:
        return False
    for other in profile_data.get("equipment_definitions") or []:
        if not isinstance(other, dict) or other is profile:
            continue
        for alias in other.get("merged_aliases") or []:
            if isinstance(alias, str) and alias.strip().lower() == name:
                return True
    return False


# Fix-kind dispatch keys.
FIX_NONE = "none"
FIX_FOLLOW_PARENT = "follow_parent"           # cat C (move into alignment)
FIX_REFRESH_ELEMENT_ID = "refresh_eid"        # cat D (refresh child's own id)
FIX_REFRESH_PARENT_ID = "refresh_parent_id"   # cat D2 (refresh stale parent id)
FIX_REFRESH_HOST_NAME = "refresh_host"        # cat F (refresh host_name)
FIX_CLEAR_LINKER = "clear_linker"             # cat B (destructive — child becomes unstamped)


# ---------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------

class QaqcFinding(object):
    __slots__ = (
        "category",
        "category_label",
        "element_id",          # int — the host-doc element to select/zoom (None for cat A).
                               # For linked-parent findings (Cat G), this is the
                               # RevitLinkInstance's id so legacy callers still
                               # have a valid host id to fall back on.
        "link_instance_id",    # int — id of the RevitLinkInstance in the host doc
                               # when the finding's target element lives in a
                               # linked doc; None for host-doc findings.
        "linked_element_id",   # int — id of the element within the linked doc;
                               # paired with ``link_instance_id`` so Select / Zoom
                               # can build a host-coord Reference and highlight
                               # the specific linked element instead of the
                               # whole link.
        "profile_id",
        "profile_name",
        "led_id",
        "led_label",
        "message",
        "fix_kind",
        "fix_payload",         # dict; varies by fix_kind
    )

    def __init__(self, category, element_id, profile_id, profile_name,
                 led_id, led_label, message, fix_kind=FIX_NONE,
                 fix_payload=None, link_instance_id=None,
                 linked_element_id=None):
        self.category = category
        self.category_label = CAT_LABELS.get(category, category)
        self.element_id = element_id
        self.link_instance_id = link_instance_id
        self.linked_element_id = linked_element_id
        self.profile_id = profile_id
        self.profile_name = profile_name
        self.led_id = led_id
        self.led_label = led_label
        self.message = message
        self.fix_kind = fix_kind
        self.fix_payload = fix_payload or {}


class QaqcResult(object):
    def __init__(self):
        self.findings = []
        self.counts = {c: 0 for c in CAT_ALL}

    def add(self, finding):
        self.findings.append(finding)
        self.counts[finding.category] = self.counts.get(finding.category, 0) + 1


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _id_value(elem_or_id):
    if elem_or_id is None:
        return None
    eid = getattr(elem_or_id, "Id", None) or elem_or_id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _location_pt_rot(elem):
    loc = getattr(elem, "Location", None)
    if not isinstance(loc, LocationPoint):
        return None, 0.0
    try:
        pt = loc.Point
    except Exception:
        return None, 0.0
    try:
        rad = loc.Rotation
    except Exception:
        rad = 0.0
    return (pt.X, pt.Y, pt.Z), geometry.normalize_angle(math.degrees(rad))


def _build_led_index(profile_data):
    out = {}
    for profile in profile_data.get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if isinstance(led, dict) and led.get("id"):
                    out[led["id"]] = (profile, set_dict, led)
    return out


def _profile_family_name(profile):
    pf = profile.get("parent_filter") or {}
    return (pf.get("family_name_pattern") or "").strip()


def _element_family_name(elem):
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        if sym is not None:
            family = getattr(sym, "Family", None)
            return family.Name if family else ""
    if isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        return gtype.Name if gtype else ""
    return ""


def _element_family_type_label(elem):
    """``"Family : Type"`` for FamilyInstance, group-type name for
    Group, "" otherwise. Mirrors ``capture.element_label`` so the
    string round-trips against ``Element_Linker.host_name`` for
    side-by-side comparison in the Cat F check.
    """
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        if sym is not None:
            family = getattr(sym, "Family", None)
            fam_name = family.Name if family else ""
            type_name = getattr(sym, "Name", "") or ""
            label = "{} : {}".format(fam_name, type_name).strip(" :")
            return label
    if isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        return gtype.Name if gtype else ""
    return ""


# Cat G scope filter — only profiles whose name starts with a
# 3-digit number prefix OR contains "HEB" / "vendor provided" are
# treated as ones the user actively tracks for missing-children
# coverage. Project-specific placeholders and generic profiles fall
# outside that net and would just be noise in the Cat G list.
#
# The prefix regex deliberately omits ``\b``: ``_`` is a word
# character in Python regex, so a boundary after the 3 digits
# fails for very common names like ``"000_Stealth_Bug_Light_..."``
# where the digits are followed by an underscore. ``\d{3,}`` at
# the start is sufficient — anything beginning with three or more
# digits in any context counts as a tracked profile.
_CAT_G_THREE_DIGIT_PREFIX = re.compile(r"^\s*\d{3,}")


def _profile_name_in_cat_g_scope(name):
    """True iff ``name`` looks like a Cat-G-tracked profile.

    Match conditions (any one):
      * Contains ``HEB`` (case-insensitive).
      * Contains ``vendor provided`` (case-insensitive).
      * Starts with three or more digits (e.g. ``"000_Stealth..."``,
        ``"305 Steam Kettle"``, ``"007 - Walk-in Cooler"``).
    """
    if not name:
        return False
    text = str(name)
    lower = text.lower()
    if "heb" in lower:
        return True
    if "vendor provided" in lower:
        return True
    if _CAT_G_THREE_DIGIT_PREFIX.match(text):
        return True
    return False


def _profile_has_leds(profile):
    """True iff the profile declares at least one LED (linked-element
    definition) anywhere in its ``linked_sets``. Profiles with no LEDs
    can never have placed children, so we don't flag them under Cat G —
    that'd be every empty-profile placeholder in the YAML.
    """
    if not isinstance(profile, dict):
        return False
    for s in profile.get("linked_sets") or []:
        if not isinstance(s, dict):
            continue
        for led in s.get("linked_element_definitions") or []:
            if isinstance(led, dict) and led.get("id"):
                return True
    return False


def _profile_is_parented_in_practice(profile_id, profile_to_parent_ids):
    """True iff the profile has at least one placed child whose
    Element_Linker carries a non-null ``parent_element_id``.

    The YAML schema's ``allow_parentless`` flag turned out to be
    unreliable in real-world data (the HEB profile set has ``true``
    on every entry regardless of whether the profile is actually
    placed with a parent), so we infer parentedness from placement
    behavior instead. If a profile has ever been placed against a
    parent, we treat it as "needs a parent" for Cat G; truly
    parentless profiles (receptacle-only, place-from-CSV, etc.) have
    placed children with null parent_element_id and won't surface
    here.
    """
    return bool(profile_to_parent_ids.get(profile_id))


def _profiles_matching_parent_family(family_name, profiles):
    """Return profiles whose ``parent_filter`` / ``merged_aliases`` /
    own name matches ``family_name``. Same two-tier rule the placement
    engine uses (``placement._match_one_linked_revit``): exact-name
    match wins over the suffix-stripped fallback so the
    ``Stinger Cart_1 / _2 / _3`` style aligns 1:1 with its profiles
    instead of producing a cross-product.
    """
    name = (family_name or "").strip()
    if not name:
        return []
    name_lower = name.lower()
    name_norm = _placement.normalize_name(name)
    if not name_norm:
        return []
    strict = [
        p for p in profiles
        if name_lower in _placement.profile_family_names_raw(p)
    ]
    if strict:
        return strict
    return [
        p for p in profiles
        if name_norm in _placement.profile_family_names(p)
    ]


def _iter_potential_parents(doc):
    """Yield every candidate parent in the host doc plus every linked
    doc as ``(elem, family_name, elem_id_int, link_inst, in_host)``.

    Host iteration skips elements that themselves carry an
    ``Element_Linker`` — those are placed children of some other
    profile and aren't parents in their own right. Linked iteration
    has no such filter (we can't read Element_Linker out of linked
    docs reliably) so every FamilyInstance / Group in a link is a
    candidate.

    For linked yields, ``link_inst`` is the host-doc RevitLinkInstance
    so callers can build a host-coord Reference back to the linked
    element via ``Reference.CreateLinkReference``.
    """
    for klass in (FamilyInstance, Group):
        try:
            coll = (
                FilteredElementCollector(doc)
                .OfClass(klass)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for elem in coll:
            try:
                if _el_io.read_from_element(elem) is not None:
                    continue
            except Exception:
                pass
            yield (
                elem,
                _element_family_name(elem),
                _id_value(elem),
                None,
                True,
            )
    try:
        link_collector = FilteredElementCollector(doc).OfClass(
            RevitLinkInstance
        )
    except Exception:
        link_collector = []
    for link_inst in link_collector:
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        for klass in (FamilyInstance, Group):
            try:
                coll = (
                    FilteredElementCollector(link_doc)
                    .OfClass(klass)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue
            for elem in coll:
                yield (
                    elem,
                    _element_family_name(elem),
                    _id_value(elem),
                    link_inst,
                    False,
                )


def _build_parent_index(doc, all_profiles):
    """Pre-build the data structures needed by Cats A, D2, E2, and G:

      * ``parent_records``: list of every candidate parent
        ``(elem, family_name, elem_id_int, link_inst, in_host)`` from
        host + linked docs.
      * ``parents_by_family_norm``: normalized family name ->
        list of parent_records. Used by Cat D2 to find a salvage
        candidate when the stored ``parent_element_id`` is stale.
      * ``matched_profiles_by_family``: raw family name -> list of
        profiles whose ``parent_filter`` / aliases match (placement-
        engine rule). Cached per unique family name so Cats E2 and G
        don't recompute the matching profile list per parent.
      * ``profile_to_parents``: profile_id -> list of parent_records
        that match this profile. Used by Cat A to distinguish
        "irrelevant profile" (no matching parent) from "matching
        parent but no children placed" (Cat G's job).
    """
    parent_records = list(_iter_potential_parents(doc))

    parents_by_family_norm = {}
    for rec in parent_records:
        fam = rec[1]
        if not fam:
            continue
        norm = _placement.normalize_name(_strip_default_decoration(fam))
        if norm:
            parents_by_family_norm.setdefault(norm, []).append(rec)

    matched_profiles_by_family = {}
    for rec in parent_records:
        fam = rec[1]
        if not fam or fam in matched_profiles_by_family:
            continue
        matched_profiles_by_family[fam] = _profiles_matching_parent_family(
            fam, all_profiles,
        )

    profile_to_parents = {}
    for rec in parent_records:
        for p in matched_profiles_by_family.get(rec[1]) or []:
            pid = p.get("id") or ""
            if pid:
                profile_to_parents.setdefault(pid, []).append(rec)

    return (
        parent_records,
        parents_by_family_norm,
        matched_profiles_by_family,
        profile_to_parents,
    )


# ---------------------------------------------------------------------
# Audit
# ---------------------------------------------------------------------

def run_audit(doc, profile_data, categories=None):
    """Walk the doc + profile_data and emit a ``QaqcResult``.

    ``categories`` can restrict the run to a subset (set of CAT_*); None
    means all nine.

    Phase order:
      1. Pre-build the parent index (host + linked) so every downstream
         check can ask "does a parent matching family X exist?" without
         re-walking the doc.
      2. Sweep every placed child, emitting per-child cats (C, D, D2, E,
         E2, F, B).
      3. Cat A — orphan profile rollup using profile_usage +
         profile_to_parents from the pre-pass.
      4. Cat G — per-parent missing-children walk, using the prebuilt
         parent_records and matched_profiles_by_family.
    """
    requested = set(categories) if categories else set(CAT_ALL)
    led_index = _build_led_index(profile_data)
    result = QaqcResult()

    profiles_by_id = {
        p.get("id"): p for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict) and p.get("id")
    }
    all_profiles = [
        p for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict)
    ]

    # ------------------------------------------------------------------
    # Phase 1: prebuilt parent index (host + linked).
    # Used by Cat A (rollup), Cat D2 (name-match salvage), Cat E2
    # (other-profile match), and Cat G (per-parent walk).
    # ------------------------------------------------------------------
    (
        parent_records,
        parents_by_family_norm,
        matched_profiles_by_family,
        profile_to_parents,
    ) = _build_parent_index(doc, all_profiles)

    # Track profile usage during the placed-child sweep so we can emit
    # Cat A at the end.
    profile_usage = {pid: 0 for pid in profiles_by_id}

    # parent_element_id -> {profile_id, profile_id, ...}
    # Records which profiles have at least one placed child for each
    # parent. Used by Cat G to detect parents that match a profile but
    # never had its LEDs fired.
    parent_to_placed_profile_ids = {}

    # profile_id -> {normalized_host_name_family, ...}
    # Secondary Cat G skip signal: when profiles get merged, children
    # placed before the merge keep the OLD parent's family name in
    # ``Element_Linker.host_name``. The new master profile's
    # ``merged_aliases`` lists that old name, so the parent (still
    # named the old name in the link) DOES correspond to placed
    # children — we just can't see it via parent_element_id alone if
    # the link was reloaded / the parent's id changed. Tracking the
    # host_name family lets Cat G recognise that coverage.
    placed_host_names_by_profile = {}

    # ------------------------------------------------------------------
    # Phase 2: placed-child sweep.
    # ------------------------------------------------------------------
    for klass in (FamilyInstance, Group):
        for elem in FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType():
            linker = _el_io.read_from_element(elem)
            if linker is None or not linker.led_id:
                continue
            entry = led_index.get(linker.led_id)
            if entry is None:
                continue
            profile, _set_dict, led = entry
            profile_id = profile.get("id") or ""
            profile_name = profile.get("name") or ""
            led_id = led.get("id") or ""
            led_label = led.get("label") or ""
            elem_id_val = _id_value(elem)

            profile_usage[profile_id] = profile_usage.get(profile_id, 0) + 1
            if linker.parent_element_id is not None:
                try:
                    pid_int = int(linker.parent_element_id)
                except (TypeError, ValueError):
                    pid_int = None
                if pid_int is not None:
                    parent_to_placed_profile_ids.setdefault(
                        pid_int, set()
                    ).add(profile_id)
            host_family = _strip_default_decoration(linker.host_name or "")
            if host_family:
                norm_family = _placement.normalize_name(host_family)
                if norm_family:
                    placed_host_names_by_profile.setdefault(
                        profile_id, set()
                    ).add(norm_family)

            # Cat D: child's own ID discrepancy.
            if CAT_D in requested:
                stored_eid = linker.element_id
                if stored_eid is not None and elem_id_val is not None:
                    try:
                        if int(stored_eid) != int(elem_id_val):
                            result.add(QaqcFinding(
                                category=CAT_D,
                                element_id=elem_id_val,
                                profile_id=profile_id,
                                profile_name=profile_name,
                                led_id=led_id,
                                led_label=led_label,
                                message="Stored element_id {} != actual {}".format(
                                    stored_eid, elem_id_val
                                ),
                                fix_kind=FIX_REFRESH_ELEMENT_ID,
                            ))
                    except (TypeError, ValueError):
                        pass

            # Parent-related checks (C, E, E2, F, B, D2) all need the
            # parent. Resolve host doc first, then fall back to links.
            parent = None
            if linker.parent_element_id is not None:
                try:
                    parent = doc.GetElement(ElementId(int(linker.parent_element_id)))
                except Exception:
                    parent = None

            if parent is None and linker.parent_element_id is not None:
                try:
                    eid = ElementId(int(linker.parent_element_id))
                except Exception:
                    eid = None
                if eid is not None:
                    for link_doc, _t in links.iter_link_documents(doc):
                        try:
                            cand = link_doc.GetElement(eid)
                        except Exception:
                            continue
                        if cand is not None:
                            parent = cand
                            break

            if parent is None:
                # B vs. D2 split. If the stored host_name name-matches a
                # parent that DOES exist somewhere in the model, the
                # linker is salvageable — emit D2 (fixable when exactly
                # one match) instead of the destructive Cat B. If no
                # name match, the child is genuinely orphan → Cat B.
                if linker.parent_element_id is not None:
                    host_norm = _placement.normalize_name(
                        _strip_default_decoration(linker.host_name or "")
                    )
                    name_matches = (
                        parents_by_family_norm.get(host_norm, [])
                        if host_norm else []
                    )
                    if name_matches and CAT_D2 in requested:
                        if len(name_matches) == 1:
                            target_rec = name_matches[0]
                            target_id = target_rec[2]
                            target_fam = target_rec[1]
                            result.add(QaqcFinding(
                                category=CAT_D2,
                                element_id=elem_id_val,
                                profile_id=profile_id,
                                profile_name=profile_name,
                                led_id=led_id,
                                led_label=led_label,
                                message=(
                                    "Stored parent_element_id={} is dead; "
                                    "exactly one '{}' in model has id {} "
                                    "— offering auto-refresh."
                                ).format(
                                    linker.parent_element_id,
                                    target_fam, target_id,
                                ),
                                fix_kind=FIX_REFRESH_PARENT_ID,
                                fix_payload={
                                    "new_parent_element_id": int(target_id),
                                    "new_host_name": target_fam,
                                },
                            ))
                        else:
                            result.add(QaqcFinding(
                                category=CAT_D2,
                                element_id=elem_id_val,
                                profile_id=profile_id,
                                profile_name=profile_name,
                                led_id=led_id,
                                led_label=led_label,
                                message=(
                                    "Stored parent_element_id={} is dead; "
                                    "{} candidates match host_name '{}' "
                                    "— manual fix needed."
                                ).format(
                                    linker.parent_element_id,
                                    len(name_matches),
                                    linker.host_name or "",
                                ),
                                fix_kind=FIX_NONE,
                            ))
                    elif CAT_B in requested:
                        result.add(QaqcFinding(
                            category=CAT_B,
                            element_id=elem_id_val,
                            profile_id=profile_id,
                            profile_name=profile_name,
                            led_id=led_id,
                            led_label=led_label,
                            message=(
                                "parent_element_id={} not found in host "
                                "or linked docs (no name match either)"
                            ).format(linker.parent_element_id),
                            fix_kind=FIX_CLEAR_LINKER,
                        ))
                continue

            # Cat E / E2 split: parent's family doesn't match profile's
            # filter AND isn't a known alias on the profile.
            #
            # ``_profile_merge_group_family_names`` widens the "known"
            # set to every family recognised by ANY profile in the
            # current profile's merge group, in either merge direction.
            # That suppresses both:
            #   * member-side noise: child's recorded profile was merged
            #     into a master, parent now wears the master's family.
            #   * master-side noise: child's profile is the master,
            #     parent wears an absorbed member's family.
            # Without this, every merged sibling pair would generate
            # E / E2 findings even though the merge intent already
            # says these families are interchangeable.
            #
            # All three sides of the comparison run through
            # ``_placement.normalize_name`` (lowercase + strip trailing
            # ``_NNN``), so the check ignores Revit's auto-incremented
            # copy suffix the same way the placement matcher does. A
            # profile expecting ``Foo_1207004`` and an actual parent
            # named ``Foo_1`` both collapse to ``foo`` and pass the
            # check — that's exactly the same rule that lets placement
            # land children on either family in the first place.
            if (CAT_E in requested) or (CAT_E2 in requested):
                expected_family = _profile_family_name(profile)
                actual_family = _element_family_name(parent)
                if expected_family and actual_family:
                    actual_norm = _placement.normalize_name(actual_family)
                    expected_norm = _placement.normalize_name(expected_family)
                    known = _profile_merge_group_family_names(
                        profile, profile_data,
                    )
                    known_norm = set(
                        _placement.normalize_name(k) for k in known if k
                    )
                    if actual_norm and actual_norm not in known_norm and \
                            expected_norm != actual_norm:
                        # Does any OTHER profile in the YAML match the
                        # new family? Pull from matched_profiles_by_family
                        # (already keyed by family name) and skip the
                        # original profile itself. If the family wasn't
                        # in the cache (i.e. the parent's family doesn't
                        # appear in the parent_records walk for some
                        # reason — extremely unlikely since the parent
                        # IS in the doc — compute on demand).
                        candidates = matched_profiles_by_family.get(actual_family)
                        if candidates is None:
                            candidates = _profiles_matching_parent_family(
                                actual_family, all_profiles,
                            )
                        other_matches = [
                            p for p in candidates
                            if (p.get("id") or "") != profile_id
                        ]
                        if other_matches:
                            if CAT_E2 in requested:
                                names = ", ".join(
                                    p.get("name") or "?"
                                    for p in other_matches[:3]
                                )
                                more = (
                                    " (+{} more)".format(len(other_matches) - 3)
                                    if len(other_matches) > 3 else ""
                                )
                                result.add(QaqcFinding(
                                    category=CAT_E2,
                                    element_id=elem_id_val,
                                    profile_id=profile_id,
                                    profile_name=profile_name,
                                    led_id=led_id,
                                    led_label=led_label,
                                    message=(
                                        "Parent family is '{}', profile "
                                        "expects '{}'. Reassign to: {}{}"
                                    ).format(
                                        actual_family, expected_family,
                                        names, more,
                                    ),
                                    fix_kind=FIX_NONE,
                                ))
                        else:
                            if CAT_E in requested:
                                result.add(QaqcFinding(
                                    category=CAT_E,
                                    element_id=elem_id_val,
                                    profile_id=profile_id,
                                    profile_name=profile_name,
                                    led_id=led_id,
                                    led_label=led_label,
                                    message=(
                                        "Parent family is '{}', profile "
                                        "expects '{}'. No profile in "
                                        "store matches '{}'."
                                    ).format(
                                        actual_family, expected_family,
                                        actual_family,
                                    ),
                                    fix_kind=FIX_NONE,
                                ))

            # Cat F: host_name mismatch — compare the FULL
            # ``"Family : Type"`` label on each side, ignoring the
            # trailing `` : Default`` / `` : Default <n>`` decoration
            # Revit auto-generates when a family has only its default
            # type. Both sides are normalised through
            # ``_strip_default_decoration`` and then compared.
            #
            # Legacy fallback: in older captures the stored
            # ``host_name`` is just the family (no ``" : Type"``
            # suffix). Comparing that to a full ``Family : Type`` label
            # would false-flag every non-default-type parent. When
            # EITHER side lacks the type suffix, we downgrade the
            # comparison to family-only so legacy data stays clean —
            # the family-mismatch check still works, but we don't
            # synthesize a missing type on the partial side.
            if CAT_F in requested:
                stored_host = (linker.host_name or "").strip()
                actual_label = _element_family_type_label(parent)
                if stored_host and actual_label:
                    stored_clean = _strip_default_decoration(stored_host)
                    actual_clean = _strip_default_decoration(actual_label)
                    stored_has_type = " : " in stored_clean
                    actual_has_type = " : " in actual_clean
                    if stored_has_type and actual_has_type:
                        # Both full Family : Type — compare full.
                        stored_norm = stored_clean.lower()
                        actual_norm = actual_clean.lower()
                        mismatch = (
                            stored_norm and actual_norm
                            and stored_norm != actual_norm
                        )
                    else:
                        # At least one side is family-only — compare
                        # only the family portion so legacy / partial
                        # host_name data doesn't generate false flags.
                        stored_fam = stored_clean.split(" : ", 1)[0].strip().lower()
                        actual_fam = actual_clean.split(" : ", 1)[0].strip().lower()
                        mismatch = (
                            stored_fam and actual_fam
                            and stored_fam != actual_fam
                        )
                    if mismatch:
                        result.add(QaqcFinding(
                            category=CAT_F,
                            element_id=elem_id_val,
                            profile_id=profile_id,
                            profile_name=profile_name,
                            led_id=led_id,
                            led_label=led_label,
                            message="host_name='{}', actual parent='{}'".format(
                                stored_host, actual_label
                            ),
                            fix_kind=FIX_REFRESH_HOST_NAME,
                            fix_payload={"new_host_name": actual_label},
                        ))

            # Cat C: far from parent.
            if CAT_C in requested:
                parent_pt, parent_rot = _fp.find_parent_pose(
                    doc, linker.parent_element_id,
                    host_name=linker.host_name,
                )
                cur_pt, _cur_rot = _location_pt_rot(elem)
                if parent_pt is not None and cur_pt is not None:
                    offsets_list = led.get("offsets") or []
                    offset = offsets_list[0] if offsets_list else {
                        "x_inches": 0.0, "y_inches": 0.0,
                        "z_inches": 0.0, "rotation_deg": 0.0,
                    }
                    target_pt = geometry.target_point_from_offsets(
                        parent_pt, parent_rot, offset
                    )
                    dx = cur_pt[0] - target_pt[0]
                    dy = cur_pt[1] - target_pt[1]
                    dz = cur_pt[2] - target_pt[2]
                    dist = math.sqrt(dx * dx + dy * dy + dz * dz)
                    if dist > geometry.Tolerances.FAR_FROM_PARENT_FT:
                        result.add(QaqcFinding(
                            category=CAT_C,
                            element_id=elem_id_val,
                            profile_id=profile_id,
                            profile_name=profile_name,
                            led_id=led_id,
                            led_label=led_label,
                            message="{:.2f} ft from expected pose (tol {:.1f} ft)".format(
                                dist, geometry.Tolerances.FAR_FROM_PARENT_FT
                            ),
                            fix_kind=FIX_FOLLOW_PARENT,
                        ))

    # ------------------------------------------------------------------
    # Phase 3: Cat A — truly orphan profile.
    # Fires only when the profile has zero placed instances AND no
    # parent in the model matches its family (per the prebuilt index).
    # The "matching parent exists but no children placed" case is
    # intentionally handled by Cat G (per-parent), so it doesn't
    # double-report here. Skips merged-member profiles (legacy
    # ced_truth_source_id or alias-target) since those are unused on
    # purpose. Mirrors 1.0 Tab 1 ("No Matching Parents").
    # ------------------------------------------------------------------
    if CAT_A in requested:
        for pid, p in profiles_by_id.items():
            if profile_usage.get(pid, 0) > 0:
                continue
            if _is_merged_member(p, profile_data):
                continue
            if profile_to_parents.get(pid):
                # Matching parent exists; Cat G surfaces it per-parent.
                continue
            result.add(QaqcFinding(
                category=CAT_A,
                element_id=None,
                profile_id=pid,
                profile_name=p.get("name") or "",
                led_id="",
                led_label="",
                message=(
                    "Profile has no placed instances and no matching "
                    "parent in host or linked docs."
                ),
                fix_kind=FIX_NONE,
            ))

    # ------------------------------------------------------------------
    # Phase 4: Cat G — per-parent missing children.
    # Walk the prebuilt parent_records (no second doc traversal), find
    # profiles that match each parent's family, and emit when a
    # matching profile has LEDs but no placed child carries those LED
    # IDs back to this parent. Restricts to profiles that aren't merged
    # members, have at least one LED, and have been placed against a
    # parent at least once (the YAML ``allow_parentless`` flag is
    # unreliable in real data, so placement behavior is the source of
    # truth here).
    # ------------------------------------------------------------------
    if CAT_G in requested:
        # Cat G candidacy: not a merged member, and has at least one LED
        # to place. Note: we deliberately do NOT require the profile to
        # have been placed against a parent before (`_profile_is_parented_in_practice`
        # was the old gate). That heuristic was meant to skip parentless
        # profiles (receptacle-only / place-from-CSV) but also silently
        # excluded brand-new or empty-after-delete profiles that have
        # parent_filter set — so a parent matching such a profile would
        # never surface in Cat G. ``_profiles_matching_parent_family``
        # already requires the profile to expose family names (which
        # come from parent_filter / aliases / own name), so a truly
        # parentless profile with no parent_filter won't appear in
        # ``matches`` anyway — making the explicit gate redundant.
        active_profile_ids = set(
            (p.get("id") or "") for p in all_profiles
            if (p.get("id") or "")
               and not _is_merged_member(p, profile_data)
               and _profile_has_leds(p)
               and _profile_name_in_cat_g_scope(p.get("name") or "")
        )

        # Dedupe so the same (parent_id, profile_id, scope) triple
        # fires once even if multiple aliasing paths matched it.
        emitted = set()
        for parent_elem, family_name, parent_id_int, link_inst, in_host in parent_records:
            if not family_name or parent_id_int is None:
                continue
            matches = matched_profiles_by_family.get(family_name) or []
            if not matches:
                continue
            placed_pids = parent_to_placed_profile_ids.get(parent_id_int, set())
            # Parent-level coverage check. If THIS parent already has at
            # least one placed child (regardless of which profile that
            # child's led_id resolves to), treat the parent as covered
            # and skip Cat G for all profiles matching this family.
            # Catches the case where a child's stored led_id maps to a
            # different-but-related profile (merged variant, re-keyed
            # YAML, sibling profile sharing the same family) — without
            # this, Cat G falsely flags a parent that clearly has a
            # child sitting on it.
            #
            # The previous secondary suppression on
            # ``placed_host_names_by_profile`` was deliberately removed:
            # it normalized the host_name family with the trailing
            # ``_NNN`` strip, so a single ``Foo_1207004`` placement
            # silently covered every unrelated ``Foo_1``, ``Foo_2``
            # parent that shared the same suffix-stripped key. The
            # link-reload-with-id-drift case it was meant to catch is
            # already surfaced by Cat D2 (parent_element_id stale +
            # name-match parent exists, with a refresh fix); once the
            # user runs that D2 fix, ``parent_to_placed_profile_ids``
            # picks up the corrected id and Cat G stops firing on the
            # next refresh.
            if placed_pids:
                continue
            for profile in matches:
                pid = profile.get("id") or ""
                if not pid or pid not in active_profile_ids:
                    continue
                key = (parent_id_int, pid, in_host)
                if key in emitted:
                    continue
                emitted.add(key)
                # Wiring for the row's Select / Zoom buttons:
                #   * Host parent — element_id is the parent's host-doc id.
                #     link_instance_id / linked_element_id are None.
                #   * Linked parent — element_id is the RevitLinkInstance's
                #     id (host-doc fallback for any legacy code path that
                #     just dispatches on element_id). The window's Select /
                #     Zoom prefers (link_instance_id, linked_element_id)
                #     when both are set, so the linked element itself gets
                #     highlighted via ``Reference.CreateLinkReference``
                #     instead of just the whole link being selected.
                where = "host" if in_host else "linked"
                if in_host:
                    row_elem_id = parent_id_int
                    fnd_link_inst_id = None
                    fnd_linked_elem_id = None
                else:
                    fnd_link_inst_id = _id_value(link_inst)
                    fnd_linked_elem_id = parent_id_int
                    row_elem_id = fnd_link_inst_id
                result.add(QaqcFinding(
                    category=CAT_G,
                    element_id=row_elem_id,
                    link_instance_id=fnd_link_inst_id,
                    linked_element_id=fnd_linked_elem_id,
                    profile_id=pid,
                    profile_name=profile.get("name") or "",
                    led_id="",
                    led_label="",
                    message=(
                        "Parent {!r} (id {}, {}) matches profile but no "
                        "child carries any of its LED IDs.".format(
                            family_name, parent_id_int, where,
                        )
                    ),
                    fix_kind=FIX_NONE,
                ))

    return result


# ---------------------------------------------------------------------
# Fixes
# ---------------------------------------------------------------------

def execute_fix(doc, profile_data, finding):
    """Apply the auto-fix for one finding. Caller manages the transaction.

    Returns ``(ok: bool, message: str)``.
    """
    kind = finding.fix_kind
    if kind == FIX_NONE:
        return False, "No automated fix for this category."

    if kind == FIX_REFRESH_ELEMENT_ID:
        elem = _get_host_element(doc, finding.element_id)
        if elem is None:
            return False, "Element {} not found.".format(finding.element_id)
        linker = _el_io.read_from_element(elem)
        if linker is None:
            return False, "Element_Linker payload missing."
        new_linker = _replace(linker, element_id=_id_value(elem))
        try:
            _el_io.write_to_element(elem, new_linker)
        except _el_io.ElementLinkerIOError as exc:
            return False, str(exc)
        return True, "Refreshed element_id."

    if kind == FIX_REFRESH_HOST_NAME:
        elem = _get_host_element(doc, finding.element_id)
        if elem is None:
            return False, "Element {} not found.".format(finding.element_id)
        linker = _el_io.read_from_element(elem)
        if linker is None:
            return False, "Element_Linker payload missing."
        new_host = (finding.fix_payload or {}).get("new_host_name") or ""
        new_linker = _replace(linker, host_name=new_host)
        try:
            _el_io.write_to_element(elem, new_linker)
        except _el_io.ElementLinkerIOError as exc:
            return False, str(exc)
        return True, "Refreshed host_name to '{}'.".format(new_host)

    if kind == FIX_REFRESH_PARENT_ID:
        # Cat D2 — child's stored parent_element_id is dead, but the
        # name-match index pinpointed a unique salvage candidate. Stamp
        # the new id (and refresh host_name to the matched parent's
        # family while we're at it, so Cat F doesn't fire on the next
        # refresh).
        elem = _get_host_element(doc, finding.element_id)
        if elem is None:
            return False, "Element {} not found.".format(finding.element_id)
        linker = _el_io.read_from_element(elem)
        if linker is None:
            return False, "Element_Linker payload missing."
        payload = finding.fix_payload or {}
        new_parent_id = payload.get("new_parent_element_id")
        if new_parent_id is None:
            return False, "Missing new_parent_element_id in fix payload."
        try:
            new_parent_id_int = int(new_parent_id)
        except (TypeError, ValueError):
            return False, "Invalid new_parent_element_id: {!r}".format(new_parent_id)
        new_host = payload.get("new_host_name") or linker.host_name
        new_linker = _replace(
            linker,
            parent_element_id=new_parent_id_int,
            host_name=new_host,
        )
        try:
            _el_io.write_to_element(elem, new_linker)
        except _el_io.ElementLinkerIOError as exc:
            return False, str(exc)
        return True, "Refreshed parent_element_id to {}.".format(new_parent_id_int)

    if kind == FIX_CLEAR_LINKER:
        elem = _get_host_element(doc, finding.element_id)
        if elem is None:
            return False, "Element {} not found.".format(finding.element_id)
        if not _el_io.clear_on_element(elem):
            return False, "Could not clear Element_Linker (read-only?)."
        return True, "Cleared Element_Linker (child is now unstamped)."

    if kind == FIX_FOLLOW_PARENT:
        elem = _get_host_element(doc, finding.element_id)
        if elem is None:
            return False, "Element {} not found.".format(finding.element_id)
        linker = _el_io.read_from_element(elem)
        if linker is None:
            return False, "Element_Linker payload missing."
        led_index = _build_led_index(profile_data)
        entry = led_index.get(linker.led_id)
        if entry is None:
            return False, "LED {} no longer in profile data.".format(linker.led_id)
        profile, _set, led = entry
        offsets_list = led.get("offsets") or []
        offset = offsets_list[0] if offsets_list else {
            "x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0,
        }
        parent_pt, parent_rot = _fp.find_parent_pose(
            doc, linker.parent_element_id, host_name=linker.host_name
        )
        if parent_pt is None:
            return False, "Parent pose unavailable."
        target_pt = geometry.target_point_from_offsets(parent_pt, parent_rot, offset)
        target_rot = geometry.child_rotation_from_offsets(parent_rot, offset)
        cur_pt, cur_rot = _location_pt_rot(elem)
        if cur_pt is None:
            return False, "Element has no usable LocationPoint."
        try:
            from Autodesk.Revit.DB import (
                ElementTransformUtils, Line, XYZ as _XYZ,
            )
            cur_xyz = _XYZ(*cur_pt)
            tgt_xyz = _XYZ(*target_pt)
            translation = tgt_xyz - cur_xyz
            if translation.GetLength() > geometry.Tolerances.POSITION_FT:
                ElementTransformUtils.MoveElement(doc, elem.Id, translation)
            rot_delta = geometry.normalize_angle(target_rot - cur_rot)
            if abs(rot_delta) > geometry.Tolerances.ROTATION_DEG:
                axis = Line.CreateBound(
                    tgt_xyz, _XYZ(tgt_xyz.X, tgt_xyz.Y, tgt_xyz.Z + 1.0)
                )
                ElementTransformUtils.RotateElement(
                    doc, elem.Id, axis, math.radians(rot_delta)
                )
        except Exception as exc:
            return False, "Move failed: {}".format(exc)
        new_linker = _replace(
            linker,
            location_ft=list(target_pt),
            rotation_deg=target_rot,
        )
        try:
            _el_io.write_to_element(elem, new_linker)
        except _el_io.ElementLinkerIOError:
            pass
        return True, "Followed parent (moved into alignment)."

    return False, "Unknown fix_kind: {}".format(kind)


def _replace(linker, **changes):
    """Return a new ElementLinker with the listed fields overwritten."""
    fields = {
        "led_id": linker.led_id,
        "set_id": linker.set_id,
        "location_ft": linker.location_ft,
        "rotation_deg": linker.rotation_deg,
        "parent_rotation_deg": linker.parent_rotation_deg,
        "parent_element_id": linker.parent_element_id,
        "level_id": linker.level_id,
        "element_id": linker.element_id,
        "facing": linker.facing,
        "host_name": linker.host_name,
        "parent_location_ft": linker.parent_location_ft,
    }
    fields.update(changes)
    return _el.ElementLinker(**fields)


def _get_host_element(doc, element_id):
    if element_id is None:
        return None
    try:
        return doc.GetElement(ElementId(int(element_id)))
    except Exception:
        return None
