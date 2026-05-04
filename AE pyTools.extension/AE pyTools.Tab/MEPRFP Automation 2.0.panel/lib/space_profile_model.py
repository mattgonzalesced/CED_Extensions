# -*- coding: utf-8 -*-
"""
Typed wrappers for ``space_profiles[*]`` YAML entries.

A space profile is the spaces-side analogue of an
``equipment_definitions[*]``. Its hierarchy mirrors the equipment shape
(``profile -> linked_sets -> leds -> annotations``) so the existing
``Tag`` / ``Annotation`` / ``Offset`` wrappers in ``profile_model`` can
be reused verbatim. The two structural differences are:

  1. **Bucket reference, not parent filter.** Each profile carries a
     ``bucket_id`` pointing at a ``space_buckets[*]``. The placement
     engine looks up classifications on each Space and unions every
     profile whose ``bucket_id`` is in that Space's assigned-bucket
     list (the "stacked templates" semantic — multiple profiles per
     bucket, all applied).

  2. **Per-LED placement rule.** Equipment LEDs anchor to a parent
     fixture; space LEDs anchor to the Space itself. The new
     ``placement_rule`` dict on each LED chooses the anchor kind
     (center / N / S / E / W / corners / door_relative) and the inset
     from the wall surface. Fine-tuning afterwards uses the same
     ``offsets[*]`` list LEDs already have.

Schema (one ``space_profiles[*]`` entry)::

    id: SP-001
    bucket_id: BUCKET-001
    name: "HEB Bakery Default"
    linked_sets:
      - id: SET-SP-001
        name: "Receptacles"
        linked_element_definitions:
          - id: SP-001-LED-001
            label: "EF-U_Receptacle_CED : Quad Wall"
            category: "Electrical Fixtures"
            placement_rule:
              kind: ne                   # see PLACEMENT_KINDS below
              inset_inches: 12           # distance off the wall surface
              door_offset_inches:        # only for kind=door_relative
                x: 12
                y: 0
            parameters: { ... }
            offsets:
              - { x_inches: 0, y_inches: 0, z_inches: 18, rotation_deg: 0 }
            annotations: [ ... ]
"""

from profile_model import (  # noqa: F401  -- re-export for callers
    _DictBacked,
    _str_or_none,
    _bool_or,
    _ensure_list,
    Offset,
    Tag,
    Annotation,
)


# ---------------------------------------------------------------------
# Placement-rule kinds
# ---------------------------------------------------------------------

KIND_CENTER = "center"
KIND_N = "n"
KIND_S = "s"
KIND_E = "e"
KIND_W = "w"
KIND_NE = "ne"
KIND_NW = "nw"
KIND_SE = "se"
KIND_SW = "sw"
KIND_DOOR_RELATIVE = "door_relative"

PLACEMENT_KINDS = (
    KIND_CENTER,
    KIND_N, KIND_S, KIND_E, KIND_W,
    KIND_NE, KIND_NW, KIND_SE, KIND_SW,
    KIND_DOOR_RELATIVE,
)

EDGE_KINDS = frozenset({KIND_N, KIND_S, KIND_E, KIND_W})
CORNER_KINDS = frozenset({KIND_NE, KIND_NW, KIND_SE, KIND_SW})


def is_valid_placement_kind(kind):
    return kind in PLACEMENT_KINDS


# ---------------------------------------------------------------------
# Placement rule
# ---------------------------------------------------------------------

class PlacementRule(_DictBacked):
    """Wraps ``placement_rule: {kind, inset_inches, door_offset_inches}``."""

    @property
    def kind(self):
        return _str_or_none(self._data.get("kind")) or KIND_CENTER

    @kind.setter
    def kind(self, value):
        self._data["kind"] = _str_or_none(value) or KIND_CENTER

    @property
    def inset_inches(self):
        v = self._data.get("inset_inches")
        if v is None:
            return 0.0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.0

    @inset_inches.setter
    def inset_inches(self, value):
        try:
            self._data["inset_inches"] = float(value)
        except (ValueError, TypeError):
            self._data["inset_inches"] = 0.0

    @property
    def door_offset_inches(self):
        """Offset *along* the door's outward normal / sideways, in inches.

        Stored as ``{x: float, y: float}`` where x is along the door
        opening direction (into the room) and y is along the door width.
        Returns the dict directly for editing; falls back to a fresh
        zero dict if missing.
        """
        d = self._data.setdefault("door_offset_inches", {})
        if not isinstance(d, dict):
            d = {}
            self._data["door_offset_inches"] = d
        d.setdefault("x", 0.0)
        d.setdefault("y", 0.0)
        return d

    @property
    def door_offset_x_inches(self):
        return float(self.door_offset_inches.get("x") or 0.0)

    @door_offset_x_inches.setter
    def door_offset_x_inches(self, value):
        self.door_offset_inches["x"] = float(value or 0.0)

    @property
    def door_offset_y_inches(self):
        return float(self.door_offset_inches.get("y") or 0.0)

    @door_offset_y_inches.setter
    def door_offset_y_inches(self, value):
        self.door_offset_inches["y"] = float(value or 0.0)

    def is_valid(self):
        return is_valid_placement_kind(self.kind)

    def __repr__(self):
        return "<PlacementRule kind={!r} inset={}>".format(
            self.kind, self.inset_inches
        )


# ---------------------------------------------------------------------
# Space LED
# ---------------------------------------------------------------------

class SpaceLED(_DictBacked):
    """A linked-element-definition placed against a Space anchor."""

    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @property
    def label(self):
        return _str_or_none(self._data.get("label"))

    @property
    def category(self):
        return _str_or_none(self._data.get("category"))

    @property
    def is_group(self):
        return _bool_or(self._data.get("is_group"), False)

    @property
    def parameters(self):
        return self._data.setdefault("parameters", {})

    @property
    def offsets(self):
        """Per-instance offset list (same shape as equipment LEDs).

        For a ``door_relative`` rule that places a copy at every door,
        the placement engine multiplies the door anchor list against
        this offset list — so a single LED with two offsets and three
        doors yields six placements.
        """
        raw = self._data.setdefault("offsets", [])
        return [Offset(o) for o in _ensure_list(raw)]

    @property
    def annotations(self):
        raw = self._data.get("annotations")
        if isinstance(raw, list):
            return [Annotation(a) for a in raw]
        return []

    @property
    def placement_rule(self):
        d = self._data.setdefault("placement_rule", {})
        if not isinstance(d, dict):
            d = {}
            self._data["placement_rule"] = d
        return PlacementRule(d)


# ---------------------------------------------------------------------
# Linked set
# ---------------------------------------------------------------------

class SpaceLinkedSet(_DictBacked):
    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @property
    def name(self):
        return _str_or_none(self._data.get("name"))

    @property
    def leds(self):
        raw = self._data.setdefault("linked_element_definitions", [])
        return [SpaceLED(l) for l in _ensure_list(raw)]


# ---------------------------------------------------------------------
# Space profile
# ---------------------------------------------------------------------

class SpaceProfile(_DictBacked):
    """A ``space_profiles[*]`` entry."""

    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @id.setter
    def id(self, value):
        self._data["id"] = _str_or_none(value)

    @property
    def name(self):
        return _str_or_none(self._data.get("name")) or ""

    @name.setter
    def name(self, value):
        self._data["name"] = _str_or_none(value) or ""

    @property
    def bucket_id(self):
        return _str_or_none(self._data.get("bucket_id"))

    @bucket_id.setter
    def bucket_id(self, value):
        self._data["bucket_id"] = _str_or_none(value)

    @property
    def linked_sets(self):
        raw = self._data.setdefault("linked_sets", [])
        return [SpaceLinkedSet(s) for s in _ensure_list(raw)]

    def __repr__(self):
        return "<SpaceProfile id={!r} bucket={!r} name={!r}>".format(
            self.id, self.bucket_id, self.name
        )


# ---------------------------------------------------------------------
# Collection helpers
# ---------------------------------------------------------------------

def wrap_profiles(raw):
    """Return ``[SpaceProfile, ...]`` from a list of dicts."""
    return [SpaceProfile(d) for d in (raw or []) if isinstance(d, dict)]


def find_profile_by_id(profiles, profile_id):
    if not profile_id:
        return None
    target = str(profile_id).strip()
    for p in profiles or ():
        pid = p.id if isinstance(p, SpaceProfile) else (p or {}).get("id")
        if str(pid or "").strip() == target:
            return p if isinstance(p, SpaceProfile) else SpaceProfile(p)
    return None


def profiles_for_bucket(profiles, bucket_id):
    """Return all profiles whose ``bucket_id`` matches.

    Multiple profiles per bucket is the stacked-templates contract:
    ``BUCKET-RESTROOM`` may have a baseline profile (toilets + sinks)
    and a "women's" profile (extra fixtures), and a Space tagged with
    that bucket gets every matching profile applied.
    """
    if not bucket_id:
        return []
    target = str(bucket_id).strip()
    out = []
    for p in profiles or ():
        wrapped = p if isinstance(p, SpaceProfile) else SpaceProfile(p)
        if str(wrapped.bucket_id or "").strip() == target:
            out.append(wrapped)
    return out


def profiles_for_buckets(profiles, bucket_ids):
    """Like ``profiles_for_bucket`` but unions multiple bucket IDs.

    Order: profiles preserve their declaration order in the YAML, and
    each bucket-id contributes once (deduped). This is what the
    placement engine uses when a Space has multiple assigned buckets.
    """
    if not bucket_ids:
        return []
    targets = set(str(b).strip() for b in bucket_ids if b)
    seen = set()
    out = []
    for p in profiles or ():
        wrapped = p if isinstance(p, SpaceProfile) else SpaceProfile(p)
        bid = str(wrapped.bucket_id or "").strip()
        if bid in targets and wrapped.id not in seen:
            seen.add(wrapped.id)
            out.append(wrapped)
    return out
