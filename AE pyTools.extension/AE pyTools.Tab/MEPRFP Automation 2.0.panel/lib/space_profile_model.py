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
#
# Cardinal kinds (n / s / e / w / ne / nw / se / sw) were dropped:
# project-north has no useful relationship to where the door of a
# space happens to land, so "north wall" was never a stable reference.
# The new vocabulary is door-anchored — every kind except ``center``
# resolves its geometry from the user-selected reference door of the
# space, so "wall opposite the door" and "corner furthest from the
# door" stay meaningful regardless of how the space is rotated.
# ---------------------------------------------------------------------

KIND_CENTER = "center"
KIND_DOOR_RELATIVE = "door_relative"
KIND_WALL_OPPOSITE_DOOR = "wall_opposite_door"
KIND_WALL_RIGHT_OF_DOOR = "wall_right_of_door"
KIND_WALL_LEFT_OF_DOOR = "wall_left_of_door"
KIND_CORNER_FURTHEST_FROM_DOOR = "corner_furthest_from_door"
KIND_CORNER_CLOSEST_TO_DOOR = "corner_closest_to_door"
# ``wall_anchored`` (deprecated capture default) — stored a single
# fraction along the closest wall. Kept for backward compatibility:
# any existing data with this kind still resolves via the old
# along-wall interpolation. New captures use ``space_anchored``
# instead.
KIND_WALL_ANCHORED = "wall_anchored"
# ``space_anchored`` is the current capture-driven kind. Position is
# stored as two fractions (``x_fraction``, ``y_fraction``) of the
# space's bounding box: a fixture at the (7', 5') mark in a 21' x
# 15' space stores (0.333, 0.333) and lands at (6', 5') in an 18' x
# 15' space, or (4.5', 4') in a 13.5' x 12' space — scaling with
# BOTH room dimensions, not just the wall length. ``wall_role`` is
# preserved on the rule purely for ROTATION resolution (so a
# captured wall-mounted fixture keeps facing into the room when
# the target space's wall roles map onto different cardinal
# directions). Door-dependent for the same reason.
KIND_SPACE_ANCHORED = "space_anchored"

# Wall-role tokens used inside ``placement_rule.wall_role`` for the
# wall_anchored kind.
WALL_ROLE_OPPOSITE_DOOR = "opposite_door"
WALL_ROLE_RIGHT_OF_DOOR = "right_of_door"
WALL_ROLE_LEFT_OF_DOOR = "left_of_door"
WALL_ROLE_BEHIND_DOOR = "behind_door"

WALL_ROLES = (
    WALL_ROLE_OPPOSITE_DOOR,
    WALL_ROLE_RIGHT_OF_DOOR,
    WALL_ROLE_LEFT_OF_DOOR,
    WALL_ROLE_BEHIND_DOOR,
)

PLACEMENT_KINDS = (
    KIND_CENTER,
    KIND_DOOR_RELATIVE,
    KIND_WALL_OPPOSITE_DOOR,
    KIND_WALL_RIGHT_OF_DOOR,
    KIND_WALL_LEFT_OF_DOOR,
    KIND_CORNER_FURTHEST_FROM_DOOR,
    KIND_CORNER_CLOSEST_TO_DOOR,
    KIND_WALL_ANCHORED,
    KIND_SPACE_ANCHORED,
)

# Every kind that needs a reference door to resolve. The placement
# workflow uses this set to decide whether a Space needs a door
# picker prompt.
DOOR_DEPENDENT_KINDS = frozenset({
    KIND_DOOR_RELATIVE,
    KIND_WALL_OPPOSITE_DOOR,
    KIND_WALL_RIGHT_OF_DOOR,
    KIND_WALL_LEFT_OF_DOOR,
    KIND_CORNER_FURTHEST_FROM_DOOR,
    KIND_CORNER_CLOSEST_TO_DOOR,
    KIND_WALL_ANCHORED,
    KIND_SPACE_ANCHORED,
})

WALL_KINDS = frozenset({
    KIND_WALL_OPPOSITE_DOOR,
    KIND_WALL_RIGHT_OF_DOOR,
    KIND_WALL_LEFT_OF_DOOR,
    KIND_WALL_ANCHORED,
    KIND_SPACE_ANCHORED,
})

CORNER_KINDS = frozenset({
    KIND_CORNER_FURTHEST_FROM_DOOR,
    KIND_CORNER_CLOSEST_TO_DOOR,
})


def is_valid_placement_kind(kind):
    return kind in PLACEMENT_KINDS


def is_door_dependent(kind):
    """True when the kind requires a reference door to resolve."""
    return kind in DOOR_DEPENDENT_KINDS


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

    # ----- wall_anchored fields -----------------------------------
    # Used only when ``kind == wall_anchored``. Stored on the rule
    # so the placement engine can resolve "which wall" (wall_role)
    # and "where along it" (position_along_wall, 0.0..1.0) without
    # the LED needing per-instance offsets in absolute inches.

    @property
    def wall_role(self):
        return _str_or_none(self._data.get("wall_role")) or WALL_ROLE_OPPOSITE_DOOR

    @wall_role.setter
    def wall_role(self, value):
        text = _str_or_none(value) or WALL_ROLE_OPPOSITE_DOOR
        self._data["wall_role"] = text

    @property
    def position_along_wall(self):
        """Fraction along the chosen wall: 0.0 at the start, 1.0 at
        the end. Stored as a fraction (not absolute inches) so a
        captured 7'/21' position scales to 6'/18' in another project.
        Defaults to 0.5 (midpoint) when missing."""
        v = self._data.get("position_along_wall")
        if v is None:
            return 0.5
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.5

    @position_along_wall.setter
    def position_along_wall(self, value):
        try:
            self._data["position_along_wall"] = float(value)
        except (ValueError, TypeError):
            self._data["position_along_wall"] = 0.5

    @property
    def distance_from_wall_inches(self):
        """Perpendicular distance inward from the wall surface, in
        inches. Used by wall-mounted fixtures that need a small
        offset from the bbox edge (cover plate thickness, etc.)."""
        v = self._data.get("distance_from_wall_inches")
        if v is None:
            return 0.0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.0

    @distance_from_wall_inches.setter
    def distance_from_wall_inches(self, value):
        try:
            self._data["distance_from_wall_inches"] = float(value)
        except (ValueError, TypeError):
            self._data["distance_from_wall_inches"] = 0.0

    # ----- space_anchored fields ----------------------------------
    # Used by ``kind == space_anchored``. Stored as two fractions
    # of the space's bounding box (0..1 in each axis). At placement
    # time:
    #   target_x = bbox.xmin + x_fraction * (bbox.xmax - bbox.xmin)
    #   target_y = bbox.ymin + y_fraction * (bbox.ymax - bbox.ymin)
    # so a fixture at the (7', 5') mark of a 21' x 15' space
    # (0.333, 0.333) lands at (6', 5') in an 18' x 15' space, and
    # (4.5', 4') in a 13.5' x 12' space — scales with both room
    # dimensions, not just one wall.

    @property
    def x_fraction(self):
        v = self._data.get("x_fraction")
        if v is None:
            return 0.5
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.5

    @x_fraction.setter
    def x_fraction(self, value):
        try:
            self._data["x_fraction"] = float(value)
        except (ValueError, TypeError):
            self._data["x_fraction"] = 0.5

    @property
    def y_fraction(self):
        v = self._data.get("y_fraction")
        if v is None:
            return 0.5
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.5

    @y_fraction.setter
    def y_fraction(self, value):
        try:
            self._data["y_fraction"] = float(value)
        except (ValueError, TypeError):
            self._data["y_fraction"] = 0.5

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
