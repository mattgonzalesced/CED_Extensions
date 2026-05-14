# -*- coding: utf-8 -*-
"""
Stage 6 — Spaces placement orchestrator.

Walks the active document, expands every classified Space's assigned
profiles into concrete placement plans, and hands the result to
``space_apply.apply_plans``. The pure-logic side
(``space_placement.expand_led_placements``) does the geometry; this
module only stitches Revit-side data together.

A *placement plan* is one prospective family-instance:

  - ``space``           SpaceInfo (from ``space_workflow.collect_spaces``)
  - ``geom``            SpaceGeometry (from ``space_placement.build_space_geometry``)
  - ``profile``         SpaceProfile (the template entry that produced this plan)
  - ``led``             SpaceLED (the linked-element-definition)
  - ``label``           "Family : Type" string for display + symbol lookup
  - ``world_pt``        (x, y, z) tuple in feet
  - ``rotation_deg``    rotation about Z, degrees
  - ``set_id``          owning ``linked_set`` id (echoed into Element_Linker)
"""

import space_bucket_model as _buckets
import space_placement as _placement
import space_profile_model as _profile_model
import space_workflow as _space_workflow


# ---------------------------------------------------------------------
# Plain-data record
# ---------------------------------------------------------------------

class SpacePlacementPlan(object):

    __slots__ = (
        "space", "geom", "profile", "led",
        "label", "set_id",
        "world_pt", "rotation_deg",
        "warnings",
        "comment",
    )

    def __init__(self, space, geom, profile, led, label, set_id,
                 world_pt, rotation_deg, warnings=None, comment=""):
        self.space = space
        self.geom = geom
        self.profile = profile
        self.led = led
        self.label = label or ""
        self.set_id = set_id or ""
        self.world_pt = tuple(world_pt) if world_pt else None
        self.rotation_deg = float(rotation_deg or 0.0)
        self.warnings = list(warnings or [])
        # Comment shown in the placement window's Comments column.
        # Non-empty when the plan is informational only — e.g. a
        # door-relative LED in a doorless space — which means
        # ``world_pt`` is None and the apply step skips it.
        self.comment = comment or ""

    @property
    def is_placeable(self):
        """True only when this plan has a real anchor point. Comment-
        only / informational plans (no door, unknown rule kind, etc.)
        return False here so the apply step skips them."""
        return self.world_pt is not None

    @property
    def space_element_id(self):
        return self.space.element_id if self.space is not None else None

    @property
    def profile_id(self):
        return self.profile.id if self.profile is not None else None

    @property
    def led_id(self):
        return self.led.id if self.led is not None else None

    def __repr__(self):
        return ("<SpacePlacementPlan space={} profile={} led={} pt={}>".format(
            self.space_element_id, self.profile_id, self.led_id, self.world_pt
        ))


# ---------------------------------------------------------------------
# Top-level
# ---------------------------------------------------------------------

class SpacePlacementRun(object):
    """One end-to-end pass. UI calls ``collect`` then ``apply``.

    For door-dependent placement kinds the UI supplies a *door
    picker* — a callable ``picker(space, door_anchors) ->
    (origin_xy, inward_xy) | None`` that returns the chosen anchor
    tuple, or ``None`` to fall back to the first door. The default
    picker is a no-op (returns None), which gives sane preview
    defaults; the typical caller pre-populates ``door_choices``
    via ``Selection.PickObject`` before opening the placement UI.

    ``door_choices`` is a ``{space_element_id: anchor_tuple}`` cache
    so the user only has to choose once per space per session.
    """

    def __init__(self, doc, profile_data=None,
                 door_picker=None, door_choices=None):
        self.doc = doc
        self.profile_data = profile_data or {}
        self.plans = []
        self.warnings = []
        # ``picker`` is invoked at most once per space per ``collect``
        # run, only when the space has multiple doors AND at least
        # one of its LEDs is door-dependent. Default no-op picker
        # returns None, falling back to the first door for any
        # ambiguous space — preview-friendly default when nothing
        # has been pre-picked.
        self._door_picker = door_picker or (lambda space, doors: None)
        self.door_choices = dict(door_choices or {})
        # Spaces touched in the latest collect() that needed (and
        # potentially still need) a door choice. Useful for the UI to
        # display "N spaces have multiple doors" diagnostics.
        self.spaces_with_multiple_doors = []

        self._buckets = _buckets.wrap_buckets(
            self.profile_data.get("space_buckets") or []
        )
        self._profiles = _profile_model.wrap_profiles(
            self.profile_data.get("space_profiles") or []
        )

    # ----- collection ------------------------------------------------

    def collect(self):
        """Walk the doc and build every placement plan."""
        self.plans = []
        self.warnings = []
        self.spaces_with_multiple_doors = []

        # 1. enumerate classified spaces
        spaces = _space_workflow.collect_spaces(self.doc)
        classifications = _space_workflow.load_classifications_indexed(self.doc)

        if not spaces:
            self.warnings.append("No placed Spaces in this document.")
            return self.plans

        if not classifications:
            self.warnings.append(
                "No saved Space classifications. Run Classify Spaces first."
            )
            return self.plans

        if not self._profiles:
            self.warnings.append(
                "No space_profiles defined in the active YAML."
            )
            return self.plans

        # Index space_id -> SpaceInfo for fast lookup.
        space_by_id = {s.element_id: s for s in spaces if s.element_id is not None}

        for sid, bucket_ids in classifications.items():
            space = space_by_id.get(sid)
            if space is None:
                # Stale classification — the Space was deleted.
                self.warnings.append(
                    "Skipped stale classification for ElementId {}.".format(sid)
                )
                continue
            if not bucket_ids:
                continue

            # 2. find every profile that targets one of this space's buckets
            matching = _profile_model.profiles_for_buckets(
                self._profiles, bucket_ids,
            )
            if not matching:
                self.warnings.append(
                    "Space '{}' (id {}) classified but no profiles match its buckets {}.".format(
                        space.name, sid, bucket_ids
                    )
                )
                continue

            # 3. compute geometry once per space
            try:
                geom = _placement.build_space_geometry(self.doc, space.element)
            except Exception as exc:
                self.warnings.append(
                    "Failed to build geometry for '{}' (id {}): {}".format(
                        space.name, sid, exc
                    )
                )
                continue
            if geom is None:
                self.warnings.append(
                    "Space '{}' (id {}) has no usable boundary.".format(
                        space.name, sid
                    )
                )
                continue

            # 4. resolve which door this space uses for door-dependent
            # placement kinds. Only call the picker when the space has
            # MORE than one door AND at least one of its LEDs needs a
            # door reference.
            chosen_door = self._resolve_door_for_space(space, geom, matching)

            # 5. expand every LED in every matching profile
            for profile in matching:
                for linked_set in profile.linked_sets:
                    set_id = linked_set.id or ""
                    for led in linked_set.leds:
                        placements = _placement.expand_led_placements(
                            led, geom, door_anchor=chosen_door,
                        )
                        # Drain any space_anchored anchor-resolution
                        # diagnostics this LED's expand call emitted, so
                        # we can stash them on every plan it produced —
                        # the preview UI's Comments column then shows
                        # exactly which bbox the engine used and which
                        # fractions it read for each LED, without the
                        # user having to apply first.
                        try:
                            led_diag_lines = (
                                _placement.drain_space_anchored_diagnostics()
                            )
                        except Exception:
                            led_diag_lines = []
                        led_diag_text = (
                            " | ".join(led_diag_lines) if led_diag_lines else ""
                        )
                        if not placements:
                            # Empty result — make a comment-only plan
                            # so the user sees the LED in the preview
                            # with a note explaining why it can't be
                            # placed. Apply step skips these (their
                            # ``is_placeable`` is False).
                            kind = led.placement_rule.kind
                            if kind == _profile_model.KIND_DOOR_RELATIVE:
                                comment = (
                                    "Door-relative placement: no host or "
                                    "linked door found at this space."
                                )
                            else:
                                comment = (
                                    "No anchor points computed (rule kind: "
                                    "{!r}).".format(kind)
                                )
                            self.warnings.append(
                                "Space '{}': LED {!r} ({}) -- {}".format(
                                    space.name, led.label, kind, comment,
                                )
                            )
                            self.plans.append(SpacePlacementPlan(
                                space=space,
                                geom=geom,
                                profile=profile,
                                led=led,
                                label=led.label or "",
                                set_id=set_id,
                                world_pt=None,
                                rotation_deg=0.0,
                                comment=comment,
                            ))
                            continue
                        for x, y, z, rot in placements:
                            self.plans.append(SpacePlacementPlan(
                                space=space,
                                geom=geom,
                                profile=profile,
                                led=led,
                                label=led.label or "",
                                set_id=set_id,
                                world_pt=(x, y, z),
                                rotation_deg=rot,
                                comment=led_diag_text,
                            ))

        return self.plans

    # ----- apply -----------------------------------------------------

    def _profile_uses_door_dependent_led(self, matching_profiles):
        """True if any LED across the matching profiles needs a door."""
        for profile in matching_profiles:
            for linked_set in profile.linked_sets:
                for led in linked_set.leds:
                    if _profile_model.is_door_dependent(led.placement_rule.kind):
                        return True
        return False

    def _resolve_door_for_space(self, space, geom, matching_profiles):
        """Pick the reference door for this space, prompting the user
        when there's more than one option AND at least one LED needs
        a door. Result is cached in ``self.door_choices`` keyed by
        space element id so repeated collect() runs don't re-prompt.

        Returns the chosen ``(origin_xy, inward_xy)`` tuple, or
        ``None`` when no doors are present (caller emits comment-only
        plans for door-dependent LEDs).
        """
        doors = list(geom.door_anchors or [])
        if not doors:
            return None
        if len(doors) == 1:
            return doors[0]

        # Multi-door space — track for UI diagnostics.
        if space.element_id not in [
            s.element_id for s, _doors in self.spaces_with_multiple_doors
        ]:
            self.spaces_with_multiple_doors.append((space, doors))

        # If no LED in this space needs a door, the choice doesn't
        # matter; default to the first one without prompting.
        if not self._profile_uses_door_dependent_led(matching_profiles):
            return doors[0]

        # Cached choice from a prior session OR from the pre-pick
        # step in the calling script (Selection.PickObject before the
        # modal opened).
        cached = self.door_choices.get(space.element_id)
        if cached is not None:
            return cached

        # Ask the picker. Convention: returns an anchor tuple
        # ``(origin_xy, inward_xy)`` or None to fall back.
        try:
            chosen = self._door_picker(space, doors)
        except Exception as exc:
            self.warnings.append(
                "Door picker failed for '{}' (id {}): {}. "
                "Defaulting to first door.".format(
                    space.name, space.element_id, exc,
                )
            )
            chosen = None
        if chosen is None:
            chosen = doors[0]
        self.door_choices[space.element_id] = chosen
        return chosen

    def apply(self, plans=None):
        """Execute placement on Revit's main thread.

        ``plans`` lets the UI hand in a filtered subset (e.g. only
        rows the user ticked); when omitted, every plan from the
        last ``collect()`` is applied.
        """
        import space_apply
        target = plans if plans is not None else self.plans
        return space_apply.apply_plans(self.doc, target)
