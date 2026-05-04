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
    )

    def __init__(self, space, geom, profile, led, label, set_id,
                 world_pt, rotation_deg, warnings=None):
        self.space = space
        self.geom = geom
        self.profile = profile
        self.led = led
        self.label = label or ""
        self.set_id = set_id or ""
        self.world_pt = tuple(world_pt) if world_pt else None
        self.rotation_deg = float(rotation_deg or 0.0)
        self.warnings = list(warnings or [])

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
    """One end-to-end pass. UI calls ``collect`` then ``apply``."""

    def __init__(self, doc, profile_data=None):
        self.doc = doc
        self.profile_data = profile_data or {}
        self.plans = []
        self.warnings = []

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

            # 4. expand every LED in every matching profile
            for profile in matching:
                for linked_set in profile.linked_sets:
                    set_id = linked_set.id or ""
                    for led in linked_set.leds:
                        placements = _placement.expand_led_placements(led, geom)
                        if not placements:
                            # Door-relative LED in a doorless space, or
                            # an unknown placement kind — note it once.
                            kind = led.placement_rule.kind
                            self.warnings.append(
                                "Space '{}': LED {!r} ({}) yielded no placements.".format(
                                    space.name, led.label, kind
                                )
                            )
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
                            ))

        return self.plans

    # ----- apply -----------------------------------------------------

    def apply(self):
        """Execute placement on Revit's main thread.

        Routed through ``space_apply`` so the Revit-API side stays in
        one place (and so this orchestrator can stay testable in the
        pure-logic layer up to ``collect``).
        """
        import space_apply
        return space_apply.apply_plans(self.doc, self.plans)
