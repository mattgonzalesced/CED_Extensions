# -*- coding: utf-8 -*-
"""
Stage 6 — Spaces annotation placement.

Walks every placed family instance in the active view whose
Element_Linker carries the ``space_id`` lineage marker, looks up the
source LED in ``space_profiles[*]``, and emits an
``AnnotationCandidate`` for each ``annotations[*]`` entry. Reuses the
equipment-side ``annotation_placement._place_tag`` /
``_place_keynote`` / ``_place_text_note`` for the Revit-API edge so
tag-vs-text-note routing matches what the equipment pipeline does.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    FamilyInstance,
    FilteredElementCollector,
    LocationPoint,
    XYZ,
)

import annotation_placement as _ap  # noqa: E402
import element_linker_io as _el_io  # noqa: E402
import geometry  # noqa: E402


# ---------------------------------------------------------------------
# LED index
# ---------------------------------------------------------------------

def _build_space_led_index(profile_data):
    """Return ``{led_id: (profile, set, led)}`` from ``space_profiles``."""
    out = {}
    for profile in profile_data.get("space_profiles") or []:
        if not isinstance(profile, dict):
            continue
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                lid = led.get("id")
                if lid:
                    out[lid] = (profile, set_dict, led)
    return out


def _fixture_pt_and_rot(fixture):
    loc = getattr(fixture, "Location", None)
    pt = None
    rad = 0.0
    if isinstance(loc, LocationPoint):
        try:
            pt = loc.Point
        except Exception:
            pt = None
        try:
            rad = loc.Rotation
        except Exception:
            rad = 0.0
    if pt is None:
        try:
            bbox = fixture.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is not None:
            pt = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
    if pt is None:
        return None, 0.0
    return (pt.X, pt.Y, pt.Z), geometry.normalize_angle(math.degrees(rad))


# ---------------------------------------------------------------------
# Collection
# ---------------------------------------------------------------------

def collect_space_candidates(doc, view, profile_data, kinds=None,
                             skip_duplicates=True):
    """Return ``[AnnotationCandidate, ...]`` for space-based fixtures.

    Filters fixtures to those where ``Element_Linker.is_space_based``
    is true (i.e. ``space_id`` was stamped at placement time).
    ``kinds`` defaults to all three (tag / keynote / text note).

    When ``skip_duplicates`` is True (the default), each candidate is
    run through ``annotation_placement.mark_duplicates`` so already-
    placed tags / keynotes / text notes get ``skip = True`` and a
    ``duplicate_reason`` populated. The UI shows the reason and
    ``execute_placement`` skips them at apply time.
    """
    if kinds is None:
        kinds = set(_ap.ALL_KINDS)
    else:
        kinds = set(kinds)

    led_index = _build_space_led_index(profile_data)
    out = []

    if view is not None:
        host_collector = FilteredElementCollector(doc, view.Id)
    else:
        host_collector = FilteredElementCollector(doc)

    for fixture in host_collector.OfClass(FamilyInstance).WhereElementIsNotElementType():
        linker = _el_io.read_from_element(fixture)
        if linker is None or not linker.is_space_based:
            continue
        if not linker.led_id:
            continue
        entry = led_index.get(linker.led_id)
        if entry is None:
            # The fixture references a LED that's no longer in the YAML.
            # Most likely the YAML was edited or replaced — skip silently.
            continue
        profile, _set, led = entry
        annotations = led.get("annotations") or []
        if not annotations:
            continue
        fixture_pt, fixture_rot = _fixture_pt_and_rot(fixture)
        if fixture_pt is None:
            continue

        for ann in annotations:
            if not isinstance(ann, dict):
                continue
            kind = ann.get("kind") or _ap.KIND_TAG
            if kind not in kinds:
                continue
            offset = ann.get("offsets") or {}
            if isinstance(offset, list):
                offset = offset[0] if offset else {}
            target_pt = geometry.target_point_from_offsets(
                fixture_pt, fixture_rot, offset,
            )
            target_rot = geometry.child_rotation_from_offsets(
                fixture_rot, offset,
            )
            out.append(_ap.AnnotationCandidate(
                fixture=fixture,
                fixture_pt=fixture_pt,
                fixture_rot=fixture_rot,
                led_id=led.get("id") or "",
                led_label=led.get("label") or "",
                profile_id=profile.get("id") or "",
                profile_name=profile.get("name") or "",
                annotation=ann,
                target_pt=target_pt,
                target_rot=target_rot,
            ))

    if skip_duplicates and out:
        try:
            _ap.mark_duplicates(doc, view, out)
        except Exception:
            # Dedup is a UX nicety — never let it block placement preview.
            pass

    return out


# ---------------------------------------------------------------------
# Apply (delegates to the equipment-side machinery)
# ---------------------------------------------------------------------

def execute_placement(doc, view, candidates):
    return _ap.execute_placement(doc, view, candidates)


def is_view_eligible(view):
    return _ap.is_view_eligible(view)
