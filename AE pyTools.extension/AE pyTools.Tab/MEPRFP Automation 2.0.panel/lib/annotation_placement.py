# -*- coding: utf-8 -*-
"""
Annotation placement engine.

For every fixture in the active view that has an ``Element_Linker``
payload, look up its LED in the active YAML, and place each of the
LED's ``annotations`` entries (tag / keynote / text_note) at the
appropriate world point relative to the fixture's current pose.

Same general shape as ``placement.py``: collect targets -> match (or
classify) -> preview -> commit. Caller manages the Revit transaction.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    Group,
    IndependentTag,
    LocationPoint,
    Reference,
    TagMode,
    TagOrientation,
    TextNote,
    TextNoteType,
    View,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType  # noqa: E402

import element_linker_io as _el_io
import geometry


# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------

KIND_TAG = "tag"
KIND_KEYNOTE = "keynote"
KIND_TEXT_NOTE = "text_note"

ALL_KINDS = (KIND_TAG, KIND_KEYNOTE, KIND_TEXT_NOTE)

# Proximity tolerance for already-placed dedup (feet).
DEDUP_RADIUS_FT = 1.0


# ---------------------------------------------------------------------
# View eligibility
# ---------------------------------------------------------------------

# View types we won't place annotations into.
_DISALLOWED_VIEW_TYPES = frozenset({
    ViewType.ThreeD,
    ViewType.Schedule,
    ViewType.Walkthrough,
    ViewType.Internal,
    ViewType.SystemBrowser,
    ViewType.ProjectBrowser,
    ViewType.Undefined,
})


def is_view_eligible(view):
    if view is None:
        return False, "No active view"
    if view.IsTemplate:
        return False, "Active view is a view template"
    if view.ViewType in _DISALLOWED_VIEW_TYPES:
        return False, "Active view type ({}) is not supported".format(view.ViewType)
    return True, ""


# ---------------------------------------------------------------------
# Match record
# ---------------------------------------------------------------------

class AnnotationCandidate(object):
    """One (fixture, annotation_descriptor) pair the user can place or skip."""

    __slots__ = (
        "fixture", "fixture_pt", "fixture_rot",
        "led_id", "led_label",
        "profile_id", "profile_name",
        "annotation",
        "target_pt", "target_rot",
        "skip", "duplicate_reason",
    )

    def __init__(self, fixture, fixture_pt, fixture_rot,
                 led_id, led_label, profile_id, profile_name,
                 annotation, target_pt, target_rot,
                 skip=False, duplicate_reason=""):
        self.fixture = fixture
        self.fixture_pt = fixture_pt
        self.fixture_rot = fixture_rot
        self.led_id = led_id
        self.led_label = led_label
        self.profile_id = profile_id
        self.profile_name = profile_name
        self.annotation = annotation
        self.target_pt = target_pt
        self.target_rot = target_rot
        self.skip = skip
        self.duplicate_reason = duplicate_reason


# ---------------------------------------------------------------------
# Profile / LED lookup helpers
# ---------------------------------------------------------------------

def _build_led_index(profile_data):
    """Return ``{led_id: (profile, set, led)}``."""
    out = {}
    for profile in profile_data.get("equipment_definitions") or []:
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
    """Return ``((x, y, z), rotation_deg)`` for a placed fixture in the
    active doc."""
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
        bbox = fixture.get_BoundingBox(None)
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
# Filters
# ---------------------------------------------------------------------

class CollectFilters(object):
    """User-driven filters that narrow the candidate list."""

    def __init__(self,
                 kinds=None,
                 profile_ids=None,
                 categories=None,
                 active_view_only=True):
        self.kinds = set(kinds) if kinds is not None else set(ALL_KINDS)
        self.profile_ids = set(profile_ids) if profile_ids else None
        self.categories = set(categories) if categories else None
        self.active_view_only = active_view_only


# ---------------------------------------------------------------------
# Candidate collection
# ---------------------------------------------------------------------

def collect_candidates(doc, view, profile_data, filters):
    """Walk fixtures (FamilyInstance + Group) in ``view`` (or doc) that
    have Element_Linker payloads, look up their LEDs, and emit one
    ``AnnotationCandidate`` per (fixture, annotation) pair that passes
    the filters."""
    led_index = _build_led_index(profile_data)
    out = []

    if filters.active_view_only and view is not None:
        host_collector = FilteredElementCollector(doc, view.Id)
    else:
        host_collector = FilteredElementCollector(doc)

    for klass in (FamilyInstance, Group):
        collector = host_collector.OfClass(klass).WhereElementIsNotElementType()
        for fixture in collector:
            linker = _el_io.read_from_element(fixture)
            if linker is None or not linker.led_id:
                continue
            entry = led_index.get(linker.led_id)
            if entry is None:
                continue
            profile, _set, led = entry
            if filters.profile_ids is not None and \
                    profile.get("id") not in filters.profile_ids:
                continue
            if filters.categories is not None:
                cat = (profile.get("parent_filter") or {}).get("category") or ""
                if cat not in filters.categories:
                    continue
            annotations = led.get("annotations") or []
            if not annotations:
                continue
            fixture_pt, fixture_rot = _fixture_pt_and_rot(fixture)
            if fixture_pt is None:
                continue
            for ann in annotations:
                if not isinstance(ann, dict):
                    continue
                kind = ann.get("kind") or KIND_TAG
                if kind not in filters.kinds:
                    continue
                offset = ann.get("offsets") or {}
                if isinstance(offset, list):
                    offset = offset[0] if offset else {}
                target_pt = geometry.target_point_from_offsets(
                    fixture_pt, fixture_rot, offset
                )
                target_rot = geometry.child_rotation_from_offsets(
                    fixture_rot, offset
                )
                out.append(AnnotationCandidate(
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
    return out


# ---------------------------------------------------------------------
# Already-placed dedup
# ---------------------------------------------------------------------

def _xy_distance(a, b):
    if a is None or b is None:
        return float("inf")
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2)


def _id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _existing_tags_for_view(doc, view):
    tags = []
    try:
        for t in FilteredElementCollector(doc, view.Id).OfClass(IndependentTag):
            tags.append(t)
    except Exception:
        pass
    return tags


def _existing_text_notes_for_view(doc, view):
    notes = []
    try:
        for n in FilteredElementCollector(doc, view.Id).OfClass(TextNote):
            notes.append(n)
    except Exception:
        pass
    return notes


def _existing_family_instances_for_view(doc, view, family_name, type_name):
    out = []
    try:
        for fi in FilteredElementCollector(doc, view.Id).OfClass(FamilyInstance):
            sym = getattr(fi, "Symbol", None)
            if sym is None:
                continue
            family = getattr(sym, "Family", None)
            if family is None:
                continue
            if family_name and family.Name != family_name:
                continue
            if type_name and sym.Name != type_name:
                continue
            out.append(fi)
    except Exception:
        pass
    return out


def _tag_targets_id(tag):
    """Local element ids the tag references."""
    out = set()
    try:
        for eid in tag.GetTaggedLocalElementIds():
            out.add(getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None))
    except Exception:
        try:
            lei = tag.TaggedElementId
            if lei is not None:
                hid = lei.HostElementId
                if hid is not None and hid != ElementId.InvalidElementId:
                    out.add(getattr(hid, "Value", None) or getattr(hid, "IntegerValue", None))
        except Exception:
            pass
    return out


def mark_duplicates(doc, view, candidates):
    """Set ``candidate.skip = True`` and ``duplicate_reason`` for each
    candidate that matches an existing annotation in the view."""
    if view is None:
        return candidates
    existing_tags = _existing_tags_for_view(doc, view)
    existing_notes = _existing_text_notes_for_view(doc, view)

    for c in candidates:
        kind = c.annotation.get("kind")
        family_name = c.annotation.get("family_name") or ""
        type_name = c.annotation.get("type_name") or ""
        target_xy = c.target_pt[:2]

        if kind == KIND_TAG:
            fid = _id_value(c.fixture)
            for t in existing_tags:
                if fid not in _tag_targets_id(t):
                    continue
                t_type = doc.GetElement(t.GetTypeId()) if t.GetTypeId() else None
                t_family = getattr(t_type, "FamilyName", "") if t_type else ""
                t_typename = getattr(t_type, "Name", "") if t_type else ""
                if family_name and family_name != t_family:
                    continue
                if type_name and type_name != t_typename:
                    continue
                c.skip = True
                c.duplicate_reason = "Tag already on this fixture"
                break

        elif kind == KIND_KEYNOTE:
            for fi in _existing_family_instances_for_view(
                doc, view, family_name, type_name
            ):
                loc = getattr(fi, "Location", None)
                pt = loc.Point if isinstance(loc, LocationPoint) else None
                if pt is None:
                    continue
                if _xy_distance((pt.X, pt.Y), target_xy) < DEDUP_RADIUS_FT:
                    c.skip = True
                    c.duplicate_reason = "Keynote symbol already nearby"
                    break

        elif kind == KIND_TEXT_NOTE:
            target_text = c.annotation.get("text") or c.annotation.get("label") or ""
            for n in existing_notes:
                try:
                    n_text = (n.Text or "").strip()
                except Exception:
                    n_text = ""
                if target_text and n_text == target_text:
                    coord = getattr(n, "Coord", None)
                    if coord is not None and _xy_distance(
                        (coord.X, coord.Y), target_xy
                    ) < DEDUP_RADIUS_FT:
                        c.skip = True
                        c.duplicate_reason = "Text note with same content nearby"
                        break

    return candidates


# ---------------------------------------------------------------------
# Type lookup
# ---------------------------------------------------------------------

def _find_family_symbol(doc, family_name, type_name):
    if not family_name:
        return None
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        family = getattr(sym, "Family", None)
        if family is None or family.Name != family_name:
            continue
        if type_name and sym.Name != type_name:
            continue
        return sym
    return None


def _find_text_note_type(doc, type_name):
    """Match TextNoteType by ``Name``. Falls back to the first available."""
    types = list(FilteredElementCollector(doc).OfClass(TextNoteType))
    if type_name:
        for t in types:
            if t.Name == type_name:
                return t
    return types[0] if types else None


def _activate_symbol(symbol):
    if symbol is None:
        return
    try:
        if not symbol.IsActive:
            symbol.Activate()
    except Exception:
        pass


# ---------------------------------------------------------------------
# Single-annotation placement
# ---------------------------------------------------------------------

def _place_tag(doc, view, candidate):
    """Place an IndependentTag at the captured offset, then add a leader.

    Two-phase to keep the head pinned to the captured offset *and* end
    up with a leader:

        1. ``IndependentTag.Create`` with ``addLeader=False`` — head
           lands exactly at ``target_pt`` because Revit isn't running
           its auto-routing pass.
        2. ``ChangeTypeId`` then re-pin ``TagHeadPosition`` (type swap
           can nudge the head a hair).
        3. Flip ``HasLeader = True`` on the placed tag — Revit adds a
           default leader from the (already-fixed) head to the tagged
           element without moving the head.

    Creating with ``addLeader=True`` is what fought the captured offset
    in earlier builds; the auto-route ran during creation and shifted
    the head before our explicit reset could take.
    """
    fixture = candidate.fixture
    target_pt = XYZ(*candidate.target_pt)
    family_name = candidate.annotation.get("family_name") or ""
    type_name = candidate.annotation.get("type_name") or ""
    sym = _find_family_symbol(doc, family_name, type_name)
    if sym is None:
        return None, "Tag type {!r} : {!r} not loaded".format(family_name, type_name)
    _activate_symbol(sym)
    ref = Reference(fixture)
    tag = None
    last_err = None
    for mode in (TagMode.TM_ADDBY_CATEGORY, TagMode.TM_ADDBY_MULTICATEGORY):
        try:
            tag = IndependentTag.Create(
                doc, view.Id, ref, False, mode,
                TagOrientation.Horizontal, target_pt,
            )
            break
        except Exception as exc:
            last_err = exc
            tag = None
    if tag is None:
        return None, "IndependentTag.Create failed: {}".format(last_err)

    try:
        tag.ChangeTypeId(sym.Id)
    except Exception:
        pass

    # Pin head before turning the leader on so HasLeader=True can't
    # take the auto-route shortcut from a stale position.
    try:
        tag.TagHeadPosition = target_pt
    except Exception:
        pass

    try:
        tag.HasLeader = True
    except Exception:
        pass

    # And again after — toggling HasLeader can re-snap the head on
    # some Revit builds.
    try:
        tag.TagHeadPosition = target_pt
    except Exception:
        pass
    return tag, ""


def _place_keynote(doc, view, candidate):
    family_name = candidate.annotation.get("family_name") or ""
    type_name = candidate.annotation.get("type_name") or ""
    sym = _find_family_symbol(doc, family_name, type_name)
    if sym is None:
        return None, "Keynote type {!r} : {!r} not loaded".format(family_name, type_name)
    _activate_symbol(sym)
    target_pt = XYZ(*candidate.target_pt)
    inst = None
    try:
        inst = doc.Create.NewFamilyInstance(target_pt, sym, view)
    except Exception:
        try:
            inst = doc.Create.NewFamilyInstance(
                target_pt, sym, StructuralType.NonStructural
            )
        except Exception as exc:
            return None, "NewFamilyInstance failed: {}".format(exc)
    if inst is not None and abs(candidate.target_rot) > geometry.Tolerances.ROTATION_DEG:
        try:
            from Autodesk.Revit.DB import ElementTransformUtils, Line
            axis = Line.CreateBound(
                target_pt,
                XYZ(target_pt.X, target_pt.Y, target_pt.Z + 1.0),
            )
            ElementTransformUtils.RotateElement(
                doc, inst.Id, axis, math.radians(candidate.target_rot)
            )
        except Exception:
            pass
    return inst, ""


def _place_text_note(doc, view, candidate):
    """Create a TextNote at the captured offset.

    ``TextNote.Create`` honours the input position for the anchor
    corner, but Revit's auto-width / alignment heuristics can shift
    ``Coord`` slightly after creation. We always re-pin ``Coord`` to
    the captured target so the offset survives placement, regardless
    of whether the note rotates.
    """
    text = candidate.annotation.get("text") or candidate.annotation.get("label") or ""
    if not text:
        return None, "Text note has no content"
    type_name = candidate.annotation.get("type_name") or ""
    note_type = _find_text_note_type(doc, type_name)
    if note_type is None:
        return None, "No TextNoteType available in the document"
    target_pt = XYZ(*candidate.target_pt)
    try:
        note = TextNote.Create(doc, view.Id, target_pt, text, note_type.Id)
    except Exception as exc:
        return None, "TextNote.Create failed: {}".format(exc)
    if note is not None:
        try:
            note.Coord = target_pt
        except Exception:
            pass
        if abs(candidate.target_rot) > geometry.Tolerances.ROTATION_DEG:
            try:
                # TextNote rotation is via its Rotation property, in radians.
                note.Rotation = math.radians(candidate.target_rot)
            except Exception:
                pass
    return note, ""


# ---------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------

class AnnotationPlacementResult(object):
    def __init__(self):
        self.placed_count_by_kind = {KIND_TAG: 0, KIND_KEYNOTE: 0, KIND_TEXT_NOTE: 0}
        self.skipped_duplicates = 0
        self.warnings = []

    @property
    def total_placed(self):
        return sum(self.placed_count_by_kind.values())


def _apply_parameters(elem, params_dict):
    if elem is None or not params_dict:
        return
    for name, value in params_dict.items():
        if name is None or value is None or value == "":
            continue
        try:
            p = elem.LookupParameter(name)
        except Exception:
            continue
        if p is None or p.IsReadOnly:
            continue
        try:
            p.Set(str(value))
        except Exception:
            try:
                p.Set(int(value))
            except Exception:
                try:
                    p.Set(float(value))
                except Exception:
                    pass


def execute_placement(doc, view, candidates):
    """Place every non-skipped candidate. Caller manages the transaction."""
    result = AnnotationPlacementResult()
    if view is None:
        result.warnings.append("No view supplied")
        return result

    for c in candidates:
        if c.skip:
            result.skipped_duplicates += 1
            continue
        kind = c.annotation.get("kind")
        if kind == KIND_TAG:
            placed, err = _place_tag(doc, view, c)
        elif kind == KIND_KEYNOTE:
            placed, err = _place_keynote(doc, view, c)
        elif kind == KIND_TEXT_NOTE:
            placed, err = _place_text_note(doc, view, c)
        else:
            placed, err = None, "Unknown annotation kind {!r}".format(kind)

        if placed is None:
            if err:
                result.warnings.append(
                    "Skipped {} on {}: {}".format(
                        kind, c.led_label or c.led_id or "?", err
                    )
                )
            continue
        _apply_parameters(placed, c.annotation.get("parameters") or {})
        result.placed_count_by_kind[kind] = result.placed_count_by_kind.get(kind, 0) + 1
    return result
