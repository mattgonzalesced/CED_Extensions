# -*- coding: utf-8 -*-
"""
Update Vector — recapture the offsets stored on a LED (or its
annotations) from the *current* positions of placed children.

Inverse of Follow Parent: there, the parent moves and we re-position
the child to match the stored offset. Here, the child has been moved
manually and we recompute the stored offset from the new pose.

Conflict policy: when multiple instances of the same LED have moved
differently, *last-moved wins* — the highest-instance-id (a stable
proxy for "most recently created/edited") overwrites the LED. A
warning is emitted any time another instance disagrees by more than
1/256 ft / 0.01°.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    FamilyInstance,
    Group,
    LocationPoint,
    XYZ,
)

import element_linker as _el
import element_linker_io as _el_io
import follow_parent_workflow  # for find_parent_pose + helpers
import geometry
import hosted_annotations


Z_TOLERANCE_FT = 1.0 / 256.0
DIVERGENCE_POSITION_FT = 1.0 / 256.0
DIVERGENCE_ROTATION_DEG = 0.01


# ---------------------------------------------------------------------
# Candidate (one row per LED-or-annotation that wants updating)
# ---------------------------------------------------------------------

class UpdateVectorCandidate(object):
    __slots__ = (
        "kind",  # 'led' or 'annotation'
        "led",
        "annotation",        # only set when kind=='annotation'
        "led_id",
        "ann_id",
        "old_offset",
        "new_offset",
        "source_element_id", # the element whose pose drove this candidate
        "diverged_others",   # list[(other_elem_id, deviation_str)]
        "skip",
        "skip_reason",
    )

    def __init__(self, kind, led, annotation, led_id, ann_id,
                 old_offset, new_offset, source_element_id,
                 diverged_others=None, skip=False, skip_reason=""):
        self.kind = kind
        self.led = led
        self.annotation = annotation
        self.led_id = led_id
        self.ann_id = ann_id
        self.old_offset = old_offset
        self.new_offset = new_offset
        self.source_element_id = source_element_id
        self.diverged_others = diverged_others or []
        self.skip = skip
        self.skip_reason = skip_reason


class UpdateVectorResult(object):
    def __init__(self):
        self.led_updates = 0
        self.ann_updates = 0
        self.warnings = []


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
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


def _offsets_diverge(a, b):
    return (
        abs(a.get("x_inches", 0) - b.get("x_inches", 0)) > DIVERGENCE_POSITION_FT * 12
        or abs(a.get("y_inches", 0) - b.get("y_inches", 0)) > DIVERGENCE_POSITION_FT * 12
        or abs(a.get("z_inches", 0) - b.get("z_inches", 0)) > DIVERGENCE_POSITION_FT * 12
        or abs(a.get("rotation_deg", 0) - b.get("rotation_deg", 0)) > DIVERGENCE_ROTATION_DEG
    )


def _preserve_z_if_within_tolerance(old_offset, new_offset):
    """Legacy behaviour: if Z movement is below tolerance, retain the
    old Z so re-runs don't accidentally re-baseline elevations."""
    old_z = old_offset.get("z_inches", 0.0)
    new_z = new_offset.get("z_inches", 0.0)
    if abs(old_z - new_z) <= Z_TOLERANCE_FT * 12.0:
        new_offset["z_inches"] = old_z
    return new_offset


# ---------------------------------------------------------------------
# Build candidates from a Revit selection
# ---------------------------------------------------------------------

def collect_candidates_from_selection(doc, profile_data, element_ids):
    """Inputs are ElementIds the user selected. Walk them, classify
    each as fixture / annotation, group by LED / ANN id, compute new
    offsets, detect divergence."""
    led_index = _build_led_index(profile_data)
    fixture_pose_by_led = {}  # led_id -> [(elem_id, pose, parent_pose, linker), ...]
    annotation_pose_by_ann = {}  # (led_id, ann_id) -> [(elem_id, ann_world_pt, ann_world_rot, fixture_world_pt, fixture_world_rot)]

    # 1. Walk picked elements; classify.
    for eid in element_ids:
        elem = doc.GetElement(eid)
        if elem is None:
            continue

        # Annotation element? (IndependentTag, TextNote, or keynote symbol family instance)
        kind = hosted_annotations.annotation_kind(elem)
        if kind is not None:
            # Need to find which LED + ANN this annotation corresponds to.
            # Heuristic: look at every LED's annotations for one whose family/type
            # matches AND whose target point is close to this element's current point.
            entry = _match_annotation_to_led(doc, elem, kind, led_index)
            if entry is None:
                continue
            led_dict, ann_dict, fixture_world_pt, fixture_world_rot = entry
            ann_world_pt = hosted_annotations.annotation_world_point(elem)
            ann_world_rot = hosted_annotations.annotation_rotation_deg(elem)
            if ann_world_pt is None:
                continue
            key = (led_dict.get("id"), ann_dict.get("id"))
            annotation_pose_by_ann.setdefault(key, []).append(
                (_id_value(elem), ann_world_pt, ann_world_rot,
                 fixture_world_pt, fixture_world_rot, ann_dict, led_dict)
            )
            continue

        # Fixture? Only FamilyInstance / Group with Element_Linker.
        if not isinstance(elem, (FamilyInstance, Group)):
            continue
        linker = _el_io.read_from_element(elem)
        if linker is None or not linker.led_id:
            continue
        if linker.led_id not in led_index:
            continue
        pose_pt, pose_rot = _location_pt_rot(elem)
        if pose_pt is None:
            continue
        parent_pt, parent_rot = follow_parent_workflow.find_parent_pose(
            doc, linker.parent_element_id, host_name=linker.host_name
        )
        if parent_pt is None:
            continue
        fixture_pose_by_led.setdefault(linker.led_id, []).append(
            (_id_value(elem), pose_pt, pose_rot, parent_pt, parent_rot, linker)
        )

    candidates = []

    # 2. Build fixture candidates.
    for led_id, entries in fixture_pose_by_led.items():
        led = led_index[led_id][2]
        # Sort by element_id descending so "highest id" approximates last-moved.
        entries.sort(key=lambda x: x[0] or 0, reverse=True)
        winner = entries[0]
        elem_id, child_pt, child_rot, parent_pt, parent_rot, linker = winner

        new_offset = geometry.compute_offsets_from_points(
            parent_pt, parent_rot, child_pt, child_rot
        )
        old_offsets = led.get("offsets") or []
        old_offset = old_offsets[0] if old_offsets else {
            "x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0,
        }
        new_offset = _preserve_z_if_within_tolerance(old_offset, dict(new_offset))

        # Divergence detection across other instances.
        diverged = []
        for other_elem_id, o_pt, o_rot, op_pt, op_rot, o_linker in entries[1:]:
            o_off = geometry.compute_offsets_from_points(op_pt, op_rot, o_pt, o_rot)
            if _offsets_diverge(o_off, new_offset):
                diverged.append((
                    other_elem_id,
                    "Δ x={:.2f} y={:.2f} z={:.2f} rot={:.1f}°".format(
                        o_off["x_inches"] - new_offset["x_inches"],
                        o_off["y_inches"] - new_offset["y_inches"],
                        o_off["z_inches"] - new_offset["z_inches"],
                        o_off["rotation_deg"] - new_offset["rotation_deg"],
                    ),
                ))

        candidates.append(UpdateVectorCandidate(
            kind="led",
            led=led,
            annotation=None,
            led_id=led_id,
            ann_id="",
            old_offset=old_offset,
            new_offset=new_offset,
            source_element_id=elem_id,
            diverged_others=diverged,
        ))

    # 3. Build annotation candidates.
    for (led_id, ann_id), entries in annotation_pose_by_ann.items():
        entries.sort(key=lambda x: x[0] or 0, reverse=True)
        elem_id, ann_pt, ann_rot, fix_pt, fix_rot, ann_dict, led_dict = entries[0]
        new_offset = geometry.compute_offsets_from_points(
            fix_pt, fix_rot, ann_pt, ann_rot
        )
        old_offset = ann_dict.get("offsets") or {}
        if isinstance(old_offset, list):
            old_offset = old_offset[0] if old_offset else {}
        new_offset = _preserve_z_if_within_tolerance(
            dict(old_offset), dict(new_offset)
        )

        diverged = []
        for other in entries[1:]:
            o_elem, o_ann, o_rot2, o_fix, o_fix_rot, _, _ = other
            o_off = geometry.compute_offsets_from_points(
                o_fix, o_fix_rot, o_ann, o_rot2
            )
            if _offsets_diverge(o_off, new_offset):
                diverged.append((
                    o_elem,
                    "Δ x={:.2f} y={:.2f} rot={:.1f}°".format(
                        o_off["x_inches"] - new_offset["x_inches"],
                        o_off["y_inches"] - new_offset["y_inches"],
                        o_off["rotation_deg"] - new_offset["rotation_deg"],
                    ),
                ))

        candidates.append(UpdateVectorCandidate(
            kind="annotation",
            led=led_dict,
            annotation=ann_dict,
            led_id=led_id,
            ann_id=ann_id,
            old_offset=dict(old_offset),
            new_offset=new_offset,
            source_element_id=elem_id,
            diverged_others=diverged,
        ))

    return candidates


def _match_annotation_to_led(doc, ann_elem, kind, led_index):
    """Best-effort: find the LED whose annotation list contains an entry
    matching this annotation's family/type, AND whose host fixture is
    near this annotation's current world point.

    Returns ``(led_dict, ann_dict, fixture_world_pt, fixture_world_rot)``
    or None.
    """
    ann_pt = hosted_annotations.annotation_world_point(ann_elem)
    if ann_pt is None:
        return None

    # Pull family/type for matching.
    type_id = ann_elem.GetTypeId()
    type_elem = doc.GetElement(type_id) if type_id else None
    target_family = getattr(type_elem, "FamilyName", "") if type_elem else ""
    target_type = getattr(type_elem, "Name", "") if type_elem else ""

    best = None
    best_dist = float("inf")

    # Walk every fixture in the doc; for each, look at its annotations
    # under the LED, and check distance.
    from Autodesk.Revit.DB import FilteredElementCollector
    for klass in (FamilyInstance, Group):
        for fixture in FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType():
            linker = _el_io.read_from_element(fixture)
            if linker is None or not linker.led_id:
                continue
            entry = led_index.get(linker.led_id)
            if entry is None:
                continue
            led = entry[2]
            fix_pt, fix_rot = _location_pt_rot(fixture)
            if fix_pt is None:
                continue
            for ann_dict in led.get("annotations") or []:
                if not isinstance(ann_dict, dict):
                    continue
                if ann_dict.get("kind") != kind:
                    continue
                fam_match = (ann_dict.get("family_name") or "") == target_family or not target_family
                typ_match = (ann_dict.get("type_name") or "") == target_type or not target_type
                if not (fam_match and typ_match):
                    continue
                # Distance from current ann point to expected ann point
                # (computed from fixture pose + stored offset).
                offsets = ann_dict.get("offsets") or {}
                if isinstance(offsets, list):
                    offsets = offsets[0] if offsets else {}
                expected = geometry.target_point_from_offsets(
                    fix_pt, fix_rot, offsets
                )
                d2 = (
                    (ann_pt[0] - expected[0]) ** 2
                    + (ann_pt[1] - expected[1]) ** 2
                    + (ann_pt[2] - expected[2]) ** 2
                )
                if d2 < best_dist:
                    best_dist = d2
                    best = (led, ann_dict, fix_pt, fix_rot)

    if best is None:
        return None
    # Cap the search radius — further than 100 ft is almost certainly a
    # false match.
    if best_dist > (100.0 * 100.0):
        return None
    return best


# ---------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------

def execute_update(doc, candidates):
    """Apply each non-skipped candidate's new offset to the YAML AND
    rewrite the placed child's Element_Linker payload (for fixtures)
    so future Follow Parent runs see the new offset. Caller manages
    the transaction."""
    result = UpdateVectorResult()

    for c in candidates:
        if c.skip:
            continue
        if c.kind == "led":
            offsets_list = c.led.get("offsets")
            if not isinstance(offsets_list, list):
                offsets_list = []
                c.led["offsets"] = offsets_list
            if not offsets_list:
                offsets_list.append({})
            offsets_list[0].update(c.new_offset)
            result.led_updates += 1

            # Also update the moved fixture's Element_Linker so its own
            # cached pose is fresh.
            try:
                source_elem = doc.GetElement(ElementId(int(c.source_element_id)))
            except Exception:
                source_elem = None
            if source_elem is not None:
                cur_pt, cur_rot = _location_pt_rot(source_elem)
                if cur_pt is not None:
                    linker = _el_io.read_from_element(source_elem)
                    if linker is not None:
                        new_linker = _el.ElementLinker(
                            led_id=linker.led_id,
                            set_id=linker.set_id,
                            location_ft=list(cur_pt),
                            rotation_deg=cur_rot,
                            parent_rotation_deg=linker.parent_rotation_deg,
                            parent_element_id=linker.parent_element_id,
                            level_id=linker.level_id,
                            element_id=linker.element_id,
                            facing=linker.facing,
                            host_name=linker.host_name,
                            parent_location_ft=linker.parent_location_ft,
                        )
                        try:
                            _el_io.write_to_element(source_elem, new_linker)
                        except Exception:
                            pass

            for other_id, msg in c.diverged_others:
                result.warnings.append(
                    "LED {}: instance {} diverged from chosen offset ({}). "
                    "Last-moved wins.".format(c.led_id, other_id, msg)
                )

        elif c.kind == "annotation":
            ann = c.annotation
            ann["offsets"] = dict(c.new_offset)
            result.ann_updates += 1
            for other_id, msg in c.diverged_others:
                result.warnings.append(
                    "ANN {}: instance {} diverged ({}). Last-moved wins.".format(
                        c.ann_id, other_id, msg
                    )
                )
    return result
