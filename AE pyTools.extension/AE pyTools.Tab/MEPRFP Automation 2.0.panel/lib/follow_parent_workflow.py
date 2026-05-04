# -*- coding: utf-8 -*-
"""
Follow Parent — re-position placed children when their parent has moved.

For every placed child with an Element_Linker payload that points at a
parent (host or linked), compute the child's *expected* current world
position from the parent's *current* pose + the stored offset, and
move the child to match if the actual position is out of alignment.

Linked-doc children are refused — we can't write to other documents.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    Line,
    LocationPoint,
    XYZ,
)

import element_linker as _el
import element_linker_io as _el_io
import geometry
import links


POSITION_TOLERANCE_FT = 1.0 / 256.0
ROTATION_TOLERANCE_DEG = 0.01


# ---------------------------------------------------------------------
# Candidate
# ---------------------------------------------------------------------

class FollowParentCandidate(object):
    __slots__ = (
        "child", "child_id",
        "profile_id", "profile_name", "led_id", "led_label",
        "linker",
        "current_pt", "current_rot",
        "target_pt", "target_rot",
        "is_linked_child",
        "skip", "skip_reason",
    )

    def __init__(self, child, child_id, profile_id, profile_name,
                 led_id, led_label, linker,
                 current_pt, current_rot, target_pt, target_rot,
                 is_linked_child=False, skip=False, skip_reason=""):
        self.child = child
        self.child_id = child_id
        self.profile_id = profile_id
        self.profile_name = profile_name
        self.led_id = led_id
        self.led_label = led_label
        self.linker = linker
        self.current_pt = current_pt
        self.current_rot = current_rot
        self.target_pt = target_pt
        self.target_rot = target_rot
        self.is_linked_child = is_linked_child
        self.skip = skip
        self.skip_reason = skip_reason


class CollectFilters(object):
    def __init__(self, profile_ids=None, categories=None):
        self.profile_ids = set(profile_ids) if profile_ids else None
        self.categories = set(categories) if categories else None


class FollowParentResult(object):
    def __init__(self):
        self.moved_count = 0
        self.skipped_aligned = 0
        self.skipped_linked = 0
        self.skipped_no_parent = 0
        self.warnings = []


class FollowParentScanStats(object):
    """Diagnostic counters emitted alongside ``collect_candidates``.

    Tells the user which gate filtered each fixture out so they can
    distinguish between data problems (placed fixtures with no
    Element_Linker, dangling led_id, missing parent) and filter mistakes
    (wrong profile / category checked).
    """

    __slots__ = (
        "elements_scanned",
        "no_element_linker",
        "led_not_in_yaml",
        "filtered_by_profile",
        "filtered_by_category",
        "parent_unresolved",
        "no_location_point",
        "candidates_built",
        "sample_orphan_led_ids",
        "profile_matches",  # {profile_id: count} — fixtures whose led_id mapped here
    )

    def __init__(self):
        self.elements_scanned = 0
        self.no_element_linker = 0
        self.led_not_in_yaml = 0
        self.filtered_by_profile = 0
        self.filtered_by_category = 0
        self.parent_unresolved = 0
        self.no_location_point = 0
        self.candidates_built = 0
        self.sample_orphan_led_ids = []  # cap small — diagnostic only
        self.profile_matches = {}

    def summary_line(self):
        bits = ["Scanned {}".format(self.elements_scanned)]
        if self.no_element_linker:
            bits.append("no Element_Linker: {}".format(self.no_element_linker))
        if self.led_not_in_yaml:
            bits.append("led_id not in YAML: {}".format(self.led_not_in_yaml))
        if self.filtered_by_profile:
            bits.append("excluded by profile filter: {}".format(self.filtered_by_profile))
        if self.filtered_by_category:
            bits.append("excluded by category filter: {}".format(self.filtered_by_category))
        if self.parent_unresolved:
            bits.append("parent unresolved: {}".format(self.parent_unresolved))
        if self.no_location_point:
            bits.append("no LocationPoint: {}".format(self.no_location_point))
        bits.append("candidates: {}".format(self.candidates_built))
        return ";  ".join(bits)


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
    """{led_id: (profile, set, led)}."""
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


def _element_family_name(elem):
    """Family name for FamilyInstance/Group; empty string otherwise."""
    if elem is None:
        return ""
    try:
        from Autodesk.Revit.DB import FamilyInstance, Group  # noqa: E402
    except Exception:
        return ""
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        if sym is None:
            return ""
        family = getattr(sym, "Family", None)
        return getattr(family, "Name", "") if family is not None else ""
    if isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        return getattr(gtype, "Name", "") if gtype is not None else ""
    return ""


def find_parent_pose(doc, parent_element_id, host_name=None):
    """Look up parent in host doc + every linked doc. Returns
    ``(world_pt, world_rot_deg)`` or ``(None, None)`` if not found.

    ElementIds aren't globally unique across documents — a linked CAD
    parent with id 12345 and a host wall with id 12345 collide. When
    ``host_name`` is supplied (Element_Linker carries the parent's
    family name), we use it to validate the candidate lookup: a doc
    only "wins" if it produces an element whose family name matches.
    Without ``host_name`` we fall back to the legacy first-found rule.
    """
    if parent_element_id is None:
        return None, None
    try:
        eid = ElementId(int(parent_element_id))
    except Exception:
        return None, None

    target_host = (host_name or "").strip().lower()

    def _local_pose(parent):
        loc = getattr(parent, "Location", None)
        if not isinstance(loc, LocationPoint):
            return None
        try:
            return loc
        except Exception:
            return None

    def _host_doc_pose():
        parent = doc.GetElement(eid)
        if parent is None:
            return None, None, None
        loc = _local_pose(parent)
        if loc is None:
            return None, None, parent
        pt, rot = _location_pt_rot(parent)
        if pt is None:
            return None, None, parent
        return pt, rot, parent

    def _linked_doc_pose():
        for link_doc, total_transform in links.iter_link_documents(doc):
            try:
                parent = link_doc.GetElement(eid)
            except Exception:
                continue
            if parent is None:
                continue
            loc = getattr(parent, "Location", None)
            if not isinstance(loc, LocationPoint):
                continue
            try:
                local_pt = loc.Point
            except Exception:
                continue
            try:
                local_rad = loc.Rotation
            except Exception:
                local_rad = 0.0
            world_pt = total_transform.OfPoint(local_pt)
            local_x = XYZ(math.cos(local_rad), math.sin(local_rad), 0.0)
            try:
                world_x = total_transform.OfVector(local_x)
                rot = geometry.normalize_angle(
                    math.degrees(math.atan2(world_x.Y, world_x.X))
                )
            except Exception:
                rot = geometry.normalize_angle(math.degrees(local_rad))
            return (world_pt.X, world_pt.Y, world_pt.Z), rot, parent
        return None, None, None

    host_pt, host_rot, host_elem = _host_doc_pose()
    link_pt, link_rot, link_elem = _linked_doc_pose()

    if target_host:
        # Prefer the doc whose element's family name matches host_name.
        if host_elem is not None and _element_family_name(host_elem).strip().lower() == target_host:
            if host_pt is not None:
                return host_pt, host_rot
        if link_elem is not None and _element_family_name(link_elem).strip().lower() == target_host:
            if link_pt is not None:
                return link_pt, link_rot

    # Fallback: first-found wins (legacy behaviour). Linked-doc
    # placements in particular rely on this when host_name is absent.
    if host_pt is not None:
        return host_pt, host_rot
    if link_pt is not None:
        return link_pt, link_rot
    return None, None


# ---------------------------------------------------------------------
# Collection
# ---------------------------------------------------------------------

def collect_candidates(doc, profile_data, filters, refuse_linked=True, stats=None):
    """Walk every host-doc + linked-doc element with an Element_Linker
    payload, compute its target pose from its parent's current state,
    and emit a candidate if the actual pose is out of alignment.

    ``refuse_linked`` raises ``ValueError`` if any linked-doc child
    appears in the candidate set; if False, those candidates are
    skipped with a warning instead.

    If ``stats`` is a ``FollowParentScanStats`` instance, per-gate
    counters are populated so the caller can report why fixtures were
    excluded (no Element_Linker, dangling led_id, filter mismatch,
    parent unresolved, etc.).
    """
    led_index = _build_led_index(profile_data)
    out = []

    # Host children.
    for klass in (FamilyInstance, Group):
        for elem in FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType():
            if stats is not None:
                stats.elements_scanned += 1
            cand = _build_candidate(
                doc, elem, led_index, filters, is_linked=False, stats=stats
            )
            if cand is not None:
                out.append(cand)
                if stats is not None:
                    stats.candidates_built += 1

    # Linked children — refuse the run if any are picked under filters.
    if refuse_linked:
        for link_doc, _t in links.iter_link_documents(doc):
            for klass in (FamilyInstance, Group):
                try:
                    iter_collector = FilteredElementCollector(link_doc).OfClass(klass).WhereElementIsNotElementType()
                except Exception:
                    continue
                for elem in iter_collector:
                    linker = _el_io.read_from_element(elem)
                    if linker is None or not linker.led_id:
                        continue
                    entry = led_index.get(linker.led_id)
                    if entry is None:
                        continue
                    profile, _set, _led = entry
                    if filters.profile_ids and profile.get("id") not in filters.profile_ids:
                        continue
                    if filters.categories:
                        cat = (profile.get("parent_filter") or {}).get("category") or ""
                        if cat not in filters.categories:
                            continue
                    raise ValueError(
                        "Filter would include linked-doc children (e.g. element {}). "
                        "Follow Parent refuses to run when linked-doc children are in scope. "
                        "Narrow the filter to host-only profiles.".format(_id_value(elem))
                    )

    return out


def _build_candidate(doc, elem, led_index, filters, is_linked=False, stats=None):
    linker = _el_io.read_from_element(elem)
    if linker is None or not linker.led_id:
        if stats is not None:
            stats.no_element_linker += 1
        return None
    entry = led_index.get(linker.led_id)
    if entry is None:
        if stats is not None:
            stats.led_not_in_yaml += 1
            if len(stats.sample_orphan_led_ids) < 8:
                stats.sample_orphan_led_ids.append(linker.led_id)
        return None
    profile, set_dict, led = entry
    if stats is not None:
        pid = profile.get("id") or "?"
        stats.profile_matches[pid] = stats.profile_matches.get(pid, 0) + 1
    if filters.profile_ids and profile.get("id") not in filters.profile_ids:
        if stats is not None:
            stats.filtered_by_profile += 1
        return None
    if filters.categories:
        cat = (profile.get("parent_filter") or {}).get("category") or ""
        if cat not in filters.categories:
            if stats is not None:
                stats.filtered_by_category += 1
            return None

    parent_pt, parent_rot = find_parent_pose(
        doc, linker.parent_element_id, host_name=linker.host_name
    )
    if parent_pt is None:
        if stats is not None:
            stats.parent_unresolved += 1
        return None  # no parent reference -> can't follow

    offsets_list = led.get("offsets") or []
    offset = offsets_list[0] if offsets_list else {
        "x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0,
    }
    target_pt = geometry.target_point_from_offsets(parent_pt, parent_rot, offset)
    target_rot = geometry.child_rotation_from_offsets(parent_rot, offset)

    current_pt, current_rot = _location_pt_rot(elem)
    if current_pt is None:
        if stats is not None:
            stats.no_location_point += 1
        return None

    return FollowParentCandidate(
        child=elem,
        child_id=_id_value(elem),
        profile_id=profile.get("id") or "",
        profile_name=profile.get("name") or "",
        led_id=led.get("id") or "",
        led_label=led.get("label") or "",
        linker=linker,
        current_pt=current_pt,
        current_rot=current_rot,
        target_pt=target_pt,
        target_rot=target_rot,
        is_linked_child=is_linked,
    )


def is_already_aligned(c):
    dx = c.current_pt[0] - c.target_pt[0]
    dy = c.current_pt[1] - c.target_pt[1]
    dz = c.current_pt[2] - c.target_pt[2]
    pos_d = math.sqrt(dx * dx + dy * dy + dz * dz)
    rot_d = abs(geometry.normalize_angle(c.current_rot - c.target_rot))
    return pos_d < POSITION_TOLERANCE_FT and rot_d < ROTATION_TOLERANCE_DEG


def mark_aligned_skips(candidates):
    for c in candidates:
        if is_already_aligned(c):
            c.skip = True
            c.skip_reason = "already aligned"


# ---------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------

def execute_follow(doc, candidates):
    """Move every non-skipped candidate to its target pose, then update
    the Element_Linker payload. Caller manages the transaction."""
    result = FollowParentResult()

    for c in candidates:
        if c.skip:
            if c.skip_reason == "already aligned":
                result.skipped_aligned += 1
            else:
                result.skipped_no_parent += 1
            continue

        try:
            cur = XYZ(*c.current_pt)
            tgt = XYZ(*c.target_pt)
            translation = tgt - cur
            if translation.GetLength() > POSITION_TOLERANCE_FT:
                ElementTransformUtils.MoveElement(doc, c.child.Id, translation)
            rot_delta = geometry.normalize_angle(c.target_rot - c.current_rot)
            if abs(rot_delta) > ROTATION_TOLERANCE_DEG:
                axis = Line.CreateBound(tgt, XYZ(tgt.X, tgt.Y, tgt.Z + 1.0))
                ElementTransformUtils.RotateElement(
                    doc, c.child.Id, axis, math.radians(rot_delta)
                )
        except Exception as exc:
            result.warnings.append(
                "Move failed for child id {}: {}".format(c.child_id, exc)
            )
            continue

        # Update Element_Linker JSON with new pose.
        try:
            new_linker = _el.ElementLinker(
                led_id=c.linker.led_id,
                set_id=c.linker.set_id,
                location_ft=list(c.target_pt),
                rotation_deg=c.target_rot,
                parent_rotation_deg=c.linker.parent_rotation_deg,
                parent_element_id=c.linker.parent_element_id,
                level_id=c.linker.level_id,
                element_id=c.linker.element_id,
                facing=c.linker.facing,
                host_name=c.linker.host_name,
                parent_location_ft=c.linker.parent_location_ft,
            )
            _el_io.write_to_element(c.child, new_linker)
        except Exception as exc:
            result.warnings.append(
                "Element_Linker rewrite failed for child id {}: {}".format(
                    c.child_id, exc
                )
            )
        result.moved_count += 1
    return result
