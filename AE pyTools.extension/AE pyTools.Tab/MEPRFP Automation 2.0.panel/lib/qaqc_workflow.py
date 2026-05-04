# -*- coding: utf-8 -*-
"""
QAQC — health-check the relationship between the YAML store and the
elements actually placed in the model.

Six categories::

    A  Orphan profile             Profile in YAML with zero placed
                                  instances anywhere in the doc.
                                  Skips profiles that are merged into
                                  another (legacy ced_truth_source_id
                                  members or profiles whose name is in
                                  another profile's merged_aliases).
    B  Missing parent             Placed child references a
                                  parent_element_id that no longer
                                  exists in the host or linked docs.
    C  Far from parent            Placed child sits more than
                                  Tolerances.FAR_FROM_PARENT_FT from
                                  the pose its stored offset implies.
    D  ID discrepancy             Element_Linker.element_id does not
                                  match the live element's Id.Value.
                                  Usually a sign of a copy/paste.
    E  Parent type change         Parent's family/type doesn't match
                                  the profile's parent_filter AND isn't
                                  in the profile's merged_aliases —
                                  i.e. a real parent reassignment, not
                                  a known alias.
    F  Host-name mismatch         parent's current family name does
                                  not match Element_Linker.host_name,
                                  ignoring trailing `` : Default`` /
                                  `` : Default <n>`` decoration.

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
)

import element_linker as _el
import element_linker_io as _el_io
import follow_parent_workflow as _fp
import geometry
import links
import truth_groups


# ---------------------------------------------------------------------
# Category constants
# ---------------------------------------------------------------------

CAT_A = "A"
CAT_B = "B"
CAT_C = "C"
CAT_D = "D"
CAT_E = "E"
CAT_F = "F"

CAT_ALL = (CAT_A, CAT_B, CAT_C, CAT_D, CAT_E, CAT_F)

CAT_LABELS = {
    CAT_A: "A  Orphan profile",
    CAT_B: "B  Missing parent",
    CAT_C: "C  Far from parent",
    CAT_D: "D  ID discrepancy",
    CAT_E: "E  Parent type change",
    CAT_F: "F  Host-name mismatch",
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
FIX_FOLLOW_PARENT = "follow_parent"     # cat D
FIX_REFRESH_ELEMENT_ID = "refresh_eid"  # cat E
FIX_REFRESH_HOST_NAME = "refresh_host"  # cat G
FIX_CLEAR_LINKER = "clear_linker"       # cat C (destructive)


# ---------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------

class QaqcFinding(object):
    __slots__ = (
        "category",
        "category_label",
        "element_id",       # int — the host-doc element to select/zoom (None for cat A)
        "profile_id",
        "profile_name",
        "led_id",
        "led_label",
        "message",
        "fix_kind",
        "fix_payload",      # dict; varies by fix_kind
    )

    def __init__(self, category, element_id, profile_id, profile_name,
                 led_id, led_label, message, fix_kind=FIX_NONE,
                 fix_payload=None):
        self.category = category
        self.category_label = CAT_LABELS.get(category, category)
        self.element_id = element_id
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


# ---------------------------------------------------------------------
# Audit
# ---------------------------------------------------------------------

def run_audit(doc, profile_data, categories=None):
    """Walk the doc + profile_data and emit a ``QaqcResult``.

    ``categories`` can restrict the run to a subset (set of CAT_*); None
    means all seven.
    """
    requested = set(categories) if categories else set(CAT_ALL)
    led_index = _build_led_index(profile_data)
    result = QaqcResult()

    profiles_by_id = {
        p.get("id"): p for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict) and p.get("id")
    }

    # Track profile usage during the placed-child sweep so we can emit
    # cat A at the end.
    profile_usage = {pid: 0 for pid in profiles_by_id}

    # Track which parent element ids are referenced so we can emit cat B.
    referenced_parent_ids = set()

    # 1. Walk every placed child with an Element_Linker payload.
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
                referenced_parent_ids.add(int(linker.parent_element_id))

            # Cat D: ID discrepancy.
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

            # Parent-related checks (C, D, F, G) all need the parent.
            parent = None
            if linker.parent_element_id is not None:
                try:
                    parent = doc.GetElement(ElementId(int(linker.parent_element_id)))
                except Exception:
                    parent = None

            parent_in_link = False
            if parent is None and linker.parent_element_id is not None:
                # Check linked docs.
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
                            parent_in_link = True
                            break

            if parent is None:
                if CAT_B in requested and linker.parent_element_id is not None:
                    result.add(QaqcFinding(
                        category=CAT_B,
                        element_id=elem_id_val,
                        profile_id=profile_id,
                        profile_name=profile_name,
                        led_id=led_id,
                        led_label=led_label,
                        message="parent_element_id={} not found in host or linked docs".format(
                            linker.parent_element_id
                        ),
                        fix_kind=FIX_CLEAR_LINKER,
                    ))
                continue

            # Cat E: parent's family no longer matches profile's filter
            # AND isn't a known alias on the profile. The merged_aliases
            # check keeps pre-merge family names (e.g. ``Stinger Cart_2``
            # absorbed into ``Stinger Cart_1``) from being flagged as a
            # type change when the placement against the absorbed family
            # is still valid.
            if CAT_E in requested:
                expected_family = _profile_family_name(profile)
                actual_family = _element_family_name(parent)
                if expected_family and actual_family:
                    actual_lower = actual_family.lower()
                    known = _profile_known_family_names(profile)
                    if actual_lower not in known and \
                            expected_family.lower() != actual_lower:
                        result.add(QaqcFinding(
                            category=CAT_E,
                            element_id=elem_id_val,
                            profile_id=profile_id,
                            profile_name=profile_name,
                            led_id=led_id,
                            led_label=led_label,
                            message="Parent family is '{}', profile expects '{}'".format(
                                actual_family, expected_family
                            ),
                            fix_kind=FIX_NONE,
                        ))

            # Cat F: host_name mismatch — ignoring trailing
            # `` : Default`` / `` : Default <n>`` decoration that Revit
            # auto-generates when a family has only its default type.
            if CAT_F in requested:
                stored_host = (linker.host_name or "").strip()
                actual_family = _element_family_name(parent)
                if stored_host and actual_family:
                    stored_norm = _strip_default_decoration(stored_host).lower()
                    actual_norm = _strip_default_decoration(actual_family).lower()
                    if stored_norm and actual_norm and stored_norm != actual_norm:
                        result.add(QaqcFinding(
                            category=CAT_F,
                            element_id=elem_id_val,
                            profile_id=profile_id,
                            profile_name=profile_name,
                            led_id=led_id,
                            led_label=led_label,
                            message="host_name='{}', actual parent family='{}'".format(
                                stored_host, actual_family
                            ),
                            fix_kind=FIX_REFRESH_HOST_NAME,
                            fix_payload={"new_host_name": actual_family},
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

    # 2. Cat A: orphan profiles. Skips profiles that are conceptually
    # merged into another (legacy ced_truth_source_id members or
    # profiles whose ``name`` lives in another profile's
    # merged_aliases) — those are unused on purpose.
    if CAT_A in requested:
        for pid, p in profiles_by_id.items():
            if profile_usage.get(pid, 0) > 0:
                continue
            if _is_merged_member(p, profile_data):
                continue
            result.add(QaqcFinding(
                category=CAT_A,
                element_id=None,
                profile_id=pid,
                profile_name=p.get("name") or "",
                led_id="",
                led_label="",
                message="Profile has no placed instances anywhere",
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
