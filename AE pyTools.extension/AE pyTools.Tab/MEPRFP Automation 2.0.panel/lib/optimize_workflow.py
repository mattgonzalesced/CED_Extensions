# -*- coding: utf-8 -*-
"""
Optimize — snap placed children to host structural features.

For every placed child with an Element_Linker payload, find the nearest
host element of the configured kind within a search radius and compute
an "optimized" pose:

    Wall          project child XY onto the nearest wall, rotate to face
                  along the wall normal, Z preserved.
    Ceiling       lift child Z to the nearest ceiling's bottom face.
    Floor         drop child Z to the nearest floor's top face.
    Door-relative re-parent the child onto the nearest door — the child
                  stays in place but Element_Linker.parent_element_id
                  flips to the door, with parent_location_ft /
                  parent_rotation_deg refreshed from the door's pose.

Mode is configured per LED label (family:type), so a single profile can
mix wall-mounted and ceiling-mounted fixtures and each follows its own
rule.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    Ceiling,
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FilteredElementCollector,
    Floor,
    Group,
    Line,
    LocationCurve,
    LocationPoint,
    Wall,
    XYZ,
)

import element_linker as _el
import element_linker_io as _el_io
import geometry
import links


DEFAULT_SEARCH_RADIUS_FT = 5.0
POSITION_TOLERANCE_FT = 1.0 / 256.0
ROTATION_TOLERANCE_DEG = 0.01


MODE_NONE = "none"
MODE_WALL = "wall"
MODE_CEILING = "ceiling"
MODE_FLOOR = "floor"
MODE_DOOR = "door_relative"

ALL_MODES = (MODE_NONE, MODE_WALL, MODE_CEILING, MODE_FLOOR, MODE_DOOR)

MODE_LABELS = {
    MODE_NONE: "(skip)",
    MODE_WALL: "Wall",
    MODE_CEILING: "Ceiling",
    MODE_FLOOR: "Floor",
    MODE_DOOR: "Door-relative",
}


def default_mode_for_label(label):
    """Pick a sensible default mode by inspecting the family:type label.

    Rules (case-insensitive substring match, first hit wins):
        ``cord drop`` -> ``ceiling``
        ``floor``     -> ``floor``
        ``wall``      -> ``wall``
        otherwise     -> ``wall`` (default)

    The user can override per row in the dialog before clicking Match.
    """
    lower = (label or "").lower()
    if "cord drop" in lower:
        return MODE_CEILING
    if "floor" in lower:
        return MODE_FLOOR
    if "wall" in lower:
        return MODE_WALL
    return MODE_WALL


# ---------------------------------------------------------------------
# Options + result objects
# ---------------------------------------------------------------------

class OptimizeOptions(object):
    def __init__(self, search_radius_ft=DEFAULT_SEARCH_RADIUS_FT,
                 mode_by_family_type=None):
        self.search_radius_ft = float(search_radius_ft)
        self.mode_by_family_type = dict(mode_by_family_type or {})


class OptimizeCandidate(object):
    __slots__ = (
        "child", "child_id",
        "profile_id", "profile_name", "led_id", "led_label",
        "linker", "mode",
        "current_pt", "current_rot",
        "target_pt", "target_rot",
        "host_element_id", "host_description",
        "skip", "skip_reason",
    )

    def __init__(self, child, child_id, profile_id, profile_name,
                 led_id, led_label, linker, mode,
                 current_pt, current_rot, target_pt, target_rot,
                 host_element_id=None, host_description="",
                 skip=False, skip_reason=""):
        self.child = child
        self.child_id = child_id
        self.profile_id = profile_id
        self.profile_name = profile_name
        self.led_id = led_id
        self.led_label = led_label
        self.linker = linker
        self.mode = mode
        self.current_pt = current_pt
        self.current_rot = current_rot
        self.target_pt = target_pt
        self.target_rot = target_rot
        self.host_element_id = host_element_id
        self.host_description = host_description
        self.skip = skip
        self.skip_reason = skip_reason


class OptimizeResult(object):
    def __init__(self):
        self.moved_count = 0
        self.reparented_count = 0
        self.skipped_no_host = 0
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


def collect_family_type_labels(doc, profile_data, selected_element_ids=None):
    """Unique LED labels (family:type strings) — one row per label in
    the per-family:type mode grid.

    When ``selected_element_ids`` is non-empty, only labels of LEDs that
    the selected placed elements link to are returned. That gives the
    user a focused grid when they pre-select a fixture before running
    Optimize.
    """
    if selected_element_ids:
        labels = set()
        led_index = _build_led_index(profile_data)
        for eid in selected_element_ids:
            elem = doc.GetElement(eid)
            if elem is None:
                continue
            linker = _el_io.read_from_element(elem)
            if linker is None or not linker.led_id:
                continue
            entry = led_index.get(linker.led_id)
            if entry is None:
                continue
            led = entry[2]
            label = (led.get("label") or "").strip()
            if label:
                labels.add(label)
        return sorted(labels)

    out = set()
    for profile in profile_data.get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if isinstance(led, dict):
                    label = (led.get("label") or "").strip()
                    if label:
                        out.add(label)
    return sorted(out)


# ---------------------------------------------------------------------
# Nearest-host searches
# ---------------------------------------------------------------------

def _find_nearest_wall(doc, point, radius):
    """Find the nearest wall in any **linked** document within ``radius``.

    Walks every linked Revit doc reachable from ``doc`` and projects the
    candidate point onto each wall's centerline curve (transformed into
    the host's world frame). Returns ``(wall_elem, proj_result, curve_world)``
    or ``None``. The host doc's own walls are ignored — host walls are
    typically our own placement geometry, not the structural source we
    want to snap against.

    ``proj_result`` is a Revit ``IntersectionResult`` from
    ``Curve.Project``; its ``XYZPoint`` is the nearest point on the
    centerline and ``Distance`` is the perpendicular distance, both in
    world coordinates.
    """
    cx, cy, cz = point
    pt_xyz = XYZ(cx, cy, cz)
    best = None
    best_dist = float("inf")
    for link_doc, total_transform in links.iter_link_documents(doc):
        try:
            walls = (
                FilteredElementCollector(link_doc)
                .OfClass(Wall)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for wall in walls:
            loc = getattr(wall, "Location", None)
            if not isinstance(loc, LocationCurve):
                continue
            try:
                local_curve = loc.Curve
                world_curve = local_curve.CreateTransformed(total_transform)
            except Exception:
                continue
            try:
                proj = world_curve.Project(pt_xyz)
            except Exception:
                continue
            if proj is None:
                continue
            d = proj.Distance
            if d is None or d > radius:
                continue
            if d < best_dist:
                best_dist = d
                best = (wall, proj, world_curve)
    return best


def _find_nearest_ceiling(doc, point, radius):
    """Return ``(ceiling_elem, bottom_face_z)`` or None. Only ceilings
    *above* the candidate Z are considered."""
    cx, cy, cz = point
    best = None
    best_dist = float("inf")
    for ceil in FilteredElementCollector(doc).OfClass(Ceiling).WhereElementIsNotElementType():
        bbox = ceil.get_BoundingBox(None)
        if bbox is None:
            continue
        if bbox.Min.Z < cz - POSITION_TOLERANCE_FT:
            continue  # ceiling must be at or above the child
        dx = max(0.0, max(bbox.Min.X - cx, cx - bbox.Max.X))
        dy = max(0.0, max(bbox.Min.Y - cy, cy - bbox.Max.Y))
        dz = max(0.0, bbox.Min.Z - cz)
        d = math.sqrt(dx * dx + dy * dy + dz * dz)
        if d > radius:
            continue
        if d < best_dist:
            best_dist = d
            best = (ceil, bbox.Min.Z)
    return best


def _find_nearest_floor(doc, point, radius):
    """Return ``(floor_elem, top_face_z)`` or None. Only floors at or
    below the candidate Z are considered."""
    cx, cy, cz = point
    best = None
    best_dist = float("inf")
    for fl in FilteredElementCollector(doc).OfClass(Floor).WhereElementIsNotElementType():
        bbox = fl.get_BoundingBox(None)
        if bbox is None:
            continue
        if bbox.Max.Z > cz + POSITION_TOLERANCE_FT:
            continue
        dx = max(0.0, max(bbox.Min.X - cx, cx - bbox.Max.X))
        dy = max(0.0, max(bbox.Min.Y - cy, cy - bbox.Max.Y))
        dz = max(0.0, cz - bbox.Max.Z)
        d = math.sqrt(dx * dx + dy * dy + dz * dz)
        if d > radius:
            continue
        if d < best_dist:
            best_dist = d
            best = (fl, bbox.Max.Z)
    return best


def _find_nearest_door(doc, point, radius):
    """Return ``(door_elem, door_world_pt, door_rotation_deg)`` or None."""
    cx, cy, cz = point
    best = None
    best_dist = float("inf")
    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Doors)
        .OfClass(FamilyInstance)
        .WhereElementIsNotElementType()
    )
    for door in collector:
        loc = getattr(door, "Location", None)
        if not isinstance(loc, LocationPoint):
            continue
        try:
            pt = loc.Point
        except Exception:
            continue
        d = math.sqrt(
            (pt.X - cx) ** 2 + (pt.Y - cy) ** 2 + (pt.Z - cz) ** 2
        )
        if d > radius:
            continue
        if d < best_dist:
            best_dist = d
            try:
                rad = loc.Rotation
            except Exception:
                rad = 0.0
            best = (
                door,
                (pt.X, pt.Y, pt.Z),
                geometry.normalize_angle(math.degrees(rad)),
            )
    return best


# ---------------------------------------------------------------------
# Candidate construction
# ---------------------------------------------------------------------

def collect_candidates(doc, profile_data, options, selected_element_ids=None):
    """Walk every placed child with an Element_Linker payload and build
    one ``OptimizeCandidate`` per element whose family:type has a mode
    other than NONE configured.

    When ``selected_element_ids`` is non-empty, only those elements are
    considered (regardless of category) — used by the dialog when the
    user pre-selects fixtures before opening Optimize.
    """
    led_index = _build_led_index(profile_data)
    out = []

    if selected_element_ids:
        elements_iter = []
        for eid in selected_element_ids:
            elem = doc.GetElement(eid)
            if elem is not None and isinstance(elem, (FamilyInstance, Group)):
                elements_iter.append(elem)
    else:
        elements_iter = []
        for klass in (FamilyInstance, Group):
            elements_iter.extend(
                FilteredElementCollector(doc)
                .OfClass(klass)
                .WhereElementIsNotElementType()
            )

    for elem in elements_iter:
        linker = _el_io.read_from_element(elem)
        if linker is None or not linker.led_id:
            continue
        entry = led_index.get(linker.led_id)
        if entry is None:
            continue
        profile, _set_dict, led = entry
        label = (led.get("label") or "").strip()
        mode = options.mode_by_family_type.get(label, MODE_NONE)
        if mode == MODE_NONE:
            continue
        cur_pt, cur_rot = _location_pt_rot(elem)
        if cur_pt is None:
            continue

        target_pt = cur_pt
        target_rot = cur_rot
        host_elem_id = None
        host_desc = ""
        skip = False
        skip_reason = ""

        if mode == MODE_WALL:
            hit = _find_nearest_wall(doc, cur_pt, options.search_radius_ft)
            if hit is None:
                skip, skip_reason = True, "no linked wall in radius"
            else:
                wall, proj, curve = hit
                p = proj.XYZPoint
                target_pt = (p.X, p.Y, cur_pt[2])
                # Facing rule: the fixture's *facing direction* must end
                # up perpendicular to the wall length and pointing away
                # from the wall, on the same side the fixture currently
                # is. We compute the desired facing direction in world
                # coords, then derive the rotation delta from the live
                # element's current FacingOrientation. This is family-
                # agnostic — works regardless of which local axis the
                # authoring family uses for "front".
                try:
                    deriv = curve.ComputeDerivatives(proj.Parameter, False)
                    tangent = deriv.BasisX
                    wall_angle = geometry.normalize_angle(
                        math.degrees(math.atan2(tangent.Y, tangent.X))
                    )
                    outward_x = cur_pt[0] - p.X
                    outward_y = cur_pt[1] - p.Y
                    if (outward_x * outward_x + outward_y * outward_y) < 1e-12:
                        # Fixture sits exactly on the centerline — keep
                        # current rotation.
                        target_rot = cur_rot
                    else:
                        outward_angle_raw = geometry.normalize_angle(
                            math.degrees(math.atan2(outward_y, outward_x))
                        )
                        option_a = geometry.normalize_angle(wall_angle + 90.0)
                        option_b = geometry.normalize_angle(wall_angle - 90.0)
                        diff_a = abs(geometry.normalize_angle(option_a - outward_angle_raw))
                        diff_b = abs(geometry.normalize_angle(option_b - outward_angle_raw))
                        target_facing_angle = option_a if diff_a <= diff_b else option_b

                        facing = getattr(elem, "FacingOrientation", None)
                        if facing is not None and (
                            facing.X * facing.X + facing.Y * facing.Y
                        ) > 1e-9:
                            cur_facing_angle = geometry.normalize_angle(
                                math.degrees(math.atan2(facing.Y, facing.X))
                            )
                            delta = geometry.normalize_angle(
                                target_facing_angle - cur_facing_angle
                            )
                            target_rot = geometry.normalize_angle(cur_rot + delta)
                        else:
                            # No usable FacingOrientation (e.g. Group
                            # instance) — assume rotation == facing angle.
                            target_rot = target_facing_angle
                except Exception:
                    target_rot = cur_rot
                host_elem_id = _id_value(wall)
                host_desc = "Linked wall {}".format(host_elem_id)

        elif mode == MODE_CEILING:
            hit = _find_nearest_ceiling(doc, cur_pt, options.search_radius_ft)
            if hit is None:
                skip, skip_reason = True, "no ceiling above in radius"
            else:
                ceil, ceil_z = hit
                target_pt = (cur_pt[0], cur_pt[1], ceil_z)
                host_elem_id = _id_value(ceil)
                host_desc = "Ceiling {}".format(host_elem_id)

        elif mode == MODE_FLOOR:
            hit = _find_nearest_floor(doc, cur_pt, options.search_radius_ft)
            if hit is None:
                skip, skip_reason = True, "no floor below in radius"
            else:
                fl, floor_z = hit
                target_pt = (cur_pt[0], cur_pt[1], floor_z)
                host_elem_id = _id_value(fl)
                host_desc = "Floor {}".format(host_elem_id)

        elif mode == MODE_DOOR:
            hit = _find_nearest_door(doc, cur_pt, options.search_radius_ft)
            if hit is None:
                skip, skip_reason = True, "no door in radius"
            else:
                door, door_pt, door_rot = hit
                # Keep the child where it is, just re-parent.
                target_pt = cur_pt
                target_rot = cur_rot
                host_elem_id = _id_value(door)
                host_desc = "Door {}".format(host_elem_id)

        else:
            continue

        out.append(OptimizeCandidate(
            child=elem,
            child_id=_id_value(elem),
            profile_id=profile.get("id") or "",
            profile_name=profile.get("name") or "",
            led_id=led.get("id") or "",
            led_label=label,
            linker=linker,
            mode=mode,
            current_pt=cur_pt,
            current_rot=cur_rot,
            target_pt=target_pt,
            target_rot=target_rot,
            host_element_id=host_elem_id,
            host_description=host_desc,
            skip=skip,
            skip_reason=skip_reason,
        ))

    return out


# ---------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------

def execute_optimize(doc, candidates):
    """Move (or re-parent) every non-skipped candidate, then refresh its
    Element_Linker payload. Caller manages the transaction."""
    result = OptimizeResult()

    for c in candidates:
        if c.skip:
            if "no " in c.skip_reason:
                result.skipped_no_host += 1
            continue

        moved_or_reparented = False

        # Move/rotate (no-op for door re-parent since target == current).
        try:
            cur = XYZ(*c.current_pt)
            tgt = XYZ(*c.target_pt)
            translation = tgt - cur
            if translation.GetLength() > POSITION_TOLERANCE_FT:
                ElementTransformUtils.MoveElement(doc, c.child.Id, translation)
                moved_or_reparented = True
            rot_delta = geometry.normalize_angle(c.target_rot - c.current_rot)
            if abs(rot_delta) > ROTATION_TOLERANCE_DEG:
                axis = Line.CreateBound(
                    tgt, XYZ(tgt.X, tgt.Y, tgt.Z + 1.0)
                )
                ElementTransformUtils.RotateElement(
                    doc, c.child.Id, axis, math.radians(rot_delta)
                )
                moved_or_reparented = True
        except Exception as exc:
            result.warnings.append(
                "Move failed for child id {}: {}".format(c.child_id, exc)
            )
            continue

        # Refresh Element_Linker. Door mode re-parents; the others keep
        # the existing parent_element_id but refresh location/rotation.
        try:
            if c.mode == MODE_DOOR and c.host_element_id is not None:
                door = doc.GetElement(ElementId(int(c.host_element_id)))
                door_pt, door_rot = _location_pt_rot(door)
                new_parent_id = c.host_element_id
                new_parent_loc = list(door_pt) if door_pt else c.linker.parent_location_ft
                new_parent_rot = (
                    door_rot if door_pt is not None else c.linker.parent_rotation_deg
                )
                moved_or_reparented = True
                result.reparented_count += 1
            else:
                new_parent_id = c.linker.parent_element_id
                new_parent_loc = c.linker.parent_location_ft
                new_parent_rot = c.linker.parent_rotation_deg
                if moved_or_reparented:
                    result.moved_count += 1

            new_linker = _el.ElementLinker(
                led_id=c.linker.led_id,
                set_id=c.linker.set_id,
                location_ft=list(c.target_pt),
                rotation_deg=c.target_rot,
                parent_rotation_deg=new_parent_rot,
                parent_element_id=new_parent_id,
                level_id=c.linker.level_id,
                element_id=c.linker.element_id,
                facing=c.linker.facing,
                host_name=c.linker.host_name,
                parent_location_ft=new_parent_loc,
            )
            _el_io.write_to_element(c.child, new_linker)
        except Exception as exc:
            result.warnings.append(
                "Element_Linker rewrite failed for child id {}: {}".format(
                    c.child_id, exc
                )
            )

    return result
