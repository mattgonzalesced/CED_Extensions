# -*- coding: utf-8 -*-
"""Capture-side workflow for space-based profiles.

Authoring path:

  1. User picks a Space in the active view.
  2. If the space has multiple doors, the user picks the reference
     door (same picker used at placement time, so capture and
     placement always agree on "which door is which").
  3. User picks child elements inside the space — receptacles,
     fixtures, keynote symbols, text notes.
  4. For each picked child the engine:
       * Identifies the closest wall by door-relative role
         (opposite_door / right_of_door / left_of_door / behind_door)
         via ``space_placement.closest_wall_for_point``.
       * Computes the proportional ``position_along_wall`` (0..1)
         and a perpendicular ``distance_from_wall_inches``.
       * Records Z elevation from the space's level, and the child's
         rotation as a delta from the wall-inward direction.
       * Emits an LED entry with
         ``placement_rule.kind = wall_anchored`` plus the wall-anchor
         fields, so the placement engine can interpolate the position
         in a different space with different wall lengths and land
         the fixture at the same proportional spot.
  5. Each captured KEYNOTE symbol becomes its own LED (per the
     user's "capture individual keynotes as separate LEDs" request)
     — its ``family_name`` / ``type_name`` come from the keynote's
     ``GA_Keynote Symbol_CED`` family, and the ``placement_rule`` is
     ``wall_anchored`` just like a fixture.
  6. The captured LEDs are bundled into a new ``space_profiles[*]``
     entry. The caller chooses a profile name + bucket assignment
     before persisting via ``active_yaml.save_active_data``.

This module is the data-shaping layer. UI plumbing (Selection.
PickObject, forms.alert, etc.) lives in the pushbutton script.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    LocationPoint,
    RevitLinkInstance,
)

import geometry
import space_placement as _placement
import space_profile_model as _profile_model


# ---------------------------------------------------------------------
# Data classes (plain, no CLR)
# ---------------------------------------------------------------------

class CapturedChild(object):
    """One picked element with the space-anchor data resolved.

    Position is stored as ``x_fraction`` / ``y_fraction`` of the
    space's bbox so it scales with both room dimensions at placement
    time. ``wall_role`` is preserved for ROTATION resolution
    (so a wall-mounted fixture keeps facing into the space when the
    target wall lands on a different cardinal direction).
    """

    __slots__ = (
        "element_id", "label", "category_name", "family_name", "type_name",
        "kind",          # "fixture" | "keynote" | "text_note"
        "x_fraction", "y_fraction",
        "wall_role", "distance_from_wall_inches",
        "z_inches", "rotation_deg",
        "parameters",
    )

    def __init__(self, **kw):
        for slot in self.__slots__:
            setattr(self, slot, kw.get(slot))


class CaptureRequest(object):
    """Inputs supplied by the calling script."""

    __slots__ = (
        "space",          # space_workflow.SpaceInfo
        "door_anchor",    # (origin_xy, inward_xy) or None
        "picked_refs",    # list of Revit References (host + linked)
        "profile_name",
        "bucket_id",
    )

    def __init__(self, space=None, door_anchor=None, picked_refs=None,
                 profile_name="", bucket_id=""):
        self.space = space
        self.door_anchor = door_anchor
        self.picked_refs = list(picked_refs or [])
        self.profile_name = profile_name or ""
        self.bucket_id = bucket_id or ""


class CaptureResult(object):

    __slots__ = ("profile", "captured", "skipped", "warnings")

    def __init__(self):
        self.profile = None          # the dict added to space_profiles
        self.captured = []           # list[CapturedChild]
        self.skipped = []            # list[(reference_or_label, reason)]
        self.warnings = []


# ---------------------------------------------------------------------
# Element identification
# ---------------------------------------------------------------------

_KEYNOTE_FAMILY_NAME = "GA_Keynote Symbol_CED"


def _classify_child(elem):
    """Return ``("fixture" | "keynote" | "text_note" | None, family_name,
    type_name, category_name)``. None means "skip — not something we
    capture as an LED"."""
    if elem is None:
        return None, "", "", ""
    cat = getattr(elem, "Category", None)
    cat_name = getattr(cat, "Name", "") if cat is not None else ""
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        family = getattr(sym, "Family", None) if sym is not None else None
        family_name = getattr(family, "Name", "") or ""
        type_name = getattr(sym, "Name", "") or ""
        if family_name == _KEYNOTE_FAMILY_NAME:
            return "keynote", family_name, type_name, cat_name
        return "fixture", family_name, type_name, cat_name
    # TextNote (annotation) — check by category name. We don't import
    # the TextNote class here to keep the module Revit-version-light.
    if cat_name in ("Text Notes", "Generic Annotations"):
        return "text_note", "", cat_name, cat_name
    return None, "", "", cat_name


def _resolve_reference(doc, reference):
    """Return ``(element, transform_or_None)``. Transform is the link
    instance's total transform when the picked element lives in a
    linked doc; ``None`` for host-doc picks. The transform is what
    lifts the element's location into the host coord frame so we can
    compute wall-relative offsets against the host-side space
    geometry.
    """
    try:
        host_id = reference.ElementId
    except Exception:
        return None, None
    if host_id is None:
        return None, None
    linked_id = None
    try:
        linked_id = reference.LinkedElementId
    except Exception:
        linked_id = None
    is_linked = (
        linked_id is not None and linked_id != ElementId.InvalidElementId
    )
    if is_linked:
        link_inst = doc.GetElement(host_id)
        if not isinstance(link_inst, RevitLinkInstance):
            return None, None
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            return None, None
        elem = link_doc.GetElement(linked_id)
        try:
            transform = link_inst.GetTotalTransform()
        except Exception:
            transform = None
        return elem, transform
    return doc.GetElement(host_id), None


def _element_location_xyz(elem, transform):
    """World-coord ``(x, y, z)`` location in feet for ``elem``, with
    ``transform`` applied when the element lives in a linked doc. Use
    LocationPoint when available; bbox center otherwise. Returns
    ``None`` when the element has no usable location."""
    pt = None
    try:
        loc = elem.Location
        if isinstance(loc, LocationPoint):
            pt = loc.Point
    except Exception:
        pt = None
    if pt is None:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            return None
        try:
            from Autodesk.Revit.DB import XYZ as _XYZ
            pt = _XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
        except Exception:
            return None
    if transform is not None:
        try:
            pt = transform.OfPoint(pt)
        except Exception:
            pass
    return (pt.X, pt.Y, pt.Z)


def _element_elevation_inches(elem, world_z_ft, floor_z_ft):
    """Return the captured Z offset for an element, in inches.

    Order of preference:

      1. ``Elevation from Level`` parameter on the element (the
         standard "height above level" knob for level-based and
         face-based families). This is the user-meaningful value
         shown in the Properties palette — capturing it directly
         avoids the world-Z math going sideways when the family's
         LocationPoint sits on its host face (which can be far
         from the space's level reference plane).
      2. ``Offset`` parameter — fallback used by some wall-hosted
         families.
      3. ``world_z_ft - floor_z_ft`` as a final fallback. Works for
         host-free instances but produces nonsense (thousands of
         inches) when the element's world Z is referenced from a
         different plane than the space's level.
    """
    if elem is None:
        return 0.0
    # Try the named parameters first — same SetValueString round-trip
    # we use elsewhere so the captured number reflects what the user
    # sees in Properties.
    for name in ("Elevation from Level", "Offset"):
        try:
            p = elem.LookupParameter(name)
        except Exception:
            p = None
        if p is None or not p.HasValue:
            continue
        try:
            return float(p.AsDouble()) * 12.0
        except Exception:
            continue
    # Fallback: world Z minus floor Z, in inches.
    try:
        return (float(world_z_ft) - float(floor_z_ft)) * 12.0
    except (TypeError, ValueError):
        return 0.0


def _wall_inward_angle_for_role(geom, door_anchor, role):
    """Inward-direction angle (degrees, math convention 0=+X, 90=+Y)
    of the wall named ``role`` in this space. Used by the capture
    side to convert an element's world rotation into a delta from
    the wall's inward direction — the same number
    ``space_placement.wall_inward_angle_deg`` adds back at placement
    time, so captures survive wall-orientation differences across
    projects.
    """
    if geom is None or door_anchor is None or not role:
        return 0.0
    walls = _placement.wall_segments_for_door(geom, door_anchor)
    seg = walls.get(role)
    if seg is None:
        return 0.0
    _, _, (inx, iny) = seg
    return math.degrees(math.atan2(iny, inx))


def _element_rotation_deg(elem, transform):
    """World-frame rotation about Z in degrees. Linked doc rotation
    is composed with the link's total transform so we end up with a
    host-frame angle. Returns 0.0 if unreadable."""
    rad = 0.0
    try:
        loc = elem.Location
        if isinstance(loc, LocationPoint):
            try:
                rad = loc.Rotation
            except Exception:
                rad = 0.0
    except Exception:
        return 0.0
    if transform is not None:
        # Approximate: take the angle of the transformed X axis.
        try:
            from Autodesk.Revit.DB import XYZ as _XYZ
            x_axis = _XYZ(math.cos(rad), math.sin(rad), 0.0)
            x_world = transform.OfVector(x_axis)
            rad = math.atan2(x_world.Y, x_world.X)
        except Exception:
            pass
    return geometry.normalize_angle(math.degrees(rad))


def _collect_element_parameters(elem):
    """Best-effort capture of an element's instance parameters as a
    plain dict. We only keep parameters that have a value to avoid
    polluting the LED with every read-only built-in slot.
    """
    if elem is None:
        return {}
    out = {}
    try:
        iter_params = elem.Parameters
    except Exception:
        iter_params = []
    for p in iter_params:
        if p is None:
            continue
        try:
            if not p.HasValue:
                continue
            name = p.Definition.Name
        except Exception:
            continue
        if not name:
            continue
        try:
            value = p.AsValueString() or p.AsString() or ""
        except Exception:
            value = ""
        if value == "":
            continue
        out.setdefault(name, value)
    return out


# ---------------------------------------------------------------------
# Capture
# ---------------------------------------------------------------------

def run_capture(doc, request):
    """Build a space profile from ``request``. Returns a
    ``CaptureResult`` with the populated ``profile`` dict (also
    appended into the caller's profile_data when the script wires it
    up). Caller persists via ``active_yaml.save_active_data``.

    ``request.space`` is the SpaceInfo for the target space.
    ``request.door_anchor`` is the (origin_xy, inward_xy) tuple
    selected by the user (None means "use the first door / no door").
    """
    result = CaptureResult()
    if doc is None or request is None or request.space is None:
        result.warnings.append("Missing doc or space; capture aborted.")
        return result

    geom = _placement.build_space_geometry(doc, request.space.element)
    if geom is None:
        result.warnings.append(
            "Space '{}' has no usable boundary; capture aborted.".format(
                request.space.name or "?"
            )
        )
        return result

    door = request.door_anchor
    if door is None and geom.door_anchors:
        door = geom.door_anchors[0]
    if door is None:
        result.warnings.append(
            "Space has no doors. Wall-anchored capture requires at "
            "least one door so walls can be labeled by role; aborting."
        )
        return result

    # Orient the door anchor so its inward points into the space —
    # mirrors what anchor_points does at placement time so capture
    # and placement see identical wall roles.
    door = _placement._orient_door_inward(
        door, (geom.x_center, geom.y_center),
    )

    floor_z = float(geom.floor_z or 0.0)

    for ref in request.picked_refs:
        elem, transform = _resolve_reference(doc, ref)
        if elem is None:
            result.skipped.append((ref, "Could not resolve picked reference."))
            continue
        child_kind, family_name, type_name, cat_name = _classify_child(elem)
        if child_kind is None:
            result.skipped.append(
                (elem, "Category '{}' isn't captured as an LED.".format(cat_name))
            )
            continue
        loc = _element_location_xyz(elem, transform)
        if loc is None:
            result.skipped.append((elem, "No usable location for this element."))
            continue
        wx, wy, wz = loc

        # Position storage — bbox-relative fractions of the space.
        # Scales with both room dimensions when this profile is placed
        # in another (differently-sized) space later.
        fractions = _placement.space_fractions_for_point(geom, (wx, wy))
        if fractions is None:
            result.skipped.append(
                (elem, "Could not resolve space fractions (degenerate bbox?)."),
            )
            continue
        x_fraction, y_fraction = fractions

        # Rotation storage — wall-relative delta. Find the closest
        # wall (by role) just for the rotation reference: a fixture
        # facing perpendicular into the space lands rotation_deg=0
        # and at placement the engine adds the target wall's inward
        # angle back so the fixture faces into the new space too.
        wall_match = _placement.closest_wall_for_point(
            geom, door, (wx, wy),
        )
        if wall_match is None:
            role = "opposite_door"
            distance_in_inches = 0.0
        else:
            role, _t_unused, distance_in_inches = wall_match

        rot_deg = _element_rotation_deg(elem, transform)
        wall_inward_deg = _wall_inward_angle_for_role(geom, door, role)
        rot_delta = geometry.normalize_angle(rot_deg - wall_inward_deg)

        if child_kind in ("keynote", "text_note"):
            # Keynotes / text notes are view-based 2D annotations.
            # Their LocationPoint.Z reflects the host view's plane,
            # not a meaningful "height above floor". Capturing it as
            # an offset would land the placed annotation thousands
            # of inches below the level on the next placement. The
            # placement side renders them on the active view's
            # plane regardless.
            z_inches = 0.0
        else:
            z_inches = _element_elevation_inches(elem, wz, floor_z)

        # Display label — "Family : Type" for fixtures and keynotes,
        # category name as a fallback for text notes.
        if family_name and type_name:
            label = "{} : {}".format(family_name, type_name)
        elif family_name:
            label = family_name
        else:
            label = cat_name or child_kind or "(unnamed)"

        captured = CapturedChild(
            element_id=_id_value(elem),
            label=label,
            category_name=cat_name,
            family_name=family_name,
            type_name=type_name,
            kind=child_kind,
            x_fraction=x_fraction,
            y_fraction=y_fraction,
            wall_role=role,
            distance_from_wall_inches=distance_in_inches,
            z_inches=z_inches,
            rotation_deg=rot_delta,
            parameters=_collect_element_parameters(elem),
        )
        result.captured.append(captured)

    if not result.captured:
        result.warnings.append("No children captured.")
        return result

    result.profile = _build_profile_dict(request, result.captured)
    return result


def commit_capture(profile_data, result):
    """Merge ``result.profile`` into ``profile_data['space_profiles']``.

    Two paths based on whether a profile with the same (case-folded,
    whitespace-trimmed) name already exists:

      * **No collision** → append ``result.profile`` as a fresh entry.
      * **Name collision** → merge captured LEDs into the existing
        profile's first ``linked_set``. LEDs that match an existing
        one on ``(family, type, placement_kind, wall_role,
        position_along_wall ~1%, z_inches ~0.5")`` are skipped to
        avoid duplicating fixtures that were already captured in a
        prior run against the same space.

    Returns ``(action, target_profile, n_added, n_skipped_duplicates)``
    where ``action`` is one of ``"noop"``, ``"created"``, or
    ``"appended"``. The caller is responsible for persisting the
    mutated ``profile_data`` (typically via
    ``active_yaml.save_active_data``).
    """
    if result is None or result.profile is None:
        return "noop", None, 0, 0
    new_profile = result.profile
    new_name = _name_key(new_profile.get("name"))
    if not new_name:
        # No name — can't match anything; just append.
        profile_data.setdefault("space_profiles", []).append(new_profile)
        return "created", new_profile, _count_leds_in_profile(new_profile), 0

    existing = _find_profile_by_name(profile_data, new_name)
    if existing is None:
        profile_data.setdefault("space_profiles", []).append(new_profile)
        return "created", new_profile, _count_leds_in_profile(new_profile), 0

    # Merge into existing profile.
    return _append_into_existing_profile(existing, new_profile)


def _name_key(s):
    return (s or "").strip().lower()


def _find_profile_by_name(profile_data, name_key):
    for p in profile_data.get("space_profiles") or []:
        if not isinstance(p, dict):
            continue
        if _name_key(p.get("name")) == name_key:
            return p
    return None


def _leds_from_profile(profile):
    out = []
    if not isinstance(profile, dict):
        return out
    for s in profile.get("linked_sets") or []:
        if not isinstance(s, dict):
            continue
        for led in s.get("linked_element_definitions") or []:
            if isinstance(led, dict):
                out.append(led)
    return out


def _count_leds_in_profile(profile):
    return sum(1 for _ in _leds_from_profile(profile))


def _led_dedup_key(led):
    """Identity tuple for dedup: same family/type at roughly the same
    space-anchored position and Z counts as the same LED.

    Position fractions are rounded to 2 decimals (~1% of room width /
    height — a 21-ft space translates to ~2.5 inches), and Z is
    snapped to the nearest half-inch. Generous enough that capturing
    the same fixture twice from slightly different selections
    collapses, tight enough that two distinct fixtures separated by
    more than ~1% of the room don't.

    Backward-compat: legacy ``wall_anchored`` LEDs that stored
    ``position_along_wall`` rather than ``x_fraction`` / ``y_fraction``
    fall through to a single-axis key based on the stored
    ``position_along_wall``; new captures use the bbox-fraction pair.
    """
    if not isinstance(led, dict):
        return None
    family = (led.get("family_name") or "").strip().lower()
    type_name = (led.get("type_name") or "").strip().lower()
    rule = led.get("placement_rule") or {}
    kind = (rule.get("kind") or "").strip().lower()
    wall_role = (rule.get("wall_role") or "").strip().lower()
    if "x_fraction" in rule or "y_fraction" in rule:
        try:
            fx = round(float(rule.get("x_fraction") or 0.0), 2)
        except (TypeError, ValueError):
            fx = 0.0
        try:
            fy = round(float(rule.get("y_fraction") or 0.0), 2)
        except (TypeError, ValueError):
            fy = 0.0
        pos_key = ("xy", fx, fy)
    else:
        try:
            pos = round(float(rule.get("position_along_wall") or 0.0), 2)
        except (TypeError, ValueError):
            pos = 0.0
        pos_key = ("alongwall", pos)
    offsets = led.get("offsets") or []
    z_in = 0.0
    if offsets and isinstance(offsets[0], dict):
        try:
            z_in = float(offsets[0].get("z_inches") or 0.0)
        except (TypeError, ValueError):
            z_in = 0.0
    z_key = round(z_in * 2.0) / 2.0
    return (family, type_name, kind, wall_role, pos_key, z_key)


def _append_into_existing_profile(existing, new_profile):
    """Move new_profile's LEDs into existing's first linked_set,
    skipping duplicates. Returns the same tuple shape as
    ``commit_capture``."""
    incoming = _leds_from_profile(new_profile)
    existing_leds = _leds_from_profile(existing)
    existing_keys = set()
    for led in existing_leds:
        k = _led_dedup_key(led)
        if k is not None:
            existing_keys.add(k)

    sets = existing.setdefault("linked_sets", [])
    if not isinstance(sets, list) or not sets or not isinstance(sets[0], dict):
        sets = [{"id": "SP-SET-001", "linked_element_definitions": []}]
        existing["linked_sets"] = sets
    target_set = sets[0]
    target_leds = target_set.setdefault("linked_element_definitions", [])
    if not isinstance(target_leds, list):
        target_leds = []
        target_set["linked_element_definitions"] = target_leds

    n_added = 0
    n_skipped = 0
    # Re-id incoming LEDs so they don't collide with the existing
    # profile's id space.
    next_idx = len(existing_leds) + 1
    for led in incoming:
        key = _led_dedup_key(led)
        if key in existing_keys:
            n_skipped += 1
            continue
        if key is not None:
            existing_keys.add(key)
        led["id"] = "SP-LED-{:03d}".format(next_idx)
        next_idx += 1
        target_leds.append(led)
        n_added += 1
    return "appended", existing, n_added, n_skipped


def _build_profile_dict(request, captured):
    """Materialise the captured children into a ``space_profiles[*]``
    entry. One linked_set holds every LED so future placements treat
    them as a cohesive group; each captured keynote becomes its own
    LED with kind=space_anchored (no annotation list — keynote stands
    alone, per the user's spec).

    Placement rule uses ``kind=space_anchored`` with ``x_fraction`` /
    ``y_fraction`` for position (so a capture in a 21'×15' space
    scales to the new space's bbox at placement) and ``wall_role``
    purely for rotation resolution (the closest wall at capture
    determines which inward direction the placed instance faces).
    """
    leds = []
    for idx, c in enumerate(captured, start=1):
        led = {
            "id": "SP-LED-{:03d}".format(idx),
            "label": c.label,
            "category": c.category_name,
            "family_name": c.family_name,
            "type_name": c.type_name,
            "placement_rule": {
                "kind": _profile_model.KIND_SPACE_ANCHORED,
                "x_fraction": float(c.x_fraction),
                "y_fraction": float(c.y_fraction),
                "wall_role": c.wall_role,
                "distance_from_wall_inches": float(c.distance_from_wall_inches),
            },
            "offsets": [{
                "x_inches": 0.0,
                "y_inches": 0.0,
                "z_inches": float(c.z_inches),
                "rotation_deg": float(c.rotation_deg),
            }],
            "parameters": dict(c.parameters or {}),
        }
        if c.kind == "keynote":
            led["is_keynote"] = True
        leds.append(led)

    return {
        "id": "SP-CAP-{:08x}".format(_short_hash(request.profile_name or "captured")),
        "name": request.profile_name or "Captured space profile",
        "bucket_id": request.bucket_id or "",
        "linked_sets": [{
            "id": "SP-SET-001",
            "linked_element_definitions": leds,
        }],
    }


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _short_hash(s):
    h = 0
    for ch in s or "":
        h = (h * 131 + ord(ch)) & 0xFFFFFFFF
    return h


def _id_value(elem):
    if elem is None:
        return None
    eid = getattr(elem, "Id", None)
    if eid is None:
        return None
    return (
        getattr(eid, "Value", None)
        or getattr(eid, "IntegerValue", None)
    )
