# -*- coding: utf-8 -*-
"""
The unified capture engine.

All four user-facing capture flows (New Profile, Add to Profile, Create
Independent Profile, and the optional directives-pass on the first two)
funnel through ``execute_capture()``. Pushbutton scripts are thin: they
build a ``CaptureRequest``, hand it here, and report the result.

Schema:
    Each LED represents a *fixture* (FamilyInstance/Group). Tags,
    keynote symbols (FamilyInstance "GA_Keynote Symbol_CED"), and text
    notes are stored under the LED as ``annotations`` entries with IDs
    of the form ``SET-NNN-LED-NNN-ANN-NNN``. Annotation offsets are
    relative to the fixture, not to the profile parent.

Coordinates:
    Parent / child / annotation world points come in as ``ElementRef``
    objects (which carry a transform from their source doc to the host
    doc). We pull location + rotation in the host coordinate frame,
    then ``geometry.compute_offsets_from_points`` for the offset math.

Element_Linker:
    For each captured fixture a JSON payload is written to the shared
    ``Element_Linker`` parameter. Linked-doc fixtures cannot have their
    parameter written (we don't own that document), so we report a
    warning and skip the write — the YAML profile still records them.
    Annotations don't carry an Element_Linker; they're sub-records of
    their host LED's element.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    FamilyInstance,
    Group,
    LocationCurve,
    LocationPoint,
    XYZ,
)

import element_linker as _el
import element_linker_io as _el_io
import geometry
import hosted_annotations
import links
import selection as _sel


# ---------------------------------------------------------------------
# request / result
# ---------------------------------------------------------------------

class CaptureRequest(object):
    """Inputs for one capture operation.

    ``profile_name``        required for "new" and "independent" modes
    ``parent``              ElementRef (or None for independent mode)
    ``children``            list[ElementRef]
    ``directives``          dict ``{led_index: {param_name: directive_value}}``
                            applied after the LED is built; pass ``{}`` for none
    ``append_to_profile_id``  EQ-### id to append children to; if set,
                              we ignore profile_name + parent + create a
                              child-only edit
    """

    def __init__(self, profile_name=None, parent=None, children=None,
                 directives=None, append_to_profile_id=None):
        self.profile_name = profile_name
        self.parent = parent
        self.children = children or []
        self.directives = directives or {}
        self.append_to_profile_id = append_to_profile_id


class CaptureResult(object):
    def __init__(self):
        self.profile_id = None
        self.profile_name = None
        self.set_id = None
        self.created_led_ids = []
        self.created_annotation_ids = []
        self.warnings = []
        self.linker_writes = 0
        self.linker_skipped = 0


# ---------------------------------------------------------------------
# Revit element introspection
# ---------------------------------------------------------------------

def _location_point(elem):
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point
    if isinstance(loc, LocationCurve):
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            pass
    bbox = elem.get_BoundingBox(None)
    if bbox is not None:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
    return None


def _location_rotation_radians(elem):
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        # ``LocationPoint.Rotation`` raises InvalidOperationException for
        # families whose orientation isn't stored as a rotation (line-based
        # hosts, some adaptive components). Fall through to the orientation
        # vectors when that happens.
        try:
            return loc.Rotation
        except Exception:
            pass
    facing = getattr(elem, "FacingOrientation", None)
    if facing is not None and (abs(facing.X) > 1e-9 or abs(facing.Y) > 1e-9):
        return math.atan2(facing.Y, facing.X)
    hand = getattr(elem, "HandOrientation", None)
    if hand is not None and (abs(hand.X) > 1e-9 or abs(hand.Y) > 1e-9):
        return math.atan2(hand.Y, hand.X)
    return 0.0


def _world_point(elem_ref):
    p = _location_point(elem_ref.element)
    if p is None:
        return None
    if elem_ref.transform is None:
        return (p.X, p.Y, p.Z)
    transformed = elem_ref.transform.OfPoint(p)
    return (transformed.X, transformed.Y, transformed.Z)


def _world_rotation_deg(elem_ref):
    rad = _location_rotation_radians(elem_ref.element)
    deg = math.degrees(rad or 0.0)
    if elem_ref.transform is not None:
        cos_r = math.cos(rad or 0.0)
        sin_r = math.sin(rad or 0.0)
        local = XYZ(cos_r, sin_r, 0.0)
        try:
            world_vec = elem_ref.transform.OfVector(local)
            deg = math.degrees(math.atan2(world_vec.Y, world_vec.X))
        except Exception:
            pass
    return geometry.normalize_angle(deg)


def _facing_world_tuple(elem_ref):
    facing = getattr(elem_ref.element, "FacingOrientation", None)
    if facing is None:
        return None
    if elem_ref.transform is None:
        return [facing.X, facing.Y, facing.Z]
    v = elem_ref.transform.OfVector(facing)
    return [v.X, v.Y, v.Z]


def _level_id_value(elem):
    lid = getattr(elem, "LevelId", None)
    if lid is None:
        try:
            param = elem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
            if param is not None:
                lid = param.AsElementId()
        except Exception:
            lid = None
    if lid is None:
        return None
    return getattr(lid, "Value", None) or getattr(lid, "IntegerValue", None)


def _id_value(elem_or_id):
    if elem_or_id is None:
        return None
    eid = getattr(elem_or_id, "Id", None) or elem_or_id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def element_label(elem):
    """Public: return ``"Family : Type"`` for FamilyInstance, group type
    name for Group, category name otherwise."""
    if isinstance(elem, FamilyInstance):
        sym = elem.Symbol
        if sym is not None:
            family = sym.Family
            return "{} : {}".format(
                family.Name if family else "",
                sym.Name or "",
            ).strip(" :")
    if isinstance(elem, Group):
        gtype = elem.GroupType
        return gtype.Name if gtype else ""
    cat = getattr(elem, "Category", None)
    return cat.Name if cat else type(elem).__name__


_element_label = element_label  # legacy alias for internal callers


def _element_category_name(elem):
    cat = getattr(elem, "Category", None)
    return cat.Name if cat else ""


# ---------------------------------------------------------------------
# Annotation world point / rotation, accounting for link transform
# ---------------------------------------------------------------------

def _annotation_world_point(ann_elem, transform):
    p = hosted_annotations.annotation_world_point(ann_elem)
    if p is None:
        return None
    if transform is None:
        return p
    xyz = XYZ(p[0], p[1], p[2])
    t = transform.OfPoint(xyz)
    return (t.X, t.Y, t.Z)


def _annotation_world_rotation_deg(ann_elem, transform):
    deg = hosted_annotations.annotation_rotation_deg(ann_elem)
    if transform is None:
        return geometry.normalize_angle(deg)
    rad = math.radians(deg)
    local = XYZ(math.cos(rad), math.sin(rad), 0.0)
    try:
        v = transform.OfVector(local)
        return geometry.normalize_angle(math.degrees(math.atan2(v.Y, v.X)))
    except Exception:
        return geometry.normalize_angle(deg)


# ---------------------------------------------------------------------
# id generation
# ---------------------------------------------------------------------

def _max_numeric_suffix(strings, prefix):
    best = 0
    for s in strings:
        if not isinstance(s, str) or not s.startswith(prefix):
            continue
        rest = s[len(prefix):]
        try:
            n = int(rest)
        except ValueError:
            continue
        if n > best:
            best = n
    return best


def _next_eq_id(profile_doc):
    profiles = profile_doc.get("equipment_definitions") or []
    ids = [p.get("id") for p in profiles if isinstance(p, dict)]
    return "EQ-{:03d}".format(_max_numeric_suffix(ids, "EQ-") + 1)


def _next_set_id(profile_doc):
    seen = []
    for p in profile_doc.get("equipment_definitions") or []:
        if not isinstance(p, dict):
            continue
        for s in p.get("linked_sets") or []:
            if isinstance(s, dict) and isinstance(s.get("id"), str):
                seen.append(s["id"])
    return "SET-{:03d}".format(_max_numeric_suffix(seen, "SET-") + 1)


def _next_led_id(set_dict):
    led_ids = []
    set_id = set_dict.get("id") or "SET-???"
    for led in set_dict.get("linked_element_definitions") or []:
        if isinstance(led, dict) and isinstance(led.get("id"), str):
            led_ids.append(led["id"])
    prefix = "{}-LED-".format(set_id)
    return "{}{:03d}".format(prefix, _max_numeric_suffix(led_ids, prefix) + 1)


def _ann_id(led_id, index):
    return "{}-ANN-{:03d}".format(led_id, index + 1)


# ---------------------------------------------------------------------
# Pick classification + annotation -> fixture matching
# ---------------------------------------------------------------------

def _split_picks_by_kind(child_refs):
    fixtures, annotations = [], []
    for ref in child_refs:
        if hosted_annotations.is_annotation_element(ref.element):
            annotations.append(ref)
        else:
            fixtures.append(ref)
    return fixtures, annotations


def _direct_match_index(ann_ref, fixture_refs):
    """Return the index of the fixture this annotation directly references
    (tagged element or hosted-on element), or None."""
    from Autodesk.Revit.DB import IndependentTag  # lazy: avoid module-load cost
    elem = ann_ref.element
    # Tag -> tagged element
    if isinstance(elem, IndependentTag):
        tagged_id_values = {
            _id_value(t) for t in hosted_annotations.tag_target_element_ids(elem)
        }
        for i, fref in enumerate(fixture_refs):
            if fref.is_linked:
                continue
            if _id_value(fref.element) in tagged_id_values:
                return i
    # Hosted FamilyInstance (e.g., a face-hosted keynote symbol)
    if isinstance(elem, FamilyInstance):
        host = getattr(elem, "Host", None)
        if host is not None:
            host_id_val = _id_value(host)
            for i, fref in enumerate(fixture_refs):
                if fref.is_linked:
                    continue
                if _id_value(fref.element) == host_id_val:
                    return i
    return None


def _proximity_match_index(ann_world_pt, fixture_world_points):
    if ann_world_pt is None:
        return None
    best_idx, best_dist = None, float("inf")
    for i, pt in enumerate(fixture_world_points):
        if pt is None:
            continue
        d = math.sqrt(
            (ann_world_pt[0] - pt[0]) ** 2
            + (ann_world_pt[1] - pt[1]) ** 2
            + (ann_world_pt[2] - pt[2]) ** 2
        )
        if d < best_dist:
            best_dist = d
            best_idx = i
    return best_idx


def _match_annotations_to_fixtures(annotation_refs, fixture_refs, fixture_world_points):
    """Returns ``{fixture_index: [annotation_ref, ...]}``."""
    matches = {i: [] for i in range(len(fixture_refs))}
    if not fixture_refs:
        return matches
    for ann_ref in annotation_refs:
        idx = _direct_match_index(ann_ref, fixture_refs)
        if idx is None:
            ann_pt = _annotation_world_point(ann_ref.element, ann_ref.transform)
            idx = _proximity_match_index(ann_pt, fixture_world_points)
        if idx is None:
            idx = 0  # final fallback: attach to first fixture
        matches[idx].append(ann_ref)
    return matches


# ---------------------------------------------------------------------
# Annotation entry construction
# ---------------------------------------------------------------------

def _build_annotation_entries(led_id, fixture_world_pt, fixture_world_rot,
                              auto_dependents, explicit_ann_refs, result):
    """Build the LED's ``annotations`` list. Dedupes by element id."""
    entries = []
    seen_ids = set()

    # 1. Auto-swept dependents (off the fixture itself).
    for elem, kind in auto_dependents:
        eid = _id_value(elem)
        if eid is None or eid in seen_ids:
            continue
        seen_ids.add(eid)
        entry = _build_one_annotation(elem, kind, transform=None,
                                      fixture_world_pt=fixture_world_pt,
                                      fixture_world_rot=fixture_world_rot)
        if entry is not None:
            entry["id"] = _ann_id(led_id, len(entries))
            entries.append(entry)
            result.created_annotation_ids.append(entry["id"])

    # 2. Explicitly-picked annotations matched to this fixture.
    for ann_ref in explicit_ann_refs:
        elem = ann_ref.element
        eid = _id_value(elem)
        if eid is None or eid in seen_ids:
            continue
        seen_ids.add(eid)
        kind = hosted_annotations.annotation_kind(elem)
        if kind is None:
            continue
        entry = _build_one_annotation(elem, kind, transform=ann_ref.transform,
                                      fixture_world_pt=fixture_world_pt,
                                      fixture_world_rot=fixture_world_rot)
        if entry is not None:
            entry["id"] = _ann_id(led_id, len(entries))
            entries.append(entry)
            result.created_annotation_ids.append(entry["id"])

    return entries


def _build_one_annotation(elem, kind, transform, fixture_world_pt, fixture_world_rot):
    descriptor = hosted_annotations.annotation_descriptor(elem, kind)
    if descriptor is None:
        return None
    ann_pt = _annotation_world_point(elem, transform)
    ann_rot = _annotation_world_rotation_deg(elem, transform)
    if ann_pt is not None and fixture_world_pt is not None:
        offsets = geometry.compute_offsets_from_points(
            fixture_world_pt, fixture_world_rot or 0.0,
            ann_pt, ann_rot or 0.0,
        )
    else:
        offsets = {"x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0,
                   "rotation_deg": 0.0}
    descriptor["offsets"] = offsets
    return descriptor


# ---------------------------------------------------------------------
# LED entry construction
# ---------------------------------------------------------------------

def _build_led_entry(child_ref, parent_world_pt, parent_world_rot_deg, led_id,
                     ann_refs_for_this_fixture, result):
    """Build a v100 LED entry dict for a captured fixture."""
    elem = child_ref.element
    child_world_pt = _world_point(child_ref)
    child_world_rot = _world_rotation_deg(child_ref)
    if child_world_pt is None or parent_world_pt is None:
        offsets = {"x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0,
                   "rotation_deg": 0.0}
    else:
        offsets = geometry.compute_offsets_from_points(
            parent_world_pt, parent_world_rot_deg or 0.0,
            child_world_pt, child_world_rot or 0.0,
        )

    auto_dependents = hosted_annotations.collect_hosted_dependents(elem)
    annotations = _build_annotation_entries(
        led_id, child_world_pt, child_world_rot,
        auto_dependents, ann_refs_for_this_fixture or [], result,
    )

    led = {
        "id": led_id,
        "label": element_label(elem),
        "category": _element_category_name(elem),
        "is_group": isinstance(elem, Group),
        "parameters": hosted_annotations.collect_element_parameters(elem),
        "offsets": [offsets],
        "annotations": annotations,
    }
    return led, child_world_pt, child_world_rot


# ---------------------------------------------------------------------
# Element_Linker payload
# ---------------------------------------------------------------------

def _build_linker_payload(led_id, set_id, child_world_pt, child_world_rot,
                          parent_world_rot, parent_ref, child_ref):
    parent_id = parent_ref.element_id_value if parent_ref else None
    return _el.ElementLinker(
        led_id=led_id,
        set_id=set_id,
        location_ft=list(child_world_pt) if child_world_pt else None,
        rotation_deg=child_world_rot,
        parent_rotation_deg=parent_world_rot if parent_ref else None,
        parent_element_id=parent_id,
        level_id=_level_id_value(child_ref.element),
        element_id=child_ref.element_id_value,
        facing=_facing_world_tuple(child_ref),
        host_name=element_label(parent_ref.element) if parent_ref else None,
        parent_location_ft=(
            list(_world_point(parent_ref))
            if parent_ref else None
        ),
    )


# ---------------------------------------------------------------------
# Directives
# ---------------------------------------------------------------------

def _apply_directives(led, child_index, directives):
    """Patch the LED's parameters dict with any directives for this child."""
    overrides = directives.get(child_index)
    if not overrides:
        return
    params = led.setdefault("parameters", {})
    for param_name, value in overrides.items():
        params[param_name] = value


# ---------------------------------------------------------------------
# Profile skeleton + lookup
# ---------------------------------------------------------------------

def _new_profile_skeleton(profile_id, profile_name, parent_ref):
    if parent_ref is not None:
        elem = parent_ref.element
        family_pattern = ""
        type_pattern = ""
        category_name = _element_category_name(elem) or ""
        if isinstance(elem, FamilyInstance) and elem.Symbol is not None:
            family_pattern = elem.Symbol.Family.Name if elem.Symbol.Family else ""
            type_pattern = elem.Symbol.Name or ""
        elif isinstance(elem, Group) and elem.GroupType is not None:
            family_pattern = elem.GroupType.Name or ""
        parent_filter = {
            "category": category_name,
            "family_name_pattern": family_pattern,
            "type_name_pattern": type_pattern,
            "parameter_filters": {},
        }
    else:
        parent_filter = {
            "category": "",
            "family_name_pattern": "",
            "type_name_pattern": "",
            "parameter_filters": {},
        }
    return {
        "id": profile_id,
        "name": profile_name,
        "schema_version": 100,
        "prompt_on_parent_mismatch": False,
        "parent_filter": parent_filter,
        "linked_sets": [],
        "allow_parentless": parent_ref is None,
        "allow_unmatched_parents": True,
        "equipment_properties": {},
    }


def _find_profile(profile_doc, profile_id):
    for p in profile_doc.get("equipment_definitions") or []:
        if isinstance(p, dict) and p.get("id") == profile_id:
            return p
    return None


def find_profile_by_name(profile_doc, name):
    if not name:
        return None
    for p in profile_doc.get("equipment_definitions") or []:
        if isinstance(p, dict) and p.get("name") == name:
            return p
    return None


def discover_parent_ref(doc, profile):
    """Find the parent that was originally captured for ``profile``.

    Searches placed FamilyInstances + Groups in ``doc`` for an element
    whose Element_Linker payload references one of this profile's sets,
    then resolves that payload's ``parent_element_id`` to a host-doc
    element. Returns an ``ElementRef`` (host coords identity transform)
    or ``None`` if no anchor element is found.
    """
    if not isinstance(profile, dict):
        return None
    sets = profile.get("linked_sets") or []
    target_set_ids = {
        s.get("id") for s in sets
        if isinstance(s, dict) and s.get("id")
    }
    if not target_set_ids:
        return None

    from Autodesk.Revit.DB import (  # noqa: E402
        ElementId,
        FilteredElementCollector,
        Transform,
    )

    for klass in (FamilyInstance, Group):
        collector = FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType()
        for elem in collector:
            linker = _el_io.read_from_element(elem)
            if linker is None:
                continue
            if linker.set_id not in target_set_ids:
                continue
            if linker.parent_element_id is None:
                continue
            try:
                parent_eid = ElementId(int(linker.parent_element_id))
            except (TypeError, ValueError):
                continue
            parent_elem = doc.GetElement(parent_eid)
            if parent_elem is None:
                continue
            return _sel.ElementRef(parent_elem, doc, None, Transform.Identity)
    return None


# ---------------------------------------------------------------------
# The engine
# ---------------------------------------------------------------------

class CaptureError(Exception):
    pass


def execute_capture(doc, profile_doc, request):
    """Mutate ``profile_doc`` to record the capture; write Element_Linker
    payloads to host fixtures. Caller manages the Revit transaction.
    """
    result = CaptureResult()

    # --- Resolve / create the target profile + set ---
    if request.append_to_profile_id:
        profile = _find_profile(profile_doc, request.append_to_profile_id)
        if profile is None:
            raise CaptureError(
                "Cannot append: profile id {} not found".format(
                    request.append_to_profile_id
                )
            )
        sets = profile.setdefault("linked_sets", [])
        if not sets:
            sets.append({
                "id": _next_set_id(profile_doc),
                "name": "{} : {}".format(profile.get("name") or "", "Default Types"),
                "linked_element_definitions": [],
            })
        target_set = sets[0]
    else:
        if not request.profile_name:
            raise CaptureError("profile_name is required for new captures")
        if find_profile_by_name(profile_doc, request.profile_name):
            raise CaptureError(
                "Profile name {!r} already exists".format(request.profile_name)
            )
        profile = _new_profile_skeleton(
            _next_eq_id(profile_doc),
            request.profile_name,
            request.parent,
        )
        profile_doc.setdefault("equipment_definitions", []).append(profile)
        set_id = _next_set_id(profile_doc)
        target_set = {
            "id": set_id,
            "name": "{} : Default Types".format(request.profile_name),
            "linked_element_definitions": [],
        }
        profile["linked_sets"] = [target_set]

    result.profile_id = profile.get("id")
    result.profile_name = profile.get("name")
    result.set_id = target_set.get("id")

    # In append mode the caller often doesn't supply a parent. Discover
    # one from the existing placed elements so offsets land in the same
    # frame as the original capture.
    if request.parent is None and request.append_to_profile_id:
        request.parent = discover_parent_ref(doc, profile)

    # --- Classify picks ---
    fixture_refs, annotation_refs = _split_picks_by_kind(request.children)

    if not fixture_refs:
        raise CaptureError(
            "Pick at least one fixture (FamilyInstance or Group). "
            "Annotations are stored as sub-entries of fixtures, not on their own."
        )

    # --- Resolve parent / centroid pose ---
    if request.parent is not None:
        parent_world_pt = _world_point(request.parent)
        parent_world_rot = _world_rotation_deg(request.parent)
    else:
        parent_world_pt, parent_world_rot = None, None

    # For independent captures, anchor offsets to the centroid of fixtures.
    if parent_world_pt is None and fixture_refs:
        pts = [p for p in (_world_point(c) for c in fixture_refs) if p is not None]
        if pts:
            parent_world_pt = (
                sum(p[0] for p in pts) / len(pts),
                sum(p[1] for p in pts) / len(pts),
                sum(p[2] for p in pts) / len(pts),
            )
            parent_world_rot = 0.0

    # --- Match annotations to fixtures ---
    fixture_world_points = [_world_point(c) for c in fixture_refs]
    matches = _match_annotations_to_fixtures(
        annotation_refs, fixture_refs, fixture_world_points
    )

    # --- Build LED + ANN entries ---
    leds = target_set.setdefault("linked_element_definitions", [])

    for index, fixture_ref in enumerate(fixture_refs):
        led_id = _next_led_id(target_set)
        led, child_world_pt, child_world_rot = _build_led_entry(
            fixture_ref, parent_world_pt, parent_world_rot or 0.0, led_id,
            ann_refs_for_this_fixture=matches.get(index, []),
            result=result,
        )
        _apply_directives(led, index, request.directives)
        leds.append(led)
        result.created_led_ids.append(led_id)

        # Element_Linker write (host fixtures only).
        if fixture_ref.is_linked:
            result.linker_skipped += 1
            result.warnings.append(
                "Element {} lives in a linked document — "
                "Element_Linker write skipped.".format(fixture_ref.element_id_value)
            )
            continue
        try:
            payload = _build_linker_payload(
                led_id, target_set.get("id"),
                child_world_pt, child_world_rot,
                parent_world_rot, request.parent, fixture_ref,
            )
            _el_io.write_to_element(fixture_ref.element, payload)
            result.linker_writes += 1
        except _el_io.ElementLinkerIOError as exc:
            result.warnings.append(str(exc))
            result.linker_skipped += 1

    return result
