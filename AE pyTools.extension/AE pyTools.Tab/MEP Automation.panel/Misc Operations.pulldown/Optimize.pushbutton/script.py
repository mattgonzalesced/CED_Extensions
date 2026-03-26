# -*- coding: utf-8 -*-
"""
Optimize
--------
Optimizes the physical placement of elements by category and family:type.
Only operates on elements that have a non-blank Element_Linker parameter.

Five optimization modes:
  Wall    - Snap element to the face of a nearby wall, facing outward.
            Only walls within 5 ft are eligible. Prefers 0/180-turn walls when
            they are within 2 ft of the best 90/270-turn option.
            Walls are searched in the host document AND all loaded linked models.
  Ceiling - Keep element X/Y and move only Z to ceiling height.
            Ceilings are searched in host + linked models.
  Floor   - Keep element X/Y and move only Z to floor level (elevation = 0).
  Door    - Place element 4 ft high, 1 ft from the nearest doorway.
            Doors are searched in host + linked models.
  Corner  - Place element at a specified corner of its parent's bounding box.
            Parent elements are searched in host + linked models (default: lower-left).
            Door-aware in-body placement is applied when door-side geometry is found.

If any elements are selected, optimization runs only on the selected elements
that have a non-blank Element_Linker parameter. Otherwise it runs on all.
"""

import imp
import math
import os
import re
import sys

from pyrevit import forms, revit, script

output = script.get_output()
output.close_others()

SCRIPT_DIR = os.path.dirname(__file__)

LIB_ROOT = os.path.abspath(
    os.path.join(SCRIPT_DIR, "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FilteredElementCollector,
    Line,
    RevitLinkInstance,
    Transaction,
    Wall,
    XYZ,
)

try:
    basestring
except NameError:
    basestring = str

TITLE = "Optimize Element Placement"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")

# Elevation-from-level BuiltInParameters tried in order.
# Built at import time, skipping names absent in this Revit version.
_ELEV_BIP_NAMES = (
    "INSTANCE_FREE_HOST_OFFSET_PARAM",
    "INSTANCE_ELEVATION_PARAM",
    "FAMILY_BASE_LEVEL_OFFSET_PARAM",
    "FAMILY_TOP_LEVEL_OFFSET_PARAM",
    "SCHEDULE_LEVEL_OFFSET_PARAM",
)
_ELEV_BIPS = tuple(
    getattr(BuiltInParameter, name)
    for name in _ELEV_BIP_NAMES
    if hasattr(BuiltInParameter, name)
)

# Door optimization constants
DOOR_HEIGHT_FT = 4.0   # elevation from level
DOOR_OFFSET_FT = 1.0   # distance from door frame in door's facing direction

# Wall optimization constants
WALL_MAX_SNAP_DIST_FT = 5.0
WALL_PARALLEL_ADVANTAGE_FT = 2.0
WALL_PARALLEL_DOT_THRESHOLD = 0.70710678  # cos(45 deg)

# Corner key -> (use_max_x, use_max_y)
_CORNER_MAP = {
    "Lower Left":  (False, False),
    "Lower Right": (True,  False),
    "Upper Left":  (False, True),
    "Upper Right": (True,  True),
}


# ---------------------------------------------------------------------------
# Linked-model helpers
# ---------------------------------------------------------------------------

def _get_loaded_links(doc):
    """
    Return list of (link_doc, transform) for all currently loaded linked models.
    transform maps points from link coordinates -> host coordinates.
    """
    result = []
    try:
        for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            link_doc = link_inst.GetLinkDocument()
            if link_doc is not None:
                result.append((link_doc, link_inst.GetTotalTransform()))
    except Exception:
        pass
    return result


def _find_element_in_links(doc, elem_id_int):
    """
    Search all loaded linked models for an element with the given integer ID.
    Returns (element, transform) or (None, None).
    transform maps link coords -> host coords.
    """
    eid = ElementId(elem_id_int)
    for link_doc, transform in _get_loaded_links(doc):
        try:
            elem = link_doc.GetElement(eid)
            if elem is not None:
                return elem, transform
        except Exception:
            continue
    return None, None


def _normalize_xy(x, y):
    """Return normalized (x, y) direction tuple, or (0.0, 0.0) if near-zero."""
    length = math.sqrt(x * x + y * y)
    if length < 1e-9:
        return 0.0, 0.0
    return x / length, y / length


def _bbox_contains_xyz(pt, bb_min, bb_max, tol=0.0):
    """Return True if pt is inside [bb_min, bb_max] with a symmetric tolerance."""
    if pt is None or bb_min is None or bb_max is None:
        return False
    return (
        bb_min.X - tol <= pt.X <= bb_max.X + tol and
        bb_min.Y - tol <= pt.Y <= bb_max.Y + tol and
        bb_min.Z - tol <= pt.Z <= bb_max.Z + tol
    )


def _bbox_proj_range(bb_min, bb_max, dir_x, dir_y):
    """
    Return (min_proj, max_proj) of an axis-aligned bbox projected onto dir.
    """
    proj_vals = (
        bb_min.X * dir_x + bb_min.Y * dir_y,
        bb_min.X * dir_x + bb_max.Y * dir_y,
        bb_max.X * dir_x + bb_min.Y * dir_y,
        bb_max.X * dir_x + bb_max.Y * dir_y,
    )
    return min(proj_vals), max(proj_vals)


# ---------------------------------------------------------------------------
# Element Linker parameter helpers
# ---------------------------------------------------------------------------

def _get_linker_text(elem):
    """Return the Element_Linker string value, or empty string if blank/missing."""
    if not elem:
        return ""
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        try:
            value = param.AsString()
        except Exception:
            value = None
        if not value:
            try:
                value = param.AsValueString()
            except Exception:
                value = None
        if value and str(value).strip():
            return str(value)
    return ""


def _set_linker_text(elem, value):
    """Write value to the Element_Linker parameter. Returns True on success."""
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(value)
            return True
        except Exception:
            continue
    return False


def _parse_full_payload(text):
    """
    Return all key:value pairs from an Element_Linker payload as a dict.

    Supports both formats used in this codebase:
      1) multiline "Key: Value" entries
      2) single-line comma-separated entries (placement engine payloads)
    """
    result = {}
    if not text:
        return result
    payload = str(text)

    # Regex pass for comma-separated and multiline payloads.
    key_pattern = re.compile(
        r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|"
        r"Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
        r"Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation|"
        r"CKT_Circuit Number_CEDT|CKT_Panel_CEDT)\s*:\s*"
    )
    matches = list(key_pattern.finditer(payload))
    if matches:
        for idx, match in enumerate(matches):
            key = match.group(1).strip()
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(payload)
            val = payload[start:end].strip().rstrip(",").strip()
            result[key] = val

    # Fallback/merge pass for simple multiline keys not covered above.
    for line in payload.splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        key, _, val = line.partition(":")
        key = key.strip()
        if key and key not in result:
            result[key] = val.strip()
    return result


def _rebuild_payload(fields):
    """Reconstruct a payload string from an ordered dict of key:value pairs."""
    lines = []
    for key, val in fields.items():
        lines.append("{}: {}".format(key, val if val is not None else ""))
    return "\n".join(lines)


def _update_payload(elem, updates):
    """
    Read the existing Element_Linker payload, apply the updates dict,
    and write the result back to the element.
    """
    text = _get_linker_text(elem)
    if not text:
        return
    fields = _parse_full_payload(text)
    fields.update(updates)
    _set_linker_text(elem, _rebuild_payload(fields))


def _get_parent_element_id(elem):
    """Return the integer parent element ID from the Element_Linker payload."""
    text = _get_linker_text(elem)
    if not text:
        return None
    fields = _parse_full_payload(text)
    val = fields.get("Parent ElementId", "").strip()
    if not val:
        return None
    try:
        return int(val)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Location / geometry helpers
# ---------------------------------------------------------------------------

def _get_point(elem):
    """Return the location point (XYZ) of an element, or None."""
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    if hasattr(loc, "Point") and loc.Point:
        return loc.Point
    if hasattr(loc, "Curve") and loc.Curve:
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            return None
    return None


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)


def _element_id_int(elem_id, default=None):
    """Safely extract an integer from a Revit ElementId."""
    if elem_id is None:
        return default
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(elem_id, attr)
            if v is not None:
                return int(v)
        except Exception:
            pass
    return default


def _get_level_element(doc, elem):
    """Return the Level element associated with elem (host doc only), or None."""
    for bip in (
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_LEVEL_PARAM,
    ):
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            continue
        if not param:
            continue
        try:
            eid = param.AsElementId()
            if eid and _element_id_int(eid, -1) > 0:
                return doc.GetElement(eid)
        except Exception:
            continue
    try:
        lid = getattr(elem, "LevelId", None)
        if lid and _element_id_int(lid, -1) > 0:
            return doc.GetElement(lid)
    except Exception:
        pass
    return None


def _move_element(doc, elem, target_pt):
    """Translate elem to target_pt. Returns True on success."""
    current_pt = _get_point(elem)
    if current_pt is None:
        return False
    delta = XYZ(
        target_pt.X - current_pt.X,
        target_pt.Y - current_pt.Y,
        target_pt.Z - current_pt.Z,
    )
    try:
        ElementTransformUtils.MoveElement(doc, elem.Id, delta)
        return True
    except Exception:
        return False


def _rotate_to_face(doc, elem, desired_facing_xy):
    """
    Rotate elem in the XY plane so its FacingOrientation aligns with
    desired_facing_xy (an XYZ with Z=0 representing a 2D direction).
    """
    current_facing = getattr(elem, "FacingOrientation", None)
    if not current_facing:
        return False
    if abs(desired_facing_xy.X) < 1e-9 and abs(desired_facing_xy.Y) < 1e-9:
        return False
    current_angle = math.atan2(current_facing.Y, current_facing.X)
    desired_angle = math.atan2(desired_facing_xy.Y, desired_facing_xy.X)
    delta = desired_angle - current_angle
    if abs(delta) < 1e-6:
        return True
    pt = _get_point(elem)
    if pt is None:
        return False
    try:
        axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
        ElementTransformUtils.RotateElement(doc, elem.Id, axis, delta)
        return True
    except Exception:
        return False


def _set_elevation_param(elem, elevation_ft):
    """Set the elevation-from-level parameter on elem to elevation_ft (in feet)."""
    for bip in _ELEV_BIPS:
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            continue
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(float(elevation_ft))
            return True
        except Exception:
            continue
    return False


# ---------------------------------------------------------------------------
# Cross-document parent / geometry lookup
# ---------------------------------------------------------------------------

def _get_parent_point_in_host(doc, parent_id):
    """
    Return the parent element's location point in host coordinates.
    Searches host doc first, then all loaded linked models.
    """
    if parent_id is None:
        return None
    # Host doc
    try:
        elem = doc.GetElement(ElementId(parent_id))
        if elem is not None:
            return _get_point(elem)
    except Exception:
        pass
    # Linked models
    elem, transform = _find_element_in_links(doc, parent_id)
    if elem is None:
        return None
    pt = _get_point(elem)
    if pt is None:
        return None
    if transform is not None:
        return transform.OfPoint(pt)
    return pt


def _get_parent_bb_corners_in_host(doc, parent_id):
    """
    Return (min_pt, max_pt) of the parent element's bounding box in host coords,
    or None if not found.  Searches host doc first, then linked models.
    """
    if parent_id is None:
        return None
    parent_elem = None
    transform = None
    # Host doc
    try:
        parent_elem = doc.GetElement(ElementId(parent_id))
    except Exception:
        parent_elem = None
    # Linked models fallback
    if parent_elem is None:
        parent_elem, transform = _find_element_in_links(doc, parent_id)
    if parent_elem is None:
        return None
    try:
        bb = parent_elem.get_BoundingBox(None)
    except Exception:
        return None
    if bb is None:
        return None
    if transform is not None:
        return transform.OfPoint(bb.Min), transform.OfPoint(bb.Max)
    return bb.Min, bb.Max


def _find_nearest_wall(doc, pt, elem=None, max_distance_ft=None,
                       prefer_parallel_within_ft=WALL_PARALLEL_ADVANTAGE_FT):
    """
    Find a preferred wall to pt, searching host doc and all loaded linked models.

    Selection policy:
      1) ignore candidates farther than max_distance_ft (if provided)
      2) if element facing is available, prefer a wall that would require
         ~0/180 deg turn over ~90/270 deg turn when that preferred wall is no
         more than prefer_parallel_within_ft farther than the best 90/270 wall
      3) otherwise pick nearest by distance

    Returns dict {"proj_pt": XYZ, "wall_dir": XYZ, "half_width": float}
    in host coordinates, or None if no wall is found.
    """
    candidates = []

    elem_facing = getattr(elem, "FacingOrientation", None) if elem is not None else None
    facing_xy = None
    if elem_facing and (abs(elem_facing.X) > 1e-9 or abs(elem_facing.Y) > 1e-9):
        fx, fy = _normalize_xy(elem_facing.X, elem_facing.Y)
        if fx != 0.0 or fy != 0.0:
            facing_xy = XYZ(fx, fy, 0)

    sources = [(doc, None)] + _get_loaded_links(doc)

    for source_doc, transform in sources:
        try:
            walls = list(
                FilteredElementCollector(source_doc)
                .OfClass(Wall)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue

        for wall in walls:
            wall_loc = getattr(wall, "Location", None)
            if not hasattr(wall_loc, "Curve"):
                continue
            curve = wall_loc.Curve
            try:
                start = curve.GetEndPoint(0)
                # Build query point in this document's coordinate space
                if transform is not None:
                    inv = transform.Inverse
                    q = inv.OfPoint(XYZ(pt.X, pt.Y, start.Z))
                else:
                    q = XYZ(pt.X, pt.Y, start.Z)

                result = curve.Project(q)
                proj_link = result.XYZPoint

                if transform is not None:
                    proj_host = transform.OfPoint(proj_link)
                else:
                    proj_host = proj_link

                dist = math.sqrt(
                    (proj_host.X - pt.X) ** 2 + (proj_host.Y - pt.Y) ** 2
                )
                if max_distance_ft is not None and dist > float(max_distance_ft):
                    continue

                end = curve.GetEndPoint(1)
                nx, ny = _normalize_xy(end.X - start.X, end.Y - start.Y)
                dir_link = XYZ(nx, ny, 0)
                if transform is not None:
                    dh = transform.OfVector(dir_link)
                    dhx, dhy = _normalize_xy(dh.X, dh.Y)
                    dir_host = XYZ(dhx, dhy, 0)
                else:
                    dir_host = dir_link
                try:
                    half_w = wall.Width / 2.0
                except Exception:
                    half_w = 0.25
                candidate = {
                    "dist": dist,
                    "proj_pt": proj_host,
                    "wall_dir": dir_host,
                    "half_width": half_w,
                }
                if facing_xy is not None:
                    wall_normal = XYZ(-dir_host.Y, dir_host.X, 0)
                    to_elem_x = pt.X - proj_host.X
                    to_elem_y = pt.Y - proj_host.Y
                    if wall_normal.X * to_elem_x + wall_normal.Y * to_elem_y < 0:
                        wall_normal = XYZ(-wall_normal.X, -wall_normal.Y, 0)
                    align_dot = abs(
                        wall_normal.X * facing_xy.X + wall_normal.Y * facing_xy.Y
                    )
                    candidate["is_parallel_turn"] = align_dot >= WALL_PARALLEL_DOT_THRESHOLD
                candidates.append(candidate)
            except Exception:
                continue

    if not candidates:
        return None

    # Always keep nearest as baseline.
    candidates.sort(key=lambda c: c["dist"])
    nearest = candidates[0]

    if facing_xy is None:
        return {
            "proj_pt": nearest["proj_pt"],
            "wall_dir": nearest["wall_dir"],
            "half_width": nearest["half_width"],
        }

    parallel = [c for c in candidates if c.get("is_parallel_turn")]
    perpendicular = [c for c in candidates if not c.get("is_parallel_turn")]
    best_parallel = parallel[0] if parallel else None
    best_perpendicular = perpendicular[0] if perpendicular else None

    chosen = nearest
    if best_parallel is not None:
        if best_perpendicular is None:
            chosen = best_parallel
        elif best_parallel["dist"] <= best_perpendicular["dist"] + float(prefer_parallel_within_ft):
            chosen = best_parallel
        else:
            chosen = best_perpendicular

    return {
        "proj_pt": chosen["proj_pt"],
        "wall_dir": chosen["wall_dir"],
        "half_width": chosen["half_width"],
    }


def _find_ceiling_z_above(doc, pt):
    """
    Return the Z of the lowest ceiling bottom face found above pt.Z.
    Searches host doc and all loaded linked models.
    Returns float or None.
    """
    ceiling_z = None
    sources = [(doc, None)] + _get_loaded_links(doc)

    for source_doc, transform in sources:
        try:
            ceilings = list(
                FilteredElementCollector(source_doc)
                .OfCategory(BuiltInCategory.OST_Ceilings)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue

        for ceiling in ceilings:
            try:
                bb = ceiling.get_BoundingBox(None)
                if bb is not None:
                    # Use bottom face center to get the Z in correct coords
                    center = XYZ(
                        (bb.Min.X + bb.Max.X) / 2.0,
                        (bb.Min.Y + bb.Max.Y) / 2.0,
                        bb.Min.Z,
                    )
                else:
                    c_pt = _get_point(ceiling)
                    if c_pt is None:
                        continue
                    center = c_pt

                if transform is not None:
                    z_host = transform.OfPoint(center).Z
                else:
                    z_host = center.Z

                if z_host > pt.Z:
                    if ceiling_z is None or z_host < ceiling_z:
                        ceiling_z = z_host
            except Exception:
                continue

    return ceiling_z


def _find_nearest_door(doc, pt):
    """
    Find the nearest door to pt, searching host doc and all loaded linked models.
    Returns dict {"door_pt": XYZ, "facing_dir": XYZ or None} in host coordinates,
    or None if no door found.
    """
    best_dist = float("inf")
    best = None
    sources = [(doc, None)] + _get_loaded_links(doc)

    for source_doc, transform in sources:
        try:
            doors = list(
                FilteredElementCollector(source_doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(FamilyInstance)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue

        for door in doors:
            door_pt_local = _get_point(door)
            if door_pt_local is None:
                continue

            if transform is not None:
                door_pt_host = transform.OfPoint(door_pt_local)
            else:
                door_pt_host = door_pt_local

            dist = math.sqrt(
                (door_pt_host.X - pt.X) ** 2 + (door_pt_host.Y - pt.Y) ** 2
            )
            if dist >= best_dist:
                continue

            best_dist = dist
            facing_local = getattr(door, "FacingOrientation", None)
            facing_dir = None
            if facing_local and (abs(facing_local.X) > 1e-9 or abs(facing_local.Y) > 1e-9):
                if transform is not None:
                    fh = transform.OfVector(facing_local)
                    fx, fy = _normalize_xy(fh.X, fh.Y)
                else:
                    fx, fy = _normalize_xy(facing_local.X, facing_local.Y)
                if fx != 0.0 or fy != 0.0:
                    facing_dir = XYZ(fx, fy, 0)
            best = {"door_pt": door_pt_host, "facing_dir": facing_dir}

    return best


# ---------------------------------------------------------------------------
# YAML profile helpers (Optimization=NO skip check)
# ---------------------------------------------------------------------------

def _build_led_params_map(data):
    """
    Build lookup maps from active YAML profile data for Optimization checks.
    Returns:
      {
        "by_host_set_led": {"host_lc|set_lc|led_lc": [raw_params, ...]},
        "by_set_led": {"set_lc|led_lc": [raw_params, ...]},
        "by_host_led": {"host_lc|led_lc": [raw_params, ...]},
        "by_led_id": {led_lc: [raw_params, ...]},
        "by_host_set_label": {"host_lc|set_lc|label_lc": [raw_params, ...]},
        "by_set_label": {"set_lc|label_lc": [raw_params, ...]},
      }
    """
    def _norm_text(value):
        return str(value or "").strip().lower()

    def _norm_label(value):
        text = str(value or "").replace(u"\uff1a", ":").strip().lower()
        if not text:
            return ""
        parts = [part.strip() for part in text.split(":")]
        if len(parts) >= 2:
            text = "{}:{}".format(parts[0], ":".join(parts[1:]).strip())
        return " ".join(text.split())

    def _ci_get(mapping, key_name):
        if not isinstance(mapping, dict):
            return None
        target = _norm_text(key_name)
        for key, value in mapping.items():
            if _norm_text(key) == target:
                return value
        return None

    def _append(bucket, key, values):
        if not key:
            return
        key = _norm_text(key)
        if not key:
            return
        for val in values:
            if val is None:
                continue
            bucket.setdefault(key, []).append(val)

    def _host_keys(eq_def, linked_set, led_def):
        keys = []
        for raw in (
            eq_def.get("name"),
            eq_def.get("id"),
            linked_set.get("name") if isinstance(linked_set, dict) else None,
            linked_set.get("id") if isinstance(linked_set, dict) else None,
            linked_set.get("host_name") if isinstance(linked_set, dict) else None,
            led_def.get("host_name") if isinstance(led_def, dict) else None,
            led_def.get("host") if isinstance(led_def, dict) else None,
        ):
            norm = _norm_text(raw)
            if norm and norm not in keys:
                keys.append(norm)
        return keys

    by_host_set_led = {}
    by_set_led = {}
    by_host_led = {}
    by_led_id = {}
    by_host_set_label = {}
    by_set_label = {}

    for eq in (data or {}).get("equipment_definitions") or []:
        if not isinstance(eq, dict):
            continue
        eq_scope = []
        eq_props = eq.get("equipment_properties")
        if eq_props is not None:
            eq_scope.append(eq_props)
        eq_params = _ci_get(eq, "parameters")
        if eq_params is not None:
            eq_scope.append(eq_params)

        linked_sets = eq.get("linked_sets") or []
        if not linked_sets and isinstance(eq.get("linked_element_definitions"), list):
            linked_sets = [{
                "id": eq.get("id"),
                "linked_element_definitions": eq.get("linked_element_definitions"),
            }]
        for ls in linked_sets or []:
            if not isinstance(ls, dict):
                continue
            set_id = _norm_text(ls.get("id") or eq.get("id"))
            set_scope = list(eq_scope)
            set_params = ls.get("parameters")
            if set_params is not None:
                set_scope.append(set_params)
            for led in (ls.get("linked_element_definitions") or []):
                if not isinstance(led, dict):
                    continue
                led_id = _norm_text(led.get("id"))
                label = led.get("label")
                if not label:
                    fam = led.get("family_name") or led.get("family") or ""
                    typ = led.get("type_name") or led.get("type") or ""
                    if fam and typ:
                        label = u"{} : {}".format(fam, typ)
                label_key = _norm_label(label)

                raw_values = list(set_scope)
                led_params = led.get("parameters")
                if led_params is not None:
                    raw_values.append(led_params)
                inst_cfg = led.get("instance_config")
                if isinstance(inst_cfg, dict):
                    inst_params = inst_cfg.get("parameters")
                    if inst_params is not None:
                        raw_values.append(inst_params)

                host_keys = _host_keys(eq, ls, led)

                _append(by_led_id, led_id, raw_values)
                if set_id and led_id:
                    _append(by_set_led, "{}|{}".format(set_id, led_id), raw_values)
                if set_id and label_key:
                    _append(by_set_label, "{}|{}".format(set_id, label_key), raw_values)

                for host in host_keys:
                    if led_id:
                        _append(by_host_led, "{}|{}".format(host, led_id), raw_values)
                    if set_id and led_id:
                        _append(by_host_set_led, "{}|{}|{}".format(host, set_id, led_id), raw_values)
                    if set_id and label_key:
                        _append(by_host_set_label, "{}|{}|{}".format(host, set_id, label_key), raw_values)

    return {
        "by_host_set_led": by_host_set_led,
        "by_set_led": by_set_led,
        "by_host_led": by_host_led,
        "by_led_id": by_led_id,
        "by_host_set_label": by_host_set_label,
        "by_set_label": by_set_label,
    }


def _get_led_id_from_elem(elem):
    """Return the Linked Element Definition ID from the Element_Linker payload."""
    text = _get_linker_text(elem)
    if not text:
        return None
    fields = _parse_full_payload(text)
    # Preferred exact key.
    val = fields.get("Linked Element Definition ID", "")
    if str(val).strip():
        return str(val).strip()
    # Fallback: case-insensitive key match.
    for key, value in fields.items():
        if str(key).strip().lower() == "linked element definition id":
            if str(value).strip():
                return str(value).strip()
    return None


def _get_payload_field_ci(elem, field_name):
    """Case-insensitive field lookup from Element_Linker payload."""
    text = _get_linker_text(elem)
    if not text:
        return None
    fields = _parse_full_payload(text)
    target = str(field_name).strip().lower()
    for key, value in fields.items():
        if str(key).strip().lower() == target:
            val = str(value or "").strip()
            return val or None
    return None


def _normalize_lookup_label(value):
    text = str(value or "").replace(u"\uff1a", ":").strip().lower()
    if not text:
        return ""
    parts = [part.strip() for part in text.split(":")]
    if len(parts) >= 2:
        text = "{}:{}".format(parts[0], ":".join(parts[1:]).strip())
    return " ".join(text.split())


def _dict_get_ci(data, key_name):
    """Case-insensitive dict key lookup."""
    if not isinstance(data, dict):
        return None
    target = str(key_name).strip().lower()
    for key, value in data.items():
        if str(key).strip().lower() == target:
            return value
    return None


def _coerce_led_params_dict(raw_params):
    """
    Coerce common YAML parameter shapes into a simple name->value dict.
    Supported inputs:
      - dict: {"Optimization": "NO"}
      - list of dict items:
          [{"name": "Optimization", "value": "NO"}, ...]
    """
    if isinstance(raw_params, dict):
        return raw_params
    result = {}
    if isinstance(raw_params, (list, tuple)):
        for item in raw_params:
            if not isinstance(item, dict):
                continue
            name = None
            for k in ("name", "parameter", "parameter_name", "key", "label"):
                v = item.get(k)
                if str(v or "").strip():
                    name = str(v).strip()
                    break
            if name:
                value = None
                for vk in ("value", "Value", "current_value", "default_value", "text", "Text"):
                    if vk in item:
                        value = item.get(vk)
                        break
                if value is None:
                    for k, v in item.items():
                        lk = str(k).strip().lower()
                        if lk not in ("name", "parameter", "parameter_name", "key", "label"):
                            value = v
                            break
                result[name] = value
            else:
                # Fallback for items shaped like {"Optimization": "NO"}
                for k, v in item.items():
                    result[str(k).strip()] = v
    return result


def _value_to_text(value):
    """Normalize various value payload shapes to a comparable uppercase string."""
    if isinstance(value, dict):
        for k in ("value", "Value", "current_value", "default_value", "text", "Text"):
            if k in value:
                return str(value.get(k) or "").strip().upper()
        return ""
    if isinstance(value, (list, tuple)):
        for item in value:
            text = _value_to_text(item)
            if text:
                return text
        return ""
    return str(value or "").strip().upper()


def _is_disabled_value(value):
    token = _value_to_text(value)
    return token in ("NO", "FALSE", "OFF", "0")


def _is_optimization_disabled(elem, led_params_map):
    """
    Return True if the element's LED profile has a parameter
    named 'Optimization' set to 'NO' (case-insensitive).
    Returns False if no YAML data is loaded or the parameter is absent.
    """
    if not led_params_map or not isinstance(led_params_map, dict):
        return False

    by_host_set_led = led_params_map.get("by_host_set_led", {})
    by_set_led = led_params_map.get("by_set_led", {})
    by_host_led = led_params_map.get("by_host_led", {})
    by_led_id = led_params_map.get("by_led_id", {})
    by_host_set_label = led_params_map.get("by_host_set_label", {})
    by_set_label = led_params_map.get("by_set_label", {})

    led_id = (_get_led_id_from_elem(elem) or "").strip().lower()
    set_id = (_get_payload_field_ci(elem, "Set Definition ID") or "").strip().lower()
    host_name = (_get_payload_field_ci(elem, "Host Name") or "").strip().lower()
    label_key = _normalize_lookup_label(_element_label(elem))

    # Strict profile association precedence from Element_Linker context.
    candidate_sets = []
    if led_id and set_id and host_name:
        candidate_sets = list(by_host_set_led.get("{}|{}|{}".format(host_name, set_id, led_id), []))
    if not candidate_sets and led_id and set_id:
        candidate_sets = list(by_set_led.get("{}|{}".format(set_id, led_id), []))
    if not candidate_sets and led_id and host_name:
        candidate_sets = list(by_host_led.get("{}|{}".format(host_name, led_id), []))
    if not candidate_sets and led_id:
        candidate_sets = list(by_led_id.get(led_id, []))
    if not candidate_sets and set_id and host_name and label_key:
        candidate_sets = list(by_host_set_label.get("{}|{}|{}".format(host_name, set_id, label_key), []))
    if not candidate_sets and set_id and label_key:
        candidate_sets = list(by_set_label.get("{}|{}".format(set_id, label_key), []))

    for raw_params in candidate_sets:
        params = _coerce_led_params_dict(raw_params)
        opt_value = _dict_get_ci(params, "Optimization")
        if _is_disabled_value(opt_value):
            return True
    return False


# ---------------------------------------------------------------------------
# Element label / category helpers
# ---------------------------------------------------------------------------

def _element_label(elem):
    """Return the 'Family : Type' string for a FamilyInstance."""
    try:
        sym = getattr(elem, "Symbol", None)
        if sym:
            fam = getattr(sym, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else None
            type_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = type_param.AsString() if type_param else getattr(sym, "Name", None)
            if fam_name and type_name:
                return u"{} : {}".format(fam_name, type_name)
    except Exception:
        pass
    return ""


def _element_category(elem):
    try:
        cat = getattr(elem, "Category", None)
        return getattr(cat, "Name", None) if cat else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Collect and classify elements (host doc only — these are what we move)
# ---------------------------------------------------------------------------

def _collect_linked_elements(doc):
    """Return all FamilyInstance elements in the host doc with a non-blank Element_Linker."""
    result = []
    try:
        instances = list(
            FilteredElementCollector(doc)
            .OfClass(FamilyInstance)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return result
    for elem in instances:
        if _get_linker_text(elem):
            result.append(elem)
    return result


def _collect_selected_linked_elements(doc):
    """
    Return selected host-model FamilyInstances with non-blank Element_Linker.
    Returns:
      (elements, had_selection)
    """
    selected = []
    had_selection = False
    uidoc = getattr(revit, "uidoc", None)
    if uidoc is None:
        return selected, had_selection
    try:
        sel_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        sel_ids = []
    if not sel_ids:
        return selected, had_selection

    had_selection = True
    for eid in sel_ids:
        try:
            elem = doc.GetElement(eid)
        except Exception:
            elem = None
        if elem is None:
            continue
        if not isinstance(elem, FamilyInstance):
            continue
        if _get_linker_text(elem):
            selected.append(elem)
    return selected, had_selection


def _build_category_map(elements):
    """Return {category_name: [unique_family_type_label, ...]}."""
    cat_map = {}
    for elem in elements:
        cat = _element_category(elem) or "Unknown"
        label = _element_label(elem)
        if not label:
            continue
        cat_map.setdefault(cat, set()).add(label)
    return {cat: list(labels) for cat, labels in cat_map.items()}


# ---------------------------------------------------------------------------
# Optimization: Wall
# ---------------------------------------------------------------------------

def _optimize_wall(doc, elem):
    """
    Move the element to the face of the nearest wall (host or linked),
    keeping the same Z height, and rotate to face away from the wall.
    """
    pt = _get_point(elem)
    if pt is None:
        return False

    wall_info = _find_nearest_wall(
        doc,
        pt,
        elem=elem,
        max_distance_ft=WALL_MAX_SNAP_DIST_FT,
        prefer_parallel_within_ft=WALL_PARALLEL_ADVANTAGE_FT,
    )
    if wall_info is None:
        return False

    proj_pt = wall_info["proj_pt"]
    wall_dir = wall_info["wall_dir"]
    half_width = wall_info["half_width"]

    # Choose the normal side facing toward the element
    wall_normal = XYZ(-wall_dir.Y, wall_dir.X, 0)
    to_elem_x = pt.X - proj_pt.X
    to_elem_y = pt.Y - proj_pt.Y
    dot = wall_normal.X * to_elem_x + wall_normal.Y * to_elem_y
    if dot < 0:
        wall_normal = XYZ(-wall_normal.X, -wall_normal.Y, 0)

    target_pt = XYZ(
        proj_pt.X + wall_normal.X * half_width,
        proj_pt.Y + wall_normal.Y * half_width,
        pt.Z,  # preserve current height
    )

    moved = _move_element(doc, elem, target_pt)
    if moved:
        _rotate_to_face(doc, elem, wall_normal)
        new_pt = _get_point(elem) or target_pt
        facing = getattr(elem, "FacingOrientation", None)
        _update_payload(elem, {
            "Location XYZ (ft)": _format_xyz(new_pt),
            "FacingOrientation": _format_xyz(facing) if facing else "",
        })
    return moved


# ---------------------------------------------------------------------------
# Optimization: Ceiling
# ---------------------------------------------------------------------------

def _optimize_ceiling(doc, elem):
    """
    Keep element X/Y and move only Z to ceiling height.
    Sets the elevation-from-level parameter to the ceiling height.
    """
    pt = _get_point(elem)
    if pt is None:
        return False

    ceiling_z = _find_ceiling_z_above(doc, pt)
    if ceiling_z is None:
        level = _get_level_element(doc, elem)
        if level:
            try:
                ceiling_z = level.Elevation + 9.0
            except Exception:
                ceiling_z = pt.Z
        else:
            ceiling_z = pt.Z

    target_pt = XYZ(pt.X, pt.Y, ceiling_z)
    moved = _move_element(doc, elem, target_pt)

    if moved:
        level = _get_level_element(doc, elem)
        level_elev = 0.0
        if level:
            try:
                level_elev = level.Elevation
            except Exception:
                pass
        _set_elevation_param(elem, ceiling_z - level_elev)
        new_pt = _get_point(elem) or target_pt
        _update_payload(elem, {"Location XYZ (ft)": _format_xyz(new_pt)})
    return moved


# ---------------------------------------------------------------------------
# Optimization: Floor
# ---------------------------------------------------------------------------

def _optimize_floor(doc, elem):
    """
    Keep element X/Y and move only Z to floor level.
    Sets elevation-from-level to 0.
    """
    pt = _get_point(elem)
    if pt is None:
        return False

    level = _get_level_element(doc, elem)
    floor_z = pt.Z
    if level:
        try:
            floor_z = level.Elevation
        except Exception:
            pass

    target_pt = XYZ(pt.X, pt.Y, floor_z)
    moved = _move_element(doc, elem, target_pt)

    if moved:
        _set_elevation_param(elem, 0.0)
        new_pt = _get_point(elem) or target_pt
        _update_payload(elem, {"Location XYZ (ft)": _format_xyz(new_pt)})
    return moved


# ---------------------------------------------------------------------------
# Optimization: Door
# ---------------------------------------------------------------------------

def _optimize_door(doc, elem):
    """
    Place the element 1 ft from the nearest door (host or linked),
    at 4 ft elevation from the associated level.
    """
    pt = _get_point(elem)
    if pt is None:
        return False

    door_info = _find_nearest_door(doc, pt)
    if door_info is None:
        return False

    door_pt = door_info["door_pt"]
    facing_dir = door_info["facing_dir"]

    if facing_dir is None:
        # Fallback: offset toward current element location
        dx = pt.X - door_pt.X
        dy = pt.Y - door_pt.Y
        fx, fy = _normalize_xy(dx, dy)
        if fx == 0.0 and fy == 0.0:
            return False
        facing_dir = XYZ(fx, fy, 0)

    level = _get_level_element(doc, elem)
    level_elev = 0.0
    if level:
        try:
            level_elev = level.Elevation
        except Exception:
            pass

    target_pt = XYZ(
        door_pt.X + facing_dir.X * DOOR_OFFSET_FT,
        door_pt.Y + facing_dir.Y * DOOR_OFFSET_FT,
        level_elev + DOOR_HEIGHT_FT,
    )

    moved = _move_element(doc, elem, target_pt)
    if moved:
        _set_elevation_param(elem, DOOR_HEIGHT_FT)
        _rotate_to_face(doc, elem, facing_dir)
        new_pt = _get_point(elem) or target_pt
        facing = getattr(elem, "FacingOrientation", None)
        _update_payload(elem, {
            "Location XYZ (ft)": _format_xyz(new_pt),
            "FacingOrientation": _format_xyz(facing) if facing else "",
        })
    return moved


# ---------------------------------------------------------------------------
# Shape detection helper
# ---------------------------------------------------------------------------

def _is_rectangular_in_plan(elem):
    """
    Return True if elem's footprint is rectangular or near-rectangular in plan.

    Strategy:
      1) Collect vertical-face normal directions clustered by parallelism.
      2) Fast path: true rectangle has <=2 directions and (if 2) they're
         perpendicular.
      3) Near-rectangle fallback: if the two dominant orthogonal directions
         account for most vertical face area, treat as rectangular. This keeps
         families with tiny chamfers/fillets from being misclassified.

    Defaults to True (treat as rectangular) if geometry cannot be read.
    """
    try:
        from Autodesk.Revit.DB import Options

        opts = Options()
        opts.ComputeReferences = False
        opts.IncludeNonVisibleObjects = False

        geo_elem = elem.get_Geometry(opts)
        if geo_elem is None:
            return True

        # Collect all solids, unwrapping GeometryInstances
        solids = []
        stack = list(geo_elem)
        while stack:
            obj = stack.pop()
            try:
                if obj.Volume > 1e-9:
                    solids.append(obj)
                    continue
            except Exception:
                pass
            try:
                inner = obj.GetInstanceGeometry()
                if inner:
                    stack.extend(list(inner))
            except Exception:
                pass

        if not solids:
            return True

        # Gather XY-plane normal clusters from vertical faces (area-weighted)
        directions = []  # [{"nx": float, "ny": float, "area": float}, ...]
        angle_tol = 0.15  # ~8.6 degrees
        cos_tol = math.cos(angle_tol)
        total_vert_area = 0.0

        for solid in solids:
            try:
                for face in solid.Faces:
                    try:
                        n = face.FaceNormal
                    except Exception:
                        continue
                    if abs(n.Z) > 0.85:          # skip top / bottom faces
                        continue
                    nx, ny = _normalize_xy(n.X, n.Y)
                    if nx == 0.0 and ny == 0.0:
                        continue
                    try:
                        area = float(face.Area)
                    except Exception:
                        area = 1.0
                    if area <= 0.0:
                        area = 1.0
                    total_vert_area += area

                    merged = False
                    for cluster in directions:
                        dot = nx * cluster["nx"] + ny * cluster["ny"]
                        if abs(dot) > cos_tol:
                            sign = 1.0 if dot >= 0.0 else -1.0
                            old_area = cluster["area"]
                            new_area = old_area + area
                            ax = cluster["nx"] * old_area + (nx * sign) * area
                            ay = cluster["ny"] * old_area + (ny * sign) * area
                            ax, ay = _normalize_xy(ax, ay)
                            cluster["nx"] = ax
                            cluster["ny"] = ay
                            cluster["area"] = new_area
                            merged = True
                            break
                    if not merged:
                        directions.append({"nx": nx, "ny": ny, "area": area})
            except Exception:
                continue

        if not directions:
            return True

        if len(directions) > 2:
            # Near-rectangular fallback:
            # two dominant orthogonal directions + minor residual faces.
            ranked = sorted(directions, key=lambda d: d["area"], reverse=True)
            if len(ranked) >= 2 and total_vert_area > 1e-6:
                d1, d2 = ranked[0], ranked[1]
                dominant_ratio = (d1["area"] + d2["area"]) / total_vert_area
                dot = abs(d1["nx"] * d2["nx"] + d1["ny"] * d2["ny"])
                if dominant_ratio >= 0.80 and dot <= 0.25:
                    return True
            return False
        if len(directions) == 2:
            dot = abs(directions[0]["nx"] * directions[1]["nx"] +
                      directions[0]["ny"] * directions[1]["ny"])
            if dot > math.sin(angle_tol):   # not perpendicular
                return False
        return True

    except Exception:
        return True     # default: treat as rectangular


# Corner placement offset (feet) when shape is rectangular
_CORNER_INSET_FT = 1.0
# Additional side-to-center offset (feet), perpendicular to container facing.
_CORNER_PERP_TO_CENTER_OFFSET_FT = 2.0

# For pairs: the opposite corner on the same level (Lower↔Lower, Upper↔Upper, flip side)
_OPPOSITE_SIDE_CORNER = {
    "Lower Left":  "Lower Right",
    "Lower Right": "Lower Left",
    "Upper Left":  "Upper Right",
    "Upper Right": "Upper Left",
}

_DOOR_SIDE_CORNER = {
    "Lower Left":  "Lower Left",
    "Lower Right": "Lower Right",
    "Upper Left":  "Upper Left",
    "Upper Right": "Upper Right",
}

# For non-door containers, invert corner selection as requested:
# LL->UR, LR->UL, UL->LR, UR->LL
_NON_DOOR_FLIPPED_CORNER = {
    "Lower Left":  "Lower Left",
    "Lower Right": "Lower Right",
    "Upper Left":  "Upper Left",
    "Upper Right": "Upper Right",
}

# Door-side controller placement depth controls:
# setback = max(min_ft, depth * ratio), clamped by max_ratio of depth.
_CASE_DOOR_SIDE_MIN_SETBACK_FT = 1.0
_CASE_DOOR_SIDE_SETBACK_RATIO = 0.45
_CASE_DOOR_SIDE_MAX_SETBACK_RATIO = 0.85
# Arc-door families (lower case style) need an extra push into housing.
# Keep modest here; explicit post-shift handles the extra 1 ft request.
_CASE_DOOR_SIDE_ARC_EXTRA_SETBACK_FT = 1.0
_CASE_DOOR_SIDE_ARC_TRIGGER_FT = 0.35
# Extra post-clamp push for arc-door cases (lower family), toward case back.
_CASE_DOOR_SIDE_ARC_POST_SHIFT_FT = 1.0
# Final direct XY offset for arc-door cases, positive facing direction.
_CASE_DOOR_SIDE_ARC_FINAL_OFFSET_FT = 0.5
# Final direct XY offset for all door-side cases in facing direction.
_CASE_DOOR_SIDE_FACING_OFFSET_FT = 2.0


def _alternate_corner(base_corner, idx):
    """
    Return the corner for the element at position idx in a same-container group.
    Even index  → base_corner (e.g. "Lower Left")
    Odd  index  → opposite side, same level (e.g. "Lower Right")
    """
    if idx % 2 == 0:
        return base_corner
    return _OPPOSITE_SIDE_CORNER.get(base_corner, base_corner)


def _corner_for_group(base_corner, idx, group_len):
    """
    Alternate only for an exact pair in the same-container same-type group.
    Otherwise keep the selected corner for every element.
    """
    if int(group_len or 0) == 2:
        return _alternate_corner(base_corner, idx)
    return base_corner


def _container_has_non_line_edges(container_elem):
    """
    Return True if container geometry has any non-line edge curves in solids.
    Arc-door case families trigger this and are used for extra back-shift.
    """
    if container_elem is None:
        return False
    try:
        from Autodesk.Revit.DB import Options
        opts = Options()
        opts.ComputeReferences = False
        opts.IncludeNonVisibleObjects = False
        geo = container_elem.get_Geometry(opts)
    except Exception:
        return False
    if geo is None:
        return False

    stack = list(geo)
    while stack:
        obj = stack.pop()
        try:
            inner = obj.GetInstanceGeometry()
            if inner is not None:
                stack.extend(list(inner))
                continue
        except Exception:
            pass
        try:
            if obj.Volume <= 1e-9:
                continue
        except Exception:
            continue
        try:
            for edge in obj.Edges:
                try:
                    if type(edge.AsCurve()).__name__ != "Line":
                        return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def _has_door_in_names(elem):
    """Return True if 'door' appears in any common name field of elem."""
    if elem is None:
        return False
    names = []
    fam_name = ""
    type_name = ""
    try:
        names.append(getattr(elem, "Name", "") or "")
    except Exception:
        pass
    try:
        sym = getattr(elem, "Symbol", None)
        type_name = getattr(sym, "Name", "") or ""
        names.append(type_name)
        fam = getattr(sym, "Family", None)
        fam_name = getattr(fam, "Name", "") or ""
        names.append(fam_name)
    except Exception:
        pass
    if fam_name or type_name:
        names.append("{}:{}".format(fam_name, type_name))
    try:
        label = _element_label(elem)
        if label:
            names.append(label)
    except Exception:
        pass
    try:
        cat = getattr(elem, "Category", None)
        names.append(getattr(cat, "Name", "") or "")
    except Exception:
        pass
    text = " ".join([str(n) for n in names if n])
    return "door" in text.lower()


def _is_probable_door_case(container_elem, bb_min=None, bb_max=None, door_dir=None,
                           doc=None, pt=None, exclude_id=None):
    """
    Door/non-door switch:
    if container family/type/instance/category name contains "door"
    (case-insensitive), it's door logic.
    Also checks larger containing families at pt to catch nested sub-family
    containers inside a door case.
    """
    if container_elem is None:
        return False
    if _has_door_in_names(container_elem):
        return True

    # Fallback: if this element sits inside a larger door-named container,
    # still treat as door logic.
    if doc is not None and pt is not None:
        try:
            instances = list(
                FilteredElementCollector(doc)
                .OfClass(FamilyInstance)
                .WhereElementIsNotElementType()
            )
        except Exception:
            instances = []
        for candidate in instances:
            try:
                if exclude_id is not None and candidate.Id == exclude_id:
                    continue
            except Exception:
                pass
            try:
                bb = candidate.get_BoundingBox(None)
            except Exception:
                bb = None
            if bb is None:
                continue
            if not _bbox_contains_xyz(pt, bb.Min, bb.Max, tol=0.1):
                continue
            if _has_door_in_names(candidate):
                return True

    return False


def _group_corner_elements_by_container(doc, elements):
    """
    Group elements by the container they belong to so paired elements inside
    the same case are distributed to opposite corners.

    Grouping key priority:
      1. Parent ElementId from Element_Linker (reliable if set correctly)
      2. Integer ID of the FamilyInstance whose bounding box contains the element
      3. Fallback key "_no_container" for unresolvable elements

    Returns {container_key_str: [elem, ...]}
    """
    groups = {}
    for elem in elements:
        pt = _get_point(elem)
        container_key = None

        # Method 1: parent ID from payload
        parent_id = _get_parent_element_id(elem)
        if parent_id is not None:
            container_key = "pid_{}".format(parent_id)

        # Method 2: bounding box containment
        if container_key is None and pt is not None:
            container_elem, _ = _find_container_by_bbox(doc, pt, exclude_id=elem.Id)
            if container_elem is not None:
                try:
                    container_key = "eid_{}".format(
                        _element_id_int(container_elem.Id, 0)
                    )
                except Exception:
                    container_key = "eid_unknown"

        if container_key is None:
            container_key = "_no_container"

        groups.setdefault(container_key, []).append(elem)
    return groups


def _corner_point_in_facing_frame(bb_min, bb_max, facing_dir, corner_name, inset_ft=0.0):
    """
    Return (target_x, target_y) for corner_name interpreted relative to
    facing_dir's orientation, optionally inset by inset_ft along each of the
    facing frame's local axes toward the bounding box center.

    Local frame (standing in front of the case, looking at it from outside):
        "Lower" = backward  — the side away from the facing direction
        "Upper" = forward   — the side toward the facing direction
        "Left"  = left when facing forward (90 deg CCW from forward in plan)
        "Right" = right when facing forward (90 deg CW from forward in plan)

    Example: case faces East (+X), corner = "Lower Left"
        Lower  = West  (-X)  → picks Min.X of bbox
        Left   = North (+Y)  → picks Max.Y of bbox
        result = world upper-left corner, which is correct for an east-facing case.
    """
    fx = facing_dir.X
    fy = facing_dir.Y
    # right = 90 deg CW from forward in XY plane
    rx = fy
    ry = -fx

    _spec = {
        "Lower Left":  (-1, -1),   # backward + left
        "Lower Right": (-1, +1),   # backward + right
        "Upper Left":  (+1, -1),   # forward  + left
        "Upper Right": (+1, +1),   # forward  + right
    }
    fwd_sign, right_sign = _spec.get(corner_name, (-1, -1))

    # Unit vector pointing toward the desired corner in world space
    desired_x = fwd_sign * fx + right_sign * rx
    desired_y = fwd_sign * fy + right_sign * ry

    # Pick the bbox corner that scores highest in that direction
    candidates = [
        (bb_min.X, bb_min.Y),
        (bb_max.X, bb_min.Y),
        (bb_min.X, bb_max.Y),
        (bb_max.X, bb_max.Y),
    ]
    best_cx, best_cy = candidates[0]
    best_score = float("-inf")
    for cx, cy in candidates:
        score = cx * desired_x + cy * desired_y
        if score > best_score:
            best_score = score
            best_cx, best_cy = cx, cy

    # Inset toward center: move opposite to fwd_sign along forward,
    # and opposite to right_sign along right
    if inset_ft != 0.0:
        inset_x = (-fwd_sign * fx + (-right_sign) * rx) * inset_ft
        inset_y = (-fwd_sign * fy + (-right_sign) * ry) * inset_ft
        best_cx += inset_x
        best_cy += inset_y

    return best_cx, best_cy

# ---------------------------------------------------------------------------
# Door-aware bounding box (shrinks bbox to exclude door sub-component areas)
# ---------------------------------------------------------------------------

def _analyze_container_geometry(container_elem, fallback_facing=None):
    """
    Walk container_elem geometry once and return (body_bb_min, body_bb_max, door_dir).

    body_bb_min / body_bb_max
        Tight bbox of body-only solids (door geometry excluded).
        Falls back to get_BoundingBox(None) on failure.
        Returns (None, None, None) if the plain bbox cannot be read.

    door_dir
        XYZ unit vector pointing FROM the case body TOWARD the door zone.
        Determined by the most reliable method available (in priority order):

        1. Bbox-extension method (most reliable when arc door swings exist):
           The direction in which the full element bbox extends BEYOND the
           body bbox (body = straight-edge-only solids).  Arc-solid door
           swings always inflate the bbox on the door side, so the extension
           direction unambiguously identifies the door side regardless of
           how FacingOrientation is defined in the family.

        2. Phase-2 thin-panel sweep (for families with no arc door solids):
           Runs Phase 2 (thin-front-panel removal) in up to three candidate
           directions: the arc-extension direction (if any), fallback_facing,
           and -fallback_facing.  The direction that produces the most
           shrinkage of the body bbox front face is selected — this is the
           door side because only door panels are thin and at the front face.

        None if door geometry cannot be reliably determined.

    fallback_facing
        Optional FacingOrientation from the container element.  Used only as
        a Phase-2 candidate direction when no arc geometry is found.
    """
    try:
        full_bb = container_elem.get_BoundingBox(None)
    except Exception:
        return None, None, None
    if full_bb is None:
        return None, None, None

    try:
        from Autodesk.Revit.DB import Options
        opts = Options()
        opts.ComputeReferences = False
        opts.IncludeNonVisibleObjects = False
        geo = container_elem.get_Geometry(opts)
    except Exception:
        return full_bb.Min, full_bb.Max, None
    if geo is None:
        return full_bb.Min, full_bb.Max, None

    # -----------------------------------------------------------------------
    # Geometry walk: collect per-solid (has_arc, pts) pairs.
    # ALL tessellated edge points are gathered (arc or straight) so Phase 2
    # can inspect the full spatial extent of each solid.
    # -----------------------------------------------------------------------
    solid_infos = []  # [(has_arc: bool, pts: [XYZ, ...])]
    stack = list(geo)
    while stack:
        obj = stack.pop()
        try:
            inner = obj.GetInstanceGeometry()
            if inner is not None:
                stack.extend(list(inner))
                continue
        except Exception:
            pass
        try:
            if obj.Volume <= 1e-9:
                continue
        except Exception:
            continue
        has_arc = False
        pts = []
        try:
            for edge in obj.Edges:
                try:
                    curve = edge.AsCurve()
                    if type(curve).__name__ != "Line":
                        has_arc = True
                except Exception:
                    pass
                try:
                    for pt in edge.Tessellate():
                        pts.append(pt)
                except Exception:
                    pass
        except Exception:
            pass
        if pts:
            solid_infos.append((has_arc, pts))

    if not solid_infos:
        return full_bb.Min, full_bb.Max, None

    # -----------------------------------------------------------------------
    # Phase 1: arc-edge exclusion
    # -----------------------------------------------------------------------
    body_infos = [(a, p) for (a, p) in solid_infos if not a]
    if not body_infos:
        return full_bb.Min, full_bb.Max, None  # all geometry had arcs

    # Body bbox after Phase 1
    body_pts_p1 = [pt for (_, pts) in body_infos for pt in pts]
    b_min_x = min(p.X for p in body_pts_p1)
    b_min_y = min(p.Y for p in body_pts_p1)
    b_min_z = min(p.Z for p in body_pts_p1)
    b_max_x = max(p.X for p in body_pts_p1)
    b_max_y = max(p.Y for p in body_pts_p1)
    b_max_z = max(p.Z for p in body_pts_p1)

    # -----------------------------------------------------------------------
    # Compute door_dir via bbox extension (Method 1 — most reliable).
    # Arc door swings inflate full_bb beyond the body bbox on the door side.
    # -----------------------------------------------------------------------
    door_dir = None
    try:
        ext_candidates = [
            (full_bb.Max.X - b_max_x, XYZ(1, 0, 0)),   # extension in +X
            (b_min_x - full_bb.Min.X, XYZ(-1, 0, 0)),  # extension in -X
            (full_bb.Max.Y - b_max_y, XYZ(0, 1, 0)),   # extension in +Y
            (b_min_y - full_bb.Min.Y, XYZ(0, -1, 0)),  # extension in -Y
        ]
        best_ext, best_dir = max(ext_candidates, key=lambda kv: kv[0])
        if best_ext > 0.08:  # allow smaller door-swing/door-panel extensions
            door_dir = best_dir
    except Exception:
        pass

    # -----------------------------------------------------------------------
    # Phase 2: thin-front-panel filter.
    # Build candidate facing directions to test.  Run Phase 2 in each; pick
    # the direction that achieves the most front-face shrinkage (that is the
    # door side).
    # -----------------------------------------------------------------------
    candidate_dirs = []

    def _append_candidate_dir(vec):
        """Normalize and append vec to candidate_dirs unless duplicated."""
        if vec is None:
            return
        try:
            nx, ny = _normalize_xy(vec.X, vec.Y)
        except Exception:
            return
        if nx == 0.0 and ny == 0.0:
            return
        for existing in candidate_dirs:
            # Keep opposite directions; drop only near-identical duplicates.
            if existing.X * nx + existing.Y * ny > 0.999:
                return
        candidate_dirs.append(XYZ(nx, ny, 0))

    _append_candidate_dir(door_dir)
    if fallback_facing is not None:
        _append_candidate_dir(fallback_facing)
        try:
            _append_candidate_dir(XYZ(-fallback_facing.X, -fallback_facing.Y, 0))
        except Exception:
            pass

    # Last resort for non-facing families (common in case controllers):
    # still try Phase 2 against world axes so top cases with rectangular doors
    # can derive a usable door direction.
    if not candidate_dirs:
        _append_candidate_dir(XYZ(1, 0, 0))
        _append_candidate_dir(XYZ(-1, 0, 0))
        _append_candidate_dir(XYZ(0, 1, 0))
        _append_candidate_dir(XYZ(0, -1, 0))

    best_shrinkage = 0.0
    best_filtered = None
    best_p2_dir = None

    for cd in candidate_dirs:
        try:
            fx, fy = cd.X, cd.Y
            all_bpts = [pt for (_, pts) in body_infos for pt in pts]
            front_proj = max(pt.X * fx + pt.Y * fy for pt in all_bpts)
            back_proj  = min(pt.X * fx + pt.Y * fy for pt in all_bpts)
            total_depth_cd = front_proj - back_proj
            # Relative threshold: catch door panels up to 60 % of the case
            # depth in this direction.  For the lower case, arc door swings
            # are already stripped by Phase 1, leaving only deep body solids
            # → no shrinkage → Phase 2 is a no-op there.  For the upper case,
            # door panels are rectangular but much shallower than the body.
            depth_threshold = max(0.4, total_depth_cd * 0.7)
            front_band = max(0.15, total_depth_cd * 0.12)
            filtered = []
            for (_, pts) in body_infos:
                projs = [pt.X * fx + pt.Y * fy for pt in pts]
                depth = max(projs) - min(projs)
                at_front = max(projs) >= front_proj - front_band
                if depth <= depth_threshold and at_front:
                    continue  # panel at front face (door / frame) → exclude
                filtered.append(pts)
            if not filtered:
                continue
            new_pts = [pt for pts in filtered for pt in pts]
            new_front = max(pt.X * fx + pt.Y * fy for pt in new_pts)
            shrinkage = front_proj - new_front
            if shrinkage > best_shrinkage:
                best_shrinkage = shrinkage
                best_filtered = filtered
                best_p2_dir = XYZ(fx, fy, 0)
        except Exception:
            pass

    if best_filtered is not None and best_shrinkage > 0.02:
        body_infos = [(False, pts) for pts in best_filtered]
        if door_dir is None:
            door_dir = best_p2_dir  # Phase 2 identified the door direction

    # -----------------------------------------------------------------------
    # Final body bbox
    # -----------------------------------------------------------------------
    all_body_pts = [pt for (_, pts) in body_infos for pt in pts]
    if not all_body_pts:
        return full_bb.Min, full_bb.Max, door_dir

    return (
        XYZ(min(p.X for p in all_body_pts),
            min(p.Y for p in all_body_pts),
            min(p.Z for p in all_body_pts)),
        XYZ(max(p.X for p in all_body_pts),
            max(p.Y for p in all_body_pts),
            max(p.Z for p in all_body_pts)),
        door_dir,
    )


# ---------------------------------------------------------------------------
# Bounding-box containment search (used by Corner)
# ---------------------------------------------------------------------------

def _find_container_by_bbox(doc, pt, exclude_id=None):
    """
    Search all host-model FamilyInstances for the one whose bounding box most
    tightly contains pt.  Excludes the element with exclude_id (the element
    being moved) so it doesn't match itself.

    Returns (element, bounding_box) or (None, None).
    The tightest match (smallest bounding box volume that still contains pt)
    is preferred so we pick the specific case, not a room or building shell.
    """
    best_elem = None
    best_bb = None
    best_vol = float("inf")
    tol = 0.1  # 0.1 ft (~1.2 in) tolerance on each face

    try:
        instances = list(
            FilteredElementCollector(doc)
            .OfClass(FamilyInstance)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return None, None

    for candidate in instances:
        if exclude_id is not None:
            try:
                if candidate.Id == exclude_id:
                    continue
            except Exception:
                pass
        try:
            bb = candidate.get_BoundingBox(None)
        except Exception:
            continue
        if bb is None:
            continue
        if _bbox_contains_xyz(pt, bb.Min, bb.Max, tol=tol):
            vol = ((bb.Max.X - bb.Min.X) *
                   (bb.Max.Y - bb.Min.Y) *
                   (bb.Max.Z - bb.Min.Z))
            if vol < best_vol:
                best_vol = vol
                best_elem = candidate
                best_bb = bb

    return best_elem, best_bb


# ---------------------------------------------------------------------------
# Optimization: Corner
# ---------------------------------------------------------------------------

def _optimize_corner(doc, elem, corner="Lower Left"):
    """
    Place the element relative to its container's bounding box and match the
    container's facing orientation.

    Container resolution (in order):
      1. Parent ElementId from the Element_Linker payload (host doc then links).
      2. Bounding-box containment search in the host doc — finds whichever
         FamilyInstance most tightly contains the element's current location.
         This is the primary path for host-model elements placed inside a
         host-model container family.

    Placement by container shape:
      Rectangular  — corner of bounding box offset _CORNER_INSET_FT (1 ft) in
                     both X and Y toward the bounding box center.
      Non-rect.    — center of bounding box (X/Y midpoint, preserve Z).

    Corner options (plan view, rectangular only):
        'Lower Left'  -> (Min.X, Min.Y)   [default]
        'Lower Right' -> (Max.X, Min.Y)
        'Upper Left'  -> (Min.X, Max.Y)
        'Upper Right' -> (Max.X, Max.Y)
    """
    pt = _get_point(elem)
    if pt is None:
        return False

    bb_min = None
    bb_max = None
    container_elem = None
    container_facing = None
    door_dir = None  # geometry-derived direction from body toward door zone

    # --- Method 1: Element_Linker parent ID ---
    parent_id = _get_parent_element_id(elem)
    bb_corners = _get_parent_bb_corners_in_host(doc, parent_id)
    if bb_corners is not None:
        bb_min, bb_max = bb_corners
        if parent_id is not None:
            try:
                container_elem = doc.GetElement(ElementId(parent_id))
            except Exception:
                container_elem = None
            if container_elem is None:
                container_elem, _ = _find_element_in_links(doc, parent_id)
        if container_elem is not None:
            container_facing = getattr(container_elem, "FacingOrientation", None)
            try:
                if container_elem.Document.IsLinked is False:
                    adj_min, adj_max, door_dir = _analyze_container_geometry(
                        container_elem, container_facing)
                    if adj_min is not None:
                        bb_min, bb_max = adj_min, adj_max
            except Exception:
                pass

    # If payload parent exists but its bbox does not contain this element's
    # current point, treat the parent as unreliable and fall back to host
    # containment search.
    if bb_min is not None and not _bbox_contains_xyz(pt, bb_min, bb_max, tol=0.25):
        bb_min = None
        bb_max = None
        container_elem = None
        container_facing = None
        door_dir = None

    # --- Method 2: Bounding-box containment (host model) ---
    if bb_min is None:
        container_elem, container_bb = _find_container_by_bbox(
            doc, pt, exclude_id=elem.Id
        )
        if container_elem is None:
            return False
        container_facing = getattr(container_elem, "FacingOrientation", None)
        adj_min, adj_max, door_dir = _analyze_container_geometry(
            container_elem, container_facing)
        if adj_min is not None:
            bb_min, bb_max = adj_min, adj_max
        else:
            bb_min = container_bb.Min
            bb_max = container_bb.Max

    # Step 1: align moved element facing to the container before corner pick.
    if container_facing and (
            abs(container_facing.X) > 1e-9 or abs(container_facing.Y) > 1e-9):
        try:
            _rotate_to_face(doc, elem, container_facing)
            pt = _get_point(elem) or pt
        except Exception:
            pass

    # --- Choose facing direction for corner-frame calculation ---
    # Corner selection should follow the container FacingOrientation directly.
    # Use geometry-derived door_dir only when FacingOrientation is unavailable.
    has_door_side_geometry = _is_probable_door_case(
        container_elem, bb_min, bb_max, door_dir, doc=doc, pt=pt, exclude_id=elem.Id
    )
    facing_for_corner = container_facing
    if (not facing_for_corner) or (
            abs(facing_for_corner.X) <= 1e-9 and abs(facing_for_corner.Y) <= 1e-9):
        facing_for_corner = door_dir

    if facing_for_corner and (
            abs(facing_for_corner.X) > 1e-9 or abs(facing_for_corner.Y) > 1e-9):
        fx, fy = _normalize_xy(facing_for_corner.X, facing_for_corner.Y)
        if fx != 0.0 or fy != 0.0:
            facing_for_corner = XYZ(fx, fy, 0)
        else:
            facing_for_corner = None
    else:
        facing_for_corner = None

    corner_for_target = corner

    # --- Determine target XY based on container shape ---
    is_rect_container = _is_rectangular_in_plan(container_elem)
    if not is_rect_container:
        # Only center for non-rectangular containers.
        target_x = (bb_min.X + bb_max.X) / 2.0
        target_y = (bb_min.Y + bb_max.Y) / 2.0
    elif facing_for_corner:
        # Rectangular container + known frame direction.
        target_x, target_y = _corner_point_in_facing_frame(
            bb_min, bb_max, facing_for_corner, corner_for_target, inset_ft=_CORNER_INSET_FT
        )
    else:
        # No facing info at all: world-axis min/max + inset.
        use_max_x, use_max_y = _CORNER_MAP.get(corner_for_target, (False, False))
        base_x = bb_max.X if use_max_x else bb_min.X
        base_y = bb_max.Y if use_max_y else bb_min.Y
        x_sign = -1.0 if use_max_x else 1.0
        y_sign = -1.0 if use_max_y else 1.0
        target_x = base_x + x_sign * _CORNER_INSET_FT
        target_y = base_y + y_sign * _CORNER_INSET_FT

    # Door-side cases: apply 2 ft offset in the case facing direction.
    if has_door_side_geometry and _CASE_DOOR_SIDE_FACING_OFFSET_FT > 0.0:
        door_ref = container_facing if container_facing else facing_for_corner
        if door_ref and (abs(door_ref.X) > 1e-9 or abs(door_ref.Y) > 1e-9):
            dfx, dfy = _normalize_xy(door_ref.X, door_ref.Y)
            if dfx != 0.0 or dfy != 0.0:
                target_x += dfx * _CASE_DOOR_SIDE_FACING_OFFSET_FT
                target_y += dfy * _CASE_DOOR_SIDE_FACING_OFFSET_FT

    # Rectangular containers: apply an additional perpendicular shift toward
    # center based on the container-facing direction.
    if (is_rect_container and (not has_door_side_geometry) and
            _CORNER_PERP_TO_CENTER_OFFSET_FT > 0.0):
        perp_ref = container_facing if container_facing else facing_for_corner
        if perp_ref and (abs(perp_ref.X) > 1e-9 or abs(perp_ref.Y) > 1e-9):
            pfx, pfy = _normalize_xy(perp_ref.X, perp_ref.Y)
            if pfx != 0.0 or pfy != 0.0:
                # right vector = 90 deg CW from facing
                rx, ry = pfy, -pfx
                cx = (bb_min.X + bb_max.X) / 2.0
                cy = (bb_min.Y + bb_max.Y) / 2.0
                to_center_x = cx - target_x
                to_center_y = cy - target_y
                proj_to_center = to_center_x * rx + to_center_y * ry
                if abs(proj_to_center) > 1e-9:
                    step = min(_CORNER_PERP_TO_CENTER_OFFSET_FT, abs(proj_to_center))
                    sign = 1.0 if proj_to_center > 0.0 else -1.0
                    target_x += rx * sign * step
                    target_y += ry * sign * step

    target_pt = XYZ(target_x, target_y, pt.Z)
    moved = _move_element(doc, elem, target_pt)

    if moved:
        if container_facing and (
                abs(container_facing.X) > 1e-9 or abs(container_facing.Y) > 1e-9):
            _rotate_to_face(doc, elem, container_facing)
        new_pt = _get_point(elem) or target_pt
        facing = getattr(elem, "FacingOrientation", None)
        _update_payload(elem, {
            "Location XYZ (ft)": _format_xyz(new_pt),
            "FacingOrientation": _format_xyz(facing) if facing else "",
        })
    return moved


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

def _apply_optimization(doc, elem, mode, corner="Lower Left"):
    if mode == "Wall":
        return _optimize_wall(doc, elem)
    elif mode == "Ceiling":
        return _optimize_ceiling(doc, elem)
    elif mode == "Floor":
        return _optimize_floor(doc, elem)
    elif mode == "Door":
        return _optimize_door(doc, elem)
    elif mode == "Corner":
        return _optimize_corner(doc, elem, corner)
    return False


def _bump_reason(counter, samples, reason, elem=None, max_samples=3):
    """Increment reason counter and keep a few sample element IDs."""
    key = reason or "Unknown"
    counter[key] = counter.get(key, 0) + 1
    if elem is None:
        return
    bucket = samples.setdefault(key, [])
    if len(bucket) >= max_samples:
        return
    try:
        eid = _element_id_int(elem.Id, None)
    except Exception:
        eid = None
    if eid is None:
        return
    eid_txt = str(eid)
    if eid_txt not in bucket:
        bucket.append(eid_txt)


def _diagnose_failure(doc, elem, mode):
    """Best-effort reason for a failed optimization call."""
    if elem is None:
        return "Element missing"
    try:
        if getattr(elem, "Pinned", False):
            return "Element is pinned"
    except Exception:
        pass

    pt = _get_point(elem)
    if pt is None:
        return "Element has no point location"

    if mode == "Wall":
        if _find_nearest_wall(
                doc,
                pt,
                elem=elem,
                max_distance_ft=WALL_MAX_SNAP_DIST_FT,
                prefer_parallel_within_ft=WALL_PARALLEL_ADVANTAGE_FT,
        ) is None:
            return "No eligible wall found within {} ft".format(int(WALL_MAX_SNAP_DIST_FT))
    elif mode == "Ceiling":
        if _find_ceiling_z_above(doc, pt) is None:
            return "No ceiling above"
    elif mode == "Door":
        if _find_nearest_door(doc, pt) is None:
            return "No nearby door found"
    elif mode == "Corner":
        parent_id = _get_parent_element_id(elem)
        if parent_id is None or _get_parent_bb_corners_in_host(doc, parent_id) is None:
            container_elem, _ = _find_container_by_bbox(doc, pt, exclude_id=elem.Id)
            if container_elem is None:
                return "No containing family found"

    return "Move/rotate blocked by host or constraints"


def _append_reason_lines(lines, title, reason_counts, reason_samples, max_rows=8):
    """Append reason breakdown lines to summary output."""
    if not reason_counts:
        return
    lines.append("")
    lines.append(title)
    ranked = sorted(reason_counts.items(), key=lambda kv: (-kv[1], kv[0]))
    for reason, count in ranked[:max_rows]:
        samples = ", ".join(reason_samples.get(reason, []))
        if samples:
            lines.append("  - {}: {} (e.g. ElementId {})".format(reason, count, samples))
        else:
            lines.append("  - {}: {}".format(reason, count))
    hidden = len(ranked) - min(len(ranked), max_rows)
    if hidden > 0:
        lines.append("  - {} more reason(s) not shown".format(hidden))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    selected_elements, had_selection = _collect_selected_linked_elements(doc)
    elements = selected_elements if had_selection else _collect_linked_elements(doc)
    if not elements:
        if had_selection:
            forms.alert(
                "Selection contains no FamilyInstances with a non-blank Element_Linker parameter.",
                title=TITLE,
            )
        else:
            forms.alert(
                "No elements with a non-blank Element_Linker parameter were found.",
                title=TITLE,
            )
        return

    category_map = _build_category_map(elements)
    if not category_map:
        forms.alert(
            "No categorized elements with Element_Linker found.",
            title=TITLE,
        )
        return

    # Load YAML profile data to check for Optimization=NO on any LED.
    # If no active YAML is loaded this silently does nothing (all elements run).
    led_params_map = {}
    try:
        from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402
        _, yaml_data = load_active_yaml_data()
        led_params_map = _build_led_params_map(yaml_data)
    except Exception:
        led_params_map = {}

    xaml_path = os.path.join(SCRIPT_DIR, "OptimizeUI.xaml")
    ui_module = imp.load_source(
        "ced_optimize_ui",
        os.path.join(SCRIPT_DIR, "OptimizeUI.py"),
    )
    window = ui_module.OptimizeWindow(xaml_path, category_map)
    window.ShowDialog()

    if not window.confirmed or not window.rules:
        return

    # Build lookup: family_type_label -> [elements]
    label_to_elems = {}
    for elem in elements:
        label = _element_label(elem)
        if label:
            label_to_elems.setdefault(label, []).append(elem)

    stats = {"attempted": 0, "succeeded": 0, "failed": 0, "skipped_no_opt": 0}
    fail_reasons = {}
    fail_samples = {}
    skip_reasons = {}
    skip_samples = {}
    txn = Transaction(doc, "Optimize Element Placement")
    try:
        txn.Start()
        for ft_label, rule in window.rules.items():
            mode = rule.get("mode", "Wall")
            corner = rule.get("corner", "Lower Left")
            elems_for_type = label_to_elems.get(ft_label, [])
            if mode == "Corner":
                # Group by container so paired elements alternate corners
                groups = _group_corner_elements_by_container(doc, elems_for_type)
                for _container_key, group_elems in groups.items():
                    for idx, elem in enumerate(group_elems):
                        stats["attempted"] += 1
                        if _is_optimization_disabled(elem, led_params_map):
                            stats["skipped_no_opt"] += 1
                            _bump_reason(
                                skip_reasons,
                                skip_samples,
                                "Optimization=NO for associated profile fixture",
                                elem,
                            )
                            continue
                        effective_corner = _corner_for_group(corner, idx, len(group_elems))
                        try:
                            ok = _apply_optimization(doc, elem, mode, effective_corner)
                            if ok:
                                stats["succeeded"] += 1
                            else:
                                stats["failed"] += 1
                                _bump_reason(
                                    fail_reasons,
                                    fail_samples,
                                    _diagnose_failure(doc, elem, mode),
                                    elem,
                                )
                        except Exception as exc:
                            stats["failed"] += 1
                            msg = str(exc).strip().splitlines()[0] if str(exc).strip() else type(exc).__name__
                            _bump_reason(
                                fail_reasons,
                                fail_samples,
                                "Exception: {}".format(msg[:140]),
                                elem,
                            )
            else:
                for elem in elems_for_type:
                    stats["attempted"] += 1
                    if _is_optimization_disabled(elem, led_params_map):
                        stats["skipped_no_opt"] += 1
                        _bump_reason(
                            skip_reasons,
                            skip_samples,
                            "Optimization=NO for associated profile fixture",
                            elem,
                        )
                        continue
                    try:
                        ok = _apply_optimization(doc, elem, mode, corner)
                        if ok:
                            stats["succeeded"] += 1
                        else:
                            stats["failed"] += 1
                            _bump_reason(
                                fail_reasons,
                                fail_samples,
                                _diagnose_failure(doc, elem, mode),
                                elem,
                            )
                    except Exception as exc:
                        stats["failed"] += 1
                        msg = str(exc).strip().splitlines()[0] if str(exc).strip() else type(exc).__name__
                        _bump_reason(
                            fail_reasons,
                            fail_samples,
                            "Exception: {}".format(msg[:140]),
                            elem,
                        )
        txn.Commit()
    except Exception as exc:
        try:
            txn.RollBack()
        except Exception:
            pass
        forms.alert(
            "Optimization failed and was rolled back.\n\n{}".format(str(exc)),
            title=TITLE,
        )
        return

    summary_lines = [
        "Optimization complete.",
        "",
        "Elements processed: {attempted}".format(**stats),
        "Succeeded: {succeeded}".format(**stats),
        "Failed: {failed}".format(**stats),
        "Skipped: {skipped_no_opt}".format(**stats),
    ]
    _append_reason_lines(summary_lines, "Skip reasons:", skip_reasons, skip_samples)
    _append_reason_lines(summary_lines, "Failure reasons:", fail_reasons, fail_samples)
    forms.alert("\n".join(summary_lines), title=TITLE)


if __name__ == "__main__":
    main()
