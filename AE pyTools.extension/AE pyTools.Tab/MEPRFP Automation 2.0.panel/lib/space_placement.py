# -*- coding: utf-8 -*-
"""
Pure-Python anchor computation for space placement rules.

Given a ``SpaceGeometry`` (axis-aligned bounding box of a Revit Space
plus the placement-direction door anchors) and a ``PlacementRule``,
``anchor_points()`` returns the list of world points where one LED
copy should be placed *before* per-instance offsets are layered on
top.

Conventions
-----------
- All distances in *feet*. ``inset_inches`` and door offsets in the
  rule are converted on the fly.
- Coordinate axes: project N/S/E/W = +Y/-Y/+E/-X. The rule's anchor
  kinds (``n``, ``s``, ``e``, ``w``, ``ne``, ``nw``, ``se``, ``sw``)
  reference the bounding-box edges/corners along these axes.
- For ``center`` and edge/corner kinds, the returned z is the space's
  floor elevation. The LED's ``offsets[*]`` list lifts the element to
  its mounting height (e.g. 18 in for wall outlets).
- For ``door_relative``, one anchor is returned per door — fulfilling
  the "place at every door" decision. ``door_offset_inches.x`` runs
  along the door's *inward* normal (into the room), and
  ``door_offset_inches.y`` runs along the door's hinge-to-knob axis
  (90° CCW from inward, so positive y is to the door's "left" when
  looking from inside the room out through the door).

The Revit-API edge that builds a ``SpaceGeometry`` from a Revit
``Space`` element lives near the bottom of this module under a
``try/except`` import guard so the pure-logic core can be unit-tested
without Revit assemblies present.
"""

import math

from space_profile_model import (
    PlacementRule,
    KIND_CENTER, KIND_N, KIND_S, KIND_E, KIND_W,
    KIND_NE, KIND_NW, KIND_SE, KIND_SW,
    KIND_DOOR_RELATIVE,
)


# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------

INCHES_PER_FOOT = 12.0


def _in_to_ft(value):
    return float(value or 0.0) / INCHES_PER_FOOT


# ---------------------------------------------------------------------
# Pure-data record
# ---------------------------------------------------------------------

class SpaceGeometry(object):
    """Plain-data view of a Revit Space for placement.

    ``bbox`` is the axis-aligned bounding box of the Space's footprint
    in *feet*: ``((xmin, ymin), (xmax, ymax))``.

    ``floor_z`` is the elevation of the Space's level in *feet*.

    ``door_anchors`` is a list of ``(origin_xy, inward_normal_xy)``
    tuples — one per door bounding the Space. ``origin_xy`` is the
    door's location point (XY only). ``inward_normal_xy`` is a unit
    vector pointing INTO the Space along the door's facing axis.
    """

    __slots__ = ("bbox", "floor_z", "door_anchors", "name", "element_id")

    def __init__(self, bbox=None, floor_z=0.0, door_anchors=None,
                 name="", element_id=None):
        self.bbox = bbox  # ((xmin, ymin), (xmax, ymax)) in feet
        self.floor_z = float(floor_z or 0.0)
        self.door_anchors = list(door_anchors or [])
        self.name = name or ""
        self.element_id = element_id

    @property
    def x_min(self):
        return float(self.bbox[0][0]) if self.bbox else 0.0

    @property
    def y_min(self):
        return float(self.bbox[0][1]) if self.bbox else 0.0

    @property
    def x_max(self):
        return float(self.bbox[1][0]) if self.bbox else 0.0

    @property
    def y_max(self):
        return float(self.bbox[1][1]) if self.bbox else 0.0

    @property
    def x_center(self):
        return (self.x_min + self.x_max) / 2.0

    @property
    def y_center(self):
        return (self.y_min + self.y_max) / 2.0

    def __repr__(self):
        return "<SpaceGeometry id={} name={!r} bbox={} doors={}>".format(
            self.element_id, self.name, self.bbox, len(self.door_anchors)
        )


# ---------------------------------------------------------------------
# Anchor computation (pure logic)
# ---------------------------------------------------------------------

def anchor_points(rule, geom):
    """Return ``[(x, y, z), ...]`` anchor points for ``rule`` in ``geom``.

    For non-door kinds returns a single-point list.
    For ``door_relative``: returns one anchor per door in
    ``geom.door_anchors`` — empty list when the space has none.
    """
    if rule is None or geom is None or geom.bbox is None:
        return []
    if not isinstance(rule, PlacementRule):
        rule = PlacementRule(dict(rule) if isinstance(rule, dict) else {})

    kind = rule.kind
    z = geom.floor_z

    if kind == KIND_CENTER:
        return [(geom.x_center, geom.y_center, z)]

    if kind in (KIND_N, KIND_S, KIND_E, KIND_W):
        return [_edge_point(kind, rule, geom, z)]

    if kind in (KIND_NE, KIND_NW, KIND_SE, KIND_SW):
        return [_corner_point(kind, rule, geom, z)]

    if kind == KIND_DOOR_RELATIVE:
        return _door_relative_points(rule, geom, z)

    # Unknown kind: don't crash placement — return empty so the
    # placement engine skips this LED and reports it.
    return []


def _edge_point(kind, rule, geom, z):
    inset = _in_to_ft(rule.inset_inches)
    if kind == KIND_N:
        return (geom.x_center, geom.y_max - inset, z)
    if kind == KIND_S:
        return (geom.x_center, geom.y_min + inset, z)
    if kind == KIND_E:
        return (geom.x_max - inset, geom.y_center, z)
    if kind == KIND_W:
        return (geom.x_min + inset, geom.y_center, z)
    return (geom.x_center, geom.y_center, z)  # unreachable


def _corner_point(kind, rule, geom, z):
    inset = _in_to_ft(rule.inset_inches)
    if kind == KIND_NE:
        return (geom.x_max - inset, geom.y_max - inset, z)
    if kind == KIND_NW:
        return (geom.x_min + inset, geom.y_max - inset, z)
    if kind == KIND_SE:
        return (geom.x_max - inset, geom.y_min + inset, z)
    if kind == KIND_SW:
        return (geom.x_min + inset, geom.y_min + inset, z)
    return (geom.x_center, geom.y_center, z)  # unreachable


def _door_relative_points(rule, geom, z):
    out = []
    door_x = _in_to_ft(rule.door_offset_x_inches)
    door_y = _in_to_ft(rule.door_offset_y_inches)
    for origin_xy, inward_xy in geom.door_anchors or ():
        ox, oy = float(origin_xy[0]), float(origin_xy[1])
        nx, ny = _normalize_xy(inward_xy)
        # Sideways = inward rotated 90° CCW: (-ny, nx).
        sx, sy = -ny, nx
        x = ox + door_x * nx + door_y * sx
        y = oy + door_x * ny + door_y * sy
        out.append((x, y, z))
    return out


def _normalize_xy(vec):
    if vec is None:
        return (1.0, 0.0)
    vx, vy = float(vec[0]), float(vec[1])
    length = math.sqrt(vx * vx + vy * vy)
    if length < 1e-9:
        return (1.0, 0.0)
    return (vx / length, vy / length)


# ---------------------------------------------------------------------
# Multi-LED expansion
# ---------------------------------------------------------------------

def expand_led_placements(led, geom):
    """Return ``[(x, y, z, rotation_deg), ...]`` for one LED in one space.

    Multiplies the rule's anchor set against the LED's per-instance
    ``offsets`` list. A door-relative LED with 3 doors and 2 offsets
    yields 6 placements; a center LED with 0 offsets yields 1
    placement at the bare anchor.
    """
    rule = led.placement_rule
    anchors = anchor_points(rule, geom)
    if not anchors:
        return []

    offsets = led.offsets or []
    out = []
    if not offsets:
        # No per-instance offsets — one element per anchor at z=anchor.z.
        for ax, ay, az in anchors:
            out.append((ax, ay, az, 0.0))
        return out

    for ax, ay, az in anchors:
        for off in offsets:
            ox = _in_to_ft(off.x_inches)
            oy = _in_to_ft(off.y_inches)
            oz = _in_to_ft(off.z_inches)
            rot = float(off.rotation_deg or 0.0)
            out.append((ax + ox, ay + oy, az + oz, rot))
    return out


# ---------------------------------------------------------------------
# Revit-edge: build SpaceGeometry from a Revit Space element
# ---------------------------------------------------------------------
#
# Wrapped in a try/except so the pure-logic above runs in plain CPython
# tests without pyrevit / Autodesk.Revit.DB on the path.

try:
    import clr  # noqa: F401
    from Autodesk.Revit.DB import (
        BuiltInCategory,
        FilteredElementCollector,
        SpatialElementBoundaryOptions,
    )
    _HAS_REVIT = True
except Exception:  # pragma: no cover -- only true outside Revit
    BuiltInCategory = None
    FilteredElementCollector = None
    SpatialElementBoundaryOptions = None
    _HAS_REVIT = False


def build_space_geometry(doc, space):
    """Build a ``SpaceGeometry`` from a Revit Space.

    Walks the Space's boundary segments to compute the XY bounding box
    and collects every door hosted in a wall bounding the Space.
    Returns ``None`` if the Space is unplaced (no boundary).
    """
    if not _HAS_REVIT:
        raise RuntimeError(
            "build_space_geometry requires Revit; only the pure-logic "
            "anchor_points/expand_led_placements run outside Revit."
        )

    if doc is None or space is None:
        return None

    bbox = _space_bbox(space)
    if bbox is None:
        return None

    floor_z = _space_floor_z(doc, space)
    door_anchors = _space_door_anchors(doc, space)

    name = ""
    try:
        name = str(getattr(space, "Name", "") or "").strip()
    except Exception:
        name = ""

    eid = None
    try:
        eid = _element_id_int(getattr(space, "Id", None))
    except Exception:
        eid = None

    return SpaceGeometry(
        bbox=bbox,
        floor_z=floor_z,
        door_anchors=door_anchors,
        name=name,
        element_id=eid,
    )


def _element_id_int(elem_id):
    if elem_id is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(elem_id, attr)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            return int(v)
        except Exception:
            continue
    return None


def _space_bbox(space):
    """XY axis-aligned bounding box (in feet) computed from boundary segments."""
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None
    if not loops:
        return None

    xs = []
    ys = []
    for loop in loops:
        for seg in loop:
            try:
                curve = seg.GetCurve()
            except Exception:
                continue
            for sample_t in (0.0, 0.25, 0.5, 0.75, 1.0):
                try:
                    pt = curve.Evaluate(sample_t, True)
                except Exception:
                    pt = None
                if pt is None:
                    continue
                xs.append(pt.X)
                ys.append(pt.Y)
    if not xs or not ys:
        return None
    return ((min(xs), min(ys)), (max(xs), max(ys)))


def _space_floor_z(doc, space):
    try:
        lvl_id = space.LevelId
    except Exception:
        return 0.0
    if lvl_id is None:
        return 0.0
    try:
        lvl = doc.GetElement(lvl_id)
    except Exception:
        return 0.0
    if lvl is None:
        return 0.0
    try:
        return float(lvl.Elevation or 0.0)
    except Exception:
        return 0.0


def _space_door_anchors(doc, space):
    """Return ``[(origin_xy, inward_normal_xy), ...]`` for doors at this space.

    Walks the bounding walls reported by ``GetBoundarySegments`` and
    pulls every Door hosted in those walls, filtering to doors whose
    location point is on the bounding loop (so doors hosted in a wall
    that just *touches* this space at a corner are excluded).
    """
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None
    if not loops:
        return []

    wall_ids = set()
    for loop in loops:
        for seg in loop:
            try:
                wid = seg.ElementId
            except Exception:
                wid = None
            if wid is None:
                continue
            wid_int = _element_id_int(wid)
            if wid_int is not None:
                wall_ids.add(wid_int)

    if not wall_ids:
        return []

    # Collect all door instances and filter by Host id.
    doors = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Doors)
        .WhereElementIsNotElementType()
    )

    out = []
    for door in doors:
        host = getattr(door, "Host", None)
        if host is None:
            continue
        host_id = _element_id_int(getattr(host, "Id", None))
        if host_id not in wall_ids:
            continue

        try:
            loc = door.Location
            pt = getattr(loc, "Point", None)
        except Exception:
            pt = None
        if pt is None:
            continue
        origin_xy = (float(pt.X), float(pt.Y))

        # The door's facing orientation tells us which side it opens
        # toward (typically OUT of the room it serves). Flipping it
        # gives an inward-pointing normal.
        try:
            facing = door.FacingOrientation
        except Exception:
            facing = None
        if facing is None:
            continue
        inward = (-float(facing.X), -float(facing.Y))
        out.append((origin_xy, inward))
    return out
