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
    KIND_CENTER,
    KIND_DOOR_RELATIVE,
    KIND_WALL_OPPOSITE_DOOR,
    KIND_WALL_RIGHT_OF_DOOR,
    KIND_WALL_LEFT_OF_DOOR,
    KIND_CORNER_FURTHEST_FROM_DOOR,
    KIND_CORNER_CLOSEST_TO_DOOR,
    KIND_WALL_ANCHORED,
    KIND_SPACE_ANCHORED,
    WALL_ROLE_OPPOSITE_DOOR,
    WALL_ROLE_RIGHT_OF_DOOR,
    WALL_ROLE_LEFT_OF_DOOR,
    WALL_ROLE_BEHIND_DOOR,
    DOOR_DEPENDENT_KINDS,
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

    ``boundary_polygon`` is the outer-loop boundary as an ordered list
    of ``(x, y)`` tuples in feet — the *actual* room shape, not the
    axis-aligned bbox. Used at placement time to clip points that the
    bbox-fraction math lands outside the visible walls (irregular
    rooms, L-shapes, alcoves).
    """

    __slots__ = (
        "bbox", "floor_z", "door_anchors", "name", "element_id",
        "boundary_polygon",
    )

    def __init__(self, bbox=None, floor_z=0.0, door_anchors=None,
                 name="", element_id=None, boundary_polygon=None):
        self.bbox = bbox  # ((xmin, ymin), (xmax, ymax)) in feet
        self.floor_z = float(floor_z or 0.0)
        self.door_anchors = list(door_anchors or [])
        self.name = name or ""
        self.element_id = element_id
        # List[(x, y)] — outer loop of the space boundary in feet.
        # Empty when the boundary could not be extracted.
        self.boundary_polygon = list(boundary_polygon or [])

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

def anchor_points(rule, geom, door_anchor=None):
    """Return ``[(x, y, z), ...]`` anchor points for ``rule`` in ``geom``.

    All non-``center`` kinds depend on a reference door. Pass it in
    via ``door_anchor`` (the ``(origin_xy, inward_xy)`` tuple). When
    omitted, the first entry of ``geom.door_anchors`` is used —
    convenient for unit tests and for spaces with exactly one door.

    Returns an empty list when:

      * ``rule`` / ``geom`` / ``geom.bbox`` is missing.
      * The kind is door-dependent and no door is available
        (caller's cue to emit a comment-only plan).
      * The kind isn't one of the recognised values.
    """
    if rule is None or geom is None or geom.bbox is None:
        return []
    if not isinstance(rule, PlacementRule):
        rule = PlacementRule(dict(rule) if isinstance(rule, dict) else {})

    kind = rule.kind
    z = geom.floor_z

    if kind == KIND_CENTER:
        return [(geom.x_center, geom.y_center, z)]

    # Resolve the reference door. Caller-supplied wins; otherwise
    # fall back to the first door in geom.
    door = door_anchor
    if door is None and geom.door_anchors:
        door = geom.door_anchors[0]
    if door is None:
        return []

    # Normalise the door's inward direction so it always points INTO
    # the space. ``door_to_anchor`` derives ``inward`` by negating
    # ``FacingOrientation``, but that heuristic only holds when the
    # door family was placed with its facing pointing out of the room.
    # Doors placed flipped (or families authored with the opposite
    # convention) come through here with ``inward`` pointing away from
    # the space center, which mirrors every wall- and corner-relative
    # anchor onto the wrong side. The dot-product check is geometry-
    # truthful — flip when the recorded inward and the actual toward-
    # center direction disagree.
    door = _orient_door_inward(door, (geom.x_center, geom.y_center))

    if kind == KIND_DOOR_RELATIVE:
        return [_door_relative_point(rule, door, z)]

    if kind in (
        KIND_WALL_OPPOSITE_DOOR,
        KIND_WALL_RIGHT_OF_DOOR,
        KIND_WALL_LEFT_OF_DOOR,
    ):
        return [_wall_relative_point(kind, rule, geom, door, z)]

    if kind == KIND_WALL_ANCHORED:
        return [_wall_anchored_point(rule, geom, door, z)]

    if kind == KIND_SPACE_ANCHORED:
        return [_space_anchored_point(rule, geom, z, door_anchor=door)]

    if kind in (
        KIND_CORNER_FURTHEST_FROM_DOOR,
        KIND_CORNER_CLOSEST_TO_DOOR,
    ):
        return [_corner_relative_point(kind, rule, geom, door, z)]

    return []


def _door_relative_point(rule, door_anchor, z):
    """Single point at the door, optionally offset along the door's
    inward (door_offset_x) and sideways (door_offset_y) axes."""
    door_x = _in_to_ft(rule.door_offset_x_inches)
    door_y = _in_to_ft(rule.door_offset_y_inches)
    origin_xy, inward_xy = door_anchor
    ox, oy = float(origin_xy[0]), float(origin_xy[1])
    nx, ny = _normalize_xy(inward_xy)
    # Sideways = inward rotated 90° CCW: (-ny, nx).
    sx, sy = -ny, nx
    x = ox + door_x * nx + door_y * sx
    y = oy + door_x * ny + door_y * sy
    return (x, y, z)


def _wall_relative_point(kind, rule, geom, door_anchor, z):
    """Anchor at the midpoint of the wall identified relative to the
    door (opposite / right / left), inset toward the room interior
    by ``rule.inset_inches``.

    "Right" / "left" are taken from the perspective of someone
    standing in the doorway facing into the room (i.e. along the
    door's inward normal). 90° clockwise from inward is "right".
    """
    inset = _in_to_ft(rule.inset_inches)
    _door_xy, inward_xy = door_anchor
    nx, ny = _normalize_xy(inward_xy)
    xmin = geom.x_min
    xmax = geom.x_max
    ymin = geom.y_min
    ymax = geom.y_max
    cx = geom.x_center
    cy = geom.y_center

    # Pick which bbox edge is nearest each cardinal of the door's
    # frame. ``axis`` is which world axis the inward normal aligns
    # most strongly with.
    if abs(nx) > abs(ny):
        # Door wall is on the X axis (east or west).
        if nx > 0:
            # Door on west wall (inward = +X). Right (90° CW from +X) = -Y.
            opposite = (xmax - inset, cy, z)
            right = (cx, ymin + inset, z)
            left = (cx, ymax - inset, z)
        else:
            # Door on east wall (inward = -X). Right (90° CW from -X) = +Y.
            opposite = (xmin + inset, cy, z)
            right = (cx, ymax - inset, z)
            left = (cx, ymin + inset, z)
    else:
        if ny > 0:
            # Door on south wall (inward = +Y). Right (90° CW from +Y) = +X.
            opposite = (cx, ymax - inset, z)
            right = (xmax - inset, cy, z)
            left = (xmin + inset, cy, z)
        else:
            # Door on north wall (inward = -Y). Right (90° CW from -Y) = -X.
            opposite = (cx, ymin + inset, z)
            right = (xmin + inset, cy, z)
            left = (xmax - inset, cy, z)

    if kind == KIND_WALL_OPPOSITE_DOOR:
        return opposite
    if kind == KIND_WALL_RIGHT_OF_DOOR:
        return right
    if kind == KIND_WALL_LEFT_OF_DOOR:
        return left
    return (cx, cy, z)  # unreachable


def wall_segments_for_door(geom, door_anchor):
    """Resolve the four bbox walls of ``geom`` to door-relative roles.

    Returns ``{role: ((sx, sy), (ex, ey), inward_xy)}`` where role is
    one of ``WALL_ROLE_OPPOSITE_DOOR / RIGHT_OF_DOOR / LEFT_OF_DOOR /
    BEHIND_DOOR``. ``(sx, sy)`` and ``(ex, ey)`` are the wall's
    endpoints in world feet, ordered so that ``position_along_wall=0``
    is at the start and ``position_along_wall=1`` is at the end. The
    inward direction is the unit vector pointing INTO the space from
    the wall surface — used for ``distance_from_wall_inches`` offsets.

    The endpoint orientation is **stable across geometry changes**:
    for the wall opposite the door we sweep in the same direction as
    "right_of_door" (so position_along_wall=0 sits on the right side),
    and for behind_door we sweep from left endpoint to right endpoint
    relative to someone standing outside the door looking in.
    Capture-side code uses this same helper to compute the fraction,
    so capture and placement agree.
    """
    if geom is None or door_anchor is None:
        return {}
    _door_xy, inward_xy = door_anchor
    nx, ny = _normalize_xy(inward_xy)
    xmin = geom.x_min
    xmax = geom.x_max
    ymin = geom.y_min
    ymax = geom.y_max

    # Each wall yields ((start), (end), inward_unit_xy). The wall
    # direction is chosen so that position_along_wall increases from
    # the door's "right" toward the door's "left" as you stand
    # inside the room facing the door.
    if abs(nx) > abs(ny):
        # Door wall is east or west; opposite wall runs along Y.
        if nx > 0:
            # Door on west wall (inward = +X).
            # Opposite (east, x=xmax): sweep from ymin (right) to
            # ymax (left). Right wall (south, y=ymin): sweep from
            # xmin (door) to xmax (far). Left wall (north, y=ymax):
            # sweep from xmin (door) to xmax (far). Behind wall is
            # the door's own wall; sweep ymin->ymax.
            return {
                WALL_ROLE_OPPOSITE_DOOR: (
                    (xmax, ymin), (xmax, ymax), (-1.0, 0.0),
                ),
                WALL_ROLE_RIGHT_OF_DOOR: (
                    (xmin, ymin), (xmax, ymin), (0.0, 1.0),
                ),
                WALL_ROLE_LEFT_OF_DOOR: (
                    (xmin, ymax), (xmax, ymax), (0.0, -1.0),
                ),
                WALL_ROLE_BEHIND_DOOR: (
                    (xmin, ymin), (xmin, ymax), (1.0, 0.0),
                ),
            }
        else:
            # Door on east wall (inward = -X).
            return {
                WALL_ROLE_OPPOSITE_DOOR: (
                    (xmin, ymax), (xmin, ymin), (1.0, 0.0),
                ),
                WALL_ROLE_RIGHT_OF_DOOR: (
                    (xmax, ymax), (xmin, ymax), (0.0, -1.0),
                ),
                WALL_ROLE_LEFT_OF_DOOR: (
                    (xmax, ymin), (xmin, ymin), (0.0, 1.0),
                ),
                WALL_ROLE_BEHIND_DOOR: (
                    (xmax, ymax), (xmax, ymin), (-1.0, 0.0),
                ),
            }
    else:
        if ny > 0:
            # Door on south wall (inward = +Y).
            return {
                WALL_ROLE_OPPOSITE_DOOR: (
                    (xmax, ymax), (xmin, ymax), (0.0, -1.0),
                ),
                WALL_ROLE_RIGHT_OF_DOOR: (
                    (xmax, ymin), (xmax, ymax), (-1.0, 0.0),
                ),
                WALL_ROLE_LEFT_OF_DOOR: (
                    (xmin, ymin), (xmin, ymax), (1.0, 0.0),
                ),
                WALL_ROLE_BEHIND_DOOR: (
                    (xmax, ymin), (xmin, ymin), (0.0, 1.0),
                ),
            }
        else:
            # Door on north wall (inward = -Y).
            return {
                WALL_ROLE_OPPOSITE_DOOR: (
                    (xmin, ymin), (xmax, ymin), (0.0, 1.0),
                ),
                WALL_ROLE_RIGHT_OF_DOOR: (
                    (xmin, ymax), (xmin, ymin), (1.0, 0.0),
                ),
                WALL_ROLE_LEFT_OF_DOOR: (
                    (xmax, ymax), (xmax, ymin), (-1.0, 0.0),
                ),
                WALL_ROLE_BEHIND_DOOR: (
                    (xmin, ymax), (xmax, ymax), (0.0, -1.0),
                ),
            }


def _wall_anchored_point(rule, geom, door_anchor, z):
    """Resolve a wall_anchored placement: pick the wall by
    ``rule.wall_role``, walk ``rule.position_along_wall`` of the way
    from start to end, then push inward by
    ``rule.distance_from_wall_inches``.

    Returns ``(x, y, z)``. Falls back to the space center on any
    malformed input so a bad capture doesn't crash the run.
    """
    walls = wall_segments_for_door(geom, door_anchor)
    if not walls:
        return (geom.x_center, geom.y_center, z)
    role = rule.wall_role
    seg = walls.get(role) or walls.get(WALL_ROLE_OPPOSITE_DOOR)
    if seg is None:
        return (geom.x_center, geom.y_center, z)
    (sx, sy), (ex, ey), (inx, iny) = seg
    t = max(0.0, min(1.0, float(rule.position_along_wall)))
    px = sx + t * (ex - sx)
    py = sy + t * (ey - sy)
    inset = _in_to_ft(rule.distance_from_wall_inches)
    return (px + inx * inset, py + iny * inset, z)


_WALL_SNAP_THRESHOLD_FT = 3.0  # 36"; fixtures captured within this
                               # distance of a wall snap to that wall
                               # in the target space.

# When a bbox-fraction point lands outside the actual boundary
# polygon, we project it back to the nearest edge. This margin is
# the *inward* offset applied after projection — keeps the placed
# fixture a hair inside the wall instead of dead on the centerline.
_POLYGON_CLIP_INSET_FT = 0.05  # ~5/8"


def _point_in_polygon(pt, poly):
    """Ray-cast point-in-polygon test for an unclosed polygon.

    ``poly`` is a list of ``(x, y)`` tuples; the implicit closing edge
    connects the last point back to the first. Returns ``True`` for
    strictly interior points; behaviour on edges is implementation-
    defined but consistent.
    """
    if not poly or len(poly) < 3:
        return False
    x, y = pt[0], pt[1]
    n = len(poly)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = poly[i]
        xj, yj = poly[j]
        # Standard ray-cast: count edges that cross the horizontal
        # ray to the right of (x, y). Half-open intervals on y avoid
        # double-counting vertices.
        if ((yi > y) != (yj > y)):
            x_intersect = (xj - xi) * (y - yi) / (yj - yi) + xi
            if x < x_intersect:
                inside = not inside
        j = i
    return inside


def _project_to_polygon_edge(pt, poly):
    """Return ``((qx, qy), inward_unit, dist_ft)`` for the closest
    point on the polygon boundary to ``pt``.

    The boundary is traversed clockwise or counter-clockwise depending
    on Revit's convention; ``inward_unit`` is the unit vector
    perpendicular to the closest edge, pointing into the polygon
    interior. ``dist_ft`` is the distance from ``pt`` to the projection.

    Used to clip out-of-polygon placements onto the nearest wall.
    """
    if not poly or len(poly) < 2:
        return None
    px, py = pt[0], pt[1]
    best = None
    n = len(poly)
    for i in range(n):
        ax, ay = poly[i]
        bx, by = poly[(i + 1) % n]
        dx = bx - ax
        dy = by - ay
        length_sq = dx * dx + dy * dy
        if length_sq < 1e-12:
            continue
        t = ((px - ax) * dx + (py - ay) * dy) / length_sq
        t_clamped = max(0.0, min(1.0, t))
        qx = ax + t_clamped * dx
        qy = ay + t_clamped * dy
        ddx = px - qx
        ddy = py - qy
        d2 = ddx * ddx + ddy * ddy
        if best is None or d2 < best[0]:
            # Edge unit normal: rotate (dx, dy) by 90°. We don't yet
            # know which side is the interior — caller resolves it
            # against the polygon centroid below.
            length = math.sqrt(length_sq)
            nx = -dy / length
            ny = dx / length
            best = (d2, qx, qy, nx, ny)
    if best is None:
        return None
    d2, qx, qy, nx, ny = best
    # Orient the normal so it points into the polygon. Use the centroid
    # as the "inside" reference — simple but robust for convex and most
    # concave shapes.
    cx = sum(p[0] for p in poly) / float(n)
    cy = sum(p[1] for p in poly) / float(n)
    if ((cx - qx) * nx + (cy - qy) * ny) < 0.0:
        nx = -nx
        ny = -ny
    return ((qx, qy), (nx, ny), math.sqrt(d2))


def _space_anchored_point(rule, geom, z, door_anchor=None):
    """Resolve a space_anchored placement.

    Three-stage resolution so wall-mounted fixtures don't drift
    outside the room when the target space's bbox doesn't tightly
    match the visible walls (irregular shapes, L-rooms, alcoves):

    1. **Bbox-fraction**: compute a world-XY point from
       ``x_fraction`` / ``y_fraction`` against the target's bbox —
       pure room-relative scaling.

    2. **Polygon clip** (always, when the boundary polygon is
       available): if the bbox-fraction point lands OUTSIDE the
       actual boundary polygon of the target space, project it
       onto the nearest boundary edge and push inward by a small
       constant inset. The bbox is the axis-aligned hull of the
       boundary, so any out-of-polygon point sits in the bbox's
       "extra" area — exactly the case where the user sees fixtures
       land "outside the new space".

    3. **Wall-snap** (when ``wall_role`` is set AND the captured
       ``distance_from_wall_inches`` is small enough to indicate
       wall-mounted): project the (post-clip) point onto the wall
       segment identified by ``wall_role`` in the target space's
       bbox and push inward by the captured perpendicular distance.

    Fixtures captured well off the wall (large
    ``distance_from_wall_inches``) skip step 3 and stay at the
    (possibly clipped) bbox-fraction point — appropriate for free-
    standing or ceiling-mounted gear.
    """
    if geom is None or geom.bbox is None:
        return (0.0, 0.0, z)
    fx = max(0.0, min(1.0, float(rule.x_fraction)))
    fy = max(0.0, min(1.0, float(rule.y_fraction)))
    width = geom.x_max - geom.x_min
    height = geom.y_max - geom.y_min
    bbox_pt = (
        geom.x_min + fx * width,
        geom.y_min + fy * height,
    )
    target_x, target_y = bbox_pt[0], bbox_pt[1]

    role = rule.wall_role if isinstance(rule, PlacementRule) else None
    dist_in_ft = (
        _in_to_ft(rule.distance_from_wall_inches)
        if isinstance(rule, PlacementRule) else 0.0
    )

    # ---- stage 2: polygon clip ---------------------------------
    polygon = getattr(geom, "boundary_polygon", None) or []
    clipped = False
    clip_dist_ft = 0.0
    if polygon and len(polygon) >= 3:
        if not _point_in_polygon((target_x, target_y), polygon):
            proj = _project_to_polygon_edge((target_x, target_y), polygon)
            if proj is not None:
                (qx, qy), (nx, ny), clip_dist_ft = proj
                # Inset inward by the captured perpendicular distance
                # when available (for wall-mounted fixtures), otherwise
                # a small constant inset just to keep the placement off
                # the wall centerline.
                if role and abs(dist_in_ft) < _WALL_SNAP_THRESHOLD_FT:
                    inset = max(_POLYGON_CLIP_INSET_FT, dist_in_ft)
                else:
                    inset = _POLYGON_CLIP_INSET_FT
                target_x = qx + nx * inset
                target_y = qy + ny * inset
                clipped = True

    # ---- stage 3: wall-snap (perpendicular offset) -------------
    snapped = False
    wall_along_t = None
    if (
        not clipped  # don't double-apply offset after polygon-clip
        and role
        and door_anchor is not None
        and abs(dist_in_ft) < _WALL_SNAP_THRESHOLD_FT
    ):
        walls = wall_segments_for_door(geom, door_anchor)
        seg = walls.get(role)
        if seg is not None:
            (sx, sy), (ex, ey), (inx, iny) = seg
            dx = ex - sx
            dy = ey - sy
            length_sq = dx * dx + dy * dy
            if length_sq > 1e-12:
                t = (
                    (target_x - sx) * dx + (target_y - sy) * dy
                ) / length_sq
                t_clamped = max(0.0, min(1.0, t))
                wall_x = sx + t_clamped * dx
                wall_y = sy + t_clamped * dy
                target_x = wall_x + inx * dist_in_ft
                target_y = wall_y + iny * dist_in_ft
                wall_along_t = t_clamped
                snapped = True

    target = (target_x, target_y, z)

    # Diagnostic — appended to a module-level list the workflow drains
    # so each placement plan's preview row can show the bbox the engine
    # actually used, the polygon-clip decision, and the wall-snap step.
    try:
        notes = []
        if clipped:
            notes.append(
                "CLIP({:.1f}\")".format(clip_dist_ft * INCHES_PER_FOOT)
            )
        if snapped:
            notes.append(
                "SNAP[{}]@t={:.2f}(in {:.1f}\")".format(
                    role, wall_along_t, dist_in_ft * INCHES_PER_FOOT,
                )
            )
        if not notes:
            notes.append("raw-bbox")
        _SPACE_ANCHORED_DIAG.append(
            "target bbox W={:.1f}' H={:.1f}' "
            "X=[{:.1f},{:.1f}] Y=[{:.1f},{:.1f}] poly={}pts | "
            "fx={:.3f} fy={:.3f} | {} -> ({:.2f}, {:.2f})".format(
                width, height,
                geom.x_min, geom.x_max,
                geom.y_min, geom.y_max,
                len(polygon),
                fx, fy,
                ",".join(notes),
                target[0], target[1],
            )
        )
    except Exception:
        pass
    return target


# Per-run diagnostic buffer for space_anchored anchor resolution. The
# placement workflow drains this into ``result.warnings`` after each
# Place Space Elements run so the user can sanity-check that the
# captured fractions and target bbox dimensions are producing the
# expected world coordinates.
_SPACE_ANCHORED_DIAG = []


def drain_space_anchored_diagnostics():
    """Return and clear the per-run space_anchored diagnostic buffer."""
    global _SPACE_ANCHORED_DIAG
    out = list(_SPACE_ANCHORED_DIAG)
    _SPACE_ANCHORED_DIAG = []
    return out


def space_fractions_for_point(geom, point_xy):
    """Reverse of ``_space_anchored_point``: given a world-XY point
    inside ``geom``, return ``(x_fraction, y_fraction)`` clamped to
    [0, 1]. Used by the capture engine to translate a picked
    element's location into the bbox-relative storage shape.
    Returns ``None`` if the geometry has no usable bbox.
    """
    if geom is None or geom.bbox is None:
        return None
    width = geom.x_max - geom.x_min
    height = geom.y_max - geom.y_min
    if width < 1e-9 or height < 1e-9:
        return None
    px, py = float(point_xy[0]), float(point_xy[1])
    fx = (px - geom.x_min) / width
    fy = (py - geom.y_min) / height
    return (max(0.0, min(1.0, fx)), max(0.0, min(1.0, fy)))


def wall_inward_angle_deg(rule, geom, door_anchor):
    """Return the wall's inward-direction angle in degrees for a
    ``wall_anchored`` rule, or 0.0 for any other kind.

    The angle is taken with the standard math convention: 0° = +X
    (east), 90° = +Y (north). Used by the placement engine to spin
    a wall-anchored LED so its captured-relative-to-wall orientation
    survives a source→target wall transplant: a fixture captured on
    an east wall (inward = west, 180°) gets re-faced to north (90°)
    when placed on a south wall, and so on.
    """
    if rule is None or geom is None:
        return 0.0
    kind = rule.kind if isinstance(rule, PlacementRule) else (rule or {}).get("kind")
    if kind not in (KIND_WALL_ANCHORED, KIND_SPACE_ANCHORED):
        return 0.0
    door = door_anchor
    if door is None and geom.door_anchors:
        door = geom.door_anchors[0]
    if door is None:
        return 0.0
    door = _orient_door_inward(door, (geom.x_center, geom.y_center))
    walls = wall_segments_for_door(geom, door)
    role = rule.wall_role if isinstance(rule, PlacementRule) else (rule or {}).get("wall_role")
    seg = walls.get(role) or walls.get(WALL_ROLE_OPPOSITE_DOOR)
    if seg is None:
        return 0.0
    _, _, (inx, iny) = seg
    return math.degrees(math.atan2(iny, inx))


def closest_wall_for_point(geom, door_anchor, point_xy):
    """Reverse of ``_wall_anchored_point``: given a world-XY point
    inside ``geom``, return ``(role, fraction, distance_in)`` where
    ``role`` is the closest wall's role, ``fraction`` is the
    ``position_along_wall`` value, and ``distance_in`` is the
    perpendicular distance from the wall in INCHES (positive = inward).

    Returns ``None`` if the geometry has no resolvable walls.
    """
    walls = wall_segments_for_door(geom, door_anchor)
    if not walls:
        return None
    px, py = float(point_xy[0]), float(point_xy[1])
    best = None
    for role, ((sx, sy), (ex, ey), (inx, iny)) in walls.items():
        dx = ex - sx
        dy = ey - sy
        L2 = dx * dx + dy * dy
        if L2 < 1e-12:
            continue
        # Project (point - start) onto (end - start) and clamp to
        # [0, 1] so the fraction is bounded.
        t = ((px - sx) * dx + (py - sy) * dy) / L2
        t_clamped = max(0.0, min(1.0, t))
        # Foot on the segment.
        fx = sx + t_clamped * dx
        fy = sy + t_clamped * dy
        # Perpendicular signed distance — positive when on the
        # inward side, negative when outside (rare, would mean the
        # point sits past the wall plane).
        # Inward unit vector dotted into (point - foot) gives the
        # signed inward distance in feet.
        signed_in_ft = (px - fx) * inx + (py - fy) * iny
        # Score: combined perpendicular distance (smaller = closer
        # to wall surface) — that's what determines "nearest wall"
        # for an arbitrary point.
        perp = abs(signed_in_ft)
        if best is None or perp < best[0]:
            best = (perp, role, t_clamped, signed_in_ft * 12.0)
    if best is None:
        return None
    _perp, role, t, distance_in = best
    return role, t, distance_in


def _corner_relative_point(kind, rule, geom, door_anchor, z):
    """Anchor at the bbox corner that's closest to / furthest from
    the door, inset diagonally toward the room interior."""
    inset = _in_to_ft(rule.inset_inches)
    (door_x, door_y), _inward = door_anchor
    xmin = geom.x_min
    xmax = geom.x_max
    ymin = geom.y_min
    ymax = geom.y_max
    cx = geom.x_center
    cy = geom.y_center

    corners = [
        (xmin, ymin),
        (xmin, ymax),
        (xmax, ymin),
        (xmax, ymax),
    ]

    def _dist_sq(c):
        return (c[0] - door_x) ** 2 + (c[1] - door_y) ** 2

    if kind == KIND_CORNER_CLOSEST_TO_DOOR:
        target = min(corners, key=_dist_sq)
    else:  # KIND_CORNER_FURTHEST_FROM_DOOR
        target = max(corners, key=_dist_sq)

    # Inset diagonally toward the room center.
    tx, ty = target
    tx += inset if tx < cx else -inset
    ty += inset if ty < cy else -inset
    return (tx, ty, z)


def _normalize_xy(vec):
    if vec is None:
        return (1.0, 0.0)
    vx, vy = float(vec[0]), float(vec[1])
    length = math.sqrt(vx * vx + vy * vy)
    if length < 1e-9:
        return (1.0, 0.0)
    return (vx / length, vy / length)


def _orient_door_inward(door_anchor, center_xy):
    """Return ``door_anchor`` with ``inward`` flipped if it points away
    from ``center_xy``.

    Used by ``anchor_points`` to make door-relative geometry insensitive
    to which side the door's family was placed on. A zero or near-zero
    dot product (door at the centroid, or inward perpendicular to the
    toward-center direction) leaves the original orientation alone —
    we can't tell which side is "in" in that degenerate case, so we
    don't guess.
    """
    if door_anchor is None or center_xy is None:
        return door_anchor
    try:
        origin_xy, inward_xy = door_anchor
        ox, oy = float(origin_xy[0]), float(origin_xy[1])
        ix, iy = float(inward_xy[0]), float(inward_xy[1])
        cx, cy = float(center_xy[0]), float(center_xy[1])
    except (TypeError, ValueError, IndexError):
        return door_anchor
    vx = cx - ox
    vy = cy - oy
    if (ix * vx + iy * vy) < 0.0:
        return ((ox, oy), (-ix, -iy))
    return door_anchor


# ---------------------------------------------------------------------
# Multi-LED expansion
# ---------------------------------------------------------------------

def expand_led_placements(led, geom, door_anchor=None):
    """Return ``[(x, y, z, rotation_deg), ...]`` for one LED in one space.

    Multiplies the rule's anchor set against the LED's per-instance
    ``offsets`` list. ``door_anchor`` is the user-chosen reference
    door for door-dependent kinds; defaults to the first door of
    ``geom`` when omitted.

    Rotation handling: for wall_anchored LEDs, the offset's
    ``rotation_deg`` is interpreted as a delta from the target wall's
    inward direction (0° = facing into the space perpendicular to
    the wall). We add ``wall_inward_angle_deg`` so the placed
    instance always faces into the space regardless of which wall
    role resolved to in this project. For non-wall-anchored kinds
    the base angle is 0° and the captured rotation lands as-is.
    """
    rule = led.placement_rule
    anchors = anchor_points(rule, geom, door_anchor=door_anchor)
    if not anchors:
        return []

    base_rot_deg = wall_inward_angle_deg(rule, geom, door_anchor)

    offsets = led.offsets or []
    out = []
    if not offsets:
        # No per-instance offsets — one element per anchor at z=anchor.z.
        for ax, ay, az in anchors:
            out.append((ax, ay, az, base_rot_deg))
        return out

    for ax, ay, az in anchors:
        for off in offsets:
            ox = _in_to_ft(off.x_inches)
            oy = _in_to_ft(off.y_inches)
            oz = _in_to_ft(off.z_inches)
            rot = float(off.rotation_deg or 0.0) + base_rot_deg
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
        RevitLinkInstance,
        SpatialElementBoundaryOptions,
        XYZ,
    )
    _HAS_REVIT = True
except Exception:  # pragma: no cover -- only true outside Revit
    BuiltInCategory = None
    FilteredElementCollector = None
    RevitLinkInstance = None
    SpatialElementBoundaryOptions = None
    XYZ = None
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
    boundary_polygon = _space_boundary_polygon(space)

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
        boundary_polygon=boundary_polygon,
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


def _space_boundary_polygon(space):
    """Return the outer-loop boundary as an ordered list of ``(x, y)``
    tuples in feet — the *actual* room shape.

    Walks the FIRST boundary loop (the outer perimeter, per Revit's
    convention) and samples each curve at 5 t-values so arcs round-trip
    as polylines. Returns an empty list if the space has no boundary.
    """
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None
    if not loops:
        return []
    # The first loop in Revit's GetBoundarySegments is the outer
    # boundary. Inner loops (holes) are ignored — they don't define
    # where the room *is*, just where it has cutouts (columns, etc.).
    try:
        outer = list(loops[0])
    except Exception:
        return []
    pts = []
    for seg in outer:
        try:
            curve = seg.GetCurve()
        except Exception:
            continue
        # Sample each curve along its length. Arcs need the extra
        # interior samples; straight lines could get by with just the
        # endpoints but the redundancy is harmless.
        for sample_t in (0.0, 0.25, 0.5, 0.75):
            try:
                pt = curve.Evaluate(sample_t, True)
            except Exception:
                pt = None
            if pt is None:
                continue
            pts.append((float(pt.X), float(pt.Y)))
    # Drop consecutive duplicates so the polygon edge list stays clean
    # for projection math.
    cleaned = []
    for p in pts:
        if not cleaned or (
            abs(cleaned[-1][0] - p[0]) > 1e-6
            or abs(cleaned[-1][1] - p[1]) > 1e-6
        ):
            cleaned.append(p)
    if cleaned and (
        abs(cleaned[0][0] - cleaned[-1][0]) < 1e-6
        and abs(cleaned[0][1] - cleaned[-1][1]) < 1e-6
    ):
        cleaned.pop()
    return cleaned


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

    Two-tier search:

      1. Doors *hosted* by walls that ``GetBoundarySegments`` reports
         as bounding this Space. Cheapest path; works for projects
         where architecture lives in the host doc.
      2. Doors in any *linked* Revit instance whose location, after
         transforming through the link's ``GetTotalTransform()``,
         falls within ~1 ft of one of the Space's boundary curves.
         Required when architecture is in a linked model — the
         host wall id check fails because the door's host wall lives
         in the link's id namespace, not this doc's.

    Both tiers are unioned. Duplicates (same physical door appearing
    in host + a link) are deduplicated by location proximity.
    """
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None
    if not loops:
        return []

    wall_ids = set()
    boundary_curves = []
    for loop in loops:
        for seg in loop:
            try:
                wid = seg.ElementId
            except Exception:
                wid = None
            if wid is not None:
                wid_int = _element_id_int(wid)
                # InvalidElementId reports as -1; ignore it.
                if wid_int is not None and wid_int > 0:
                    wall_ids.add(wid_int)
            try:
                curve = seg.GetCurve()
            except Exception:
                curve = None
            if curve is not None:
                boundary_curves.append(curve)

    out = []

    # ----- Tier 1: host doors hosted in our boundary walls ----------
    if wall_ids:
        try:
            host_doors = (
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .WhereElementIsNotElementType()
            )
        except Exception:
            host_doors = []
        for door in host_doors:
            host = getattr(door, "Host", None)
            if host is None:
                continue
            host_id = _element_id_int(getattr(host, "Id", None))
            if host_id not in wall_ids:
                continue
            anchor = door_to_anchor(door, transform=None)
            if anchor is not None:
                out.append(anchor)

    # ----- Tier 2: linked doors near the space's boundary -----------
    if boundary_curves and RevitLinkInstance is not None:
        try:
            link_collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
            link_instances = list(link_collector)
        except Exception:
            link_instances = []
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                continue
            try:
                transform = link.GetTotalTransform()
            except Exception:
                transform = None
            try:
                linked_doors = (
                    FilteredElementCollector(link_doc)
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue
            for door in linked_doors:
                anchor = door_to_anchor(door, transform=transform)
                if anchor is None:
                    continue
                origin_xy = anchor[0]
                if not _point_near_any_curve(origin_xy, boundary_curves, tol=1.0):
                    continue
                if _origin_already_seen(origin_xy, out, tol=0.1):
                    continue
                out.append(anchor)

    return out


def door_to_anchor(door, transform=None):
    """Return ``(origin_xy, inward_xy)`` for a single Door element.

    ``transform`` is the link's total transform when the door lives
    in a linked doc; ``None`` for host doors. Returns ``None`` when
    the door has no Location.Point or no FacingOrientation.

    Public so the pre-placement door-picker (which calls
    ``Selection.PickObject``) can resolve the picked Reference to
    the same anchor tuple shape the workflow expects.
    """
    try:
        loc = door.Location
        pt = getattr(loc, "Point", None)
    except Exception:
        pt = None
    if pt is None:
        return None
    try:
        facing = door.FacingOrientation
    except Exception:
        facing = None
    if facing is None:
        return None
    if transform is not None:
        try:
            pt = transform.OfPoint(pt)
        except Exception:
            pass
        try:
            facing = transform.OfVector(facing)
        except Exception:
            pass
    origin_xy = (float(pt.X), float(pt.Y))
    inward = (-float(facing.X), -float(facing.Y))
    return (origin_xy, inward)


def _point_near_any_curve(point_xy, curves, tol=1.0):
    """True if ``point_xy`` (X, Y in feet) is within ``tol`` ft of
    the closest point on any of ``curves`` (boundary segments)."""
    if not curves or XYZ is None:
        return False
    px, py = point_xy
    for curve in curves:
        try:
            # Use the curve's start Z so Curve.Project doesn't
            # disqualify a coplanar test point on the Z axis.
            start = curve.Evaluate(0.0, True)
            test = XYZ(px, py, start.Z if start is not None else 0.0)
        except Exception:
            continue
        try:
            res = curve.Project(test)
        except Exception:
            res = None
        if res is None:
            continue
        try:
            cp = res.XYZPoint
        except Exception:
            cp = None
        if cp is None:
            continue
        dx = cp.X - px
        dy = cp.Y - py
        if (dx * dx + dy * dy) <= (tol * tol):
            return True
    return False


def _origin_already_seen(point_xy, anchors, tol=0.1):
    """True if any existing anchor's origin is within ``tol`` ft of
    ``point_xy`` — used to dedupe host vs linked sightings of the
    same physical door."""
    px, py = point_xy
    for (origin_xy, _inward) in anchors or ():
        try:
            ox, oy = origin_xy
        except Exception:
            continue
        dx = ox - px
        dy = oy - py
        if (dx * dx + dy * dy) <= (tol * tol):
            return True
    return False
