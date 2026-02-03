# -*- coding: utf-8 -*-
__title__ = "Space Coils"
__doc__ = "Evenly space selected coil families against a chosen wall direction."

import math

from pyrevit import revit, DB, forms, script


logger = script.get_logger()
doc = revit.doc

WALL_OFFSET_FT = 15.0 / 12.0  # 15 inches


def _is_zero_xy(vec, tol=1e-9):
    return abs(vec.X) + abs(vec.Y) < tol


def _angle_xy(v1, v2):
    dot = v1.X * v2.X + v1.Y * v2.Y
    cross = v1.X * v2.Y - v1.Y * v2.X
    return math.atan2(cross, dot)


def _get_location_point(elem):
    loc = getattr(elem, "Location", None)
    if loc and hasattr(loc, "Point"):
        return loc.Point
    return None


def _get_bbox(elem):
    bbox = None
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(revit.active_view)
        except Exception:
            bbox = None
    return bbox


def _find_spatial_element(doc, point):
    if not point:
        return None
    getter = getattr(doc, "GetSpaceAtPoint", None)
    if callable(getter):
        try:
            space = getter(point)
            if space:
                return space
        except Exception:
            pass
    getter = getattr(doc, "GetRoomAtPoint", None)
    if callable(getter):
        try:
            room = getter(point)
            if room:
                return room
        except Exception:
            pass
    return None


def _bounds_from_spatial(spatial):
    options = DB.SpatialElementBoundaryOptions()
    options.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
    segments = spatial.GetBoundarySegments(options)
    if not segments:
        return None
    points = []
    for seg_list in segments:
        for seg in seg_list:
            curve = seg.GetCurve()
            points.append(curve.GetEndPoint(0))
            points.append(curve.GetEndPoint(1))
    if not points:
        return None
    xs = [pt.X for pt in points]
    ys = [pt.Y for pt in points]
    return min(xs), max(xs), min(ys), max(ys)


def _wall_line(wall):
    loc = getattr(wall, "Location", None)
    if not isinstance(loc, DB.LocationCurve):
        return None
    curve = loc.Curve
    if not isinstance(curve, DB.Line):
        return None
    return curve


def _wall_midpoint(line):
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    return (p0 + p1) * 0.5


def _classify_wall(line):
    direction = line.Direction
    if abs(direction.X) >= abs(direction.Y):
        return "H"  # horizontal wall (east-west)
    return "V"      # vertical wall (north-south)


def _closest_coord(coords, target, side):
    if not coords:
        return None
    if side == "pos":
        candidates = [c for c in coords if c >= target]
        return min(candidates, key=lambda c: c - target) if candidates else None
    candidates = [c for c in coords if c <= target]
    return min(candidates, key=lambda c: target - c) if candidates else None


def _bounds_from_walls(center, direction):
    walls = DB.FilteredElementCollector(doc, revit.active_view.Id) \
        .OfCategory(DB.BuiltInCategory.OST_Walls) \
        .WhereElementIsNotElementType()

    horiz = []
    vert = []
    for wall in walls:
        line = _wall_line(wall)
        if not line:
            continue
        mid = _wall_midpoint(line)
        if _classify_wall(line) == "H":
            horiz.append(mid.Y)
        else:
            vert.append(mid.X)

    if direction in ("North", "South"):
        wall_coord = _closest_coord(horiz, center.Y, "pos" if direction == "North" else "neg")
        left = _closest_coord(vert, center.X, "neg")
        right = _closest_coord(vert, center.X, "pos")
        if wall_coord is None or left is None or right is None:
            return None
        return {
            "axis": "X",
            "perp": "Y",
            "left": min(left, right),
            "right": max(left, right),
            "wall": wall_coord,
        }

    wall_coord = _closest_coord(vert, center.X, "pos" if direction == "East" else "neg")
    left = _closest_coord(horiz, center.Y, "neg")
    right = _closest_coord(horiz, center.Y, "pos")
    if wall_coord is None or left is None or right is None:
        return None
    return {
        "axis": "Y",
        "perp": "X",
        "left": min(left, right),
        "right": max(left, right),
        "wall": wall_coord,
    }


def _resolve_bounds(coils, direction):
    point = _get_location_point(coils[0])
    spatial = _find_spatial_element(doc, point)
    if spatial:
        bounds = _bounds_from_spatial(spatial)
        if bounds:
            min_x, max_x, min_y, max_y = bounds
            if direction in ("North", "South"):
                return {
                    "axis": "X",
                    "perp": "Y",
                    "left": min_x,
                    "right": max_x,
                    "wall": max_y if direction == "North" else min_y,
                    "source": "space",
                }
            return {
                "axis": "Y",
                "perp": "X",
                "left": min_y,
                "right": max_y,
                "wall": max_x if direction == "East" else min_x,
                "source": "space",
            }

    pts = [p for p in (_get_location_point(c) for c in coils) if p]
    if pts:
        cx = sum(p.X for p in pts) / float(len(pts))
        cy = sum(p.Y for p in pts) / float(len(pts))
        cz = sum(p.Z for p in pts) / float(len(pts))
        center = DB.XYZ(cx, cy, cz)
    else:
        center = DB.XYZ(0, 0, 0)
    bounds = _bounds_from_walls(center, direction)
    if bounds:
        bounds["source"] = "walls"
    return bounds


def _target_facing(direction):
    if direction == "North":
        return DB.XYZ(0, 1, 0)
    if direction == "South":
        return DB.XYZ(0, -1, 0)
    if direction == "East":
        return DB.XYZ(1, 0, 0)
    return DB.XYZ(-1, 0, 0)


def _rotate_to_facing(elem, target_vec):
    facing = getattr(elem, "FacingOrientation", None)
    if facing is None:
        return False
    current = DB.XYZ(facing.X, facing.Y, 0)
    target = DB.XYZ(target_vec.X, target_vec.Y, 0)
    if _is_zero_xy(current) or _is_zero_xy(target):
        return False
    angle = _angle_xy(current, target)
    if abs(angle) < 1e-7:
        return True
    loc = getattr(elem, "Location", None)
    if not loc or not hasattr(loc, "Point"):
        return False
    axis = DB.Line.CreateUnbound(loc.Point, DB.XYZ(0, 0, 1))
    try:
        DB.ElementTransformUtils.RotateElement(doc, elem.Id, axis, angle)
        return True
    except Exception:
        return False


def _build_item_data(coils, axis, direction):
    items = []
    for coil in coils:
        bbox = _get_bbox(coil)
        if not bbox:
            logger.warning("Skipping coil {}: no bounding box".format(coil.Id.IntegerValue))
            continue
        center = (bbox.Min + bbox.Max) * 0.5
        if axis == "X":
            length = bbox.Max.X - bbox.Min.X
            center_axis = center.X
            back_coord = bbox.Max.Y if direction == "North" else bbox.Min.Y
        else:
            length = bbox.Max.Y - bbox.Min.Y
            center_axis = center.Y
            back_coord = bbox.Max.X if direction == "East" else bbox.Min.X
        items.append({
            "id": coil.Id,
            "length": length,
            "center_axis": center_axis,
            "back_coord": back_coord,
        })
    return sorted(items, key=lambda d: d["center_axis"])


def main():
    selection = revit.get_selection()
    coils = [el for el in selection if isinstance(el, DB.FamilyInstance)]

    if not coils:
        forms.alert("Select coil family instances before running Space Coils.", exitscript=True)

    direction = forms.CommandSwitchWindow.show(
        ["North", "South", "East", "West"],
        message="Which wall direction should the coils align to?",
    )
    if not direction:
        script.exit()

    bounds = _resolve_bounds(coils, direction)
    if not bounds:
        forms.alert(
            "Could not find bounding walls or a Space/Room for the selection.\n"
            "Try selecting coils inside a Space/Room or ensure nearby walls are visible.",
            exitscript=True,
        )

    axis = bounds["axis"]
    left = bounds["left"]
    right = bounds["right"]
    wall = bounds["wall"]

    target_facing = _target_facing(direction)

    with revit.Transaction("Space Coils"):
        for coil in coils:
            _rotate_to_facing(coil, target_facing)

        doc.Regenerate()

        items = _build_item_data(coils, axis, direction)
        if not items:
            forms.alert("No valid coil bounding boxes were found.", exitscript=True)

        total_len = sum(i["length"] for i in items)
        available = right - left
        gap = (available - total_len) / float(len(items) + 1)
        if available <= 0 or gap < 0:
            forms.alert(
                "Not enough space between the bounding walls to place the coils evenly.",
                exitscript=True,
            )

        if direction in ("North", "East"):
            target_back = wall - WALL_OFFSET_FT
        else:
            target_back = wall + WALL_OFFSET_FT

        cursor = left + gap
        for item in items:
            target_center = cursor + item["length"] * 0.5
            cursor += item["length"] + gap

            delta_axis = target_center - item["center_axis"]
            delta_perp = target_back - item["back_coord"]

            if axis == "X":
                move_vec = DB.XYZ(delta_axis, delta_perp, 0)
            else:
                move_vec = DB.XYZ(delta_perp, delta_axis, 0)

            try:
                DB.ElementTransformUtils.MoveElement(doc, item["id"], move_vec)
            except Exception as ex:
                logger.warning("Failed to move coil {}: {}".format(item["id"], ex))


if __name__ == "__main__":
    main()
