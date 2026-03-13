# -*- coding: utf-8 -*-
__title__ = "Space Coils"
__doc__ = "Place selected coils by wall or centered grid."

import math

from pyrevit import revit, DB, forms, script


logger = script.get_logger()
doc = revit.doc

DEFAULT_WALL_OFFSET_IN = 15.0
STYLE_WALL = "Wall distribution"
STYLE_CENTER = "Center distribution"
SIDE_SPAN_COORD_TOL_FT = 1.0 / 96.0


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


def _bbox_center(bbox):
    return (bbox.Min + bbox.Max) * 0.5


def _bbox_size_xy(bbox):
    return bbox.Max.X - bbox.Min.X, bbox.Max.Y - bbox.Min.Y


def _find_spatial_element(doc, point):
    if not point:
        return None
    def _try_point(pt):
        getter = getattr(doc, "GetSpaceAtPoint", None)
        if callable(getter):
            try:
                space = getter(pt)
                if space:
                    return space
            except Exception:
                pass
        getter = getattr(doc, "GetRoomAtPoint", None)
        if callable(getter):
            try:
                room = getter(pt)
                if room:
                    return room
            except Exception:
                pass
        return None

    spatial = _try_point(point)
    if spatial:
        return spatial

    try:
        view = revit.active_view
        level = getattr(view, "GenLevel", None)
        if level:
            z = level.Elevation + 0.1
            spatial = _try_point(DB.XYZ(point.X, point.Y, z))
            if spatial:
                return spatial
    except Exception:
        pass

    spatial = _try_point(DB.XYZ(point.X, point.Y, 0))
    if spatial:
        return spatial
    return None


def _segment_count(seg_list):
    try:
        return seg_list.Count
    except Exception:
        return len(seg_list) if seg_list else 0


def _loop_points(seg_list):
    pts = []
    count = _segment_count(seg_list)
    if count <= 0:
        return pts
    for i in range(count):
        seg = seg_list[i]
        curve = seg.GetCurve()
        pts.append(curve.GetEndPoint(0))
    try:
        last_seg = seg_list[count - 1]
        pts.append(last_seg.GetCurve().GetEndPoint(1))
    except Exception:
        pass
    return pts


def _poly_area_xy(pts):
    if len(pts) < 3:
        return 0.0
    area = 0.0
    for i in range(len(pts)):
        x1, y1 = pts[i].X, pts[i].Y
        x2, y2 = pts[(i + 1) % len(pts)].X, pts[(i + 1) % len(pts)].Y
        area += (x1 * y2) - (x2 * y1)
    return area * 0.5


def _spatial_boundary_loops(spatial):
    options = DB.SpatialElementBoundaryOptions()
    options.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
    segments = spatial.GetBoundarySegments(options)
    if not segments:
        return []
    return [seg_list for seg_list in segments if _segment_count(seg_list) > 0]


def _primary_boundary_loop(loops):
    if not loops:
        return None
    loop_areas = []
    for seg_list in loops:
        pts = _loop_points(seg_list)
        if not pts:
            continue
        loop_areas.append((abs(_poly_area_xy(pts)), seg_list))
    if loop_areas:
        return max(loop_areas, key=lambda item: item[0])[1]
    return loops[0]


def _full_collinear_span(candidates, chosen, axis):
    if not candidates or not chosen:
        return None, None, None

    if axis == "X":
        target_coord = (chosen["p0"].Y + chosen["p1"].Y) * 0.5
        peers = [
            c for c in candidates
            if abs(((c["p0"].Y + c["p1"].Y) * 0.5) - target_coord) <= SIDE_SPAN_COORD_TOL_FT
        ]
        if not peers:
            peers = [chosen]
        left = min(min(c["p0"].X, c["p1"].X) for c in peers)
        right = max(max(c["p0"].X, c["p1"].X) for c in peers)
        return left, right, target_coord

    target_coord = (chosen["p0"].X + chosen["p1"].X) * 0.5
    peers = [
        c for c in candidates
        if abs(((c["p0"].X + c["p1"].X) * 0.5) - target_coord) <= SIDE_SPAN_COORD_TOL_FT
    ]
    if not peers:
        peers = [chosen]
    left = min(min(c["p0"].Y, c["p1"].Y) for c in peers)
    right = max(max(c["p0"].Y, c["p1"].Y) for c in peers)
    return left, right, target_coord


def _find_spatial_for_coils(coils):
    for coil in coils:
        point = _get_location_point(coil)
        spatial = _find_spatial_element(doc, point)
        if spatial:
            return spatial
        bbox = _get_bbox(coil)
        if bbox:
            spatial = _find_spatial_element(doc, _bbox_center(bbox))
            if spatial:
                return spatial
    return None


def _bounds_box_from_spatial(spatial):
    loops = _spatial_boundary_loops(spatial)
    if not loops:
        return None

    loop_bounds = []
    for seg_list in loops:
        pts = _loop_points(seg_list)
        if not pts:
            continue
        min_x = min(p.X for p in pts)
        max_x = max(p.X for p in pts)
        min_y = min(p.Y for p in pts)
        max_y = max(p.Y for p in pts)
        area = abs(_poly_area_xy(pts))
        loop_bounds.append((area, min_x, max_x, min_y, max_y))

    if loop_bounds:
        _, min_x, max_x, min_y, max_y = max(loop_bounds, key=lambda b: b[0])
    else:
        points = []
        for seg_list in loops:
            for seg in seg_list:
                curve = seg.GetCurve()
                points.append(curve.GetEndPoint(0))
                points.append(curve.GetEndPoint(1))
        if not points:
            return None
        min_x = min(p.X for p in points)
        max_x = max(p.X for p in points)
        min_y = min(p.Y for p in points)
        max_y = max(p.Y for p in points)

    center = None
    try:
        loc = spatial.Location
        if loc and hasattr(loc, "Point"):
            center = loc.Point
    except Exception:
        center = None
    if center is None:
        center = DB.XYZ((min_x + max_x) * 0.5, (min_y + max_y) * 0.5, 0)

    return {
        "min_x": min_x,
        "max_x": max_x,
        "min_y": min_y,
        "max_y": max_y,
        "center": center,
        "source": "space",
    }


def _bounds_from_spatial(spatial, direction):
    loops = _spatial_boundary_loops(spatial)
    if not loops:
        return None

    primary_loop = _primary_boundary_loop(loops)
    if not primary_loop:
        return None

    points = []
    for seg in primary_loop:
        curve = seg.GetCurve()
        points.append(curve.GetEndPoint(0))
        points.append(curve.GetEndPoint(1))
    if not points:
        return None

    center = None
    try:
        loc = spatial.Location
        if loc and hasattr(loc, "Point"):
            center = loc.Point
    except Exception:
        center = None
    if center is None:
        cx = sum(p.X for p in points) / float(len(points))
        cy = sum(p.Y for p in points) / float(len(points))
        center = DB.XYZ(cx, cy, 0)

    candidates = []
    for seg in primary_loop:
        curve = seg.GetCurve()
        if not isinstance(curve, DB.Line):
            continue
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        mid = (p0 + p1) * 0.5
        wall_type = _classify_wall(curve)
        candidates.append({
            "line": curve,
            "p0": p0,
            "p1": p1,
            "mid": mid,
            "type": wall_type,
        })

    if not candidates:
        return None

    if direction in ("North", "South"):
        horiz = [c for c in candidates if c["type"] == "H"]
        if not horiz:
            return None
        if direction == "North":
            horiz = [c for c in horiz if c["mid"].Y >= center.Y]
            if not horiz:
                return None
        else:
            horiz = [c for c in horiz if c["mid"].Y <= center.Y]
            if not horiz:
                return None
        chosen = max(horiz, key=lambda c: c["line"].Length)
        left, right, wall_coord = _full_collinear_span(horiz, chosen, axis="X")
        if left is None or right is None or wall_coord is None:
            return None
        return {
            "axis": "X",
            "perp": "Y",
            "left": left,
            "right": right,
            "wall": wall_coord,
        }

    vert = [c for c in candidates if c["type"] == "V"]
    if not vert:
        return None
    if direction == "East":
        vert = [c for c in vert if c["mid"].X >= center.X]
        if not vert:
            return None
    else:
        vert = [c for c in vert if c["mid"].X <= center.X]
        if not vert:
            return None

    chosen = max(vert, key=lambda c: c["line"].Length)
    left, right, wall_coord = _full_collinear_span(vert, chosen, axis="Y")
    if left is None or right is None or wall_coord is None:
        return None
    return {
        "axis": "Y",
        "perp": "X",
        "left": left,
        "right": right,
        "wall": wall_coord,
    }


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
            horiz.append((line, mid))
        else:
            vert.append((line, mid))

    if direction in ("North", "South"):
        candidates = []
        for line, mid in horiz:
            if direction == "North" and mid.Y >= center.Y:
                candidates.append((line, mid))
            elif direction == "South" and mid.Y <= center.Y:
                candidates.append((line, mid))
        if not candidates:
            return None
        line, mid = max(candidates, key=lambda c: c[0].Length)
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        left = min(p0.X, p1.X)
        right = max(p0.X, p1.X)
        wall_coord = (p0.Y + p1.Y) * 0.5
        return {
            "axis": "X",
            "perp": "Y",
            "left": left,
            "right": right,
            "wall": wall_coord,
        }

    candidates = []
    for line, mid in vert:
        if direction == "East" and mid.X >= center.X:
            candidates.append((line, mid))
        elif direction == "West" and mid.X <= center.X:
            candidates.append((line, mid))
    if not candidates:
        return None
    line, mid = max(candidates, key=lambda c: c[0].Length)
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    left = min(p0.Y, p1.Y)
    right = max(p0.Y, p1.Y)
    wall_coord = (p0.X + p1.X) * 0.5
    return {
        "axis": "Y",
        "perp": "X",
        "left": left,
        "right": right,
        "wall": wall_coord,
    }


def _bounds_from_walls_extreme(center, direction):
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
            horiz.append((line, mid))
        else:
            vert.append((line, mid))

    if direction in ("North", "South"):
        candidates = []
        for line, mid in horiz:
            if direction == "North" and mid.Y >= center.Y:
                candidates.append((line, mid))
            elif direction == "South" and mid.Y <= center.Y:
                candidates.append((line, mid))
        if not candidates:
            return None
        if direction == "North":
            line, mid = max(candidates, key=lambda c: (c[1].Y, c[0].Length))
        else:
            line, mid = min(candidates, key=lambda c: (c[1].Y, -c[0].Length))
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        left = min(p0.X, p1.X)
        right = max(p0.X, p1.X)
        wall_coord = (p0.Y + p1.Y) * 0.5
        return {
            "axis": "X",
            "perp": "Y",
            "left": left,
            "right": right,
            "wall": wall_coord,
        }

    candidates = []
    for line, mid in vert:
        if direction == "East" and mid.X >= center.X:
            candidates.append((line, mid))
        elif direction == "West" and mid.X <= center.X:
            candidates.append((line, mid))
    if not candidates:
        return None
    if direction == "East":
        line, mid = max(candidates, key=lambda c: (c[1].X, c[0].Length))
    else:
        line, mid = min(candidates, key=lambda c: (c[1].X, -c[0].Length))
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    left = min(p0.Y, p1.Y)
    right = max(p0.Y, p1.Y)
    wall_coord = (p0.X + p1.X) * 0.5
    return {
        "axis": "Y",
        "perp": "X",
        "left": left,
        "right": right,
        "wall": wall_coord,
    }


def _resolve_bounds(coils, direction):
    spatial = _find_spatial_for_coils(coils)
    if spatial:
        bounds = _bounds_from_spatial(spatial, direction)
        if bounds:
            bounds["source"] = "space"
            return bounds

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


def _bounds_box_from_walls(center):
    north = _bounds_from_walls_extreme(center, "North")
    south = _bounds_from_walls_extreme(center, "South")
    east = _bounds_from_walls_extreme(center, "East")
    west = _bounds_from_walls_extreme(center, "West")
    if not north or not south or not east or not west:
        return None
    min_x = west["wall"]
    max_x = east["wall"]
    min_y = south["wall"]
    max_y = north["wall"]
    return {
        "min_x": min_x,
        "max_x": max_x,
        "min_y": min_y,
        "max_y": max_y,
        "center": DB.XYZ((min_x + max_x) * 0.5, (min_y + max_y) * 0.5, center.Z),
        "source": "walls",
    }


def _resolve_room_bounds(coils):
    spatial = _find_spatial_for_coils(coils)
    if spatial:
        bounds = _bounds_box_from_spatial(spatial)
        if bounds:
            return bounds

    centers = []
    for coil in coils:
        bbox = _get_bbox(coil)
        if bbox:
            centers.append(_bbox_center(bbox))
    if not centers:
        point = _get_location_point(coils[0])
        if point:
            centers.append(point)
    if centers:
        cx = sum(p.X for p in centers) / float(len(centers))
        cy = sum(p.Y for p in centers) / float(len(centers))
        cz = sum(p.Z for p in centers) / float(len(centers))
        center = DB.XYZ(cx, cy, cz)
    else:
        center = DB.XYZ(0, 0, 0)

    bounds = _bounds_box_from_walls(center)
    if bounds:
        return bounds
    return None


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


def _coil_bbox_data(coil):
    bbox = None
    try:
        bbox = coil.get_BoundingBox(revit.active_view)
    except Exception:
        bbox = None
    if not bbox:
        bbox = _get_bbox(coil)
    if not bbox:
        return None
    center = _bbox_center(bbox)
    size_x, size_y = _bbox_size_xy(bbox)
    loc = _get_location_point(coil)
    symbol_id = None
    try:
        symbol_id = coil.Symbol.Id.IntegerValue
    except Exception:
        symbol_id = None
    return {
        "id": coil.Id,
        "elem": coil,
        "loc": loc,
        "center": center,
        "min_x": bbox.Min.X,
        "max_x": bbox.Max.X,
        "min_y": bbox.Min.Y,
        "max_y": bbox.Max.Y,
        "size_x": size_x,
        "size_y": size_y,
        "symbol_id": symbol_id,
    }


def _prompt_wall_offset():
    raw = forms.ask_for_string(
        prompt="Offset from wall (inches):",
        default=str(DEFAULT_WALL_OFFSET_IN),
        title="Wall Offset",
    )
    if raw is None:
        script.exit()
    try:
        value = float(str(raw).strip())
    except Exception:
        forms.alert("Offset must be a number (inches).", exitscript=True)
    if value < 0:
        forms.alert("Offset must be zero or positive.", exitscript=True)
    return value / 12.0


def _prompt_grid_size():
    options = ["{}x{}".format(r, c) for r in range(1, 6) for c in range(1, 6)]
    picked = forms.CommandSwitchWindow.show(
        options,
        message="Select grid size (rows x columns):",
    )
    if not picked:
        script.exit()
    try:
        parts = picked.lower().split("x")
        rows = int(parts[0])
        cols = int(parts[1])
    except Exception:
        forms.alert("Invalid grid selection.", exitscript=True)
    return rows, cols


def _prompt_mixed_sizes(rows, cols):
    choice = forms.CommandSwitchWindow.show(
        ["No", "Yes"],
        message="Does the {}x{} grid include different coil sizes?".format(rows, cols),
    )
    if not choice:
        script.exit()
    if choice != "Yes":
        return False

    forms.alert(
        (
            "Arrange the selected coils into the desired {}x{} layout.\n"
            "Then click Continue.\n\n"
            "This tells the tool which size belongs in each grid position."
        ).format(rows, cols),
        title="Different Coil Sizes",
    )

    proceed = forms.CommandSwitchWindow.show(
        ["Continue", "Cancel"],
        message="Ready to use current layout as the grid template?",
    )
    if proceed != "Continue":
        script.exit()
    return True


def _place_wall_distribution(coils):
    direction = forms.CommandSwitchWindow.show(
        ["North", "South", "East", "West"],
        message="Which wall direction should the coils align to?",
    )
    if not direction:
        script.exit()

    offset_ft = _prompt_wall_offset()

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
            target_back = wall - offset_ft
        else:
            target_back = wall + offset_ft

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


def _place_center_distribution(coils, rows, cols, use_current_layout_template=False):
    bounds = _resolve_room_bounds(coils)
    if not bounds:
        forms.alert(
            "Could not find a Space/Room or surrounding walls for the selection.\n"
            "Try selecting coils inside a Space/Room or ensure nearby walls are visible.",
            exitscript=True,
        )

    data = []
    for coil in coils:
        item = _coil_bbox_data(coil)
        if not item:
            logger.warning("Skipping coil {}: no bounding box".format(coil.Id.IntegerValue))
            continue
        data.append(item)

    if not data:
        forms.alert("No valid coil bounding boxes were found.", exitscript=True)

    capacity = rows * cols
    if len(data) > capacity:
        forms.alert(
            "Selected {} coil(s), but grid only has {} slots.".format(len(data), capacity),
            exitscript=True,
        )

    min_x = bounds["min_x"]
    max_x = bounds["max_x"]
    min_y = bounds["min_y"]
    max_y = bounds["max_y"]
    available_x = max_x - min_x
    available_y = max_y - min_y
    if available_x <= 0 or available_y <= 0:
        forms.alert("Not enough space between the walls to place the coils.", exitscript=True)

    if rows == 1 and cols == 1 and len(data) == 1:
        item = data[0]
        if item["size_x"] > available_x or item["size_y"] > available_y:
            forms.alert(
                "Not enough space between the walls to place the coil.",
                exitscript=True,
            )
        target_min_x = min_x + (available_x - item["size_x"]) * 0.5
        target_min_y = min_y + (available_y - item["size_y"]) * 0.5
        delta_x = target_min_x - item["min_x"]
        delta_y = target_min_y - item["min_y"]
        with revit.Transaction("Space Coils"):
            try:
                DB.ElementTransformUtils.MoveElement(
                    doc, item["id"], DB.XYZ(delta_x, delta_y, 0)
                )
            except Exception as ex:
                logger.warning("Failed to move coil {}: {}".format(item["id"], ex))
        return

    if use_current_layout_template:
        layout_seed = sorted(data, key=lambda d: (d["center"].Y, d["center"].X))
        move_seed = layout_seed
    else:
        # Preserve legacy behavior: size matrix from incoming order,
        # then assign to row/col targets by current spatial order.
        layout_seed = data[:]
        move_seed = sorted(data, key=lambda d: (d["center"].Y, d["center"].X))

    rows_data = []
    idx = 0
    for _ in range(rows):
        row_items = []
        for _ in range(cols):
            row_items.append(layout_seed[idx] if idx < len(layout_seed) else None)
            idx += 1
        rows_data.append(row_items)

    col_widths = []
    for c in range(cols):
        widths = [item["size_x"] for item in (row[c] for row in rows_data) if item]
        col_widths.append(max(widths) if widths else 0.0)

    row_heights = []
    for r in range(rows):
        row_items = [item for item in rows_data[r] if item]
        heights = [item["size_y"] for item in row_items]
        row_heights.append(max(heights) if heights else 0.0)

    gap_x = (available_x - sum(col_widths)) / float(cols + 1)
    gap_y = (available_y - sum(row_heights)) / float(rows + 1)
    if gap_x < 0 or gap_y < 0:
        forms.alert(
            "Not enough space between the walls to fit the grid with equal spacing.",
            exitscript=True,
        )

    col_min_x = []
    cursor_x = min_x + gap_x
    for w in col_widths:
        col_min_x.append(cursor_x)
        cursor_x += w + gap_x

    row_min_y = []
    cursor_y = min_y + gap_y
    for h in row_heights:
        row_min_y.append(cursor_y)
        cursor_y += h + gap_y

    col_centers = [col_min_x[i] + (col_widths[i] * 0.5) for i in range(cols)]
    row_centers = [row_min_y[i] + (row_heights[i] * 0.5) for i in range(rows)]

    family_offsets_y = {}
    for item in data:
        loc = item.get("loc")
        if not loc:
            continue
        key = item.get("symbol_id")
        if key is None:
            continue
        offset = item["center"].Y - loc.Y
        family_offsets_y.setdefault(key, []).append(offset)
    family_offsets_y = {k: sum(v) / float(len(v)) for k, v in family_offsets_y.items()}

    targets = []
    for row_idx in range(rows):
        for col_idx in range(cols):
            targets.append((row_idx, col_idx))

    with revit.Transaction("Space Coils"):
        for item, target in zip(move_seed, targets):
            row_idx, col_idx = target
            target_min_x = col_min_x[col_idx]
            delta_x = target_min_x - item["min_x"]
            loc = item.get("loc")
            base_y = item["center"].Y
            key = item.get("symbol_id")
            if loc and key in family_offsets_y:
                base_y = loc.Y + family_offsets_y[key]
            target_y = row_centers[row_idx]
            delta_y = target_y - base_y
            logger.info(
                "Center distribution: coil=%s row=%s col=%s locY=%s centerY=%s targetY=%s baseY=%s deltaY=%s",
                item["id"].IntegerValue,
                row_idx,
                col_idx,
                "{:.4f}".format(loc.Y) if loc else "None",
                "{:.4f}".format(item["center"].Y),
                "{:.4f}".format(target_y),
                "{:.4f}".format(base_y),
                "{:.4f}".format(delta_y),
            )
            move_vec = DB.XYZ(delta_x, delta_y, 0)
            try:
                DB.ElementTransformUtils.MoveElement(doc, item["id"], move_vec)
            except Exception as ex:
                logger.warning("Failed to move coil {}: {}".format(item["id"], ex))


def main():
    selection = revit.get_selection()
    coils = [el for el in selection if isinstance(el, DB.FamilyInstance)]

    if not coils:
        forms.alert("Select coil family instances before running Space Coils.", exitscript=True)

    style = forms.CommandSwitchWindow.show(
        [STYLE_WALL, STYLE_CENTER],
        message="How should the coils be placed?",
    )
    if not style:
        script.exit()

    if style == STYLE_WALL:
        _place_wall_distribution(coils)
    else:
        rows, cols = _prompt_grid_size()
        mixed_sizes = _prompt_mixed_sizes(rows, cols)
        _place_center_distribution(
            coils,
            rows,
            cols,
            use_current_layout_template=mixed_sizes,
        )


if __name__ == "__main__":
    main()
