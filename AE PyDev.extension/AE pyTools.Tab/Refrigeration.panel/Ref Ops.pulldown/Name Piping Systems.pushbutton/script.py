# -*- coding: utf-8 -*-
__title__ = "Name Piping Systems"
__doc__ = "Place text notes centered on pipe segments using rack-based naming rules."

import math
import re

from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
DEBUG = True
MAX_PIPE_EQUIP_DIST = 3.0


PIPE_CATEGORY_IDS = set()
for _name in ("OST_PipeCurves", "OST_FlexPipe", "OST_PipePlaceholder", "OST_FabricationPipework"):
    try:
        _bic = getattr(DB.BuiltInCategory, _name)
    except Exception:
        _bic = None
    if _bic is not None:
        try:
            PIPE_CATEGORY_IDS.add(int(_bic))
        except Exception:
            pass


class _PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):  # noqa: N802
        return _is_pipe_like(elem)

    def AllowReference(self, reference, position):  # noqa: N802
        return False


def _is_pipe_like(elem):
    if elem is None:
        return False
    try:
        cat = elem.Category
    except Exception:
        cat = None
    if cat is None:
        return False
    try:
        return cat.Id.IntegerValue in PIPE_CATEGORY_IDS
    except Exception:
        return False


def _element_connectors(elem):
    try:
        mgr = elem.ConnectorManager
    except Exception:
        mgr = None
    if mgr is None:
        try:
            mgr = elem.MEPModel.ConnectorManager
        except Exception:
            mgr = None
    if mgr is None:
        return []
    try:
        return list(mgr.Connectors)
    except Exception:
        return []


def _pipe_connectors(pipe):
    return _element_connectors(pipe)


def _pipe_open_points(pipe):
    points = []
    for conn in _pipe_connectors(pipe):
        try:
            if not conn.IsConnected:
                points.append(conn.Origin)
                continue
        except Exception:
            pass
        try:
            refs = list(conn.AllRefs)
        except Exception:
            refs = None
        if refs is not None and len(refs) == 0:
            try:
                points.append(conn.Origin)
            except Exception:
                pass
    return points


def _pipe_has_open_connector(pipe):
    return bool(_pipe_open_points(pipe))


def _connected_pipe_ids(elem):
    pipes = set()
    for conn in _element_connectors(elem):
        for ref in conn.AllRefs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if _is_pipe_like(owner):
                pipes.add(owner.Id.IntegerValue)
    return pipes


def _pipe_neighbors(pipe):
    neighbors = set()
    for conn in _pipe_connectors(pipe):
        for ref in conn.AllRefs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None:
                continue
            if _is_pipe_like(owner):
                if owner.Id.IntegerValue != pipe.Id.IntegerValue:
                    neighbors.add(owner)
            else:
                for pid in _connected_pipe_ids(owner):
                    if pid != pipe.Id.IntegerValue:
                        try:
                            neighbors.add(doc.GetElement(DB.ElementId(pid)))
                        except Exception:
                            pass
    return sorted(neighbors, key=lambda e: e.Id.IntegerValue)


def _pipe_curve(pipe):
    try:
        loc = pipe.Location
    except Exception:
        loc = None
    if loc is None or not hasattr(loc, "Curve"):
        return None
    return loc.Curve


def _pipe_diameter(pipe):
    if pipe is None:
        return 0.0
    try:
        param = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    except Exception:
        param = None
    if param:
        try:
            val = float(param.AsDouble())
            if val > 0:
                return val
        except Exception:
            pass
    for alt in (
        DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER,
        DB.BuiltInParameter.RBS_PIPE_INNER_DIAMETER,
    ):
        try:
            alt_param = pipe.get_Parameter(alt)
        except Exception:
            alt_param = None
        if alt_param:
            try:
                val = float(alt_param.AsDouble())
                if val > 0:
                    return val
            except Exception:
                pass
    try:
        return float(pipe.Diameter)
    except Exception:
        return 0.0


def _is_tee_fitting(elem):
    part_type = None
    try:
        part_type = elem.MEPModel.PartType
    except Exception:
        part_type = None
    if part_type is not None:
        try:
            if part_type in (
                DB.PartType.Tee,
                DB.PartType.Cross,
                DB.PartType.Wye,
                DB.PartType.Tap,
                DB.PartType.Takeoff,
            ):
                return True
            if part_type in (
                DB.PartType.Elbow,
                DB.PartType.Union,
                DB.PartType.Transition,
                DB.PartType.Coupling,
            ):
                return False
        except Exception:
            pass

    name_bits = []
    try:
        name_bits.append(elem.Name)
    except Exception:
        pass
    try:
        sym = elem.Symbol
    except Exception:
        sym = None
    if sym is not None:
        try:
            name_bits.append(sym.Name)
        except Exception:
            pass
        try:
            name_bits.append(sym.FamilyName)
        except Exception:
            pass
    name_str = " ".join([b for b in name_bits if b])
    name_lower = name_str.lower()
    if any(token in name_lower for token in ("tee", "wye", "cross", "tap", "takeoff", "olet", "saddle")):
        return True
    if any(token in name_lower for token in ("elbow", "union", "transition", "coupling")):
        return False

    conns = _element_connectors(elem)
    if conns:
        owners = set()
        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            elem_id = None
        for conn in conns:
            for ref in conn.AllRefs:
                try:
                    owner = ref.Owner
                except Exception:
                    owner = None
                if owner is None:
                    continue
                try:
                    owner_id = owner.Id.IntegerValue
                except Exception:
                    continue
                if elem_id is not None and owner_id == elem_id:
                    continue
                if _is_pipe_like(owner):
                    owners.add(owner_id)
        # Tees typically connect 3+ pipes.
        if len(owners) >= 3 and not any(token in name_lower for token in ("elbow", "union", "transition", "coupling")):
            return True
    return False


def _tee_ids_in_pipe(pipe):
    tee_ids = set()
    for conn in _pipe_connectors(pipe):
        for ref in conn.AllRefs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or _is_pipe_like(owner):
                continue
            if not _is_tee_fitting(owner):
                continue
            try:
                tee_ids.add(owner.Id.IntegerValue)
            except Exception:
                pass
    return sorted(tee_ids)


def _connector_to_neighbor(pipe, neighbor_id):
    for conn in _pipe_connectors(pipe):
        for ref in conn.AllRefs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None:
                continue
            if _is_pipe_like(owner):
                if owner.Id.IntegerValue == neighbor_id:
                    return conn
            else:
                connected = _connected_pipe_ids(owner)
                if neighbor_id in connected:
                    return conn
    return None


def _direction_from_connector(pipe, conn):
    curve = _pipe_curve(pipe)
    if curve is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return None
    try:
        origin = conn.Origin
    except Exception:
        return None
    try:
        d0 = origin.DistanceTo(p0)
        d1 = origin.DistanceTo(p1)
    except Exception:
        return None
    vec = (p1 - p0) if d0 <= d1 else (p0 - p1)
    try:
        return vec.Normalize()
    except Exception:
        return None


def _pick_trunk_neighbor(current_pipe, parent_id, neighbor_ids, pipe_map, adjacency, allowed_ids=None):
    if parent_id is None:
        return None
    parent_conn = _connector_to_neighbor(current_pipe, parent_id)
    if parent_conn is None:
        return None
    incoming = _direction_from_connector(current_pipe, parent_conn)
    if incoming is None:
        return None
    cont_dir = incoming
    best_id = None
    best_dot = None
    for nbr_id in neighbor_ids:
        nbr_pipe = pipe_map.get(nbr_id)
        if nbr_pipe is None:
            continue
        nbr_conn = _connector_to_neighbor(nbr_pipe, current_pipe.Id.IntegerValue)
        if nbr_conn is None:
            continue
        nbr_dir = _direction_from_connector(nbr_pipe, nbr_conn)
        if nbr_dir is None:
            continue
        dot = cont_dir.DotProduct(nbr_dir)
        if best_dot is None or dot > best_dot:
            best_dot = dot
            best_id = nbr_id
    return best_id


def _downstream_metrics(
    start_pid,
    blocked_pid,
    adjacency,
    pipe_map,
    allowed_ids=None,
    blocked_ids=None,
    extra_blocked=None,
    size_cache=None,
):
    if start_pid is None:
        return (0, 0, 0.0)
    cache = size_cache if size_cache is not None else {}
    extra = tuple(sorted(extra_blocked)) if extra_blocked else ()
    key = (start_pid, blocked_pid, extra)
    if key in cache:
        return cache[key]
    blocked_ids = set(blocked_ids or [])
    if extra_blocked:
        blocked_ids.update(extra_blocked)
    visited = set()
    stack = [start_pid]
    count = 0
    total_len = 0.0
    max_depth = 0
    stack = [(start_pid, 0)]
    while stack:
        pid, depth = stack.pop()
        if pid == blocked_pid or pid in blocked_ids or pid in visited:
            continue
        if allowed_ids is not None and pid not in allowed_ids:
            continue
        visited.add(pid)
        count += 1
        if depth > max_depth:
            max_depth = depth
        pipe = pipe_map.get(pid) if pipe_map else None
        curve = _pipe_curve(pipe) if pipe is not None else None
        if curve is not None:
            try:
                total_len += float(curve.Length)
            except Exception:
                pass
        for nbr in adjacency.get(pid, []):
            if nbr == blocked_pid or nbr in blocked_ids:
                continue
            if allowed_ids is not None and nbr not in allowed_ids:
                continue
            if nbr not in visited:
                stack.append((nbr, depth + 1))
    cache[key] = (count, max_depth, total_len)
    return cache[key]


def _downstream_component(start_pid, blocked_pid, adjacency, allowed_ids=None, blocked_ids=None, extra_blocked=None):
    if start_pid is None:
        return set(), 0
    blocked_ids = set(blocked_ids or [])
    if extra_blocked:
        blocked_ids.update(extra_blocked)
    visited = set()
    max_depth = 0
    stack = [(start_pid, 0)]
    while stack:
        pid, depth = stack.pop()
        if pid == blocked_pid or pid in blocked_ids or pid in visited:
            continue
        if allowed_ids is not None and pid not in allowed_ids:
            continue
        visited.add(pid)
        if depth > max_depth:
            max_depth = depth
        for nbr in adjacency.get(pid, []):
            if nbr == blocked_pid or nbr in blocked_ids:
                continue
            if allowed_ids is not None and nbr not in allowed_ids:
                continue
            if nbr not in visited:
                stack.append((nbr, depth + 1))
    return visited, max_depth


def _leaf_count(comp, adjacency, blocked_ids=None):
    blocked_ids = set(blocked_ids or [])
    leaves = 0
    for pid in comp:
        deg = 0
        for nbr in adjacency.get(pid, []):
            if nbr in blocked_ids:
                continue
            if nbr in comp:
                deg += 1
        if deg <= 1:
            leaves += 1
    return leaves


def _build_network(start_pipes):
    pipe_map = {}
    adjacency = {}
    queue = list(start_pipes)
    while queue:
        pipe = queue.pop()
        pid = pipe.Id.IntegerValue
        if pid in pipe_map:
            continue
        pipe_map[pid] = pipe
        neighbors = _pipe_neighbors(pipe)
        adjacency[pid] = [n.Id.IntegerValue for n in neighbors]
        for n in neighbors:
            if n.Id.IntegerValue not in pipe_map:
                queue.append(n)
    return pipe_map, adjacency


def _collect_component(start_pid, adjacency, blocked_ids):
    allowed = set()
    stack = [start_pid]
    blocked_ids = set(blocked_ids or [])
    while stack:
        pid = stack.pop()
        if pid in blocked_ids or pid in allowed:
            continue
        allowed.add(pid)
        for nbr in adjacency.get(pid, []):
            if nbr not in blocked_ids and nbr not in allowed:
                stack.append(nbr)
    return allowed


def _collect_traversal_records(start_pid, adjacency, pipe_map, allowed_ids, base_is_trunk=True, blocked_ids=None):
    records = []
    visited = set()
    blocked_ids = set(blocked_ids or [])
    # Stack frames: (pid, parent_id, t_count, is_trunk, seen_tees, dec_after)
    stack = [(start_pid, None, 0, base_is_trunk, set(), False)]
    while stack:
        pid, parent_id, t_count, is_trunk, seen_tees, dec_after = stack.pop()
        if pid in blocked_ids:
            continue
        if allowed_ids is not None and pid not in allowed_ids:
            continue
        if pid in visited:
            continue
        visited.add(pid)
        pipe = pipe_map.get(pid)
        if pipe is None:
            continue
        curr_d = _pipe_diameter(pipe)
        records.append((pid, t_count, is_trunk, dec_after))
        neighbors = [n for n in adjacency.get(pid, []) if n != parent_id and n not in blocked_ids]
        if allowed_ids is not None:
            neighbors = [n for n in neighbors if n in allowed_ids]
        neighbors.sort()
        if not neighbors:
            continue

        new_tee = None
        tee_neighbors = None
        for tid in _tee_ids_in_pipe(pipe):
            if tid in seen_tees:
                continue
            tee_elem = doc.GetElement(DB.ElementId(tid))
            if tee_elem is None:
                continue
            tee_connected = _connected_pipe_ids(tee_elem)
            cand_neighbors = [n for n in neighbors if n in tee_connected]
            if len(cand_neighbors) >= 2:
                new_tee = tid
                tee_neighbors = cand_neighbors
                break

        if new_tee is None or len(neighbors) == 1:
            for nbr in reversed(neighbors):
                stack.append((nbr, pid, t_count, is_trunk, set(seen_tees), False))
            continue

        trunk = None
        if tee_neighbors:
            size_cache = {}
            sizes = [
                (n, _downstream_metrics(n, pid, adjacency, pipe_map, allowed_ids, blocked_ids, set(tee_neighbors) - {n}, size_cache))
                for n in tee_neighbors
            ]
            max_count = max(sz[0] for _, sz in sizes) if sizes else 0
            cand = [n for n, sz in sizes if sz[0] == max_count]
            if len(cand) > 1:
                max_depth = max(sz[1] for n, sz in sizes if n in cand)
                cand = [n for n, sz in sizes if n in cand and sz[1] == max_depth]
            if len(cand) > 1:
                diameters = {n: _pipe_diameter(pipe_map.get(n)) for n in cand}
                max_d = max(diameters.values()) if diameters else 0.0
                tol = 1e-6
                cand = [n for n, d in diameters.items() if abs(d - max_d) <= tol]
            if len(cand) == 1:
                trunk = cand[0]
            else:
                trunk = _pick_trunk_neighbor(pipe, parent_id, cand, pipe_map, adjacency, allowed_ids)
                if trunk is None:
                    trunk = min(cand)
        if trunk is None and tee_neighbors:
            trunk = min(tee_neighbors)

        branches = [(n, 0.0) for n in tee_neighbors if n != trunk]
        branches.sort(key=lambda x: x[0])

        next_count = max(t_count, 0) + 1
        next_seen = set(seen_tees)
        next_seen.add(new_tee)

        # Push branches first, then trunk so traversal labels from root down the trunk.
        for nbr, _ in reversed(branches):
            stack.append((nbr, pid, next_count, False, set(next_seen), True))
        stack.append((trunk, pid, next_count, is_trunk, set(next_seen), False))
    return records


def _apply_traversal_labels(records, base_number, letter, label_map):
    prefix = str(base_number)
    trunk_counts = [t for _, t, is_trunk, _ in records if is_trunk]
    max_t = max(trunk_counts) if trunk_counts else 0
    for pid, t_count, is_trunk, _ in records:
        if is_trunk:
            if t_count <= 0:
                label = "{}{}".format(prefix, letter)
            else:
                adj = max_t - t_count
                if adj <= 0:
                    adj = max_t
                label = "{}.{}{}".format(prefix, adj, letter)
        else:
            adj = max_t - t_count
            if adj <= 0:
                adj = max_t
            label = "{}.{:02d}{}".format(prefix, adj, letter)
        label_map[pid] = label


def _get_element_center(elem, view):
    if not elem:
        return None
    bbox = None
    try:
        bbox = elem.get_BoundingBox(view)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _distance_point_to_bbox(pt, bbox):
    if pt is None or bbox is None:
        return None
    try:
        x = min(max(pt.X, bbox.Min.X), bbox.Max.X)
        y = min(max(pt.Y, bbox.Min.Y), bbox.Max.Y)
        z = min(max(pt.Z, bbox.Min.Z), bbox.Max.Z)
        closest = DB.XYZ(x, y, z)
        return pt.DistanceTo(closest)
    except Exception:
        return None


def _get_identity_mark(elem):
    if elem is None:
        return None
    try:
        param = elem.LookupParameter("Identity Mark")
    except Exception:
        param = None
    if not param:
        try:
            param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        except Exception:
            param = None
    if param:
        try:
            value = param.AsString()
        except Exception:
            value = None
        if value is None:
            try:
                value = param.AsValueString()
            except Exception:
                value = None
        if value:
            return value.strip()
    return None


def _normalize_system_label(text):
    if not text:
        return text
    raw = text.strip()
    match = re.match(r"^([A-Za-z]+)(\d+)([A-Za-z]+)?$", raw)
    if not match:
        return raw
    prefix = match.group(1) or ""
    number = match.group(2) or ""
    if len(prefix) > 2:
        prefix = prefix[:-2]
    return "{}{}".format(prefix, number)


def _collect_equipment_labels(view):
    equip = []
    try:
        collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(
            DB.BuiltInCategory.OST_MechanicalEquipment
        ).WhereElementIsNotElementType()
    except Exception:
        collector = DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_MechanicalEquipment
        ).WhereElementIsNotElementType()
    equip_map = {e.Id.IntegerValue: e for e in collector}

    for elem_id, elem in equip_map.items():
        bbox = None
        try:
            bbox = elem.get_BoundingBox(view)
        except Exception:
            bbox = None
        if not bbox:
            try:
                bbox = elem.get_BoundingBox(None)
            except Exception:
                bbox = None
        if not bbox:
            continue

        label = _normalize_system_label(_get_identity_mark(elem))

        if not label:
            continue
        equip.append((elem_id, bbox, label))
    return equip


def _apply_equipment_labels(label_map, pipe_map, adjacency, view):
    equip_labels = _collect_equipment_labels(view)
    if not equip_labels:
        if DEBUG:
            logger.info("equipment labels: none found")
        return
    matched = 0
    for pid, pipe in pipe_map.items():
        pipe = pipe_map.get(pid)
        if pipe is None:
            continue
        open_points = _pipe_open_points(pipe)
        if not open_points:
            continue
        best_dist = None
        best_text = None
        for _, bbox, text in equip_labels:
            for pt in open_points:
                dist = _distance_point_to_bbox(pt, bbox)
                if dist is None:
                    continue
                if dist > MAX_PIPE_EQUIP_DIST:
                    continue
                if best_dist is None or dist < best_dist:
                    best_dist = dist
                    best_text = text
        if best_text:
            label_map[pid] = best_text
            matched += 1
            if DEBUG:
                logger.info("equipment label: pipe %s -> %s", pid, best_text)
    if DEBUG:
        logger.info("equipment labels: applied to %s pipe(s)", matched)


def _select_trunk_neighbor(current_pipe, parent_id, neighbor_ids, pipe_map, adjacency, allowed_ids=None, blocked_ids=None, size_cache=None):
    if not neighbor_ids or current_pipe is None:
        return None
    curr_id = current_pipe.Id.IntegerValue
    size_cache = size_cache if size_cache is not None else {}
    best_id = None
    best_count = None
    best_depth = None
    best_d = None
    neighbor_set = set(neighbor_ids)
    tol = 1e-6
    comp_map = {}
    depth_map = {}
    leaf_map = {}
    for nid in neighbor_ids:
        comp, depth = _downstream_component(
            nid,
            curr_id,
            adjacency,
            allowed_ids,
            blocked_ids,
            neighbor_set - {nid},
        )
        comp_map[nid] = comp
        depth_map[nid] = depth
        leaf_map[nid] = _leaf_count(comp, adjacency, blocked_ids)
    for nid in neighbor_ids:
        other_union = set()
        for oid, comp in comp_map.items():
            if oid == nid:
                continue
            other_union.update(comp)
        unique_count = len(comp_map.get(nid, set()) - other_union)
        size = _downstream_metrics(
            nid,
            curr_id,
            adjacency,
            pipe_map,
            allowed_ids,
            blocked_ids,
            neighbor_set - {nid},
            size_cache,
        )
        diam = _pipe_diameter(pipe_map.get(nid))
        count = unique_count if unique_count > 0 else (size[0] if isinstance(size, tuple) else size)
        depth = depth_map.get(nid, 0)
        leafs = leaf_map.get(nid, 0)
        if (
            best_count is None
            or leafs > best_count
            or (leafs == best_count and (best_d is None or diam > best_d + tol))
        ):
            best_count = leafs
            best_depth = depth
            best_d = diam
            best_id = nid
        elif (
            best_count is not None
            and leafs == best_count
            and best_d is not None
            and abs(diam - best_d) <= tol
        ):
            pick = _pick_trunk_neighbor(current_pipe, parent_id, [best_id, nid], pipe_map, adjacency, allowed_ids)
            if pick is not None:
                best_id = pick
            else:
                best_id = min(best_id, nid)
    return best_id


def _find_tee_neighbors(pipe, neighbors, seen_tees):
    for tid in _tee_ids_in_pipe(pipe):
        if tid in seen_tees:
            continue
        tee_elem = doc.GetElement(DB.ElementId(tid))
        if tee_elem is None:
            continue
        tee_connected = _connected_pipe_ids(tee_elem)
        cand_neighbors = [n for n in neighbors if n in tee_connected]
        if len(cand_neighbors) >= 2:
            return tid, cand_neighbors
    return None, []


def _tee_between_pipes(pipe_a, pipe_b):
    if pipe_a is None or pipe_b is None:
        return None
    for conn in _pipe_connectors(pipe_a):
        for ref in conn.AllRefs:
            try:
                owner = ref.Owner
            except Exception:
                owner = None
            if owner is None or _is_pipe_like(owner):
                continue
            if not _is_tee_fitting(owner):
                continue
            connected = _connected_pipe_ids(owner)
            if pipe_b.Id.IntegerValue in connected:
                return owner
    return None

def _label_tree_from_trunk(start_pid, adjacency, pipe_map, blocked_ids, base_number, letter, view):
    allowed_ids = _collect_component(start_pid, adjacency, blocked_ids)
    trunk_order = []
    visited = set(blocked_ids)
    size_cache = {}

    curr = start_pid
    parent = None
    while curr is not None and curr not in visited:
        visited.add(curr)
        trunk_order.append(curr)
        walk_blocked = set(blocked_ids)
        walk_blocked.update(visited)
        if curr in walk_blocked:
            walk_blocked.remove(curr)
        neighbors = [n for n in adjacency.get(curr, []) if n != parent and n not in walk_blocked]
        if allowed_ids is not None:
            neighbors = [n for n in neighbors if n in allowed_ids]
        # Always include tee-connected pipes to avoid missing trunk continuation.
        curr_pipe = pipe_map.get(curr)
        if curr_pipe is not None:
            tee_extra = set()
            for tid in _tee_ids_in_pipe(curr_pipe):
                tee_elem = doc.GetElement(DB.ElementId(tid))
                if tee_elem is None:
                    continue
                for n in _connected_pipe_ids(tee_elem):
                    if n == curr or n == parent:
                        continue
                    if allowed_ids is not None and n not in allowed_ids:
                        continue
                    if n in walk_blocked:
                        continue
                    tee_extra.add(n)
            if tee_extra:
                neighbors = sorted(set(neighbors).union(tee_extra))
        if not neighbors and parent is not None:
            tee_elem = _tee_between_pipes(pipe_map.get(parent), pipe_map.get(curr))
            if tee_elem is not None:
                tee_connected = _connected_pipe_ids(tee_elem)
                neighbors = [n for n in tee_connected if n not in (parent, curr)]
                if allowed_ids is not None:
                    neighbors = [n for n in neighbors if n in allowed_ids]
                if walk_blocked:
                    neighbors = [n for n in neighbors if n not in walk_blocked]
        if not neighbors:
            break

        if len(neighbors) == 1:
            single = neighbors[0]
            # If the only forward neighbor is the third leg of a tee (parent+curr+single),
            # stop the trunk walk here to avoid walking into a branch.
            if parent is not None and curr_pipe is not None:
                stop_at_tee = False
                for tid in _tee_ids_in_pipe(curr_pipe):
                    tee_elem = doc.GetElement(DB.ElementId(tid))
                    if tee_elem is None:
                        continue
                    tee_connected = _connected_pipe_ids(tee_elem)
                    if parent in tee_connected and single in tee_connected:
                        stop_at_tee = True
                        break
                if stop_at_tee:
                    break
            parent, curr = curr, neighbors[0]
        else:
            trunk = _select_trunk_neighbor(
                pipe_map.get(curr),
                parent,
                neighbors,
                pipe_map,
                adjacency,
                allowed_ids,
                walk_blocked,
                size_cache,
            )
            parent, curr = curr, trunk

    trunk_set = set(trunk_order)
    label_map = {}

    # Preserve actual trunk walk order (root -> end).
    trunk_order_sorted = list(trunk_order)
    index_map = {pid: idx for idx, pid in enumerate(trunk_order_sorted)}

    # Walk trunk in order and assign tee indices + branch roots.
    tee_seen = set()
    branch_idx = 0
    branch_roots = []  # list of (tee_index, root_pid, blocked_ids)
    if DEBUG:
        logger.info("Name Piping Systems Debug - Branch %s", base_number)
        logger.info("Start pipe: %s", start_pid)
        logger.info("Trunk order (walk order): %s", trunk_order_sorted)

    trunk_index = {trunk_order_sorted[0]: 0} if trunk_order_sorted else {}
    for i in range(1, len(trunk_order_sorted)):
        prev_pid = trunk_order_sorted[i - 1]
        curr_pid = trunk_order_sorted[i]
        prev_pipe = pipe_map.get(prev_pid)
        curr_pipe = pipe_map.get(curr_pid)
        idx = trunk_index.get(prev_pid, 0)
        tee_elem = _tee_between_pipes(prev_pipe, curr_pipe)
        if tee_elem is None:
            trunk_index[curr_pid] = idx
            continue
        tid = tee_elem.Id.IntegerValue
        tee_connected = _connected_pipe_ids(tee_elem)
        upstream_ids = set(trunk_order_sorted[:i])
        tee_neighbors = [
            n for n in tee_connected
            if n in allowed_ids
            and n not in blocked_ids
            and n != prev_pid
            and n not in upstream_ids
        ]
        trunk_candidates = [n for n in tee_neighbors if n in trunk_set]

        size_cache = {}
        comp_map = {}
        leaf_map = {}
        sizes = {}
        for n in tee_neighbors:
            extra_blocked = set(tee_neighbors) - {n}
            extra_blocked.update(upstream_ids)
            extra_blocked.add(prev_pid)
            if n != curr_pid:
                extra_blocked.add(curr_pid)
            comp, depth = _downstream_component(
                n,
                prev_pid,
                adjacency,
                allowed_ids,
                blocked_ids,
                extra_blocked,
            )
            comp_map[n] = comp
            leaf_map[n] = _leaf_count(comp, adjacency, blocked_ids)
            sizes[n] = _downstream_metrics(
                n,
                prev_pid,
                adjacency,
                pipe_map,
                allowed_ids,
                blocked_ids,
                extra_blocked,
                size_cache,
            )

        if tid in tee_seen or not trunk_candidates:
            trunk_index[curr_pid] = idx
            if DEBUG:
                logger.info(
                    "tee %s skipped between %s-%s: trunk_candidates=%s branch_candidates=%s sizes=%s",
                    tid,
                    prev_pid,
                    curr_pid,
                    trunk_candidates,
                    branch_candidates,
                    sizes,
                )
            continue

        tee_seen.add(tid)
        trunk_index[curr_pid] = idx + 1

        trunk_pick = None
        if leaf_map:
            max_leaf = max(leaf_map.values())
            trunk_side = [n for n, cnt in leaf_map.items() if cnt == max_leaf]
            if len(trunk_side) > 1:
                max_count = max(sizes.get(n, (0, 0, 0.0))[0] for n in trunk_side)
                trunk_side = [n for n in trunk_side if sizes.get(n, (0, 0, 0.0))[0] == max_count]
            if len(trunk_side) > 1:
                diameters = {n: _pipe_diameter(pipe_map.get(n)) for n in trunk_side}
                max_d = max(diameters.values()) if diameters else 0.0
                tol = 1e-6
                trunk_side = [n for n, d in diameters.items() if abs(d - max_d) <= tol]
            if len(trunk_side) > 1:
                trunk_pick = _pick_trunk_neighbor(prev_pipe, prev_pid, trunk_side, pipe_map, adjacency, allowed_ids)
                if trunk_pick is None:
                    if curr_pid in trunk_side:
                        trunk_pick = curr_pid
                    else:
                        trunk_pick = min(trunk_side)
            else:
                trunk_pick = trunk_side[0] if trunk_side else None
        if trunk_pick is None and tee_neighbors:
            trunk_pick = min(tee_neighbors)
        branch_candidates = [n for n in tee_neighbors if n != trunk_pick]

        existing_branch_roots = set(b for _, b, _ in branch_roots)
        new_branch_candidates = [b for b in branch_candidates if b not in existing_branch_roots]
        if new_branch_candidates:
            branch_idx += 1
            for b in new_branch_candidates:
                blocked_for_branch = set(upstream_ids)
                blocked_for_branch.add(prev_pid)
                if trunk_pick is not None:
                    blocked_for_branch.add(trunk_pick)
                branch_roots.append((branch_idx, b, blocked_for_branch))
        else:
            trunk_index[curr_pid] = idx
            if DEBUG:
                logger.info(
                    "tee %s skipped between %s-%s (no new branch): trunk_candidates=%s branch_candidates=%s sizes=%s",
                    tid,
                    prev_pid,
                    curr_pid,
                    trunk_candidates,
                    branch_candidates,
                    sizes,
                )
            continue

        if DEBUG:
            logger.info(
                "tee %s counted between %s-%s: branch_idx=%s trunk_candidates=%s branch_candidates=%s sizes=%s",
                tid,
                prev_pid,
                curr_pid,
                branch_idx,
                trunk_candidates,
                branch_candidates,
                sizes,
            )

    for pid in trunk_order_sorted:
        idx = trunk_index.get(pid, 0)
        if idx <= 0:
            label = "{}{}".format(base_number, letter)
        else:
            label = "{}.{}{}".format(base_number, idx, letter)
        label_map[pid] = label
        if DEBUG:
            logger.info("trunk label: pipe %s -> %s", pid, label)

    # Assign each branch pipe to the nearest branch root (FIFO),
    # keeping branch indices in trunk order.
    branch_roots_sorted = sorted(branch_roots, key=lambda x: (x[0], x[1]))
    all_branch_roots = set(root for _, root, _ in branch_roots_sorted)
    branch_assign = {}
    for idx, root_pid, blocked_for_branch in branch_roots_sorted:
        if root_pid in blocked_ids:
            continue
        if allowed_ids is not None and root_pid not in allowed_ids:
            continue
        blocked = set(blocked_ids)
        blocked.update(blocked_for_branch)
        blocked.update(trunk_set)
        blocked.discard(root_pid)
        blocked.update(all_branch_roots - {root_pid})
        comp = _collect_component(root_pid, adjacency, blocked)
        for pid in comp:
            if allowed_ids is not None and pid not in allowed_ids:
                continue
            if pid in blocked:
                continue
            if pid in branch_assign:
                continue
            branch_assign[pid] = idx

    for pid, idx in branch_assign.items():
        label_map[pid] = "{}.{:02d}{}".format(base_number, idx, letter)
        if DEBUG:
            logger.info("branch label: pipe %s -> %s", pid, label_map[pid])

    return label_map


def _pick_text_type():
    types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    if not types:
        return None
    for t in types:
        try:
            name = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            font = t.get_Parameter(DB.BuiltInParameter.TEXT_FONT).AsString() or ""
            if ("3/32" in name) and ("Arial" in font):
                return t
        except Exception:
            continue
    return types[0]


def _pipe_tag_types():
    tag_types = []
    for cat_name in ("OST_PipeTags", "OST_MultiCategoryTags"):
        try:
            cat = getattr(DB.BuiltInCategory, cat_name)
        except Exception:
            cat = None
        if cat is None:
            continue
        tag_types.extend(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilySymbol)
            .OfCategory(cat)
            .ToElements()
        )
    return tag_types


def _tag_type_label(tag_type):
    try:
        fam_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
    except Exception:
        fam_name = None
    try:
        type_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except Exception:
        type_name = None
    try:
        cat_name = tag_type.Category.Name if tag_type.Category else "Tag"
    except Exception:
        cat_name = "Tag"
    return "[{}] {} : {}".format(cat_name, fam_name or "?", type_name or "?")


def _pick_pipe_tag_type():
    tag_types = _pipe_tag_types()
    if not tag_types:
        return None
    options = [_tag_type_label(t) for t in tag_types]
    picked = forms.SelectFromList.show(
        options,
        multiselect=False,
        title="Select Pipe Tag Type",
    )
    if not picked:
        return None
    try:
        idx = options.index(picked)
    except Exception:
        return None
    return tag_types[idx]


def _place_text_notes(label_map, pipe_map, text_type, view):
    if not label_map:
        return 0
    count = 0
    with revit.Transaction("Name Piping Systems - Text Notes"):
        for pid, label in label_map.items():
            pipe = pipe_map.get(pid)
            if pipe is None:
                continue
            curve = _pipe_curve(pipe)
            if curve is None:
                continue
            try:
                pt = curve.Evaluate(0.5, True)
            except Exception:
                continue
            try:
                opts = DB.TextNoteOptions()
                opts.TypeId = text_type.Id
                try:
                    opts.HorizontalAlignment = DB.HorizontalTextAlignment.Center
                except Exception:
                    pass
                try:
                    opts.VerticalAlignment = DB.VerticalTextAlignment.Middle
                except Exception:
                    pass
                DB.TextNote.Create(doc, view.Id, pt, label, opts)
                count += 1
            except Exception:
                continue
    return count


def _pipe_direction_xy(pipe):
    curve = _pipe_curve(pipe)
    if curve is None:
        return None
    try:
        deriv = curve.ComputeDerivatives(0.5, True)
        tangent = deriv.BasisX
    except Exception:
        tangent = None
    if tangent is None:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            tangent = p1 - p0
        except Exception:
            return None
    try:
        vec = DB.XYZ(tangent.X, tangent.Y, 0.0)
        if abs(vec.X) + abs(vec.Y) < 1e-9:
            return None
        return vec.Normalize()
    except Exception:
        return None


def _rotate_element_about_z(elem_id, origin, angle):
    if elem_id is None or origin is None:
        return
    if abs(angle) < 1e-7:
        return
    axis = DB.Line.CreateUnbound(origin, DB.XYZ(0, 0, 1))
    DB.ElementTransformUtils.RotateElement(doc, elem_id, axis, angle)


def _place_pipe_tags(label_map, pipe_map, tag_type, view):
    if not label_map:
        return 0
    count = 0
    with revit.Transaction("Name Piping Systems - Tags"):
        if not tag_type.IsActive:
            tag_type.Activate()
            doc.Regenerate()
        for pid in label_map.keys():
            pipe = pipe_map.get(pid)
            if pipe is None:
                continue
            curve = _pipe_curve(pipe)
            if curve is None:
                continue
            try:
                pt = curve.Evaluate(0.5, True)
            except Exception:
                continue
            try:
                reference = DB.Reference(pipe)
                tag = DB.IndependentTag.Create(
                    doc,
                    tag_type.Id,
                    view.Id,
                    reference,
                    False,
                    DB.TagOrientation.Horizontal,
                    pt,
                )
                if tag:
                    dir_xy = _pipe_direction_xy(pipe)
                    if dir_xy is not None:
                        angle = math.atan2(dir_xy.Y, dir_xy.X)
                        try:
                            _rotate_element_about_z(tag.Id, tag.TagHeadPosition, angle)
                        except Exception:
                            _rotate_element_about_z(tag.Id, pt, angle)
                    count += 1
            except Exception:
                continue
    return count


def _apply_identity_mark(label_map, pipe_map):
    if not label_map:
        return
    with revit.Transaction("Name Piping Systems - Identity Mark"):
        for pid, label in label_map.items():
            pipe = pipe_map.get(pid)
            if pipe is None:
                continue
            try:
                param = pipe.LookupParameter("Identity Mark")
                if not param:
                    param = pipe.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
                if param and not param.IsReadOnly:
                    if param.StorageType == DB.StorageType.String:
                        param.Set(label)
                    else:
                        param.SetValueString(label)
            except Exception as ex:
                logger.warning("Failed to set Identity Mark for {}: {}".format(pid, ex))


def _ask_branch_letter():
    letters = [chr(code) for code in range(ord("A"), ord("H") + 1)]
    letter = forms.ask_for_one_item(
        letters,
        default="A",
        prompt="Select the rack letter (A-H) to append to labels.",
        title="Select Rack Letter",
    )
    if not letter:
        forms.alert("No letter selected.", exitscript=True)
    return letter


def _prompt_start_pipes():
    try:
        forms.alert(
            "Select the start pipes for this rack in order (2-5).",
            title="Select Start Pipes",
            warn_icon=False,
        )
    except Exception:
        pass
    start_pipes = []
    selected_ids = []
    while True:
        if len(start_pipes) >= 5:
            break
        prompt = "Select start pipe #{} (ESC to finish)".format(len(start_pipes) + 1)
        try:
            with forms.WarningBar(title=prompt):
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    _PipeSelectionFilter(),
                    prompt,
                )
        except Exception:
            break
        try:
            elem = doc.GetElement(ref.ElementId)
        except Exception:
            elem = None
        if not _is_pipe_like(elem):
            continue
        elem_id = elem.Id.IntegerValue
        if elem_id in selected_ids:
            continue
        start_pipes.append(elem)
        selected_ids.append(elem_id)
        try:
            uidoc.Selection.SetElementIds(List[DB.ElementId]([p.Id for p in start_pipes]))
        except Exception:
            pass

    if len(start_pipes) < 2 or len(start_pipes) > 5:
        forms.alert(
            "Select 2 to 5 start pipes. Detected {} pipe(s).".format(len(start_pipes)),
            exitscript=True,
        )
    return start_pipes


def main():
    active_view = revit.active_view
    if active_view.IsTemplate:
        forms.alert("Active view is a template. Open a working view first.", exitscript=True)

    start_pipes = _prompt_start_pipes()
    rack_letter = _ask_branch_letter()

    pipe_map, adjacency = _build_network(start_pipes)
    start_ids = [p.Id.IntegerValue for p in start_pipes]
    label_map = {}
    for idx, pipe in enumerate(start_pipes, start=1):
        pid = pipe.Id.IntegerValue
        blocked_ids = set(start_ids)
        blocked_ids.discard(pid)
        allowed_ids = _collect_component(pid, adjacency, blocked_ids)
        local_labels = _label_tree_from_trunk(
            pid,
            adjacency,
            pipe_map,
            blocked_ids,
            idx,
            rack_letter,
            active_view,
        )
        for lpid, lbl in local_labels.items():
            label_map[lpid] = lbl

    _apply_equipment_labels(label_map, pipe_map, adjacency, active_view)

    output_mode = forms.CommandSwitchWindow.show(
        ["Text Notes", "Tags", "Both"],
        message="Place system names as:",
    )
    if not output_mode:
        script.exit()

    _apply_identity_mark(label_map, pipe_map)

    placed_notes = 0
    placed_tags = 0
    if output_mode in ("Tags", "Both"):
        tag_type = _pick_pipe_tag_type()
        if tag_type is None:
            forms.alert("No Pipe Tag type selected or available.", exitscript=True)
        placed_tags = _place_pipe_tags(label_map, pipe_map, tag_type, active_view)
    if output_mode in ("Text Notes", "Both"):
        text_type = _pick_text_type()
        if text_type is None:
            forms.alert("No TextNoteType available in this project.", exitscript=True)
        placed_notes = _place_text_notes(label_map, pipe_map, text_type, active_view)

    if output_mode == "Both":
        forms.alert("Placed {} tags and {} text notes.".format(placed_tags, placed_notes))
    elif output_mode == "Tags":
        forms.alert("Placed {} tags.".format(placed_tags))
    else:
        forms.alert("Placed {} text notes.".format(placed_notes))


if __name__ == "__main__":
    main()
