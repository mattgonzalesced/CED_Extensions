# -*- coding: utf-8 -*-
__title__ = "Name Piping Systems"
__doc__ = "Place text notes centered on pipe segments using rack-based naming rules."

from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, DB, forms, script


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
DEBUG = True


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

        info = [(n, _pipe_diameter(pipe_map.get(n))) for n in tee_neighbors]
        curr_d = _pipe_diameter(pipe)
        trunk = None
        tol = 1e-6
        if curr_d > 0:
            same_as_parent = [n for n, d in info if abs(d - curr_d) <= tol]
            if len(same_as_parent) == 1:
                trunk = same_as_parent[0]
            elif len(same_as_parent) > 1:
                trunk = _pick_trunk_neighbor(pipe, parent_id, same_as_parent, pipe_map, adjacency, allowed_ids)
        if trunk is None:
            max_d = max(d for _, d in info)
            cand = [n for n, d in info if abs(d - max_d) <= tol]
            if len(cand) == 1 and max_d > 0:
                trunk = cand[0]
            else:
                trunk = _pick_trunk_neighbor(pipe, parent_id, cand or tee_neighbors, pipe_map, adjacency, allowed_ids)
            if trunk is None:
                trunk = cand[0] if cand else tee_neighbors[0]
        if trunk is None and tee_neighbors:
            trunk = min(tee_neighbors)

        branches = [(n, d) for n, d in info if n != trunk]
        branches.sort(key=lambda x: (x[1], x[0]))

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


def _select_trunk_neighbor(current_pipe, parent_id, neighbor_ids, pipe_map, adjacency, allowed_ids=None):
    if not neighbor_ids:
        return None
    curr_d = _pipe_diameter(current_pipe)
    tol = 1e-6
    info = [(nid, _pipe_diameter(pipe_map.get(nid))) for nid in neighbor_ids]
    if curr_d > 0:
        same = [nid for nid, d in info if abs(d - curr_d) <= tol]
        if len(same) == 1:
            return same[0]
        if len(same) > 1:
            pick = _pick_trunk_neighbor(current_pipe, parent_id, same, pipe_map, adjacency, allowed_ids)
            if pick is not None:
                return pick
            return same[0]
    max_d = max(d for _, d in info)
    cand = [nid for nid, d in info if abs(d - max_d) <= tol]
    if len(cand) == 1 and max_d > 0:
        return cand[0]
    pick = _pick_trunk_neighbor(current_pipe, parent_id, cand or neighbor_ids, pipe_map, adjacency, allowed_ids)
    return pick if pick is not None else (cand[0] if cand else neighbor_ids[0])


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

def _label_tree_from_trunk(start_pid, adjacency, pipe_map, blocked_ids, base_number, letter):
    allowed_ids = _collect_component(start_pid, adjacency, blocked_ids)
    trunk_order = []
    visited = set(blocked_ids)

    curr = start_pid
    parent = None
    while curr is not None and curr not in visited:
        visited.add(curr)
        trunk_order.append(curr)
        neighbors = [n for n in adjacency.get(curr, []) if n != parent and n not in blocked_ids]
        if allowed_ids is not None:
            neighbors = [n for n in neighbors if n in allowed_ids]
        if not neighbors:
            break

        if len(neighbors) == 1:
            parent, curr = curr, neighbors[0]
        else:
            trunk = _select_trunk_neighbor(pipe_map.get(curr), parent, neighbors, pipe_map, adjacency, allowed_ids)
            parent, curr = curr, trunk

    trunk_set = set(trunk_order)
    label_map = {}

    # Preserve actual trunk walk order (root -> end).
    trunk_order_sorted = list(trunk_order)
    index_map = {pid: idx for idx, pid in enumerate(trunk_order_sorted)}

    # Walk trunk in order and assign tee indices + branch roots.
    tee_seen = set()
    branch_idx = 0
    branch_roots = []  # list of (tee_index, root_pid)
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
        trunk_candidates = [n for n in tee_connected if n in trunk_set]
        branch_candidates = [
            n for n in tee_connected
            if n in allowed_ids and n not in trunk_set and n not in blocked_ids
        ]
        trunk_d = _pipe_diameter(curr_pipe)
        if trunk_d > 0:
            branch_candidates = [
                n for n in branch_candidates
                if _pipe_diameter(pipe_map.get(n)) > 0
                and _pipe_diameter(pipe_map.get(n)) < trunk_d
            ]
        if tid in tee_seen or not branch_candidates or len(trunk_candidates) < 2:
            trunk_index[curr_pid] = idx
            if DEBUG:
                logger.info(
                    "tee %s skipped between %s-%s: trunk_candidates=%s branch_candidates=%s",
                    tid,
                    prev_pid,
                    curr_pid,
                    trunk_candidates,
                    branch_candidates,
                )
            continue
        branch_idx += 1
        tee_seen.add(tid)
        trunk_index[curr_pid] = idx + 1
        if DEBUG:
            logger.info(
                "tee %s counted between %s-%s: branch_idx=%s trunk_candidates=%s branch_candidates=%s",
                tid,
                prev_pid,
                curr_pid,
                branch_idx,
                trunk_candidates,
                branch_candidates,
            )
        for b in branch_candidates:
            branch_roots.append((branch_idx, b))

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
    branch_assign = {}
    queue = []
    for idx, root_pid in branch_roots_sorted:
        if root_pid in blocked_ids or root_pid in trunk_set:
            continue
        if allowed_ids is not None and root_pid not in allowed_ids:
            continue
        if root_pid in branch_assign:
            continue
        branch_assign[root_pid] = idx
        queue.append(root_pid)

    q_index = 0
    while q_index < len(queue):
        pid = queue[q_index]
        q_index += 1
        idx = branch_assign.get(pid)
        if idx is None:
            continue
        for nbr in adjacency.get(pid, []):
            if nbr in trunk_set or nbr in blocked_ids:
                continue
            if allowed_ids is not None and nbr not in allowed_ids:
                continue
            if nbr in branch_assign:
                continue
            branch_assign[nbr] = idx
            queue.append(nbr)

    for pid, idx in branch_assign.items():
        if pid in label_map:
            continue
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
            "Select the start pipes for this rack (2-5), then press Finish.",
            title="Select Start Pipes",
            warn_icon=False,
        )
    except Exception:
        pass
    try:
        picked = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _PipeSelectionFilter(),
            "Select start pipes (ESC to cancel).",
        )
    except Exception:
        picked = None
    if not picked:
        forms.alert("No start pipes selected.", exitscript=True)
    start_pipes = []
    for ref in picked:
        try:
            elem = doc.GetElement(ref.ElementId)
        except Exception:
            elem = None
        if _is_pipe_like(elem):
            start_pipes.append(elem)
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
        local_labels = _label_tree_from_trunk(pid, adjacency, pipe_map, blocked_ids, idx, rack_letter)
        for lpid, lbl in local_labels.items():
            label_map[lpid] = lbl

    text_type = _pick_text_type()
    if text_type is None:
        forms.alert("No TextNoteType available in this project.", exitscript=True)
    placed = _place_text_notes(label_map, pipe_map, text_type, active_view)
    forms.alert("Placed {} text notes.".format(placed))


if __name__ == "__main__":
    main()
