from Autodesk.Revit import DB
from Autodesk.Revit.DB import XYZ

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str


def _normalize(value):
    if isinstance(value, basestring):
        return value.strip().upper()
    if value is None:
        return ""
    try:
        return str(value).strip().upper()
    except Exception:
        return ""


def classify_items(items):
    dedicated = []
    nongrouped = []
    tvtruss = []
    normal = []

    for item in items:
        circuit_upper = _normalize(item.get("circuit_number"))

        if circuit_upper == "DEDICATED":
            dedicated.append(item)
        elif circuit_upper == "NONGROUPEDBLOCK":
            nongrouped.append(item)
        elif circuit_upper == "TVTRUSS":
            tvtruss.append(item)
        else:
            normal.append(item)

    return dedicated, nongrouped, tvtruss, normal


def split_combined_circuit(panel_name, circuit_number, members, make_group, logger=None, parse_int=None):
    if not circuit_number or "&" not in circuit_number:
        return None

    parts = [part.strip() for part in circuit_number.split("&") if part.strip()]
    if len(parts) < 2:
        return None

    total = len(members)
    segment_count = len(parts)
    base = total // segment_count
    remainder = total % segment_count

    special_load_names = None
    try:
        normalized_parts = [str(int(part)) for part in parts]
    except ValueError:
        normalized_parts = parts[:]

    if set(normalized_parts) == {"6", "9"} and len(normalized_parts) == 2:
        special_load_names = {"6": "SINK 1", "9": "SINK 2"}

    groups = []
    index = 0
    for i, part in enumerate(parts):
        size = base + (1 if i < remainder else 0)
        if i == segment_count - 1:
            group_members = members[index:]
        else:
            group_members = members[index : index + size]
        index += size

        if special_load_names:
            normalized_part = None
            if parse_int:
                normalized_part = parse_int(part)
            if normalized_part is None:
                try:
                    normalized_part = int(part)
                except Exception:
                    normalized_part = None
            normalized_part = str(normalized_part) if normalized_part is not None else part
            label = special_load_names.get(normalized_part)
            if label:
                for member in group_members:
                    member["load_name"] = label

        key = "{}{}".format(panel_name, part)
        groups.append(make_group(key, group_members))

    return groups


_LOAD_PRIORITY = {
    "TREADMILL": 0,
    "POWERED BIKE": 1,
    "POWERED BIKE1": 1,
    "POWERED BIKE2": 1,
    "STAIRMASTER": 2,
    "SINK 1": 3,
    "SINK 2": 3,
}

SPACE_TOLERANCE_FT = 2.0
PANEL_SPACE_MAP = {
    "BA": ["BAKERY"],
    "DA": ["DELI"],
    "FL": ["FLORAL"],
    "PR": ["PRODUCE"],
    "MT": ["MEAT", "SEAFOOD"],
}

_POSITION_RULES = [
    {"keyword": "CHECKSTAND RECEPT", "group_size": 2},
    {"keyword": "CHECKSTAND JBOX", "group_size": 2},
    {"keyword": "SELF CHECKOUT", "group_size": 2},
    {"keyword": "TABLE", "group_size": 3, "label": "Grouped Tables"},
    {"keyword": "DESK QUAD", "group_size": 3, "label": "Grouped Desks - Quad"},
    {"keyword": "DESK DUPLEX", "group_size": 3, "label": "Grouped Desks - Duplex"},
    {"keyword": "ARTISAN BREAD", "group_size": 3},
    {"keyword": "MADIX CLEAN", "group_size": 3},
    {"keyword": "MADIX DIRTY", "group_size": 3},
]


def _normalize_panel_space_map(panel_space_map):
    normalized = {}
    for panel, spaces in (panel_space_map or {}).items():
        panel_key = _normalize(panel)
        if not panel_key:
            continue
        tokens = []
        for space_name in spaces or []:
            token = _normalize(space_name)
            if token:
                tokens.append(token)
        if tokens:
            normalized[panel_key] = tokens
    return normalized


_PANEL_SPACE_MAP = _normalize_panel_space_map(PANEL_SPACE_MAP)
_POSITION_RULES_SORTED = sorted(_POSITION_RULES, key=lambda r: -len(r.get("keyword") or ""))


def get_load_priority(group):
    load_name = group.get("load_name")
    if load_name:
        key = load_name.strip().upper()
        if key in _LOAD_PRIORITY:
            return _LOAD_PRIORITY[key]
    return 3


def _split_panel_choices(value):
    if not value:
        return []
    text = value if isinstance(value, basestring) else str(value)
    for sep in (",", ";", "|", "\n", "\r"):
        text = text.replace(sep, " ")
    candidates = [part.strip() for part in text.split(" ") if part.strip()]
    unique = []
    seen = set()
    for name in candidates:
        upper = name.upper()
        if upper in seen:
            continue
        seen.add(upper)
        unique.append(name)
    return unique


def _extract_space_attr(space, *names):
    if not space:
        return None
    for name in names:
        value = None
        try:
            value = getattr(space, name, None)
        except Exception:
            value = None
        if not value:
            try:
                param = space.LookupParameter(name)
                if param and param.HasValue:
                    value = param.AsString()
            except Exception:
                value = None
        if value:
            text = str(value).strip()
            if text:
                return text
    return None


def _space_label(space):
    label = _extract_space_attr(space, "Name", "Space Name", "Room Name")
    if not label:
        label = _extract_space_attr(space, "Number", "Space Number", "Room Number")
    return label


def _get_bounding_box(elem):
    if not elem:
        return None
    try:
        return elem.get_BoundingBox(None)
    except Exception:
        return None


def _get_point(elem):
    if not elem:
        return None
    location = getattr(elem, "Location", None)
    if location:
        point = getattr(location, "Point", None)
        if point:
            return point
        curve = getattr(location, "Curve", None)
        if curve:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                pass
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )
    return None


def _point_in_bbox(point, bbox, tolerance):
    if not point or not bbox:
        return False
    tol = float(tolerance or 0.0)
    min_x = bbox.Min.X - tol
    min_y = bbox.Min.Y - tol
    min_z = bbox.Min.Z - tol
    max_x = bbox.Max.X + tol
    max_y = bbox.Max.Y + tol
    max_z = bbox.Max.Z + tol
    return (
        min_x <= point.X <= max_x
        and min_y <= point.Y <= max_y
        and min_z <= point.Z <= max_z
    )


def _point_in_space(space, point, tolerance):
    if not space or point is None:
        return False
    checker = getattr(space, "IsPointInSpace", None)
    if not callable(checker):
        checker = getattr(space, "IsPointInRoom", None)
    if callable(checker):
        try:
            if checker(point):
                return True
        except Exception:
            pass
    bbox = _get_bounding_box(space)
    if bbox and tolerance:
        return _point_in_bbox(point, bbox, tolerance)
    return False


def _collect_spaces(doc):
    if not doc:
        return []
    categories = []
    try:
        categories.append(DB.BuiltInCategory.OST_MEPSpaces)
    except Exception:
        pass
    try:
        categories.append(DB.BuiltInCategory.OST_Rooms)
    except Exception:
        pass
    spaces = []
    for category in categories:
        try:
            collector = DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType()
        except Exception:
            continue
        for space in collector:
            label = _space_label(space)
            space_id = getattr(getattr(space, "Id", None), "IntegerValue", None)
            spaces.append(
                {
                    "element": space,
                    "id": space_id,
                    "name": label,
                    "upper": _normalize(label),
                    "point": _get_point(space),
                    "bbox": _get_bounding_box(space),
                }
            )
    return spaces


def _build_panel_cache(panel_lookup):
    cache = {}
    for name, info in (panel_lookup or {}).items():
        panel_elem = info.get("element")
        point = _get_point(panel_elem)
        if not point:
            continue
        upper = (name or "").strip().upper()
        if not upper or upper in cache:
            continue
        cache[upper] = {
            "upper": upper,
            "name": name,
            "point": point,
            "info": info,
        }
    return cache


def _distance(point_a, point_b):
    if point_a is None or point_b is None:
        return None
    try:
        return point_a.DistanceTo(point_b)
    except Exception:
        try:
            dx = point_a.X - point_b.X
            dy = point_a.Y - point_b.Y
            dz = point_a.Z - point_b.Z
            return (dx * dx + dy * dy + dz * dz) ** 0.5
        except Exception:
            return None


def _nearest_panel(location, panel_cache, allowed=None):
    if not location:
        return None
    allowed_set = {item.upper() for item in (allowed or []) if item}
    best = None
    best_dist = None
    for entry in panel_cache.values():
        if allowed_set and entry.get("upper") not in allowed_set:
            continue
        dist = _distance(location, entry.get("point"))
        if dist is None:
            continue
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best = entry
    return best


def _candidate_spaces_for_panels(spaces, allowed_panels, panel_space_map):
    if not spaces or not allowed_panels or not panel_space_map:
        return []
    matches = {}
    for panel in allowed_panels:
        tokens = panel_space_map.get(panel)
        if not tokens:
            continue
        for space_entry in spaces:
            space_upper = space_entry.get("upper") or ""
            if not space_upper:
                continue
            if not any(token in space_upper for token in tokens):
                continue
            space_id = space_entry.get("id")
            key = space_id if space_id is not None else id(space_entry.get("element"))
            entry = matches.get(key)
            if not entry:
                entry = {"space": space_entry, "panels": set()}
                matches[key] = entry
            entry["panels"].add(panel)
    return list(matches.values())


def _nearest_space(location, candidates):
    if not location or not candidates:
        return None
    best = None
    best_dist = None
    for candidate in candidates:
        point = candidate.get("space", {}).get("point")
        dist = _distance(location, point)
        if dist is None:
            continue
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best = candidate
    return best or (candidates[0] if candidates else None)


def _select_space_for_point(location, candidates, tolerance):
    if not location or not candidates:
        return None
    matches = []
    for candidate in candidates:
        space = candidate.get("space", {}).get("element")
        if _point_in_space(space, location, tolerance):
            matches.append(candidate)
    if matches:
        return _nearest_space(location, matches)
    return None


def _find_space_entry(location, spaces, tolerance):
    if not location or not spaces:
        return None
    matches = []
    for entry in spaces:
        if _point_in_space(entry.get("element"), location, tolerance):
            matches.append(entry)
    if matches:
        best = None
        best_dist = None
        for entry in matches:
            dist = _distance(location, entry.get("point"))
            if dist is None:
                continue
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best = entry
        return best or matches[0]
    nearest = _nearest_space(location, [{"space": entry} for entry in spaces])
    return nearest.get("space") if nearest else None


def _apply_space_load_name(item, space_label):
    if not item or not space_label:
        return
    label = str(space_label).strip()
    if not label:
        return
    existing = (item.get("load_name") or "").strip()
    if not existing:
        item["load_name"] = label
        return
    suffix = " - {}".format(label)
    if existing.lower().endswith(suffix.lower()):
        return
    item["load_name"] = "{}{}".format(existing, suffix)


def _match_position_rule(circuit_number):
    text = _normalize(circuit_number)
    if not text:
        return None
    for rule in _POSITION_RULES_SORTED:
        keyword = rule.get("keyword")
        if keyword and keyword in text:
            return rule
    return None


def _position_sort_key(item):
    point = item.get("location")
    if not point:
        return (1e9, 1e9, 1e9)
    return (point.X, point.Y, point.Z)


def _cluster_by_nearest(items, group_size):
    remaining = sorted(list(items), key=_position_sort_key)
    groups = []
    while remaining:
        seed = remaining.pop(0)
        seed_point = seed.get("location")
        if group_size <= 1 or not remaining or seed_point is None:
            groups.append([seed])
            continue
        distances = []
        for idx, other in enumerate(remaining):
            dist = _distance(seed_point, other.get("location"))
            if dist is None:
                dist = 1e12
            distances.append((dist, idx))
        distances.sort(key=lambda entry: entry[0])
        pick_count = min(group_size - 1, len(remaining))
        pick_indices = sorted((distances[i][1] for i in range(pick_count)), reverse=True)
        group = [seed]
        for idx in pick_indices:
            group.append(remaining.pop(idx))
        groups.append(group)
    return groups


def _cluster_by_radius(items, radius):
    if not items:
        return [], []
    if radius is None:
        return [], list(items)
    radius_val = float(radius)
    grouped = []
    singles = []
    remaining = list(items)
    visited = set()
    while remaining:
        seed = remaining.pop(0)
        if seed in visited:
            continue
        cluster = [seed]
        queue = [seed]
        visited.add(seed)
        while queue:
            current = queue.pop(0)
            current_point = current.get("location")
            if current_point is None:
                continue
            for other in list(remaining):
                if other in visited:
                    continue
                dist = _distance(current_point, other.get("location"))
                if dist is None or dist > radius_val:
                    continue
                visited.add(other)
                queue.append(other)
                cluster.append(other)
                remaining.remove(other)
        if len(cluster) > 1:
            grouped.append(cluster)
        else:
            singles.extend(cluster)
    return grouped, singles


def _sanitize_token(value, fallback="X"):
    if not value:
        return fallback
    token = "".join(ch for ch in str(value) if ch.isalnum())
    return token or fallback


def _apply_group_label(members, label, index, space_label):
    if not label:
        return None
    label_text = space_label or "Unassigned Space"
    group_name = "{} {} - {}".format(label, index, label_text)
    for member in members:
        member["load_name"] = group_name
    return group_name


def create_position_groups(items, make_group, logger=None):
    if not items:
        return [], items
    grouped_items = []
    remaining = []
    for item in items:
        rule = _match_position_rule(item.get("circuit_number"))
        if rule:
            item["_heb_position_rule"] = rule
            grouped_items.append(item)
        else:
            remaining.append(item)

    if not grouped_items:
        return [], remaining

    buckets = {}
    for item in grouped_items:
        rule = item.get("_heb_position_rule") or {}
        keyword = rule.get("keyword") or ""
        panel_name = item.get("panel_name") or "NO_PANEL"
        space_key = item.get("space_id") or item.get("space_label") or "NO_SPACE"
        bucket_key = (panel_name, space_key, keyword)
        buckets.setdefault(bucket_key, []).append(item)

    groups = []
    counters = {}
    for bucket_key in sorted(buckets.keys(), key=lambda k: (str(k[1]), str(k[0]), str(k[2]))):
        members = buckets[bucket_key]
        if not members:
            continue
        rule = members[0].get("_heb_position_rule") or {}
        group_size = int(rule.get("group_size") or 1)
        cluster_radius = rule.get("cluster_radius")
        label = rule.get("label")
        space_label = members[0].get("space_label")
        keyword = rule.get("keyword") or "KEY"
        if cluster_radius:
            position_groups, remainder = _cluster_by_radius(members, cluster_radius)
            remaining.extend(remainder)
        else:
            position_groups = _cluster_by_nearest(members, group_size)
        for pos_index, cluster in enumerate(position_groups, start=1):
            count_key = (bucket_key[1], keyword)
            counters[count_key] = counters.get(count_key, 0) + 1
            group_index = counters[count_key]
            group_name = _apply_group_label(cluster, label, group_index, space_label)
            token = _sanitize_token(keyword, fallback="POS")
            space_token = _sanitize_token(space_label, fallback="SPACE")
            key = "{}{}_{}_{}".format(bucket_key[0], token, space_token, group_index)
            group = make_group(key, cluster, group_type="position")
            if group_name:
                group["load_name"] = group_name
            groups.append(group)

    return groups, remaining


def preprocess_items(items, doc, panel_lookup, logger=None):
    if not items:
        return items
    panel_cache = _build_panel_cache(panel_lookup)
    if not panel_cache:
        return items
    spaces = _collect_spaces(doc)
    for item in items:
        location = item.get("location")
        allowed_raw = _split_panel_choices(item.get("panel_name"))
        fallback_space = _find_space_entry(location, spaces, SPACE_TOLERANCE_FT)
        if fallback_space:
            item["space_label"] = fallback_space.get("name")
            item["space_id"] = fallback_space.get("id")
        if not location or not allowed_raw:
            continue
        allowed = {name.strip().upper(): name.strip() for name in allowed_raw if name.strip()}
        if not allowed:
            continue
        allowed_panels = set(allowed.keys())
        candidates = _candidate_spaces_for_panels(spaces, allowed_panels, _PANEL_SPACE_MAP)
        chosen = _select_space_for_point(location, candidates, SPACE_TOLERANCE_FT)
        if not chosen and candidates:
            chosen = _nearest_space(location, candidates)

        if chosen:
            space_label = chosen.get("space", {}).get("name")
            item["space_label"] = space_label or item.get("space_label")
            item["space_id"] = chosen.get("space", {}).get("id")
            _apply_space_load_name(item, space_label)
            allowed_panels = set(chosen.get("panels") or allowed_panels)

        nearest = _nearest_panel(location, panel_cache, allowed_panels)
        if not nearest:
            continue
        if nearest["upper"] not in allowed:
            continue
        panel_info = nearest.get("info") or {}
        item["panel_name"] = nearest["name"]
        item["panel_element"] = panel_info.get("element")
        item["panel_distribution_system_ids"] = list(panel_info.get("distribution_system_ids") or [])
        if logger:
            element = item.get("element")
            elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
            logger.debug(
                "HEB helper assigned panel %s to element %s",
                nearest["name"],
                elem_id if elem_id is not None else "unknown",
            )
    return items
