from collections import defaultdict

try:
    from pyrevit import DB
except Exception:  # pragma: no cover
    DB = None

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str

TVTRUSS_KEYWORD = "TVTRUSS"
TVTRUSS_MAX_NEIGHBOR_DISTANCE_FT = 10.0

try:
    ELECTRICAL_FIXTURE_CATEGORY_ID = (
        DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures).IntegerValue if DB else None
    )
except Exception:
    ELECTRICAL_FIXTURE_CATEGORY_ID = None


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
    normal = []

    for item in items:
        circuit_upper = _normalize(item.get("circuit_number"))

        if circuit_upper == "DEDICATED":
            dedicated.append(item)
        else:
            # TVTRUSS now handled by position rules in create_position_groups()
            normal.append(item)

    return dedicated, [], [], normal


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
            group_members = members[index:index + size]
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

_POSITION_RULES = [
    {"keyword": TVTRUSS_KEYWORD, "group_size": 3},
]

_POSITION_RULES_SORTED = sorted(_POSITION_RULES, key=lambda r: -len(r.get("keyword") or ""))


def get_load_priority(group):
    load_name = group.get("load_name")
    if load_name:
        key = load_name.strip().upper()
        if key in _LOAD_PRIORITY:
            return _LOAD_PRIORITY[key]
    return 3


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


def _distance_between_points(point_a, point_b):
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


def _is_electrical_fixture_item(item):
    element = item.get("element")
    if not element:
        return False
    category = getattr(element, "Category", None)
    if not category:
        return False
    category_id = getattr(getattr(category, "Id", None), "IntegerValue", None)
    if ELECTRICAL_FIXTURE_CATEGORY_ID is not None and category_id is not None:
        return category_id == ELECTRICAL_FIXTURE_CATEGORY_ID
    return _normalize(getattr(category, "Name", "")) == "ELECTRICAL FIXTURES"


def _item_id_text(item):
    element = item.get("element")
    element_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
    return str(element_id) if element_id is not None else "unknown"


def _filter_tvtruss_fixture_outliers(grouped_items, logger=None):
    tvtruss_items = []
    for item in grouped_items:
        rule = item.get("_pf_position_rule") or {}
        if _normalize(rule.get("keyword")) == TVTRUSS_KEYWORD:
            tvtruss_items.append(item)

    fixture_items = [item for item in tvtruss_items if _is_electrical_fixture_item(item)]
    if not fixture_items:
        return grouped_items

    keep_ids = set()
    skipped = []
    for item in fixture_items:
        point = item.get("location")
        nearest_distance = None
        for other in fixture_items:
            if other is item:
                continue
            distance = _distance_between_points(point, other.get("location"))
            if distance is None:
                continue
            if nearest_distance is None or distance < nearest_distance:
                nearest_distance = distance
            if distance <= TVTRUSS_MAX_NEIGHBOR_DISTANCE_FT:
                keep_ids.add(id(item))
                break
        if id(item) not in keep_ids:
            skipped.append((item, nearest_distance))

    if not skipped:
        return grouped_items

    skipped_item_ids = set(id(item) for item, _ in skipped)
    filtered = [item for item in grouped_items if id(item) not in skipped_item_ids]

    if logger:
        logger.info(
            "TVTRUSS proximity filter skipped {} electrical fixture(s) with no TVTRUSS fixture within {:.1f} ft.".format(
                len(skipped), TVTRUSS_MAX_NEIGHBOR_DISTANCE_FT
            )
        )
        for item, nearest in skipped:
            logger.info(
                "  Skipped element {} (nearest TVTRUSS fixture distance: {}).".format(
                    _item_id_text(item),
                    "{:.2f} ft".format(nearest) if nearest is not None else "unknown",
                )
            )

    return filtered


def _cluster_by_nearest(items, group_size):
    import math
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
            other_point = other.get("location")
            if seed_point and other_point:
                try:
                    dist = seed_point.DistanceTo(other_point)
                except Exception:
                    try:
                        dx = seed_point.X - other_point.X
                        dy = seed_point.Y - other_point.Y
                        dz = seed_point.Z - other_point.Z
                        dist = (dx * dx + dy * dy + dz * dz) ** 0.5
                    except Exception:
                        dist = 1e12
            else:
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


def _sanitize_token(value, fallback="X"):
    if not value:
        return fallback
    token = "".join(ch for ch in str(value) if ch.isalnum())
    return token or fallback


def create_position_groups(items, make_group, logger=None):
    if not items:
        return [], items
    grouped_items = []
    remaining = []
    for item in items:
        rule = _match_position_rule(item.get("circuit_number"))
        if rule:
            item["_pf_position_rule"] = rule
            grouped_items.append(item)
        else:
            remaining.append(item)

    if not grouped_items:
        return [], remaining

    grouped_items = _filter_tvtruss_fixture_outliers(grouped_items, logger=logger)
    if not grouped_items:
        return [], remaining

    buckets = {}
    for item in grouped_items:
        rule = item.get("_pf_position_rule") or {}
        keyword = rule.get("keyword") or ""
        panel_name = item.get("panel_name") or "NO_PANEL"
        bucket_key = (panel_name, keyword)
        buckets.setdefault(bucket_key, []).append(item)

    groups = []
    counters = {}
    for bucket_key in sorted(buckets.keys()):
        members = buckets[bucket_key]
        if not members:
            continue
        rule = members[0].get("_pf_position_rule") or {}
        group_size = int(rule.get("group_size") or 1)
        keyword = rule.get("keyword") or "KEY"

        position_groups = _cluster_by_nearest(members, group_size)

        for cluster in position_groups:
            count_key = (bucket_key[0], keyword)
            counters[count_key] = counters.get(count_key, 0) + 1
            group_index = counters[count_key]
            token = _sanitize_token(keyword, fallback="POS")
            key = "{}{}{}".format(bucket_key[0], token, group_index)
            group = make_group(key, cluster, group_type="position")
            groups.append(group)

    return groups, remaining
