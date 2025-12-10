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


def _nearest_panel(location, panel_cache):
    if not location:
        return None
    best = None
    best_dist = None
    for entry in panel_cache.values():
        dist = _distance(location, entry.get("point"))
        if dist is None:
            continue
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best = entry
    return best


def preprocess_items(items, doc, panel_lookup, logger=None):
    if not items:
        return items
    panel_cache = _build_panel_cache(panel_lookup)
    if not panel_cache:
        return items
    for item in items:
        location = item.get("location")
        allowed_raw = _split_panel_choices(item.get("panel_name"))
        if not location or not allowed_raw:
            continue
        allowed = {name.strip().upper(): name.strip() for name in allowed_raw if name.strip()}
        if not allowed:
            continue
        nearest = _nearest_panel(location, panel_cache)
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
