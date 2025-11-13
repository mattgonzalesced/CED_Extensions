from collections import defaultdict

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


def get_load_priority(group):
    load_name = group.get("load_name")
    if load_name:
        key = load_name.strip().upper()
        if key in _LOAD_PRIORITY:
            return _LOAD_PRIORITY[key]
    return 3
