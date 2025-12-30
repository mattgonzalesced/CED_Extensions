from collections import defaultdict
import math

from libGeneral.common import try_parse_int


def make_group(key, members, group_type=None, parent_key=None):
    sample = members[0]
    rating = sample.get("rating")
    load_name = sample.get("load_name")
    circuit_number = sample.get("circuit_number")
    circuit_notes = sample.get("circuit_notes")
    number_of_poles = sample.get("number_of_poles")

    if not rating:
        for item in members:
            if item.get("rating"):
                rating = item["rating"]
                break
    if not load_name:
        for item in members:
            if item.get("load_name"):
                load_name = item["load_name"]
                break
    if not circuit_notes:
        for item in members:
            if item.get("circuit_notes"):
                circuit_notes = item["circuit_notes"]
                break

    poles_candidates = [item.get("number_of_poles") for item in members if item.get("number_of_poles")]
    connector_candidates = [item.get("connector_poles") for item in members if item.get("connector_poles")]
    aggregated_poles = []
    if poles_candidates:
        aggregated_poles.extend(int(p) for p in poles_candidates)
    if connector_candidates:
        aggregated_poles.extend(int(p) for p in connector_candidates)
    if aggregated_poles:
        number_of_poles = max(aggregated_poles)
    number_of_poles = int(number_of_poles) if number_of_poles else 1
    connector_poles = max(connector_candidates) if connector_candidates else None

    panel_distribution_ids = []
    seen_ids = set()
    for item in members:
        for ds_id in item.get("panel_distribution_system_ids") or []:
            if ds_id in seen_ids:
                continue
            seen_ids.add(ds_id)
            panel_distribution_ids.append(ds_id)

    return {
        "key": key,
        "members": members,
        "panel_name": sample.get("panel_name"),
        "panel_element": sample.get("panel_element"),
        "circuit_number": circuit_number,
        "rating": rating,
        "load_name": load_name,
        "circuit_notes": circuit_notes,
        "number_of_poles": number_of_poles,
        "group_type": group_type,
        "parent_key": parent_key,
        "connector_poles": connector_poles,
        "panel_distribution_system_ids": panel_distribution_ids,
    }


def _split_combined(client_helpers, panel_name, circuit_number, members, logger, parse_int):
    if client_helpers and hasattr(client_helpers, "split_combined_circuit"):
        try:
            result = client_helpers.split_combined_circuit(
                panel_name,
                circuit_number,
                members,
                make_group,
                logger=logger,
                parse_int=parse_int,
            )
            if result:
                for group in result:
                    group["group_type"] = "special"
            return result
        except Exception as ex:
            if logger:
                logger.warning("Client split_combined_circuit failed: {}".format(ex))
    return None


def create_dedicated_groups(items):
    counters = defaultdict(lambda: defaultdict(int))
    groups = []
    for item in items:
        panel_name = item.get("panel_name") or "NO_PANEL"
        pole_count = try_parse_int(item.get("number_of_poles")) or 1
        label = {1: "DEDICATED", 2: "DEDICATED2POLE", 3: "DEDICATED3POLE"}.get(pole_count, "DEDICATED")
        counters[panel_name][label] += 1
        key = "{}{}{}".format(panel_name, label, counters[panel_name][label])
        groups.append(make_group(key, [item], group_type="dedicated"))
    return groups


def create_nongroupedblock_groups(items):
    groups_by_panel = defaultdict(list)
    for item in items:
        panel_name = item.get("panel_name") or "NO_PANEL"
        groups_by_panel[panel_name].append(item)

    groups = []
    for panel_name in sorted(groups_by_panel.keys(), key=lambda x: x or ""):
        members = groups_by_panel[panel_name]
        key = "{}NONGROUPEDBLOCK".format(panel_name)
        groups.append(make_group(key, members, group_type="nongrouped"))
    return groups


def group_by_key(items, client_helpers, logger):
    grouped = defaultdict(list)
    for item in items:
        panel_name = item.get("panel_name")
        circuit_number = item.get("circuit_number")
        if not panel_name or not circuit_number:
            if logger:
                logger.debug(
                    "Skipping element {} missing panel or circuit number for key grouping.".format(
                        item["element"].Id
                    )
                )
            continue
        grouped[(panel_name, circuit_number)].append(item)

    groups = []
    for panel_name, circuit_number in sorted(
        grouped.keys(),
        key=lambda k: (
            (k[0] or "").lower(),
            try_parse_int(k[1]) if try_parse_int(k[1]) is not None else (k[1] or "").lower(),
        ),
    ):
        members = grouped[(panel_name, circuit_number)]
        split_groups = _split_combined(client_helpers, panel_name, circuit_number, members, logger, try_parse_int)
        if split_groups:
            groups.extend(split_groups)
        else:
            key = "{}{}".format(panel_name, circuit_number)
            groups.append(make_group(key, members, group_type="normal"))
    return groups


def _position_sort_key(item):
    location = item.get("location")
    if not location:
        return (math.inf, math.inf, math.inf)
    return (location.X, location.Y, location.Z)


def group_by_position(items, group_size, logger):
    buckets = defaultdict(list)
    for item in items:
        panel_name = item.get("panel_name")
        load_name = item.get("load_name") or ""
        if not panel_name:
            if logger:
                logger.debug(
                    "Skipping element {} missing panel for position grouping.".format(
                        item["element"].Id
                    )
                )
            continue
        buckets[(panel_name, load_name)].append(item)

    groups = []
    for (panel_name, load_name), members in buckets.items():
        sorted_members = sorted(members, key=_position_sort_key)
        chunk_count = int(math.ceil(len(sorted_members) / float(group_size)))
        for index in range(chunk_count):
            chunk = sorted_members[index * group_size : (index + 1) * group_size]
            if not chunk:
                continue
            sanitized = "".join(ch for ch in (load_name or "") if ch.isalnum()) or "UNSPECIFIED"
            key = "{}{}_POS{}".format(panel_name, sanitized, index + 1)
            groups.append(make_group(key, chunk, group_type="position"))

    return groups


def _split_group_by_poles(base_group, logger=None):
    members = base_group.get("members") or []
    if not members:
        return [base_group]

    buckets = defaultdict(list)
    for member in members:
        pole_value = member.get("connector_poles") or member.get("number_of_poles")
        pole_value = int(pole_value) if pole_value else base_group.get("number_of_poles") or 1
        buckets[pole_value].append(member)

    if len(buckets) <= 1:
        return [base_group]

    split_groups = []
    for pole_value in sorted(buckets.keys()):
        pole_members = buckets[pole_value]
        new_key = "{}_{:d}P".format(base_group.get("key", ""), pole_value)
        new_group = make_group(
            new_key,
            pole_members,
            group_type=base_group.get("group_type"),
            parent_key=base_group.get("key"),
        )
        split_groups.append(new_group)
        if logger:
            logger.info(
                "Split group {} into {} by pole count {} ({} member(s)).".format(
                    base_group.get("key"), new_key, pole_value, len(pole_members)
                )
            )
    return split_groups


def split_groups_by_poles(groups, logger=None):
    result = []
    for group in groups:
        result.extend(_split_group_by_poles(group, logger))
    return result


def assemble_groups(items, client_helpers, position_group_size, logger):
    working_items = list(items)
    groups = []

    if client_helpers and hasattr(client_helpers, "create_position_groups"):
        try:
            position_groups, remaining = client_helpers.create_position_groups(
                working_items, make_group, logger=logger
            )
            if position_groups:
                groups.extend(position_groups)
            if remaining is not None:
                working_items = list(remaining)
        except Exception as ex:
            if logger:
                logger.warning("Client create_position_groups failed: {}".format(ex))

    dedicated, nongrouped, tvtruss, normal = [], [], [], list(working_items)
    if client_helpers and hasattr(client_helpers, "classify_items"):
        try:
            dedicated, nongrouped, tvtruss, normal = client_helpers.classify_items(working_items)
        except Exception as ex:
            if logger:
                logger.warning("Client classify_items failed: {}".format(ex))

    if dedicated:
        groups.extend(create_dedicated_groups(dedicated))

    if nongrouped:
        groups.extend(create_nongroupedblock_groups(nongrouped))

    if tvtruss:
        groups.extend(group_by_position(tvtruss, position_group_size, logger))

    if normal:
        groups.extend(group_by_key(normal, client_helpers, logger))

    groups = split_groups_by_poles(groups, logger)
    return groups
