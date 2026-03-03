# -*- coding: utf-8 -*-
__title__ = "Circuit By Space"

import math
import os
import sys
from collections import OrderedDict

from pyrevit import forms, revit, script


def _bundle_path():
    try:
        getter = getattr(script, "get_bundle_path", None)
        if callable(getter):
            return getter()
    except Exception:
        pass
    try:
        bundle_file = getattr(script, "get_bundle_file", None)
        if callable(bundle_file):
            bundle_root = bundle_file()
            if bundle_root:
                return os.path.dirname(bundle_root)
    except Exception:
        pass
    return os.path.dirname(__file__)


BUNDLE_PATH = _bundle_path()
SUPER_CIRCUIT_PATH = os.path.normpath(
    os.path.join(BUNDLE_PATH, "..", "SuperCircuitV3.pushbutton")
)
if BUNDLE_PATH not in sys.path:
    sys.path.append(BUNDLE_PATH)
if SUPER_CIRCUIT_PATH not in sys.path:
    sys.path.append(SUPER_CIRCUIT_PATH)

from Snippets._elecutils import get_all_light_fixtures, get_all_panels
from libGeneral import circuits, common, data, grouping, transactions
from organized.MEPKit.electrical import selection as spatial_selection


CLIENT_CHOICES = OrderedDict(
    [
        ("General", None),
        ("Planet Fitness", "planet_fitness"),
    ]
)
SPACE_FALLBACK_LABEL = "Unassigned Space"

logger = script.get_logger()


def _select_client_key():
    if len(CLIENT_CHOICES) <= 1:
        return next(iter(CLIENT_CHOICES.values()))

    selection = forms.CommandSwitchWindow.show(
        list(CLIENT_CHOICES.keys()),
        message="Select client-specific space grouping (Esc for general).",
    )
    if not selection:
        return None
    return CLIENT_CHOICES.get(selection)


def _load_client_helper(client_key):
    if client_key == "planet_fitness":
        try:
            import PFspacehelper

            return PFspacehelper
        except ImportError as ex:
            logger.warning("PF space helper unavailable: {}".format(ex))
    return None


def _collect_lighting(doc):
    collectors = (get_all_light_fixtures,)
    return data.collect_target_elements(doc, collectors, revit.get_selection, logger)


def _describe_fixture_type(element):
    symbol = getattr(element, "Symbol", None)
    family = getattr(getattr(symbol, "Family", None), "Name", None) if symbol else None
    type_name = getattr(symbol, "Name", None) if symbol else None
    family = (family or "").strip()
    type_name = (type_name or "").strip()
    if family and type_name:
        label = "{} - {}".format(family, type_name)
    else:
        label = family or type_name or "Lighting Fixture"
    key = "{}::{}".format(family or "FAM", type_name or "TYPE")
    return key.upper(), label


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
                value = common.get_param_value(space, name)
            except Exception:
                value = None
        if value:
            text = str(value).strip()
            if text:
                return text
    return None


def _format_space_label(space):
    if not space:
        return SPACE_FALLBACK_LABEL

    label = _extract_space_attr(space, "Name", "Space Name", "Room Name")
    if not label:
        label = _extract_space_attr(space, "Number", "Space Number", "Room Number")
    return label or SPACE_FALLBACK_LABEL


def _extract_space_number(space):
    if not space:
        return None
    return _extract_space_attr(space, "Number", "Space Number", "Room Number")


def _resolve_space_bucket(space, label, circuit_class, helper):
    display_label = label or SPACE_FALLBACK_LABEL

    if helper and hasattr(helper, "get_space_bucket"):
        try:
            override = helper.get_space_bucket(space, display_label, circuit_class)
        except Exception as ex:
            logger.warning("Client space helper failed: {}".format(ex))
        else:
            if override and override.get("bucket_id"):
                custom_label = override.get("label") or display_label
                return override["bucket_id"], custom_label

    if space is not None:
        try:
            space_id = space.Id.IntegerValue
        except Exception:
            space_id = None
        if space_id:
            return "SPACE_{}".format(space_id), display_label

    return "SPACE_UNASSIGNED", display_label


def _classify_circuit(item):
    marker = "{} {}".format(item.get("circuit_number") or "", item.get("load_name") or "")
    if "EMER" in marker.upper():
        return "emergency"
    return "standard"


def _prepare_items(doc, info_items, helper):
    prepared = []
    for item in info_items:
        panel_name = common.safe_strip(item.get("panel_name"))
        panel_element = item.get("panel_element")
        if not panel_name or panel_element is None:
            host = item.get("element")
            try:
                host_id = host.Id.IntegerValue
            except Exception:
                host_id = "?"
            logger.warning("Skipping element {} without a resolvable panel.".format(host_id))
            continue

        space = spatial_selection.element_space_or_room(item.get("element"), doc)
        space_label = _format_space_label(space)
        fixture_key, fixture_label = _describe_fixture_type(item.get("element"))
        space_number = _extract_space_number(space)
        circuit_class = _classify_circuit(item)
        bucket_id, display_label = _resolve_space_bucket(space, space_label, circuit_class, helper)

        enriched = dict(item)
        enriched["panel_name"] = panel_name
        enriched["space_element"] = space
        enriched["space_display_label"] = display_label
        enriched["space_bucket_id"] = bucket_id
        enriched["space_number"] = space_number
        enriched["circuit_class"] = circuit_class
        enriched["fixture_type_key"] = fixture_key
        enriched["fixture_type_label"] = fixture_label
        prepared.append(enriched)

    return prepared


def _apply_space_label_to_load(load_name, space_label, circuit_class):
    label = space_label or SPACE_FALLBACK_LABEL
    if load_name and "(Space)" in load_name:
        return load_name.replace("(Space)", label)
    if load_name:
        return load_name
    if circuit_class == "emergency":
        return "Emergency - {}".format(label)
    return label


def _group_by_space(prepared):
    grouped = OrderedDict()
    for item in prepared:
        key = (item["panel_name"], item["space_bucket_id"], item["circuit_class"])
        if key not in grouped:
            grouped[key] = {
                "members": [],
                "space_label": item["space_display_label"],
                "bucket_id": item["space_bucket_id"],
                "space_number": item.get("space_number"),
                "circuit_class": item["circuit_class"],
            }
        grouped[key]["members"].append(item)

    groups = []
    for (panel_name, bucket_id, circuit_class), payload in grouped.items():
        group_key = "{}|{}|{}".format(panel_name, circuit_class.upper(), bucket_id)
        group = grouping.make_group(group_key, payload["members"], group_type="space")
        _assign_space_metadata(
            group,
            payload["bucket_id"],
            payload["space_label"],
            circuit_class,
            payload.get("space_number"),
        )
        groups.append(group)
    return groups


def _assign_space_metadata(group, bucket_id, space_label, circuit_class, space_number=None):
    group["space_bucket_id"] = bucket_id
    group["space_display_label"] = space_label
    group["circuit_class"] = circuit_class
    group["space_number"] = space_number
    group["load_name"] = _apply_space_label_to_load(
        group.get("load_name"),
        space_label,
        circuit_class,
    )
    group["base_load_name"] = group.get("load_name")


def _sanitize_token(value):
    if not value:
        return "X"
    token = "".join(ch for ch in str(value) if ch.isalnum())
    return token[:24] or "X"


def _clone_group(base_group, members, suffix):
    new_key = "{}|{}".format(base_group.get("key"), suffix)
    clone = grouping.make_group(
        new_key,
        list(members),
        group_type=base_group.get("group_type"),
        parent_key=base_group.get("key"),
    )
    _assign_space_metadata(
        clone,
        base_group.get("space_bucket_id"),
        base_group.get("space_display_label"),
        base_group.get("circuit_class"),
        base_group.get("space_number"),
    )
    return clone


def _split_group_by_fixture_types(group):
    members = group.get("members") or []
    buckets = OrderedDict()
    for member in members:
        type_key = member.get("fixture_type_key") or "UNKNOWN"
        buckets.setdefault(type_key, []).append(member)

    if len(buckets) <= 1:
        return [group]

    split_groups = []
    for index, (type_key, items) in enumerate(buckets.items(), start=1):
        suffix = "FT{}{}".format(index, _sanitize_token(type_key))
        split_groups.append(_clone_group(group, items, suffix))

    return split_groups


def _split_group_members(group, attempt_index):
    members = list(group.get("members") or [])
    if len(members) <= 1:
        return None

    split_size = int(math.ceil(len(members) / 2.0))
    first = members[:split_size]
    second = members[split_size:]
    if not second:
        second = [first.pop()]

    first_suffix = "S{}A".format(attempt_index)
    second_suffix = "S{}B".format(attempt_index)
    return [
        _clone_group(group, first, first_suffix),
        _clone_group(group, second, second_suffix),
    ]


def _enforce_minimum_groups(groups, min_required):
    if min_required <= 1 or len(groups) >= min_required:
        return groups

    working = list(groups)
    attempt = 1
    while len(working) < min_required:
        target = max(working, key=lambda g: len(g.get("members") or []))
        split_pair = _split_group_members(target, attempt)
        if not split_pair:
            logger.warning(
                "Unable to reach required circuit count ({}) for space {} {} due to limited fixtures.".format(
                    min_required,
                    target.get("space_display_label"),
                    target.get("circuit_class"),
                )
            )
            break
        working.remove(target)
        working.extend(split_pair)
        attempt += 1

    return working


def _apply_space_rules(groups, helper):
    if not helper or not hasattr(helper, "get_space_rules"):
        return groups

    adjusted = []
    for group in groups:
        rules = helper.get_space_rules(group.get("space_display_label"))
        if not rules:
            adjusted.append(group)
            continue

        processed = [group]
        if rules.get("split_by_fixture"):
            temp = []
            for item in processed:
                temp.extend(_split_group_by_fixture_types(item))
            processed = temp

        min_map = rules.get("min_circuits") or {}
        required = max(int(min_map.get(group.get("circuit_class"), 1)), 1)
        processed = _enforce_minimum_groups(processed, required)

        adjusted.extend(processed)

    return adjusted


def _apply_cardio_threeway(groups):
    CARDIO_SPACE_NUMBERS = {"211"}
    TARGET_GROUPS = 3
    preserved = []
    bucket = {}

    for group in groups:
        space_number = (group.get("space_number") or "").strip()
        circuit_class = group.get("circuit_class")
        if space_number in CARDIO_SPACE_NUMBERS and circuit_class == "standard":
            key = (space_number, circuit_class)
            bucket.setdefault(key, []).append(group)
        else:
            preserved.append(group)

    for (space_number, circuit_class), related_groups in bucket.items():
        all_members = []
        for grp in related_groups:
            all_members.extend(grp.get("members") or [])

        if not all_members:
            preserved.extend(related_groups)
            continue

        base_group = related_groups[0]
        chunk_size = int(math.ceil(len(all_members) / float(TARGET_GROUPS)))
        new_groups = []
        for idx in range(TARGET_GROUPS):
            start = idx * chunk_size
            chunk = all_members[start : start + chunk_size]
            if not chunk:
                continue
            suffix = str(idx + 1)
            clone = _clone_group(base_group, chunk, "CARDIO{}".format(suffix))
            base_name = clone.get("base_load_name") or clone.get("load_name") or clone.get("space_display_label") or "Space {}".format(space_number)
            clone["load_name"] = "{} {}".format(base_name, suffix)
            new_groups.append(clone)

        if len(new_groups) < TARGET_GROUPS:
            logger.warning(
                "Space {} ({}) only produced {} circuit(s) out of {} due to limited fixtures.".format(
                    base_group.get("space_display_label"),
                    circuit_class,
                    len(new_groups),
                    TARGET_GROUPS,
                )
            )

        preserved.extend(new_groups)

    return preserved


def _apply_functional_pair(groups):
    TARGET_NUMBERS = {"107"}
    TARGET_LABEL_TOKENS = ("FUNCTIONAL", "TRAIN")
    TARGET_COUNT = 2

    preserved = []
    bucket = {}

    for group in groups:
        space_number = (group.get("space_number") or "").strip()
        space_label = (group.get("space_display_label") or "").upper()
        circuit_class = group.get("circuit_class")
        matches_label = all(token in space_label for token in TARGET_LABEL_TOKENS)
        if circuit_class == "standard" and (space_number in TARGET_NUMBERS or matches_label):
            key = (space_number or space_label, circuit_class)
            bucket.setdefault(key, []).append(group)
        else:
            preserved.append(group)

    for (space_key, circuit_class), related_groups in bucket.items():
        all_members = []
        for grp in related_groups:
            all_members.extend(grp.get("members") or [])

        if not all_members:
            preserved.extend(related_groups)
            continue

        base_group = related_groups[0]
        chunk_size = int(math.ceil(len(all_members) / float(TARGET_COUNT)))
        new_groups = []
        for idx in range(TARGET_COUNT):
            chunk = all_members[idx * chunk_size : (idx + 1) * chunk_size]
            if not chunk:
                continue
            suffix = str(idx + 1)
            clone = _clone_group(base_group, chunk, "FUNC{}".format(suffix))
            base_name = (
                clone.get("base_load_name")
                or clone.get("load_name")
                or clone.get("space_display_label")
                or "Space {}".format(space_key)
            )
            clone["load_name"] = "{} {}".format(base_name, suffix)
            new_groups.append(clone)

        if len(new_groups) < TARGET_COUNT:
            logger.warning(
                "Space {} ({}) only produced {} pair circuit(s) out of {} due to limited fixtures.".format(
                    base_group.get("space_display_label"),
                    circuit_class,
                    len(new_groups),
                    TARGET_COUNT,
                )
            )

        preserved.extend(new_groups)

    return preserved


def _sort_groups(groups):
    def sort_key(group):
        panel = (group.get("panel_name") or "").lower()
        circuit_class = group.get("circuit_class") or ""
        space_label = (group.get("space_display_label") or "").lower()
        return (panel, circuit_class, space_label)

    return sorted(groups, key=sort_key)


def main():
    doc = revit.doc
    client_key = _select_client_key()
    helper = _load_client_helper(client_key)

    panels = list(get_all_panels(doc))
    panel_lookup = data.build_panel_lookup(panels)

    elements = _collect_lighting(doc)
    if not elements:
        logger.info("No lighting fixtures selected or found.")
        return

    info_items = data.gather_element_info(doc, elements, panel_lookup, logger)
    if not info_items:
        logger.info("No lighting fixtures with circuit data were found.")
        return

    prepared = _prepare_items(doc, info_items, helper)
    if not prepared:
        logger.info("No lighting fixtures qualified for space grouping.")
        return

    groups = _group_by_space(prepared)
    groups = _apply_space_rules(groups, helper)
    if helper:
        groups = _apply_cardio_threeway(groups)
        groups = _apply_functional_pair(groups)
    if not groups:
        logger.info("Grouping produced no circuit batches.")
        return

    groups = _sort_groups(groups)

    created_systems = transactions.run_creation(
        doc,
        groups,
        lambda d, g: circuits.create_circuit(d, g, logger),
        logger,
        transaction_label="CircuitBySpace - Create Circuits",
    )
    if not created_systems:
        logger.info("No circuits were created.")
        return

    transactions.run_apply_data(
        doc,
        created_systems,
        lambda system, group: circuits.apply_circuit_data(system, group, logger),
        logger,
        transaction_label="CircuitBySpace - Apply Circuit Data",
    )

    std_count = sum(1 for group in groups if group.get("circuit_class") == "standard")
    em_count = sum(1 for group in groups if group.get("circuit_class") == "emergency")
    logger.info(
        "CircuitBySpace created {} circuits ({} standard / {} emergency).".format(
            len(created_systems), std_count, em_count
        )
    )


if __name__ == "__main__":
    main()
