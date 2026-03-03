# -*- coding: utf-8 -*-
__title__ = "SUPER CIRCUIT V3"

from collections import OrderedDict

from pyrevit import revit, script, forms, DB

from Snippets._elecutils import (
    get_all_data_devices,
    get_all_elec_fixtures,
    get_all_light_devices,
    get_all_light_fixtures,
    get_all_panels,
)

from libGeneral import data, grouping, circuits, transactions, common

logger = script.get_logger()
POSITION_GROUP_SIZE = 3
CLIENT_CHOICES = OrderedDict([
    ("Planet Fitness", "planet_fitness"),
    ("HEB", "heb"),
])
client_helpers = None
EXCLUDED_CATEGORY_IDS = {
    DB.ElementId(DB.BuiltInCategory.OST_LightingDevices).IntegerValue,
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures).IntegerValue,
}


def _select_client():
    selection = forms.CommandSwitchWindow.show(
        list(CLIENT_CHOICES.keys()), message="Select client configuration"
    )
    if not selection:
        return None
    return CLIENT_CHOICES.get(selection)


def _load_client_helpers(client_key):
    if client_key == "planet_fitness":
        try:
            from PFlib import PFhelpers

            return PFhelpers
        except ImportError as ex:
            logger.warning("PF helpers unavailable: {}".format(ex))
    elif client_key == "heb":
        try:
            from HEBlib import HEBhelper

            return HEBhelper
        except ImportError as ex:
            logger.warning("HEB helpers unavailable: {}".format(ex))
    return None


def _collect_elements(doc):
    collectors = (
        get_all_elec_fixtures,
        get_all_light_devices,
        get_all_light_fixtures,
        get_all_data_devices,
    )
    return data.collect_target_elements(doc, collectors, revit.get_selection, logger)


def _filter_disallowed_elements(elements):
    filtered = []
    skipped = 0
    for element in elements or []:
        category = getattr(element, "Category", None)
        category_id = category.Id.IntegerValue if category and category.Id else None
        if category_id in EXCLUDED_CATEGORY_IDS:
            skipped += 1
            continue
        filtered.append(element)

    if skipped:
        logger.info("Skipped {} lighting element(s); use the dedicated lighting tool.".format(skipped))
    return filtered


def _run_creation(doc, groups):
    return transactions.run_creation(
        doc,
        groups,
        lambda d, g: circuits.create_circuit(d, g, logger),
        logger,
    )


def _run_apply_data(doc, created_systems):
    transactions.run_apply_data(
        doc,
        created_systems,
        lambda system, group: circuits.apply_circuit_data(system, group, logger),
        logger,
    )


def _group_priority(group_type):
    priority_map = {
        "dedicated": 0,
        "nongrouped": 1,
        "special": 2,
        "position": 2,  # tv truss / positional groupings stay ahead of general load order
    }
    return priority_map.get(group_type or "normal", 3)


def _load_priority(group, group_priority_value):
    if group_priority_value < 3:
        return 0
    if client_helpers and hasattr(client_helpers, "get_load_priority"):
        try:
            return client_helpers.get_load_priority(group)
        except Exception as ex:
            logger.warning("Client load priority lookup failed: {}".format(ex))
    return 99


def _sort_groups(groups):
    def sort_key(group):
        priority = _group_priority(group.get("group_type"))
        panel = (group.get("panel_name") or "").lower()
        load_priority = _load_priority(group, priority)
        circuit_number = group.get("circuit_number")
        circuit_sort = common.try_parse_int(circuit_number)
        if circuit_sort is None:
            circuit_sort = circuit_number or group.get("key") or ""
        return (priority, panel, load_priority, circuit_sort, group.get("key"))

    return sorted(groups, key=sort_key)


def main():
    client_key = _select_client()
    if not client_key:
        logger.info("No client selected; aborting.")
        return

    global client_helpers
    client_helpers = _load_client_helpers(client_key)

    doc = revit.doc
    panels = list(get_all_panels(doc))
    panel_lookup = data.build_panel_lookup(panels)

    elements = _collect_elements(doc)
    elements = _filter_disallowed_elements(elements)
    if not elements:
        logger.info("No elements found for processing.")
        return

    info_items = data.gather_element_info(doc, elements, panel_lookup, logger)
    if not info_items:
        logger.info("No elements with circuit data were found.")
        return
    if client_helpers and hasattr(client_helpers, "preprocess_items"):
        try:
            processed = client_helpers.preprocess_items(info_items, doc, panel_lookup, logger)
            if processed:
                info_items = processed
        except Exception as ex:
            logger.warning("Client preprocess_items failed: {}".format(ex))

    groups = grouping.assemble_groups(info_items, client_helpers, POSITION_GROUP_SIZE, logger)
    if not groups:
        logger.info("Grouping produced no circuit batches.")
        return

    groups = _sort_groups(groups)

    created_systems = _run_creation(doc, groups)
    if not created_systems:
        logger.info("No circuits were created.")
        return

    _run_apply_data(doc, created_systems)

    logger.info("Created {} circuits.".format(len(created_systems)))


if __name__ == "__main__":
    main()
