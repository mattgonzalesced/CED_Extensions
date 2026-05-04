# -*- coding: utf-8 -*-
"""Headless Super Circuit V5 runner — no UI prompts, no preview window.

Called from the MCP automation pipeline via execute_revit_code.
Hardcodes client=planet_fitness, scope=all elements.

Usage (from execute_revit_code):
    import sys, os
    pushbutton_dir = r"C:\\CED_Extensions\\..."
    if pushbutton_dir not in sys.path:
        sys.path.insert(0, pushbutton_dir)
    from headless_run import run_headless
    result = run_headless(doc)
    print(result)
"""

import logging
import os
import sys

# Ensure CEDLib.lib is on the path (for Snippets._elecutils)
_LIB_ROOT = r"C:\CED_Extensions\CEDLib.lib"
if _LIB_ROOT not in sys.path:
    sys.path.append(_LIB_ROOT)

from collections import OrderedDict

from pyrevit import revit, script, DB

from Snippets._elecutils import (
    get_all_data_devices,
    get_all_elec_fixtures,
    get_all_light_devices,
    get_all_light_fixtures,
    get_all_mech_control_devices,
    get_all_panels,
)

import circuits

try:
    import importlib
    circuits = importlib.reload(circuits)
except Exception:
    try:
        circuits = reload(circuits)
    except Exception:
        pass

from PFlib import PFhelpers

try:
    import importlib
    PFhelpers = importlib.reload(PFhelpers)
except Exception:
    try:
        PFhelpers = reload(PFhelpers)
    except Exception:
        pass

logger = script.get_logger()
logger.setLevel(logging.INFO)

# Categories to exclude (lighting — handled by a separate tool)
EXCLUDED_CATEGORY_IDS = {
    DB.ElementId(DB.BuiltInCategory.OST_LightingDevices).IntegerValue,
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures).IntegerValue,
}


def _group_priority(group_type):
    priority_map = {
        "dedicated": 0,
        "special": 2,
        "position": 2,
    }
    return priority_map.get(group_type or "normal", 3)


def _load_priority(group, group_priority_value):
    if group_priority_value < 3:
        return 0
    if hasattr(PFhelpers, "get_load_priority"):
        try:
            return PFhelpers.get_load_priority(group)
        except Exception:
            pass
    return 99


def _sort_groups(groups):
    def sort_key(group):
        members = group.get("members") or []
        choice_counts = [
            m.get("panel_choice_count")
            for m in members
            if m.get("panel_choice_count") is not None
        ]
        panel_choice_count = min(choice_counts) if choice_counts else 99
        panel_choice_priority = 0 if panel_choice_count == 1 else 1

        priority = _group_priority(group.get("group_type"))
        panel = (group.get("panel_name") or "").lower()
        load_pri = _load_priority(group, priority)
        circuit_number = group.get("circuit_number")
        circuit_sort = circuits.try_parse_int(circuit_number)
        if circuit_sort is None:
            circuit_sort = circuit_number or group.get("key") or ""
        return (
            panel_choice_priority,
            panel_choice_count,
            priority,
            panel,
            load_pri,
            circuit_sort,
            group.get("key"),
        )

    return sorted(groups, key=sort_key)


def _split_panel_choices(value):
    try:
        basestring
    except NameError:
        basestring = str
    if not value:
        return []
    text = value if isinstance(value, basestring) else str(value)
    for sep in (",", ";", "|", "/", "\n", "\r"):
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


def _update_panel_choice_counts(items):
    for item in items or []:
        panel_choices = item.get("panel_choices")
        if panel_choices:
            item["panel_choice_count"] = len(panel_choices)
            continue
        panel_raw = item.get("panel_name")
        choices = _split_panel_choices(panel_raw)
        if choices:
            unique = {c.strip().upper() for c in choices if c and c.strip()}
            item["panel_choice_count"] = len(unique)
        else:
            item["panel_choice_count"] = None


def _collect_elements(doc):
    """Collect all eligible electrical elements (no selection filter)."""
    collectors = (
        get_all_elec_fixtures,
        get_all_light_devices,
        get_all_light_fixtures,
        get_all_data_devices,
        get_all_mech_control_devices,
    )
    return circuits.collect_target_elements(
        doc, collectors, lambda: [], logger
    )


def _filter_disallowed(elements):
    """Remove lighting elements (handled by dedicated lighting tool)."""
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
        logger.info(
            "Skipped {} lighting element(s).".format(skipped)
        )
    return filtered


def _patch_overflow_prompt():
    """Monkey-patch the overflow panel prompt so it never blocks.

    In headless mode there is no user to pick an overflow panel.
    Instead, log a warning and leave the circuit unassigned.
    """
    def _no_prompt(doc, base_panel_name, exclude_ids=None, logger=None):
        if logger:
            logger.warning(
                "Headless mode: overflow panel prompt skipped for '{}'. "
                "Circuit left unassigned.".format(base_panel_name)
            )
        return None

    circuits._get_overflow_panel = _no_prompt


def run_headless(doc):
    """Run Super Circuit V5 headless for Planet Fitness.

    Args:
        doc: The active Revit document (injected by MCP route).

    Returns:
        dict with created_count and skipped info.
    """
    _patch_overflow_prompt()

    # 1. Collect panels and build lookup
    panels = list(get_all_panels(doc))
    panel_lookup = circuits.build_panel_lookup(panels)
    logger.info("Panels found: {}".format(len(panel_lookup)))

    # 2. Collect all eligible elements
    elements = _collect_elements(doc)
    elements = _filter_disallowed(elements)
    if not elements:
        msg = "No eligible elements found."
        logger.info(msg)
        return {"created_count": 0, "message": msg}
    logger.info("Eligible elements: {}".format(len(elements)))

    # 3. Gather circuit info from element parameters
    info_items = circuits.gather_element_info(doc, elements, panel_lookup, logger)
    logger.info("Items with circuit data: {}".format(len(info_items)))
    if not info_items:
        msg = "No elements with circuit data (CKT_Panel_CEDT / CKT_Circuit Number_CEDT)."
        logger.info(msg)
        return {"created_count": 0, "message": msg}

    _update_panel_choice_counts(info_items)

    # 4. PF client preprocess (classify_items is called inside assemble_groups,
    #    but PFhelpers has no preprocess_items — only HEB does)
    if hasattr(PFhelpers, "preprocess_items"):
        try:
            processed = PFhelpers.preprocess_items(info_items, doc, panel_lookup, logger)
            if processed:
                info_items = processed
            _update_panel_choice_counts(info_items)
        except Exception as ex:
            logger.warning("PF preprocess_items failed: {}".format(ex))

    # 5. Assemble groups (dedicated, position/TVTRUSS, BYPARENT, normal)
    groups = circuits.assemble_groups(info_items, PFhelpers, logger)
    if not groups:
        msg = "Grouping produced no circuit batches."
        logger.info(msg)
        return {"created_count": 0, "message": msg}

    groups = _sort_groups(groups)
    logger.info("Circuit groups assembled: {}".format(len(groups)))

    # 6. Create circuits (transaction managed by circuits.run_creation)
    created_systems = circuits.run_creation(
        doc,
        groups,
        lambda d, g: circuits.create_circuit(d, g, logger),
        logger,
        transaction_label="SuperCircuitV5 Headless - Create Circuits",
    )

    # 7. Apply circuit data (load name, rating, notes)
    circuits.run_apply_data(
        doc,
        created_systems,
        lambda system, group: circuits.apply_circuit_data(system, group, logger),
        logger,
        transaction_label="SuperCircuitV5 Headless - Apply Circuit Data",
    )

    created_count = len(created_systems)
    msg = "Created {} circuit(s) from {} group(s).".format(created_count, len(groups))
    logger.info(msg)
    print(msg)
    return {"created_count": created_count, "message": msg}


# When exec'd directly by MCP, `doc` is injected by the route handler
if __name__ == "__main__" or "doc" in dir():
    try:
        result = run_headless(doc)
        print("SuperCircuitV5 headless complete: " + str(result))
    except Exception as ex:
        import traceback
        traceback.print_exc()
        print("SuperCircuitV5 headless FAILED: " + str(ex))
