# -*- coding: utf-8 -*-
__title__ = "SUPER CIRCUIT V4"

from collections import OrderedDict
import logging
import re

from pyrevit import revit, script, forms, DB

from Snippets._elecutils import (
    get_all_data_devices,
    get_all_elec_fixtures,
    get_all_light_devices,
    get_all_light_fixtures,
    get_all_mech_control_devices,
    get_all_panels,
)

import circuits

logger = script.get_logger()
logger.setLevel(logging.INFO)

try:
    basestring
except NameError:  # Python 3 fallback
    basestring = str
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
            try:
                import importlib

                importlib.reload(HEBhelper)
                logger.info("Reloaded HEBhelper module.")
            except Exception:
                try:
                    reload(HEBhelper)  # IronPython fallback
                    logger.info("Reloaded HEBhelper module via reload().")
                except Exception as ex:
                    logger.warning("HEBhelper reload failed: {}".format(ex))

            return HEBhelper
        except ImportError as ex:
            logger.warning("HEB helpers unavailable: {}".format(ex))
    return None


def _split_panel_choices(value):
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


def _tokenize_panel_value(value):
    if not value:
        return []
    text = value if isinstance(value, basestring) else str(value)
    tokens = re.findall(r"[A-Za-z0-9]+", text)
    return [token.strip() for token in tokens if token.strip()]


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
        return DB.XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )
    return None


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


def _build_panel_point_cache(panel_lookup):
    cache = {}
    for name, info in (panel_lookup or {}).items():
        panel_elem = info.get("element")
        point = _get_point(panel_elem)
        if not point:
            continue
        upper = (name or "").strip().upper()
        if not upper or upper in cache:
            continue
        cache[upper] = {"name": name, "point": point}
    return cache


def _emit_debug(output, text):
    try:
        output.print_md(text)
    except Exception:
        pass
    try:
        print(text)
    except Exception:
        pass


def _debug_ba_da(label, items, panel_lookup, output):
    panel_point_cache = _build_panel_point_cache(panel_lookup)
    for item in items:
        panel_value = item.get("panel_name")
        tokens = _tokenize_panel_value(panel_value)
        token_set = {token.upper() for token in tokens}
        if not (("BA" in token_set) or ("DA" in token_set)):
            continue
        element = item.get("element")
        elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
        location = item.get("location")
        load_name = item.get("load_name") or "None"
        if location:
            loc_text = "{:.3f},{:.3f},{:.3f}".format(location.X, location.Y, location.Z)
        else:
            loc_text = "None"
        _emit_debug(
            output,
            "INFO SuperCircuitV4 {} BA/DA debug | element {} | location {} | CKT_Panel_CEDT {} | CKT_Load Name_CEDT {} | tokens {}".format(
                label,
                elem_id if elem_id is not None else "unknown",
                loc_text,
                panel_value or "None",
                load_name,
                ", ".join(sorted(token_set)) if token_set else "none",
            ),
        )
        for panel_key in ("BA", "DA"):
            if panel_key not in token_set:
                _emit_debug(
                    output,
                    "INFO SuperCircuitV4 {} BA/DA debug | panel {} not listed in CKT_Panel_CEDT".format(
                        label, panel_key
                    ),
                )
                continue
            entry = panel_point_cache.get(panel_key)
            if not entry:
                _emit_debug(
                    output,
                    "INFO SuperCircuitV4 {} BA/DA debug | panel {} missing from cache".format(
                        label, panel_key
                    ),
                )
                continue
            point = entry.get("point")
            dist = _distance(location, point) if location and point else None
            if point:
                ptext = "{:.3f},{:.3f},{:.3f}".format(point.X, point.Y, point.Z)
            else:
                ptext = "None"
            _emit_debug(
                output,
                "INFO SuperCircuitV4 {} BA/DA debug | panel {} point {} | distance {}".format(
                    label,
                    panel_key,
                    ptext,
                    "{:.3f}".format(dist) if dist is not None else "None",
                ),
            )


def _collect_elements(doc):
    collectors = (
        get_all_elec_fixtures,
        get_all_light_devices,
        get_all_light_fixtures,
        get_all_data_devices,
        get_all_mech_control_devices,
    )
    return circuits.collect_target_elements(doc, collectors, revit.get_selection, logger)


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
    return circuits.run_creation(
        doc,
        groups,
        lambda d, g: circuits.create_circuit(d, g, logger),
        logger,
    )


def _run_apply_data(doc, created_systems):
    circuits.run_apply_data(
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
        members = group.get("members") or []
        choice_counts = [m.get("panel_choice_count") for m in members if m.get("panel_choice_count") is not None]
        panel_choice_count = min(choice_counts) if choice_counts else 99
        panel_choice_priority = 0 if panel_choice_count == 1 else 1

        priority = _group_priority(group.get("group_type"))
        panel = (group.get("panel_name") or "").lower()
        load_priority = _load_priority(group, priority)
        circuit_number = group.get("circuit_number")
        circuit_sort = circuits.try_parse_int(circuit_number)
        if circuit_sort is None:
            circuit_sort = circuit_number or group.get("key") or ""
        return (
            panel_choice_priority,
            panel_choice_count,
            priority,
            panel,
            load_priority,
            circuit_sort,
            group.get("key"),
        )

    return sorted(groups, key=sort_key)


def main():
    client_key = _select_client()
    if not client_key:
        logger.info("No client selected; aborting.")
        return

    global client_helpers
    client_helpers = _load_client_helpers(client_key)
    if client_helpers:
        logger.info(
            "Client helpers loaded: %s",
            getattr(client_helpers, "__file__", "unknown"),
        )
        logger.info(
            "Client helpers has preprocess_items: %s",
            hasattr(client_helpers, "preprocess_items"),
        )

    doc = revit.doc
    panels = list(get_all_panels(doc))
    panel_lookup = circuits.build_panel_lookup(panels)

    elements = _collect_elements(doc)
    elements = _filter_disallowed_elements(elements)
    if not elements:
        logger.info("No elements found for processing.")
        return

    info_items = circuits.gather_element_info(doc, elements, panel_lookup, logger)
    logger.info("Gathered circuit info items: {}".format(len(info_items)))
    if not info_items:
        logger.info("No elements with circuit data were found.")
        return

    for item in info_items:
        panel_raw = item.get("panel_name")
        choices = _split_panel_choices(panel_raw)
        if choices:
            unique = {c.strip().upper() for c in choices if c and c.strip()}
            item["panel_choice_count"] = len(unique)

    output = script.get_output()
    _debug_ba_da("pre", info_items, panel_lookup, output)
    if client_helpers and hasattr(client_helpers, "preprocess_items"):
        try:
            logger.info("Calling client_helpers.preprocess_items")
            processed = client_helpers.preprocess_items(info_items, doc, panel_lookup, logger)
            if processed:
                info_items = processed
        except Exception as ex:
            logger.warning("Client preprocess_items failed: {}".format(ex))

    _debug_ba_da("post", info_items, panel_lookup, output)

    groups = circuits.assemble_groups(info_items, client_helpers, logger)
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
