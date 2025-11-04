# -*- coding: utf-8 -*-
__title__ = "SUPER CIRCUIT V3"

from collections import defaultdict, OrderedDict
import math

from System.Collections.Generic import List

from pyrevit import revit, DB, script
from pyrevit.revit.db import query

from Snippets._elecutils import (
    get_all_data_devices,
    get_all_elec_fixtures,
    get_all_light_devices,
    get_all_light_fixtures,
    get_all_panels,
)

# Global switches that determine how circuits will be grouped.
CIRCUITBYKEY = True
CIRCUITBYPOSITION = False

# Configuration for position-based grouping.
POSITION_GROUP_SIZE = 3  # how many devices to lump together when grouping by proximity

logger = script.get_logger()


def _safe_strip(value):
    return value.strip() if isinstance(value, basestring) else value


def _get_param_value(element, param_name):
    param = query.get_param(element, param_name)
    return _safe_strip(query.get_param_value(param)) if param else None


def _get_element_location(element):
    location = getattr(element, "Location", None)
    if location is not None:
        point = getattr(location, "Point", None)
        if point:
            return point
        curve = getattr(location, "Curve", None)
        if curve:
            return curve.Evaluate(0.5, True)
    bbox = element.get_BoundingBox(None)
    if bbox:
        return DB.XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )
    return DB.XYZ.Zero


def _collect_target_elements(doc):
    selection = revit.get_selection()
    if selection:
        return list(selection)

    elements = []
    collectors = (
        get_all_elec_fixtures,
        get_all_light_devices,
        get_all_light_fixtures,
        get_all_data_devices,
    )
    for fn in collectors:
        try:
            elements.extend(list(fn(doc)))
        except Exception as ex:
            logger.warning("Collector {} failed: {}".format(fn.__name__, ex))
    return elements


def _build_panel_lookup(panels):
    lookup = {}
    for panel in panels:
        panel_name_param = query.get_param(panel, "Panel Name")
        panel_name = query.get_param_value(panel_name_param) if panel_name_param else None
        if panel_name:
            lookup[panel_name.strip()] = panel
    return lookup


def _gather_element_info(elements, panel_lookup):
    info_items = []
    for element in elements:
        panel_name = _get_param_value(element, "CKT_Panel_CEDT")
        circuit_number = _get_param_value(element, "CKT_Circuit Number_CEDT")
        rating = _get_param_value(element, "CKT_Rating_CED")
        load_name = _get_param_value(element, "CKT_Load Name_CEDT")

        if not panel_name and not circuit_number:
            continue

        info_items.append(
            {
                "element": element,
                "panel_name": panel_name,
                "panel_element": panel_lookup.get(panel_name),
                "circuit_number": circuit_number,
                "rating": rating,
                "load_name": load_name,
                "location": _get_element_location(element),
            }
        )
    return info_items


def _separate_dedicated(items):
    dedicated = []
    normal = []
    for item in items:
        circuit_number = item.get("circuit_number")
        if circuit_number and circuit_number.upper() == "DEDICATED":
            dedicated.append(item)
        else:
            normal.append(item)
    return dedicated, normal


def _make_group(key, members):
    sample = members[0]
    rating = sample.get("rating")
    load_name = sample.get("load_name")
    circuit_number = sample.get("circuit_number")

    # Prefer the first non-empty rating/load name within the group.
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

    return OrderedDict(
        [
            ("key", key),
            ("members", members),
            ("panel_name", sample.get("panel_name")),
            ("panel_element", sample.get("panel_element")),
            ("circuit_number", circuit_number),
            ("rating", rating),
            ("load_name", load_name),
        ]
    )


def _create_dedicated_groups(items):
    counters = defaultdict(int)
    groups = []
    for item in items:
        panel_name = item.get("panel_name") or "NO_PANEL"
        counters[panel_name] += 1
        key = "{}DEDICATED{}".format(panel_name, counters[panel_name])
        groups.append(_make_group(key, [item]))
    return groups


def _group_by_key(items):
    grouped = defaultdict(list)
    for item in items:
        panel_name = item.get("panel_name")
        circuit_number = item.get("circuit_number")
        if not panel_name or not circuit_number:
            logger.debug(
                "Skipping element {} missing panel or circuit number for key grouping.".format(
                    item["element"].Id
                )
            )
            continue
        key = "{}{}".format(panel_name, circuit_number)
        grouped[key].append(item)

    groups = []
    for key in sorted(grouped.keys(), key=lambda x: x.lower()):
        groups.append(_make_group(key, grouped[key]))
    return groups


def _position_sort_key(item):
    location = item.get("location")
    if not location:
        return (math.inf, math.inf, math.inf)
    return (location.X, location.Y, location.Z)


def _sanitize_for_key(value):
    if not value:
        return "UNSPECIFIED"
    return "".join(ch for ch in value if ch.isalnum())


def _group_by_position(items, group_size):
    buckets = defaultdict(list)
    for item in items:
        panel_name = item.get("panel_name")
        load_name = item.get("load_name") or ""
        if not panel_name:
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
            key = "{}{}_POS{}".format(panel_name, _sanitize_for_key(load_name), index + 1)
            groups.append(_make_group(key, chunk))
    return groups


def _assemble_groups(items):
    dedicated, normal = _separate_dedicated(items)
    groups = _create_dedicated_groups(dedicated)

    if CIRCUITBYPOSITION:
        groups.extend(_group_by_position(normal, POSITION_GROUP_SIZE))
    elif CIRCUITBYKEY:
        groups.extend(_group_by_key(normal))
    else:
        logger.warning("No grouping mode selected. Enable CIRCUITBYKEY or CIRCUITBYPOSITION.")

    return groups


def _remove_from_existing_systems(item):
    element = item.get("element")
    mep_model = getattr(element, "MEPModel", None)
    if not mep_model:
        return
    systems = getattr(mep_model, "ElectricalSystems", None)
    if not systems:
        return

    for system in systems:
        try:
            system.Remove(element.Id)
        except Exception as ex:
            logger.warning("Failed removing {} from system {}: {}".format(element.Id, system.Id, ex))


def _supports_power_circuit(element):
    mep_model = getattr(element, "MEPModel", None)
    if not mep_model:
        return False

    can_assign = getattr(mep_model, "CanAssignToElectricalCircuit", None)
    if isinstance(can_assign, bool):
        return can_assign

    if callable(can_assign):
        try:
            return bool(can_assign())
        except TypeError:
            try:
                return bool(can_assign(DB.Electrical.ElectricalSystemType.PowerCircuit))
            except Exception:
                pass

    connector_manager = getattr(mep_model, "ConnectorManager", None)
    connectors = getattr(connector_manager, "Connectors", None) if connector_manager else None

    if connectors and connectors.Size > 0:
        for connector in connectors:
            try:
                system_type = getattr(connector, "ElectricalSystemType", None)
                if system_type == DB.Electrical.ElectricalSystemType.PowerCircuit:
                    return True
            except Exception:
                pass

            try:
                all_types = getattr(connector, "AllSystemTypes", None)
                if all_types:
                    for sys_type in all_types:
                        if sys_type == DB.Electrical.ElectricalSystemType.PowerCircuit:
                            return True
            except Exception:
                pass

        # If connectors exist but none matched, dump diagnostic information for the element.
        connector_details = []
        try:
            for connector in connectors:
                try:
                    domain = getattr(connector, "Domain", None)
                except Exception:
                    domain = None
                try:
                    ctype = getattr(connector, "ConnectorType", None)
                except Exception:
                    ctype = None

                system_type = getattr(connector, "ElectricalSystemType", None)
                types_list = []
                try:
                    all_types = getattr(connector, "AllSystemTypes", None)
                    if all_types:
                        types_list = [str(t) for t in all_types]
                except Exception:
                    types_list = []

                connector_details.append(
                    "domain={} type={} sys={} all={}".format(domain, ctype, system_type, types_list)
                )
        except Exception:
            connector_details.append("failed to inspect connectors")

        logger.warning(
            "Element {} has connectors but none advertise PowerCircuit. Details: {}".format(
                element.Id.IntegerValue, "; ".join(connector_details)
            )
        )

    return False


def _create_circuit(doc, group):
    valid_members = []
    skipped = []

    for member in group["members"]:
        element = member["element"]
        if _supports_power_circuit(element):
            valid_members.append(member)
        else:
            skipped.append(element.Id)
            logger.debug("Skipping element {} in group {} (no power connector).".format(element.Id, group["key"]))

    if skipped:
        logger.warning(
            "Group {} had {} element(s) without power connectors; they were skipped.".format(group["key"], len(skipped))
        )

    element_ids = []
    for member in valid_members:
        _remove_from_existing_systems(member)
        element_ids.append(member["element"].Id)

    if not element_ids:
        return None

    id_list = List[DB.ElementId](element_ids)
    try:
        system = DB.Electrical.ElectricalSystem.Create(
            doc, id_list, DB.Electrical.ElectricalSystemType.PowerCircuit
        )
    except Exception as ex:
        logger.error("Circuit creation failed for {}: {}".format(group["key"], ex))
        for member in valid_members:
            element = member["element"]
            category = element.Category.Name if element.Category else "No Category"
            family_name = getattr(element, "Name", None)
            can_assign = getattr(getattr(element, "MEPModel", None), "CanAssignToElectricalCircuit", None)
            if callable(can_assign):
                try:
                    can_assign_value = bool(can_assign())
                except TypeError:
                    try:
                        can_assign_value = bool(can_assign(DB.Electrical.ElectricalSystemType.PowerCircuit))
                    except Exception:
                        can_assign_value = "error"
                except Exception:
                    can_assign_value = "error"
            else:
                can_assign_value = can_assign
            try:
                logger.error(
                    "  Element {} | {} | {} | CanAssignToElectricalCircuit: {}".format(
                        element.Id.IntegerValue, category, family_name, can_assign_value
                    )
                )
            except Exception:
                logger.error("  Element {} failed during diagnostics.".format(element.Id.IntegerValue))
        return None

    panel_element = group.get("panel_element")
    if panel_element:
        try:
            system.SelectPanel(panel_element)
        except Exception as ex:
            logger.warning(
                "Unable to select panel {} for circuit {}: {}".format(
                    panel_element.Id, group["key"], ex
                )
            )
    else:
        logger.warning("No panel found for group {}; circuit left unassigned.".format(group["key"]))

    return system


def _parse_rating(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    digits = []
    for ch in str(value):
        if ch.isdigit() or ch == ".":
            digits.append(ch)
    try:
        return float("".join(digits)) if digits else None
    except ValueError:
        return None


def _set_string_param(param, value):
    if param and value and not param.IsReadOnly:
        try:
            param.Set(str(value))
        except Exception as ex:
            logger.warning("Failed to set string parameter {}: {}".format(param.Definition.Name, ex))


def _set_double_param(param, value):
    if param and value is not None and not param.IsReadOnly:
        try:
            param.Set(value)
        except Exception as ex:
            logger.warning("Failed to set numeric parameter {}: {}".format(param.Definition.Name, ex))


def _apply_circuit_data(system, group):
    load_name = group.get("load_name")
    circuit_number = group.get("circuit_number")
    rating_value = _parse_rating(group.get("rating"))

    name_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
    notes_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
    number_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
    rating_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)

    _set_string_param(name_param, load_name)
    _set_string_param(notes_param, group.get("panel_name"))
    _set_string_param(number_param, circuit_number)
    _set_double_param(rating_param, rating_value)


def main():
    doc = revit.doc
    panels = list(get_all_panels(doc))
    panel_lookup = _build_panel_lookup(panels)

    elements = _collect_target_elements(doc)
    info_items = _gather_element_info(elements, panel_lookup)

    if not info_items:
        logger.info("No elements with circuit data were found.")
        return

    groups = _assemble_groups(info_items)
    if not groups:
        logger.info("Grouping produced no circuit batches.")
        return

    created_systems = OrderedDict()

    with revit.Transaction("SuperCircuitV3 - Create Circuits"):
        for group in groups:
            system = _create_circuit(doc, group)
            if not system:
                logger.warning("Circuit creation skipped for {}.".format(group["key"]))
                continue
            doc.Regenerate()
            created_systems[system.Id] = group

    if not created_systems:
        logger.info("No circuits were created.")
        return

    with revit.Transaction("SuperCircuitV3 - Apply Circuit Data"):
        for system_id, group in created_systems.items():
            system = doc.GetElement(system_id)
            if not system:
                logger.warning("Could not locate system {} for data application.".format(system_id))
                continue
            _apply_circuit_data(system, group)

    logger.info("Created {} circuits.".format(len(created_systems)))


if __name__ == "__main__":
    main()
