from System.Collections.Generic import List
from pyrevit import revit, DB
from pyrevit.revit.db import query

from libGeneral.common import safe_strip, get_param_value, get_element_location, try_parse_int
from libGeneral import panels as panel_utils


def collect_target_elements(doc, collectors, selection_getter, logger=None):
    selection = selection_getter()
    if selection:
        return list(selection)

    elements = []
    for fn in collectors:
        try:
            elements.extend(list(fn(doc)))
        except Exception as ex:
            if logger:
                logger.warning("Collector {} failed: {}".format(fn.__name__, ex))
    return elements


def build_panel_lookup(panels):
    lookup = {}
    for panel in panels:
        panel_name_param = query.get_param(panel, "Panel Name")
        panel_name = query.get_param_value(panel_name_param) if panel_name_param else None
        if panel_name:
            panel_info = panel_utils.describe(panel) or {"element": panel, "distribution_system_ids": []}
            lookup[panel_name.strip()] = panel_info
    return lookup


def get_power_connectors(element, log_details=True, logger=None):
    connectors_result = []
    if not element:
        return connectors_result

    mep_model = getattr(element, "MEPModel", None)
    if not mep_model:
        return connectors_result

    connector_manager = getattr(mep_model, "ConnectorManager", None)
    connectors = getattr(connector_manager, "Connectors", None) if connector_manager else None

    if not connectors:
        return connectors_result

    all_connectors = []
    try:
        iterator = connectors.ForwardIterator()
        while iterator.MoveNext():
            all_connectors.append(iterator.Current)
    except AttributeError:
        for conn in connectors:
            all_connectors.append(conn)

    for connector in all_connectors:
        try:
            if (
                connector.Domain == DB.Domain.DomainElectrical
                and connector.ElectricalSystemType == DB.Electrical.ElectricalSystemType.PowerCircuit
            ):
                connectors_result.append(connector)
        except Exception:
            continue

    if log_details and not connectors_result and all_connectors and logger:
        connector_details = []
        try:
            for connector in all_connectors:
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

    return connectors_result


def infer_connector_poles(element):
    max_poles = 0
    connectors = get_power_connectors(element, log_details=False)
    for connector in connectors:
        poles = None
        for attr in ("NumberOfPoles", "PolesNumber", "NumberOfPhases"):
            try:
                value = getattr(connector, attr, None)
            except Exception:
                value = None
            if isinstance(value, (int, float)) and value > 0:
                poles = max(poles or 0, int(value))
        if poles and poles > max_poles:
            max_poles = poles
    return max_poles or None


def find_assignable_element(doc, element, logger=None):
    doc = doc or element.Document
    to_process = [element]
    visited = set()

    while to_process:
        current = to_process.pop()
        if not current:
            continue

        current_id = current.Id.IntegerValue if hasattr(current, "Id") else id(current)
        if current_id in visited:
            continue
        visited.add(current_id)

        if get_power_connectors(current, log_details=False):
            return current, "direct"

        if isinstance(current, DB.FamilyInstance):
            try:
                sub_ids = current.GetSubComponentIds()
            except Exception:
                sub_ids = []

            if sub_ids:
                for sid in sub_ids:
                    try:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            to_process.append(sub_elem)
                    except Exception:
                        continue

    if logger:
        logger.debug("No assignable power connector found for element {}.".format(element.Id))
    return None, "no power connector"


def gather_element_info(doc, elements, panel_lookup, logger=None):
    info_items = []
    for element in elements:
        panel_name = get_param_value(element, "CKT_Panel_CEDT")
        circuit_number = get_param_value(element, "CKT_Circuit Number_CEDT")
        rating = get_param_value(element, "CKT_Rating_CED")
        load_name = get_param_value(element, "CKT_Load Name_CEDT")
        circuit_notes = get_param_value(element, "CKT_Schedule Notes_CEDT")
        voltage_value = get_param_value(element, "Voltage_CED")
        poles_value = get_param_value(element, "Number of Poles_CED")
        number_of_poles = try_parse_int(poles_value)

        if not panel_name and not circuit_number:
            continue

        circuit_element, circuit_reason = find_assignable_element(doc, element, logger)
        connector_poles = infer_connector_poles(circuit_element) if circuit_element else None

        panel_info = panel_lookup.get(panel_name) if panel_lookup else None
        if panel_info:
            panel_element = panel_info.get("element")
            panel_distribution_system_ids = panel_info.get("distribution_system_ids") or []
        else:
            panel_element = None
            panel_distribution_system_ids = []

        info_items.append(
            {
                "element": element,
                "circuit_element": circuit_element,
                "assignment_reason": circuit_reason,
                "panel_name": panel_name,
                "panel_element": panel_element,
                "circuit_number": circuit_number,
                "rating": rating,
                "load_name": load_name,
                "location": get_element_location(element),
                "circuit_notes": circuit_notes,
                "number_of_poles": number_of_poles or 1,
                "connector_poles": connector_poles,
                "voltage_ced": voltage_value,
                "panel_distribution_system_ids": list(panel_distribution_system_ids),
            }
        )
    return info_items
