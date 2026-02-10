# -*- coding: utf-8 -*-
"""
Consolidated circuits library - combines all libGeneral modules:
- common: Utility functions
- data: Element collection and info gathering
- panels: Panel distribution system handling
- grouping: Element grouping strategies
- circuits: Circuit creation and configuration
- transactions: Transaction management
"""

from collections import defaultdict, OrderedDict
import math

from System.Collections.Generic import List
from pyrevit import revit, DB, forms
from pyrevit.revit.db import query

_OVERFLOW_PANEL_CACHE = {}

# =============================================================================
# COMMON UTILITIES
# =============================================================================

try:
    basestring
except NameError:
    basestring = str

try:
    long
except NameError:
    long = int


def safe_strip(value):
    return value.strip() if isinstance(value, basestring) else value


def get_param_value(element, param_name):
    param = query.get_param(element, param_name)
    return safe_strip(query.get_param_value(param)) if param else None


def get_element_location(element):
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


def try_parse_int(value):
    if value is None:
        return None
    if isinstance(value, (int, long)):
        return int(value)
    try:
        text = str(value).strip()
        digits = []
        for ch in text:
            if ch.isdigit():
                digits.append(ch)
            elif digits:
                break
        if not digits:
            return None
        return int("".join(digits))
    except Exception:
        return None


def try_parse_float(value):
    if value is None:
        return None
    if isinstance(value, (int, long, float)):
        return float(value)
    try:
        text = str(value).strip()
        cleaned = []
        decimal_found = False
        for ch in text:
            if ch.isdigit():
                cleaned.append(ch)
            elif ch in (".", ","):
                if not decimal_found:
                    cleaned.append(".")
                    decimal_found = True
            elif cleaned:
                break
        if not cleaned:
            return None
        return float("".join(cleaned))
    except Exception:
        return None


def iterate_collection(collection):
    if not collection:
        return
    try:
        iterator = collection.ForwardIterator()
        while iterator.MoveNext():
            yield iterator.Current
        return
    except AttributeError:
        pass
    try:
        for item in collection:
            yield item
    except TypeError:
        if collection:
            yield collection


def extract_parent_location(element):
    """Extract parent location coordinates from Element_Linker parameter."""
    if not element:
        return None

    try:
        linker_param = element.LookupParameter("Element_Linker")
    except Exception:
        return None

    if not linker_param or not linker_param.HasValue:
        return None

    try:
        linker_text = linker_param.AsString()
    except Exception:
        return None

    if not linker_text:
        return None

    # Look for "Parent_location: X,Y,Z" pattern (case-insensitive)
    for line in linker_text.split('\n'):
        line = line.strip()
        if line.lower().startswith("parent_location:"):
            # Extract "44.072917,12.598958,0.000000"
            coords = line.split(":", 1)[1].strip()
            return coords

    return None


def extract_parent_element_id(element):
    """Extract parent element id from Element_Linker parameter."""
    if not element:
        return None

    try:
        linker_param = element.LookupParameter("Element_Linker")
    except Exception:
        return None

    if not linker_param or not linker_param.HasValue:
        return None

    try:
        linker_text = linker_param.AsString()
    except Exception:
        return None

    if not linker_text:
        return None

    for line in linker_text.split('\n'):
        line = line.strip()
        if line.lower().startswith("parent elementid:"):
            value = line.split(":", 1)[1].strip()
            try:
                return int(value)
            except Exception:
                return None
    return None


# =============================================================================
# PANEL UTILITIES
# =============================================================================

def _add_element_id(target_ids, seen_ids, elem_id):
    if not elem_id:
        return
    if isinstance(elem_id, DB.ElementId):
        numeric = elem_id.IntegerValue
        element_id = elem_id
    else:
        try:
            numeric = int(elem_id)
        except Exception:
            return
        element_id = DB.ElementId(numeric)
    if numeric <= 0 or numeric in seen_ids:
        return
    seen_ids.add(numeric)
    target_ids.append(numeric)


def _get_connector_manager(owner):
    if not owner:
        return None
    connector_manager = getattr(owner, "ConnectorManager", None)
    if connector_manager:
        return connector_manager
    mep_model = getattr(owner, "MEPModel", None)
    if mep_model:
        return getattr(mep_model, "ConnectorManager", None)
    return None


def _distribution_id_from_connector(connector):
    if not connector:
        return None
    try:
        dist_elem = getattr(connector, "DistributionSystem", None)
        if dist_elem and getattr(dist_elem, "Id", None):
            return dist_elem.Id.IntegerValue
    except Exception:
        pass
    try:
        ds_elem = getattr(connector, "MEPDistributionSystem", None)
        if ds_elem and getattr(ds_elem, "Id", None):
            return ds_elem.Id.IntegerValue
    except Exception:
        pass
    try:
        ds_id = connector.GetMEPDistributionSystemId()
        if ds_id and ds_id.IntegerValue > 0:
            return ds_id.IntegerValue
    except Exception:
        pass
    return None


def _collect_distribution_ids_from_owner(owner):
    ids = set()
    if not owner:
        return ids

    candidate_bips = (
        "RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS",
        "RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM",
    )
    for bip_name in candidate_bips:
        bip = getattr(DB.BuiltInParameter, bip_name, None)
        if not bip:
            continue
        try:
            param = owner.get_Parameter(bip)
        except Exception:
            param = None
        if param and param.HasValue:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue > 0:
                ids.add(elem_id.IntegerValue)

    try:
        lookup_param = owner.LookupParameter("Distribution System")
    except Exception:
        lookup_param = None
    if lookup_param and lookup_param.HasValue:
        elem_id = lookup_param.AsElementId()
        if elem_id and elem_id.IntegerValue > 0:
            ids.add(elem_id.IntegerValue)

    connector_manager = _get_connector_manager(owner)
    if connector_manager:
        connectors = getattr(connector_manager, "Connectors", None)
        for connector in iterate_collection(connectors):
            ds_id = _distribution_id_from_connector(connector)
            if ds_id:
                ids.add(ds_id)

    return ids


def gather_distribution_system_ids(panel_element):
    """Return a list of integer ElementIds for distribution systems related to the panel."""
    if not panel_element:
        return []

    doc = getattr(panel_element, "Document", None)
    owners = [panel_element]
    symbol = getattr(panel_element, "Symbol", None)
    if symbol and symbol not in owners:
        owners.append(symbol)
    if doc:
        try:
            type_id = panel_element.GetTypeId()
        except Exception:
            type_id = None
        if type_id:
            try:
                type_elem = doc.GetElement(type_id)
            except Exception:
                type_elem = None
            if type_elem and type_elem not in owners:
                owners.append(type_elem)

    collected_ids = []
    seen_ids = set()
    for owner in owners:
        for ds_id in _collect_distribution_ids_from_owner(owner):
            _add_element_id(collected_ids, seen_ids, ds_id)

    return collected_ids


def describe_panel(panel_element):
    """Return a metadata dictionary for the supplied panel element."""
    if not panel_element:
        return None
    return {
        "element": panel_element,
        "distribution_system_ids": gather_distribution_system_ids(panel_element),
    }


# =============================================================================
# DATA COLLECTION
# =============================================================================

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
            panel_info = describe_panel(panel) or {"element": panel, "distribution_system_ids": []}
            lookup[panel_name.strip()] = panel_info
    return lookup


def _panel_display_name(panel_element, panel_name=None):
    name = panel_name or get_param_value(panel_element, "Panel Name") or getattr(panel_element, "Name", None)
    if name:
        return str(name).strip()
    return "Unnamed Panel"


def _collect_panel_option_map(doc, exclude_ids=None):
    exclude_ids = set(exclude_ids or [])
    options = {}
    collector = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_ElectricalEquipment
    ).WhereElementIsNotElementType()
    for panel in collector:
        panel_id = getattr(getattr(panel, "Id", None), "IntegerValue", None)
        if panel_id is not None and panel_id in exclude_ids:
            continue
        display = "{} (Id {})".format(_panel_display_name(panel), panel_id if panel_id is not None else "NA")
        if display in options:
            display = "{} [{}]".format(display, getattr(panel, "Name", "Unknown"))
        options[display] = panel
    return options


def _prompt_overflow_panel(doc, base_panel_name, exclude_ids=None, logger=None):
    options = _collect_panel_option_map(doc, exclude_ids)
    if not options:
        if logger:
            logger.warning("No panels available for overflow selection.")
        return None
    panel_name = base_panel_name or "Selected panel"
    prompt_msg = (
        "Panel '{}' is full. Select a similar panel for overflow circuits "
        "(avoid unrelated panels).".format(panel_name)
    )
    choice = forms.SelectFromList.show(
        sorted(options.keys(), key=lambda value: value.lower()),
        title="Select Overflow Panel (Full: {})".format(panel_name),
        prompt=prompt_msg,
        multiselect=False,
    )
    if not choice:
        if logger:
            logger.warning("Overflow panel selection cancelled.")
        return None
    return options.get(choice)


def _get_overflow_panel(doc, base_panel_name, exclude_ids=None, logger=None):
    cache_key = (base_panel_name or "").strip().upper() or "UNKNOWN"
    if cache_key in _OVERFLOW_PANEL_CACHE:
        return _OVERFLOW_PANEL_CACHE[cache_key]
    selected = _prompt_overflow_panel(doc, base_panel_name, exclude_ids, logger)
    if selected:
        _OVERFLOW_PANEL_CACHE[cache_key] = selected
    return selected


def _panel_candidates_from_group(group):
    candidates = []
    for choice in group.get("panel_choices") or []:
        panel_elem = choice.get("element")
        if not panel_elem:
            continue
        candidates.append(
            {
                "name": choice.get("name") or _panel_display_name(panel_elem),
                "element": panel_elem,
                "distribution_system_ids": list(choice.get("distribution_system_ids") or []),
            }
        )

    if not candidates:
        panel_element = group.get("panel_element")
        if panel_element:
            candidates.append(
                {
                    "name": group.get("panel_name") or _panel_display_name(panel_element),
                    "element": panel_element,
                    "distribution_system_ids": list(group.get("panel_distribution_system_ids") or []),
                }
            )

    unique = []
    seen = set()
    for candidate in candidates:
        panel_element = candidate.get("element")
        panel_id = getattr(getattr(panel_element, "Id", None), "IntegerValue", None)
        if panel_id is not None and panel_id in seen:
            continue
        if panel_id is not None:
            seen.add(panel_id)
        unique.append(candidate)
    return unique


def _attempt_assign_to_panel(doc, system, group, candidate, desired_poles, logger=None):
    panel_element = candidate.get("element")
    if not panel_element:
        return False

    panel_name = candidate.get("name") or _panel_display_name(panel_element)
    panel_distribution_ids = candidate.get("distribution_system_ids") or group.get("panel_distribution_system_ids")
    aligned = _align_distribution_system(
        doc,
        system,
        panel_element,
        desired_poles,
        panel_distribution_ids,
        logger,
        group.get("key"),
    )
    if not aligned and logger:
        logger.debug(
            "Distribution system alignment failed for circuit {} (panel {}).".format(
                group.get("key"), panel_name
            )
        )

    can_assign = True
    can_assign_method = getattr(system, "CanAssignToPanel", None)
    if callable(can_assign_method):
        try:
            can_assign = bool(can_assign_method(panel_element))
        except Exception as ex:
            if logger:
                logger.debug(
                    "CanAssignToPanel check failed for panel {} and circuit {}: {}".format(
                        panel_element.Id, group.get("key"), ex
                    )
                )

    if can_assign:
        try:
            system.SelectPanel(panel_element)
            group["panel_name"] = panel_name
            group["panel_element"] = panel_element
            group["panel_distribution_system_ids"] = list(panel_distribution_ids or [])
            return True
        except Exception as ex:
            system_ds_id, system_ds_name = _describe_distribution(system, doc)
            panel_ds_id = None
            panel_ds_name = None
            for attr in ("RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS", "RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM"):
                bip = getattr(DB.BuiltInParameter, attr, None)
                if not bip:
                    continue
                try:
                    param = panel_element.get_Parameter(bip)
                except Exception:
                    param = None
                if param and param.HasValue:
                    try:
                        panel_ds_id = param.AsElementId()
                    except Exception:
                        panel_ds_id = None
                    if panel_ds_id and panel_ds_id.IntegerValue > 0:
                        try:
                            ds_elem = doc.GetElement(panel_ds_id)
                            panel_ds_name = getattr(ds_elem, "Name", None)
                        except Exception:
                            panel_ds_name = None
                        break
            if logger:
                logger.warning(
                    "Unable to select panel {} (Id {}) for circuit {}: {}. Requested poles {} system poles {} | system DS {} (Id {}) vs panel DS {} (Id {}).".format(
                        panel_name,
                        panel_element.Id,
                        group.get("key"),
                        ex,
                        desired_poles or "unknown",
                        getattr(system, "PolesNumber", None) or "unknown",
                        system_ds_name or "unknown",
                        system_ds_id.IntegerValue if system_ds_id else "None",
                        panel_ds_name or "unknown",
                        panel_ds_id.IntegerValue if panel_ds_id else "None",
                    )
                )
    else:
        panel_phases = None
        try:
            phases_param = panel_element.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NUMPHASES_PARAM)
            if phases_param and phases_param.HasValue:
                panel_phases = phases_param.AsInteger()
        except Exception:
            panel_phases = None
        if logger:
            logger.warning(
                "Panel {} (Id {}) is not compatible with circuit {} (requested poles: {}, system poles: {}, panel phases: {}). Leaving circuit unassigned.".format(
                    panel_name or "Unnamed",
                    panel_element.Id,
                    group.get("key"),
                    desired_poles or "unknown",
                    getattr(system, "PolesNumber", None) or "unknown",
                    panel_phases or "unknown",
                )
            )
    return False


def _assign_panel_with_fallback(doc, system, group, logger=None):
    connector_poles = group.get("connector_poles")
    desired_poles = try_parse_int(group.get("number_of_poles")) or connector_poles
    candidates = _panel_candidates_from_group(group)
    if not candidates:
        if logger:
            logger.warning("No panel found for group {}; circuit left unassigned.".format(group.get("key")))
        return False

    for candidate in candidates:
        if _attempt_assign_to_panel(doc, system, group, candidate, desired_poles, logger):
            return True

    if len(candidates) <= 1:
        base_panel_name = group.get("panel_name") or candidates[0].get("name")
        exclude_ids = [getattr(getattr(candidates[0].get("element"), "Id", None), "IntegerValue", None)]
        overflow_panel = _get_overflow_panel(doc, base_panel_name, exclude_ids, logger)
        if overflow_panel:
            panel_info = describe_panel(overflow_panel) or {"element": overflow_panel, "distribution_system_ids": []}
            overflow_candidate = {
                "name": _panel_display_name(overflow_panel),
                "element": overflow_panel,
                "distribution_system_ids": list(panel_info.get("distribution_system_ids") or []),
            }
            if _attempt_assign_to_panel(doc, system, group, overflow_candidate, desired_poles, logger):
                return True

    if logger:
        logger.warning("Unable to assign panel for circuit {}.".format(group.get("key")))
    return False


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

        parent_location = extract_parent_location(element)
        parent_element_id = extract_parent_element_id(element)

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
                "parent_location": parent_location,
                "parent_element_id": parent_element_id,
            }
        )
    return info_items


# =============================================================================
# GROUPING
# =============================================================================

def make_group(key, members, group_type=None, parent_key=None):
    sample = members[0]
    rating = sample.get("rating")
    load_name = sample.get("load_name")
    circuit_number = sample.get("circuit_number")
    circuit_notes = sample.get("circuit_notes")
    number_of_poles = sample.get("number_of_poles")
    panel_choices = None
    for item in members:
        if item.get("panel_choices"):
            panel_choices = item.get("panel_choices")
            break

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
        "panel_choices": panel_choices,
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


def create_circuitbyparent_groups(items, keyword_suffix=""):
    """Group items by panel and parent_element_id from Element_Linker parameter."""
    buckets = defaultdict(list)

    for item in items:
        panel_name = item.get("panel_name") or "NO_PANEL"
        parent_element_id = item.get("parent_element_id")
        if parent_element_id is None:
            element = item.get("element")
            elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
            parent_element_id = "NO_PARENT_{}".format(elem_id if elem_id is not None else id(item))

        # Group by panel + parent_element_id
        key = (panel_name, parent_element_id)
        buckets[key].append(item)

    groups = []
    counters = defaultdict(int)

    for (panel_name, parent_element_id), members in sorted(buckets.items()):
        counters[panel_name] += 1
        # Key format: PANEL1BYPARENT1 or PANEL1SECONDBYPARENT1
        base_name = "BYPARENT" if not keyword_suffix else keyword_suffix
        group_key = "{}{}{}".format(panel_name, base_name, counters[panel_name])
        group_type = "circuitbyparent" if not keyword_suffix else "secondcircuitbyparent"
        groups.append(make_group(group_key, members, group_type=group_type))

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


def assemble_groups(items, client_helpers, logger):
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
    circuitbyparent = []
    secondcircuitbyparent = []

    if client_helpers and hasattr(client_helpers, "classify_items"):
        try:
            dedicated, nongrouped, tvtruss, normal = client_helpers.classify_items(working_items)
        except Exception as ex:
            if logger:
                logger.warning("Client classify_items failed: {}".format(ex))

    # Extract BYPARENT and SECONDBYPARENT items from normal bucket
    filtered_normal = []
    for item in normal:
        circuit_upper = (item.get("circuit_number") or "").strip().upper()
        if circuit_upper in ("CIRCUITBYPARENT", "BYPARENT"):
            circuitbyparent.append(item)
        elif circuit_upper in ("SECONDCIRCUITBYPARENT", "SECONDBYPARENT"):
            secondcircuitbyparent.append(item)
        else:
            filtered_normal.append(item)
    normal = filtered_normal

    if dedicated:
        groups.extend(create_dedicated_groups(dedicated))

    if nongrouped:
        groups.extend(create_nongroupedblock_groups(nongrouped))

    if circuitbyparent:
        groups.extend(create_circuitbyparent_groups(circuitbyparent))

    if secondcircuitbyparent:
        groups.extend(create_circuitbyparent_groups(secondcircuitbyparent, "SECONDBYPARENT"))

    # tvtruss is now handled by client position rules, this is just a fallback
    if tvtruss:
        if logger:
            logger.warning("{} items classified as tvtruss but not handled by position rules".format(len(tvtruss)))
        groups.extend(group_by_position(tvtruss, 3, logger))

    if normal:
        groups.extend(group_by_key(normal, client_helpers, logger))

    groups = split_groups_by_poles(groups, logger)
    return groups


# =============================================================================
# CIRCUIT CREATION AND CONFIGURATION
# =============================================================================

try:
    THREE_PHASE_ENUM = DB.Electrical.ElectricalPhase.ThreePhase
except Exception:
    THREE_PHASE_ENUM = None


def _convert_to_element_ids(id_values):
    element_ids = []
    seen = set()
    for value in id_values or []:
        if isinstance(value, DB.ElementId):
            numeric = value.IntegerValue
            element_id = value
        else:
            try:
                numeric = int(value)
            except Exception:
                continue
            element_id = DB.ElementId(numeric)
        if numeric <= 0 or numeric in seen:
            continue
        seen.add(numeric)
        element_ids.append(element_id)
    return element_ids


def _is_three_phase_required(desired_poles):
    if not desired_poles:
        return False
    try:
        poles_value = int(desired_poles)
    except Exception:
        return False
    return poles_value >= 3


def _is_three_phase_distribution(dist_elem):
    if not dist_elem or THREE_PHASE_ENUM is None:
        return False
    try:
        return getattr(dist_elem, "ElectricalPhase", None) == THREE_PHASE_ENUM
    except Exception:
        return False


def _prioritize_distribution_elements(elements, desired_poles):
    elements = list(elements or [])
    if not elements or not _is_three_phase_required(desired_poles):
        return elements
    primary = [elem for elem in elements if _is_three_phase_distribution(elem)]
    secondary = [elem for elem in elements if elem not in primary]
    return primary + secondary


def _apply_distribution_element(system, dist_elem, logger=None):
    setter = getattr(system, "SetDistributionSystem", None)
    if not callable(setter) or not dist_elem:
        return False
    try:
        setter(dist_elem)
        if logger:
            logger.debug(
                "SetDistributionSystem applied {} to circuit {}.".format(
                    dist_elem.Id.IntegerValue, system.Id.IntegerValue
                )
            )
        return True
    except Exception as ex:
        if logger:
            logger.debug(
                "SetDistributionSystem failed for circuit {} with DS {}: {}".format(
                    system.Id.IntegerValue, dist_elem.Id.IntegerValue, ex
                )
            )
        return False


def _set_distribution_parameter(system, element_id, logger=None):
    if not element_id:
        return False
    if not isinstance(element_id, DB.ElementId):
        try:
            element_id = DB.ElementId(int(element_id))
        except Exception:
            return False

    candidate_params = []
    for name in ("RBS_ELEC_DISTRIBUTION_SYSTEM", "RBS_ELEC_CIRCUIT_DISTRIBUTION_SYSTEM"):
        bip = getattr(DB.BuiltInParameter, name, None)
        if bip:
            try:
                candidate_params.append(system.get_Parameter(bip))
            except Exception:
                candidate_params.append(None)
    try:
        candidate_params.append(system.LookupParameter("Distribution System"))
    except Exception:
        candidate_params.append(None)

    for param in candidate_params:
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(element_id)
            if logger:
                param_name = param.Definition.Name if param.Definition else "Unknown"
                logger.debug(
                    "Set {} for circuit {} to distribution {}.".format(
                        param_name, system.Id.IntegerValue, element_id.IntegerValue
                    )
                )
            return True
        except Exception as ex:
            if logger:
                param_name = param.Definition.Name if param.Definition else "Unknown"
                logger.debug(
                    "Failed to set {} for circuit {}: {}".format(
                        param_name, system.Id.IntegerValue, ex
                    )
                )
    return False


def _describe_distribution(owner, doc):
    ds_id = None
    ds_name = None
    distrib_obj = getattr(owner, "DistributionSystem", None)
    if distrib_obj:
        try:
            ds_id = distrib_obj.Id
        except Exception:
            ds_id = None
        try:
            ds_name = getattr(distrib_obj, "Name", None)
        except Exception:
            ds_name = None
    if not ds_id:
        bip_candidates = []
        for name in ("RBS_ELEC_DISTRIBUTION_SYSTEM", "RBS_ELEC_CIRCUIT_DISTRIBUTION_SYSTEM"):
            bip = getattr(DB.BuiltInParameter, name, None)
            if bip:
                bip_candidates.append(bip)
        for bip in bip_candidates:
            try:
                param = owner.get_Parameter(bip)
            except Exception:
                param = None
            if param and param.HasValue:
                try:
                    ds_id = param.AsElementId()
                except Exception:
                    ds_id = None
                if ds_id and ds_id.IntegerValue > 0:
                    break
    if not ds_id:
        try:
            param = owner.LookupParameter("Distribution System")
        except Exception:
            param = None
        if param and param.HasValue:
            try:
                ds_id = param.AsElementId()
            except Exception:
                ds_id = None
    if ds_id and doc:
        try:
            ds_elem = doc.GetElement(ds_id)
            ds_name = getattr(ds_elem, "Name", None)
        except Exception:
            ds_name = None
    return ds_id, ds_name


def _align_distribution_system(
    doc,
    system,
    panel_element,
    desired_poles=None,
    extra_distribution_ids=None,
    logger=None,
    group_key=None,
):
    if not panel_element:
        return False

    doc = doc or panel_element.Document

    candidate_ids = _convert_to_element_ids(extra_distribution_ids)
    supplemental_ids = gather_distribution_system_ids(panel_element)
    if supplemental_ids:
        supplemental = _convert_to_element_ids(supplemental_ids)
        existing = {elem_id.IntegerValue for elem_id in candidate_ids}
        for elem_id in supplemental:
            if elem_id.IntegerValue not in existing:
                candidate_ids.append(elem_id)
                existing.add(elem_id.IntegerValue)

    available_ids = []
    available_elems = []
    get_available = getattr(system, "GetAvailableDistributionSystems", None)
    if callable(get_available):
        try:
            available_elems = [ds for ds in get_available() if ds]
            available_ids = [
                ds.Id.IntegerValue
                for ds in available_elems
                if getattr(ds, "Id", None) and ds.Id.IntegerValue > 0
            ]
        except Exception as ex:
            if logger:
                logger.debug(
                    "GetAvailableDistributionSystems failed for circuit {}: {}".format(
                        system.Id.IntegerValue, ex
                    )
                )
    candidate_elems = []
    if doc and candidate_ids:
        for elem_id in candidate_ids:
            try:
                candidate = doc.GetElement(elem_id)
            except Exception:
                candidate = None
            if candidate:
                candidate_elems.append(candidate)

    if logger:
        logger.debug(
            "Circuit {} group {} distribution candidates {} available {}.".format(
                system.Id.IntegerValue,
                group_key or "unknown",
                [elem.Id.IntegerValue for elem in candidate_elems],
                available_ids,
            )
        )

    for dist_elem in _prioritize_distribution_elements(candidate_elems, desired_poles):
        if _apply_distribution_element(system, dist_elem, logger):
            return True

    for elem_id in candidate_ids:
        if _set_distribution_parameter(system, elem_id, logger):
            return True

    for dist_elem in _prioritize_distribution_elements(available_elems, desired_poles):
        if _apply_distribution_element(system, dist_elem, logger):
            return True

    if logger:
        logger.debug(
            "Unable to align distribution system for circuit {} (panel {}, group {}).".format(
                system.Id.IntegerValue, panel_element.Id, group_key or "unknown"
            )
        )
    return False


def _remove_from_existing_systems(item, logger=None):
    element = item.get("circuit_element") or item.get("element")
    if not element:
        return

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
            if logger:
                logger.warning("Failed removing {} from system {}: {}".format(element.Id, system.Id, ex))


def create_circuit(doc, group, logger=None):
    candidates = []
    skipped_ids = set()

    for member in group["members"]:
        circuit_element = member.get("circuit_element")
        if not circuit_element:
            host = member["element"]
            skipped_ids.add(host.Id.IntegerValue)
            if logger:
                logger.debug(
                    "Skipping element {} in group {} ({}).".format(
                        host.Id, group["key"], member.get("assignment_reason")
                    )
                )
            continue

        if not get_power_connectors(circuit_element, log_details=True, logger=logger):
            skipped_ids.add(circuit_element.Id.IntegerValue)
            if logger:
                logger.warning(
                    "Element {} in group {} does not expose a PowerCircuit connector.".format(
                        circuit_element.Id, group["key"]
                    )
                )
            continue

        candidates.append(member)

    if skipped_ids and logger:
        logger.warning(
            "Group {} had {} element(s) without power connectors; they were skipped.".format(
                group["key"], len(skipped_ids)
            )
        )

    if not candidates:
        return None

    for member in candidates:
        _remove_from_existing_systems(member, logger)

    element_ids = List[DB.ElementId]()
    for member in candidates:
        circuit_element = member.get("circuit_element")
        element_ids.Add(circuit_element.Id)

    if element_ids.Count == 0:
        return None

    try:
        system = DB.Electrical.ElectricalSystem.Create(
            doc, element_ids, DB.Electrical.ElectricalSystemType.PowerCircuit
        )
    except Exception as ex:
        if logger:
            logger.error("Circuit creation failed for {}: {}".format(group.get("key"), ex))
        return None

    _assign_panel_with_fallback(doc, system, group, logger)

    return system


def _set_string_param(param, value, logger=None):
    if param and value is not None and not param.IsReadOnly:
        try:
            param.Set(str(value))
        except Exception as ex:
            if logger:
                logger.warning("Failed to set string parameter {}: {}".format(param.Definition.Name, ex))


def _set_double_param(param, value, logger=None):
    if param and value is not None and not param.IsReadOnly:
        try:
            param.Set(value)
        except Exception as ex:
            if logger:
                logger.warning("Failed to set numeric parameter {}: {}".format(param.Definition.Name, ex))


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


def apply_circuit_data(system, group, logger=None):
    load_name = group.get("load_name")
    rating_value = _parse_rating(group.get("rating"))
    circuit_notes = group.get("circuit_notes")

    name_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
    notes_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
    rating_param = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)

    _set_string_param(name_param, load_name, logger)
    _set_string_param(notes_param, circuit_notes, logger)
    _set_double_param(rating_param, rating_value, logger)


def configure_circuit(doc, system, group, logger=None):
    if not system or not group:
        return False

    doc = doc or getattr(system, "Document", None)
    if not doc:
        return False
    return _assign_panel_with_fallback(doc, system, group, logger)


# =============================================================================
# TRANSACTION MANAGEMENT
# =============================================================================

DEFAULT_CREATE_LABEL = "SuperCircuitV4 - Create Circuits"
DEFAULT_APPLY_LABEL = "SuperCircuitV4 - Apply Circuit Data"


def run_creation(doc, groups, create_func, logger, transaction_label=None):
    created_systems = OrderedDict()
    if not groups:
        return created_systems

    label = transaction_label or DEFAULT_CREATE_LABEL
    with revit.Transaction(label):
        for group in groups:
            system = create_func(doc, group)
            if not system:
                if logger:
                    logger.warning("Circuit creation skipped for {}.".format(group.get("key")))
                continue
            created_systems[system.Id] = group

    return created_systems


def run_apply_data(doc, created_systems, apply_func, logger, transaction_label=None):
    if not created_systems:
        if logger:
            logger.info("No circuits were created.")
        return

    label = transaction_label or DEFAULT_APPLY_LABEL
    with revit.Transaction(label):
        for system_id, group in created_systems.items():
            system = doc.GetElement(system_id)
            if not system:
                if logger:
                    logger.warning("Could not locate system {} for data application.".format(system_id))
                continue
            apply_func(system, group)
