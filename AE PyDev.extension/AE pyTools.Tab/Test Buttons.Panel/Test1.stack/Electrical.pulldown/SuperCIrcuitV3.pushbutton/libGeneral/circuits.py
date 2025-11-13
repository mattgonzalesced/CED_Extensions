from System.Collections.Generic import List
from pyrevit import DB

from libGeneral.common import try_parse_int
from libGeneral.data import get_power_connectors, find_assignable_element
from libGeneral.panels import gather_distribution_system_ids

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

    connector_poles = group.get("connector_poles")
    desired_poles = try_parse_int(group.get("number_of_poles")) or connector_poles

    panel_element = group.get("panel_element")
    if panel_element:
        panel_distribution_ids = group.get("panel_distribution_system_ids")
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
                "Distribution system alignment failed for circuit {}.".format(group.get("key"))
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
            except Exception as ex:
                panel_name = group.get("panel_name") or getattr(panel_element, "Name", None) or panel_element.Id
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
            panel_name = getattr(panel_element, "Name", None)
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
    else:
        if logger:
            logger.warning("No panel found for group {}; circuit left unassigned.".format(group.get("key")))

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

    connector_poles = group.get("connector_poles")
    desired_poles = try_parse_int(group.get("number_of_poles")) or connector_poles

    panel_element = group.get("panel_element")
    if not panel_element:
        if logger:
            logger.warning("No panel found for group {}; circuit left unassigned.".format(group.get("key")))
        return False

    aligned = _align_distribution_system(
        doc,
        system,
        panel_element,
        desired_poles,
        group.get("panel_distribution_system_ids"),
        logger,
        group.get("key"),
    )
    if not aligned and logger:
        logger.debug(
            "Distribution system alignment failed for circuit {}.".format(group.get("key"))
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
            return True
        except Exception as ex:
            panel_name = group.get("panel_name") or getattr(panel_element, "Name", None) or panel_element.Id
            if logger:
                logger.warning(
                    "Unable to select panel {} (Id {}) for circuit {}: {}. Verify distribution system and pole settings.".format(
                        panel_name, panel_element.Id, group.get("key"), ex
                    )
                )
    else:
        panel_phases = None
        panel_name = getattr(panel_element, "Name", None)
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

