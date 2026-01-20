from pyrevit import DB

from libGeneral.common import iterate_collection


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


def describe(panel_element):
    """Return a metadata dictionary for the supplied panel element."""
    if not panel_element:
        return None
    return {
        "element": panel_element,
        "distribution_system_ids": gather_distribution_system_ids(panel_element),
    }
