# -*- coding: utf-8 -*-
from pyrevit import DB

from Snippets import revit_helpers


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def wire_connected_unconnected_connectors(wire):
    connected = []
    unconnected = []
    for connector in wire.ConnectorManager.Connectors:
        if connector.IsConnected:
            connected.append(connector)
        else:
            unconnected.append(connector)
    return connected, unconnected


def is_homerun_wire(wire):
    connected, unconnected = wire_connected_unconnected_connectors(wire)
    return len(connected) == 1 and len(unconnected) == 1


def get_element_connector_from_wire_connector(wire_connector):
    for ref in wire_connector.AllRefs:
        owner = getattr(ref, "Owner", None)
        if not owner:
            continue
        if isinstance(owner, DB.Electrical.Wire):
            continue
        if isinstance(owner, DB.Electrical.ElectricalSystem):
            continue
        if hasattr(owner, "Id"):
            return ref
    return None


def get_wire_type_id(wire):
    type_param = wire.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    if type_param:
        return type_param.AsElementId()
    return None


def distance_xy(point1, point2):
    return abs(point1.X - point2.X) + abs(point1.Y - point2.Y)


def safe_direction(from_point, to_point):
    vector = DB.XYZ(to_point.X - from_point.X, to_point.Y - from_point.Y, 0)
    if abs(vector.X) < 0.0001 and abs(vector.Y) < 0.0001:
        return DB.XYZ(1, 0, 0)
    return vector.Normalize()


def perpendicular_xy(direction):
    perp = DB.XYZ(-direction.Y, direction.X, 0)
    if abs(perp.X) < 0.0001 and abs(perp.Y) < 0.0001:
        return DB.XYZ(0, 1, 0)
    return perp.Normalize()


def are_points_coincident(point1, point2, tolerance=0.001):
    return abs(point1.X - point2.X) < tolerance and abs(point1.Y - point2.Y) < tolerance


def collect_selected_electrical_circuits(doc, uidoc, logger=None):
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        if logger:
            logger.warning("No elements selected.")
        return []

    circuits_by_id = {}
    for sel_id in selected_ids:
        element = doc.GetElement(sel_id)
        mep_model = getattr(element, "MEPModel", None)
        if not mep_model:
            continue
        systems = mep_model.GetElectricalSystems() or []
        for system in systems:
            # Keep all electrical system types; caller can filter.
            circuits_by_id[_elid_value(system.Id)] = system
    return list(circuits_by_id.values())


def collect_selected_power_circuits(doc, uidoc, logger=None):
    systems = collect_selected_electrical_circuits(doc, uidoc, logger=logger)
    result = []
    for sys in systems:
        if sys.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit:
            result.append(sys)
    return result


def resolve_wire_type_id(doc, config, logger=None):
    wire_types = list(DB.FilteredElementCollector(doc).OfClass(DB.Electrical.WireType).ToElements())
    if not wire_types:
        if logger:
            logger.error("No wire types were found in this project.")
        return None

    options_by_name = {}
    for wire_type in wire_types:
        name_param = wire_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        name = name_param.AsString() if name_param else None
        if name:
            options_by_name[name] = wire_type.Id

    configured_name = getattr(config, "default_wire_type", None)
    if configured_name and configured_name in options_by_name:
        if logger:
            logger.info("Using saved default wire type: " + configured_name)
        return options_by_name[configured_name]

    fallback_wire_type = wire_types[0]
    fallback_name_param = fallback_wire_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
    fallback_name = fallback_name_param.AsString() if fallback_name_param else "<unnamed wire type>"
    if logger:
        logger.warning("Configured wire type is missing. Using first available wire type: " + fallback_name)
    return fallback_wire_type.Id


def find_previous_connector_for_homerun(start_element_connector, homerun_wire_id=None):
    for ref in start_element_connector.AllRefs:
        owner = getattr(ref, "Owner", None)
        if not owner or not isinstance(owner, DB.Electrical.Wire):
            continue
        if homerun_wire_id and owner.Id == homerun_wire_id:
            continue

        best = None
        best_dist = -1.0
        for wire_connector in owner.ConnectorManager.Connectors:
            elem_connector = get_element_connector_from_wire_connector(wire_connector)
            if not elem_connector:
                continue
            dist = distance_xy(start_element_connector.Origin, elem_connector.Origin)
            if dist > best_dist:
                best_dist = dist
                best = elem_connector
        if best:
            return best
    return None


def build_local_frame(start_point, previous_point=None, fallback_end_point=None):
    if previous_point:
        along = safe_direction(previous_point, start_point)
    elif fallback_end_point:
        along = safe_direction(start_point, fallback_end_point)
    else:
        along = DB.XYZ(1, 0, 0)
    perp = perpendicular_xy(along)
    return along, perp


def project_point_to_frame(point, origin, along, perp):
    rel = DB.XYZ(point.X - origin.X, point.Y - origin.Y, point.Z - origin.Z)
    return rel.DotProduct(along), rel.DotProduct(perp), rel.Z


def point_from_frame(origin, along, perp, along_dist, perp_dist, z_dist):
    return DB.XYZ(
        origin.X + along.X * along_dist + perp.X * perp_dist,
        origin.Y + along.Y * along_dist + perp.Y * perp_dist,
        origin.Z + z_dist
    )


def get_wire_circuit_id(wire):
    # Preferred path for wires: GetMEPSystems()
    try:
        systems = wire.GetMEPSystems()
        if systems:
            for sys in systems:
                if isinstance(sys, DB.Electrical.ElectricalSystem):
                    return sys.Id
    except Exception:
        pass

    mep_system = getattr(wire, "MEPSystem", None)
    if mep_system and isinstance(mep_system, DB.Electrical.ElectricalSystem):
        return mep_system.Id

    for wire_connector in wire.ConnectorManager.Connectors:
        for ref in wire_connector.AllRefs:
            owner = getattr(ref, "Owner", None)
            if owner and isinstance(owner, DB.Electrical.ElectricalSystem):
                return owner.Id
    return None


def collect_active_view_wires_by_circuit(doc, view_id):
    wire_map = {}
    wires = DB.FilteredElementCollector(doc, view_id).OfClass(DB.Electrical.Wire).ToElements()
    for wire in wires:
        circuit_id = get_wire_circuit_id(wire)
        if not circuit_id:
            continue
        key = _elid_value(circuit_id)
        if key not in wire_map:
            wire_map[key] = []
        wire_map[key].append(wire)
    return wire_map


def delete_element_ids(doc, element_ids):
    deleted = 0
    for eid in element_ids:
        try:
            doc.Delete(eid)
            deleted += 1
        except Exception:
            pass
    return deleted

