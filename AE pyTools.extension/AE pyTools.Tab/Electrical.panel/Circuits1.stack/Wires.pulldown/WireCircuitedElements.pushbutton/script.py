# -*- coding: utf-8 -*-
__title__ = "Wire Circuited Elements"

from Snippets.wireutils import (
    wire_connected_unconnected_connectors,
    is_homerun_wire,
    get_element_connector_from_wire_connector,
    get_wire_type_id,
    are_points_coincident,
    collect_selected_electrical_circuits,
    resolve_wire_type_id,
    find_previous_connector_for_homerun,
    build_local_frame,
    collect_active_view_wires_by_circuit,
    delete_element_ids,
)
from pyrevit import script, DB, forms

logger = script.get_logger()
HOME_RUN_LENGTH = 4.0
GEOM_TOL = 1e-6
WIRING_TYPE_NAMES = {
    "Arc": DB.Electrical.WiringType.Arc,
    "Chamfer": DB.Electrical.WiringType.Chamfer
}


def wiring_type_from_config(config_key, default_name, config):
    configured = getattr(config, config_key, default_name)
    return WIRING_TYPE_NAMES.get(configured, WIRING_TYPE_NAMES[default_name])


def homerun_length_from_config(config):
    value = getattr(config, "homerun_length", HOME_RUN_LENGTH)
    try:
        parsed = float(value)
        if parsed > 0:
            return parsed
    except Exception:
        pass
    return HOME_RUN_LENGTH


def xyz_text(point):
    return "({:.3f}, {:.3f}, {:.3f})".format(point.X, point.Y, point.Z)


def remove_existing_wires_in_active_view(doc, view_id, circuits):
    circuits_by_key = {c.Id.IntegerValue: c for c in circuits}
    wire_map = collect_active_view_wires_by_circuit(doc, view_id)
    delete_ids = []
    by_circuit_count = {}

    for key in circuits_by_key.keys():
        wires = wire_map.get(key, [])
        if not wires:
            continue
        by_circuit_count[key] = len(wires)
        for wire in wires:
            delete_ids.append(wire.Id)

    deleted_count = delete_element_ids(doc, delete_ids) if delete_ids else 0
    return deleted_count, by_circuit_count


def resolve_target_system_type(circuits):
    types_present = {}
    for c in circuits:
        type_name = str(c.SystemType)
        if type_name not in types_present:
            types_present[type_name] = c.SystemType

    if not types_present:
        return None
    if len(types_present) == 1:
        return list(types_present.values())[0]

    selected = forms.SelectFromList.show(
        sorted(types_present.keys()),
        title="Select Electrical System Type to Wire",
        button_name="Use Type"
    )
    if not selected:
        return None
    return types_present[selected]


class WireGenerator(object):
    def __init__(
        self,
        doc,
        view,
        wire_type_id,
        branch_wiring_type,
        homerun_wiring_type,
        home_run_length=HOME_RUN_LENGTH
    ):
        self.doc = doc
        self.view = view
        self.wire_type_id = wire_type_id
        self.branch_wiring_type = branch_wiring_type
        self.homerun_wiring_type = homerun_wiring_type
        self.home_run_length = home_run_length

    def _clamp(self, value, low, high):
        return max(low, min(high, value))

    def _vector_length(self, vec):
        try:
            return vec.GetLength()
        except Exception:
            return 0.0

    def _build_homerun_end(self, start, native_end, previous_connector):
        native_vec = native_end.Subtract(start)
        native_len = self._vector_length(native_vec)

        previous = previous_connector.Origin if previous_connector else None
        along, _ = build_local_frame(start, previous_point=previous, fallback_end_point=native_end)

        if native_len > GEOM_TOL:
            direction = native_vec.Normalize()
        else:
            direction = along

        target_len = native_len if native_len > GEOM_TOL else self.home_run_length
        if target_len > self.home_run_length:
            target_len = self.home_run_length

        end = start.Add(direction.Multiply(target_len))
        return end, direction, target_len, native_len

    def _build_homerun_vertex(self, start, end, direction, native_vertex, previous_connector):
        segment = end.Subtract(start)
        seg_len = self._vector_length(segment)
        if seg_len <= GEOM_TOL:
            seg_len = self.home_run_length
            end = start.Add(direction.Multiply(seg_len))

        previous = previous_connector.Origin if previous_connector else None
        _, fallback_perp = build_local_frame(start, previous_point=previous, fallback_end_point=end)

        perp_dir = fallback_perp
        perp_mag = max(seg_len * 0.2, 0.15)
        if native_vertex:
            v = native_vertex.Subtract(start)
            along_dist = v.DotProduct(direction)
            projected = direction.Multiply(along_dist)
            perp_vec = v.Subtract(projected)
            candidate_mag = self._vector_length(perp_vec)
            if candidate_mag > GEOM_TOL:
                perp_dir = perp_vec.Normalize()
                perp_mag = candidate_mag

        min_perp = max(seg_len * 0.08, 0.05)
        max_perp = max(seg_len * 0.45, 0.15)
        perp_mag = self._clamp(perp_mag, min_perp, max_perp)

        # Keep the vertex projection centered on the run.
        base = start.Add(direction.Multiply(seg_len * 0.5))
        vertex = base.Add(perp_dir.Multiply(perp_mag))

        if are_points_coincident(start, vertex) or are_points_coincident(vertex, end):
            base = start.Add(direction.Multiply(seg_len * 0.5))
            vertex = base.Add(perp_dir.Multiply(min_perp))
        return vertex

    def _build_homerun_points(self, start, native_end, native_vertex, previous_connector):
        end, direction, target_len, native_len = self._build_homerun_end(start, native_end, previous_connector)
        vertex = self._build_homerun_vertex(start, end, direction, native_vertex, previous_connector)
        logger.info(
            "HomeRun shape | native_len={:.3f} target_len={:.3f} native_vertex={} start={} vertex={} end={}".format(
                native_len,
                target_len,
                "yes" if native_vertex else "no",
                xyz_text(start),
                xyz_text(vertex),
                xyz_text(end)
            )
        )
        return [start, vertex, end]

    def _replace_homerun(self, homerun_wire):
        connected, unconnected = wire_connected_unconnected_connectors(homerun_wire)
        if not connected or not unconnected:
            return False

        start_connector = get_element_connector_from_wire_connector(connected[0])
        if not start_connector:
            return False
        start = start_connector.Origin
        native_end = unconnected[0].Origin

        previous_connector = find_previous_connector_for_homerun(
            start_connector,
            homerun_wire_id=homerun_wire.Id
        )
        native_vertex = None
        try:
            if getattr(homerun_wire, "NumberOfVertices", 0) > 0:
                native_vertex = homerun_wire.GetVertex(0)
        except Exception:
            native_vertex = None

        logger.info(
            "HomeRun replace | start={} native_end={} previous={} native_vertex={}".format(
                xyz_text(start),
                xyz_text(native_end),
                xyz_text(previous_connector.Origin) if previous_connector else "<none>"
                ,
                xyz_text(native_vertex) if native_vertex else "<none>"
            )
        )
        homerun_points = self._build_homerun_points(start, native_end, native_vertex, previous_connector)
        new_wire_type_id = get_wire_type_id(homerun_wire) or self.wire_type_id

        # Always recreate so we reliably force an arc homerun with a control vertex
        # even when native NewWires produced chamfer-style homerun geometry.
        self.doc.Delete(homerun_wire.Id)
        DB.Electrical.Wire.Create(
            self.doc,
            new_wire_type_id,
            self.view.Id,
            self.homerun_wiring_type,
            homerun_points,
            start_connector,
            None
        )
        return True

    def generate_for_circuit(self, circuit):
        cnum_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        cnum = cnum_param.AsString() if cnum_param else "<unknown>"
        logger.info("Generating wires for circuit {}".format(cnum))
        wire_set = circuit.NewWires(self.view, self.branch_wiring_type)
        if not wire_set:
            logger.info("No wires returned by NewWires for circuit {}.".format(cnum))
            return 0, 0

        wire_count = 0
        homerun_wire = None
        for wire in wire_set:
            wire_count += 1
            if not homerun_wire and is_homerun_wire(wire):
                homerun_wire = wire

        replaced = 0
        if homerun_wire and self._replace_homerun(homerun_wire):
            replaced = 1
        else:
            logger.info("No homerun replacement performed for circuit {}.".format(cnum))
        return wire_count, replaced


def run_generator_for_circuits(doc, circuits, generator):
    transaction = DB.Transaction(doc, "Create wires with NewWires + custom homerun")
    native_created = 0
    homeruns_replaced = 0
    removed_existing = 0
    try:
        transaction.Start()
        removed_existing, removed_by_circuit = remove_existing_wires_in_active_view(doc, generator.view.Id, circuits)
        if removed_existing:
            logger.info("Removed {} existing wire(s) in active view before regeneration.".format(removed_existing))
            for circuit_key, count in removed_by_circuit.items():
                logger.info("  CircuitId {} -> removed {} wire(s)".format(circuit_key, count))
        else:
            logger.info("No existing active-view wires found for selected circuits.")

        for circuit in circuits:
            created_count, replaced_count = generator.generate_for_circuit(circuit)
            native_created += created_count
            homeruns_replaced += replaced_count
        transaction.Commit()
        logger.info(
            "Wire generation complete. Removed existing: {} | Native wires: {} | Homeruns replaced: {}.".format(
                removed_existing, native_created, homeruns_replaced
            )
        )
    except Exception as ex:
        logger.error("Wire generation failed: {}".format(str(ex)))
        transaction.RollBack()


def main():
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    config = script.get_config("wire_type_config")

    all_circuits = collect_selected_electrical_circuits(doc, uidoc, logger=logger)
    if not all_circuits:
        logger.warning("No electrical circuits found from selected elements.")
        return

    target_type = resolve_target_system_type(all_circuits)
    if target_type is None:
        logger.warning("No system type selected. Cancelled.")
        return

    circuits = [c for c in all_circuits if c.SystemType == target_type]
    if not circuits:
        logger.warning("No circuits matched selected system type.")
        return

    wire_type_id = resolve_wire_type_id(doc, config, logger=logger)
    if not wire_type_id:
        script.exit()

    generator = WireGenerator(
        doc=doc,
        view=doc.ActiveView,
        wire_type_id=wire_type_id,
        branch_wiring_type=wiring_type_from_config("branch_wiring_type", "Chamfer", config),
        homerun_wiring_type=wiring_type_from_config("homerun_wiring_type", "Arc", config),
        home_run_length=homerun_length_from_config(config)
    )
    run_generator_for_circuits(doc, circuits, generator)


if __name__ == "__main__":
    main()
