# -*- coding: utf-8 -*-
import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import revit, DB, forms, script

from Snippets import _elecutils as eu
from Snippets import revit_helpers

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()
output.close_others()
def _idval(item):
    return int(revit_helpers.get_elementid_value(item))


def _idfrom(value):
    return revit_helpers.elementid_from_value(value)


class SwallowMergeFailures(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failures_accessor):
        try:
            for failure in failures_accessor.GetFailureMessages():
                try:
                    if failure.GetSeverity() == DB.FailureSeverity.Warning:
                        failures_accessor.DeleteWarning(failure)
                except Exception:
                    pass
        except Exception:
            pass
        return DB.FailureProcessingResult.Continue


def get_circuit_voltage_poles(circuit):
    voltage_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
    poles_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)

    voltage = None
    if voltage_param and voltage_param.HasValue:
        raw_voltage = voltage_param.AsDouble()
        try:
            # Convert Revit internal units to Volts using ForgeTypeId-backed unit id.
            voltage = DB.UnitUtils.ConvertFromInternalUnits(raw_voltage, DB.UnitTypeId.Volts)
        except Exception:
            voltage = raw_voltage
    poles = poles_param.AsInteger() if poles_param and poles_param.HasValue else None
    return voltage, poles


def get_start_slot(circuit):
    start_slot_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT)
    if start_slot_param and start_slot_param.HasValue:
        return start_slot_param.AsInteger()
    return 0


def format_circuit_label(circuit):
    ckt_id = _idval(circuit.Id)
    panel_name = circuit.BaseEquipment.Name if circuit.BaseEquipment else "No Panel"
    circuit_number = circuit.CircuitNumber or ""
    load_name = (circuit.LoadName or "").strip()

    rating = "N/A"
    pole = "?"
    if circuit.SystemType == DBE.ElectricalSystemType.PowerCircuit:
        try:
            rating = int(round(circuit.Rating, 0))
        except Exception:
            rating = "N/A"
        try:
            pole = circuit.PolesNumber
        except Exception:
            pole = "?"

    return "[{}]  {}/{} - {}  ({} A/{}P)".format(ckt_id, panel_name, circuit_number, load_name, rating, pole)


def format_panel_circuit(circuit):
    panel_name = circuit.BaseEquipment.Name if circuit.BaseEquipment else "No Panel"
    circuit_number = circuit.CircuitNumber or ""
    return "{}/{}".format(panel_name, circuit_number)


def format_voltage_pole(circuit):
    voltage, poles = get_circuit_voltage_poles(circuit)
    if voltage is None:
        voltage_text = "N/A"
    else:
        rounded = int(round(voltage))
        voltage_text = str(rounded) if abs(voltage - rounded) < 0.01 else "{:.1f}".format(voltage)
    pole_text = str(poles) if poles is not None else "?"
    return "{}V/{}P".format(voltage_text, pole_text)


def make_result_row(circuit, elements_moved, status, detail="", source_link=""):
    return [
        format_panel_circuit(circuit),
        (circuit.LoadName or "").strip() or "N/A",
        format_voltage_pole(circuit),
        elements_moved,
        status,
        detail,
        source_link
    ]


def make_result_row_values(panel_circuit, load_name, voltage_pole, elements_moved, status, detail="", source_link=""):
    return [
        panel_circuit,
        load_name,
        voltage_pole,
        elements_moved,
        status,
        detail,
        source_link
    ]


def build_circuit_groups(circuits):
    grouped_options = {" All": []}
    panel_groups = {}
    all_labels = []
    label_lookup = {}

    for ckt in circuits:
        panel_name = ckt.BaseEquipment.Name if ckt.BaseEquipment else "No Panel"
        load_name = (ckt.LoadName or "").strip()
        start_slot = get_start_slot(ckt)
        sort_key = (panel_name, start_slot, load_name)

        label = format_circuit_label(ckt)
        label_lookup[label] = ckt
        all_labels.append((sort_key, label))

        if panel_name not in panel_groups:
            panel_groups[panel_name] = []
        panel_groups[panel_name].append((sort_key, label))

    grouped_options[" All"] = [label for _, label in sorted(all_labels)]

    for panel_name, label_list in panel_groups.items():
        grouped_options[panel_name] = [label for _, label in sorted(label_list)]

    return grouped_options, label_lookup


def select_circuit(circuits, title, multiselect):
    grouped_options, label_lookup = build_circuit_groups(circuits)
    selected = forms.SelectFromList.show(
        grouped_options,
        title=title,
        group_selector_title="Panel:",
        multiselect=multiselect
    )

    if not selected:
        script.exit()

    if not isinstance(selected, list):
        selected = [selected]

    return [label_lookup[label] for label in selected]


def is_circuit_compatible(main_circuit, other_circuit, main_voltage, main_poles):
    if other_circuit.SystemType != main_circuit.SystemType:
        return False, "System type mismatch"

    voltage, poles = get_circuit_voltage_poles(other_circuit)
    if voltage is None or poles is None or main_voltage is None or main_poles is None:
        return False, "Missing voltage or poles"

    if abs(voltage - main_voltage) > 1.0:
        return False, "Voltage mismatch"

    if poles != main_poles:
        return False, "Poles mismatch"

    return True, ""


def get_circuit_elements(circuit):
    elements = []
    try:
        for el in circuit.Elements:
            if isinstance(el, DB.Element):
                elements.append(el)
    except Exception:
        pass
    return elements


def build_element_set(elements):
    element_set = DB.ElementSet()
    for el in elements:
        element_set.Insert(el)
    return element_set


def dedupe_circuits(circuits):
    unique = []
    seen_ids = set()
    for ckt in circuits or []:
        if not isinstance(ckt, DBE.ElectricalSystem):
            continue
        cid = _idval(ckt.Id)
        if cid in seen_ids:
            continue
        unique.append(ckt)
        seen_ids.add(cid)
    return unique


def get_selected_circuits(active_view=None):
    selected_circuits = []
    if active_view is None:
        active_view = uidoc.ActiveView

    # In panel schedule views, highlighted rows/cells resolve to circuit element ids.
    if isinstance(active_view, DBE.PanelScheduleView):
        try:
            selected_ids = uidoc.Selection.GetElementIds()
        except Exception:
            selected_ids = []

        for element_id in selected_ids:
            try:
                element = doc.GetElement(element_id)
            except Exception:
                element = None
            if isinstance(element, DBE.ElectricalSystem):
                selected_circuits.append(element)

        if selected_circuits:
            return dedupe_circuits(selected_circuits)

    selection = revit.get_selection().elements
    if not selection:
        return []

    try:
        selected_circuits = eu.get_circuits_from_selection(selection)
    except Exception:
        selected_circuits = []

    return dedupe_circuits(selected_circuits)


def get_circuits_from_panel_schedule_view(panel_schedule_view):
    circuits_by_id = {}
    try:
        table_data = panel_schedule_view.GetTableData()
        body = table_data.GetSectionData(DB.SectionType.Body)
        if not body:
            return []

        for row in range(body.NumberOfRows):
            for col in range(body.NumberOfColumns):
                circuit_id = panel_schedule_view.GetCircuitIdByCell(row, col)
                if not circuit_id or circuit_id == DB.ElementId.InvalidElementId:
                    continue
                cid = _idval(circuit_id)
                if cid in circuits_by_id:
                    continue
                circuit = doc.GetElement(circuit_id)
                if isinstance(circuit, DBE.ElectricalSystem):
                    circuits_by_id[cid] = circuit
    except Exception:
        return []

    return list(circuits_by_id.values())


def main():
    active_view = uidoc.ActiveView
    panel_schedule_mode = isinstance(active_view, DBE.PanelScheduleView)
    if panel_schedule_mode:
        all_circuits = get_circuits_from_panel_schedule_view(active_view)
    else:
        all_circuits = DB.FilteredElementCollector(doc) \
            .OfClass(DBE.ElectricalSystem) \
            .WhereElementIsNotElementType() \
            .WherePasses(eu.option_filter) \
            .ToElements()

    all_circuits = [c for c in all_circuits if c.CircuitType not in [DBE.CircuitType.Spare, DBE.CircuitType.Space]]

    if not all_circuits:
        if panel_schedule_mode:
            forms.alert("No circuits found in the active panel schedule.", exitscript=True)
        else:
            forms.alert("No circuits found in the model.", exitscript=True)

    all_circuit_ids = set([_idval(c.Id) for c in all_circuits])

    selected_circuits = get_selected_circuits(active_view=active_view)
    selected_circuits = [
        c for c in selected_circuits
        if _idval(c.Id) in all_circuit_ids
        and c.CircuitType not in [DBE.CircuitType.Spare, DBE.CircuitType.Space]
    ]

    if not selected_circuits:
        forms.alert(
            "No selection found. Please select circuit(s) or circuited element(s), then run MergeCircuits again.",
            exitscript=True
        )

    used_selection = bool(selected_circuits)

    main_circuit = select_circuit(selected_circuits, "Select Main Circuit (from Selection)", multiselect=False)[0]

    if main_circuit.CircuitType in [DBE.CircuitType.Spare, DBE.CircuitType.Space]:
        forms.alert("Main circuit cannot be a spare or space.", exitscript=True)

    main_voltage, main_poles = get_circuit_voltage_poles(main_circuit)
    if main_voltage is None or main_poles is None:
        forms.alert("Main circuit is missing voltage or poles.", exitscript=True)

    compatible_circuits = []

    for ckt in all_circuits:
        if ckt.Id == main_circuit.Id:
            continue
        ok, _ = is_circuit_compatible(main_circuit, ckt, main_voltage, main_poles)
        if ok:
            compatible_circuits.append(ckt)

    if not compatible_circuits:
        forms.alert("No compatible circuits found to merge into the main circuit.", exitscript=True)

    report_rows_by_id = None
    report_order_ids = []

    if used_selection:
        report_scope_circuits = selected_circuits
        report_rows_by_id = {}
        merge_candidates = []
        for ckt in report_scope_circuits:
            cid = _idval(ckt.Id)
            report_order_ids.append(cid)
            if ckt.Id == main_circuit.Id:
                report_rows_by_id[cid] = make_result_row(
                    ckt,
                    "",
                    "Source",
                    "Selected as merge target",
                    ""
                )
                continue
            ok, reason = is_circuit_compatible(main_circuit, ckt, main_voltage, main_poles)
            if ok:
                merge_candidates.append(ckt)
                report_rows_by_id[cid] = make_result_row(
                    ckt,
                    0,
                    "Not merged",
                    "Not selected for merging",
                    ""
                )
            else:
                report_rows_by_id[cid] = make_result_row(
                    ckt,
                    0,
                    "Not merged",
                    "Incompatible: {}".format(reason),
                    output.linkify(ckt.Id)
                )
        if not merge_candidates:
            forms.alert("No compatible selected circuits found to merge into the main circuit.", exitscript=True)
        circuits_to_merge = select_circuit(
            merge_candidates,
            "Select Circuits to Merge Into Main (from Selection)",
            multiselect=True
        )
    else:
        circuits_to_merge = select_circuit(
            compatible_circuits,
            "Select Circuits to Merge Into Main",
            multiselect=True
        )

    if not circuits_to_merge:
        script.exit()

    main_elements = set([_idval(el.Id) for el in get_circuit_elements(main_circuit)])

    result_rows = []
    merged_count = 0

    tg = DB.TransactionGroup(doc, "Merge Circuits")
    tg.Start()

    try:
        for src in circuits_to_merge:
            src_id_int = _idval(src.Id)
            src_id = _idfrom(src_id_int)
            src_panel_circuit = format_panel_circuit(src)
            src_load_name = (src.LoadName or "").strip() or "N/A"
            src_voltage_pole = format_voltage_pole(src)

            src_elements = [el for el in get_circuit_elements(src) if _idval(el.Id) not in main_elements]

            if not src_elements:
                row = make_result_row_values(
                    src_panel_circuit,
                    src_load_name,
                    src_voltage_pole,
                    0,
                    "Not merged",
                    "No elements to move",
                    ""
                )
                if report_rows_by_id is not None:
                    report_rows_by_id[src_id_int] = row
                else:
                    result_rows.append(row)
                continue

            try:
                t = DB.Transaction(doc, "Merge Circuit {}".format(src_id_int))
                t.Start()
                try:
                    fail_opts = t.GetFailureHandlingOptions()
                    fail_opts.SetFailuresPreprocessor(SwallowMergeFailures())
                    fail_opts.SetClearAfterRollback(True)
                    try:
                        fail_opts.SetForcedModalHandling(False)
                    except Exception:
                        pass
                    t.SetFailureHandlingOptions(fail_opts)
                except Exception:
                    pass

                try:
                    element_set = build_element_set(src_elements)
                    success = main_circuit.AddToCircuit(element_set)
                    if not success:
                        raise Exception("AddToCircuit returned False.")

                    doc.Regenerate()

                    for el in src_elements:
                        main_elements.add(_idval(el.Id))

                    pending_row = make_result_row_values(
                        src_panel_circuit,
                        src_load_name,
                        src_voltage_pole,
                        len(src_elements),
                        "Merged",
                        "",
                        ""
                    )
                    if t.Commit() != DB.TransactionStatus.Committed:
                        raise Exception("Merge transaction did not commit.")
                    merged_count += 1
                    if report_rows_by_id is not None:
                        report_rows_by_id[src_id_int] = pending_row
                    else:
                        result_rows.append(pending_row)
                except Exception:
                    if t.GetStatus() == DB.TransactionStatus.Started:
                        t.RollBack()
                    raise

            except Exception as ex:
                row = make_result_row_values(
                    src_panel_circuit,
                    src_load_name,
                    src_voltage_pole,
                    len(src_elements),
                    "Not merged",
                    str(ex),
                    output.linkify(src_id)
                )
                if report_rows_by_id is not None:
                    report_rows_by_id[src_id_int] = row
                else:
                    result_rows.append(row)

        if merged_count:
            tg.Assimilate()
        else:
            tg.RollBack()

    except Exception as ex:
        tg.RollBack()
        forms.alert("Merge failed: {}".format(str(ex)), exitscript=True)

    panel_name = main_circuit.BaseEquipment.Name if main_circuit.BaseEquipment else "No Panel"
    circuit_number = main_circuit.CircuitNumber or ""
    output.print_md("Main circuit: {} / {} ({})".format(panel_name, circuit_number, output.linkify(main_circuit.Id)))

    if used_selection:
        output.print_md("Selection mode: {} circuit(s) detected from current selection.".format(len(selected_circuits)))

    if report_rows_by_id is not None:
        result_rows = [report_rows_by_id[cid] for cid in report_order_ids if cid in report_rows_by_id]

    if result_rows:
        output.print_table(
            result_rows,
            [
                "Original Circuit",
                "Load Name",
                "Voltage/Pole",
                "Elements Merged",
                "Status",
                "Detail",
                "Source Link (failed only)"
            ]
        )


if __name__ == "__main__":
    main()

