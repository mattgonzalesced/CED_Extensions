# -*- coding: utf-8 -*-
import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import revit, DB, forms, script
from pyrevit.compat import get_elementid_value_func, get_elementid_from_value_func

from Snippets import _elecutils as eu

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()
output.close_others()
_get_elid_value = get_elementid_value_func()
_get_elid_from_value = get_elementid_from_value_func()


def _idval(item):
    try:
        return int(_get_elid_value(item))
    except Exception:
        return int(getattr(item, "IntegerValue", 0))


def _idfrom(value):
    return _get_elid_from_value(int(value))


def get_circuit_voltage_poles(circuit):
    voltage_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
    poles_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)

    voltage = voltage_param.AsDouble() if voltage_param and voltage_param.HasValue else None
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


def get_selected_circuits():
    selection = revit.get_selection().elements
    if not selection:
        return []

    try:
        selected_circuits = eu.get_circuits_from_selection(selection)
    except Exception:
        selected_circuits = []

    return dedupe_circuits(selected_circuits)


def main():
    all_circuits = DB.FilteredElementCollector(doc) \
        .OfClass(DBE.ElectricalSystem) \
        .WhereElementIsNotElementType() \
        .WherePasses(eu.option_filter) \
        .ToElements()

    all_circuits = [c for c in all_circuits if c.CircuitType not in [DBE.CircuitType.Spare, DBE.CircuitType.Space]]

    if not all_circuits:
        forms.alert("No circuits found in the model.", exitscript=True)

    all_circuit_ids = set([_idval(c.Id) for c in all_circuits])

    selected_circuits = get_selected_circuits()
    selected_circuits = [
        c for c in selected_circuits
        if _idval(c.Id) in all_circuit_ids
        and c.CircuitType not in [DBE.CircuitType.Spare, DBE.CircuitType.Space]
    ]

    used_selection = bool(selected_circuits)

    if used_selection and len(selected_circuits) == 1:
        main_circuit = selected_circuits[0]
    elif used_selection:
        main_circuit = select_circuit(selected_circuits, "Select Main Circuit (from Selection)", multiselect=False)[0]
    else:
        main_circuit = select_circuit(all_circuits, "Select Main Circuit", multiselect=False)[0]

    if main_circuit.CircuitType in [DBE.CircuitType.Spare, DBE.CircuitType.Space]:
        forms.alert("Main circuit cannot be a spare or space.", exitscript=True)

    main_voltage, main_poles = get_circuit_voltage_poles(main_circuit)
    if main_voltage is None or main_poles is None:
        forms.alert("Main circuit is missing voltage or poles.", exitscript=True)

    compatible_circuits = []
    incompat_count = 0

    for ckt in all_circuits:
        if ckt.Id == main_circuit.Id:
            continue
        ok, _ = is_circuit_compatible(main_circuit, ckt, main_voltage, main_poles)
        if ok:
            compatible_circuits.append(ckt)
        else:
            incompat_count += 1

    if not compatible_circuits:
        forms.alert("No compatible circuits found to merge into the main circuit.", exitscript=True)

    selected_candidates = []
    selected_compat_to_merge = []
    selected_incompat_count = 0

    if used_selection:
        selected_candidates = [c for c in selected_circuits if c.Id != main_circuit.Id]
        for ckt in selected_candidates:
            ok, _ = is_circuit_compatible(main_circuit, ckt, main_voltage, main_poles)
            if ok:
                selected_compat_to_merge.append(ckt)
            else:
                selected_incompat_count += 1

    if selected_compat_to_merge:
        circuits_to_merge = selected_compat_to_merge
    else:
        circuits_to_merge = select_circuit(
            compatible_circuits,
            "Select Circuits to Merge Into Main",
            multiselect=True
        )

    if not circuits_to_merge:
        script.exit()

    main_elements = set([_idval(el.Id) for el in get_circuit_elements(main_circuit)])

    merged_rows = []
    skipped_rows = []
    failed_rows = []

    tg = DB.TransactionGroup(doc, "Merge Circuits")
    tg.Start()

    try:
        for src in circuits_to_merge:
            src_id_int = _idval(src.Id)
            src_id = _idfrom(src_id_int)
            src_link = output.linkify(src_id)

            src_elements = [el for el in get_circuit_elements(src) if _idval(el.Id) not in main_elements]

            if not src_elements:
                skipped_rows.append([src_link, 0, "No elements to move"])
                continue

            try:
                with revit.Transaction("Merge Circuit {}".format(src_id_int)):
                    element_set = build_element_set(src_elements)
                    success = main_circuit.AddToCircuit(element_set)
                    if not success:
                        raise Exception("AddToCircuit returned False.")

                    doc.Regenerate()

                    for el in src_elements:
                        main_elements.add(_idval(el.Id))

                    if not src.IsValidObject:
                        merged_rows.append([src_link, len(src_elements), "Merged (source invalid/deleted)"])
                    else:
                        remaining = get_circuit_elements(src)
                        if not remaining:
                            doc.Delete(src_id)
                            merged_rows.append([src_link, len(src_elements), "Merged + deleted empty circuit"])
                        else:
                            merged_rows.append([src_link, len(src_elements), "Merged"])

            except Exception as ex:
                failed_rows.append([src_link, len(src_elements), "Failed: {}".format(str(ex))])

        if merged_rows:
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
        if selected_compat_to_merge:
            output.print_md("Merged {} compatible selected circuit(s) directly.".format(len(selected_compat_to_merge)))
        if selected_incompat_count:
            output.print_md("Ignored {} incompatible selected circuit(s).".format(selected_incompat_count))

    if incompat_count:
        output.print_md("Filtered out {} incompatible circuit(s).".format(incompat_count))

    if merged_rows:
        output.print_table(merged_rows, ["Source Circuit", "Elements Moved", "Status"])

    if skipped_rows:
        output.print_table(skipped_rows, ["Source Circuit", "Elements Moved", "Status"])

    if failed_rows:
        output.print_table(failed_rows, ["Source Circuit", "Elements Moved", "Status"])


if __name__ == "__main__":
    main()
