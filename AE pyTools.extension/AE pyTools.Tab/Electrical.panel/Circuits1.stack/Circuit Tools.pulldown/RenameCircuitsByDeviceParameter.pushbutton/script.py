# -*- coding: utf-8 -*-
import os
from collections import OrderedDict

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import revit, DB, forms, script
from Snippets import _elecutils as eu
from System.Windows import Thickness
from System.Windows.Controls import StackPanel, Orientation, TextBlock, ComboBox


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
output.close_others()

XAML_PATH = os.path.join(os.path.dirname(__file__), "ParameterSelectionWindow.xaml")
NO_CHANGE_OPTION = "<No Change>"
NO_PARAMETER_OPTION = "<No common text parameters available>"

try:
    text_type = unicode
except NameError:
    text_type = str


def _safe_text(value):
    if value is None:
        return u""
    try:
        return value.strip()
    except Exception:
        try:
            return text_type(value).strip()
        except Exception:
            return u""


def _dedupe_elements(elements):
    unique = []
    seen_ids = set()
    for element in elements or []:
        if not element:
            continue
        try:
            eid = element.Id.IntegerValue
        except Exception:
            continue
        if eid in seen_ids:
            continue
        seen_ids.add(eid)
        unique.append(element)
    return unique


def _is_valid_circuit(circuit):
    if not isinstance(circuit, DBE.ElectricalSystem):
        return False
    if not circuit.IsValidObject:
        return False
    if circuit.CircuitType in [DBE.CircuitType.Spare, DBE.CircuitType.Space]:
        return False
    return True


def _get_selected_elements():
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        forms.alert("Select device(s) first, then run the tool.", exitscript=True)
    selected_elements = []
    for element_id in selected_ids:
        element = doc.GetElement(element_id)
        if not element:
            continue
        if getattr(element, "ViewSpecific", False):
            continue
        selected_elements.append(element)
    if not selected_elements:
        forms.alert("No valid model elements found in the current selection.", exitscript=True)
    return selected_elements


def _get_element_circuits(element):
    if isinstance(element, DBE.ElectricalSystem):
        return [element]

    circuits = []
    try:
        circuits = eu.get_circuits_from_selection([element]) or []
    except Exception:
        circuits = []

    if not circuits:
        try:
            mep_model = element.MEPModel
            if mep_model:
                circuits = list(mep_model.GetElectricalSystems() or [])
        except Exception:
            circuits = []

    valid = []
    for circuit in circuits:
        if _is_valid_circuit(circuit):
            valid.append(circuit)
    return valid


def _get_circuit_load_elements(circuit):
    elements = []
    try:
        for element in circuit.Elements:
            if isinstance(element, DB.Element) and not isinstance(element, DBE.ElectricalSystem):
                elements.append(element)
    except Exception:
        pass
    return _dedupe_elements(elements)


def _build_circuit_map(selected_elements):
    circuit_map = OrderedDict()

    for element in selected_elements:
        for circuit in _get_element_circuits(element):
            cid = circuit.Id.IntegerValue
            if cid not in circuit_map:
                circuit_map[cid] = {
                    "circuit": circuit,
                    "source_elements": []
                }
            if not isinstance(element, DBE.ElectricalSystem):
                circuit_map[cid]["source_elements"].append(element)

    if not circuit_map:
        forms.alert("No connected circuits were found from the current selection.", exitscript=True)

    for cid, data in circuit_map.items():
        src = _dedupe_elements(data["source_elements"])
        if not src:
            src = _get_circuit_load_elements(data["circuit"])
        data["source_elements"] = src

    return circuit_map


def _collect_text_param_names(param_owner):
    names = set()
    if not param_owner:
        return names
    try:
        parameters = param_owner.Parameters
    except Exception:
        return names

    for param in parameters:
        try:
            if param.StorageType != DB.StorageType.String:
                continue
            definition = param.Definition
            if not definition:
                continue
            pname = _safe_text(definition.Name)
            if pname:
                names.add(pname)
        except Exception:
            continue
    return names


def _get_text_param_names(element):
    names = set()
    names.update(_collect_text_param_names(element))

    type_element = None
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != DB.ElementId.InvalidElementId:
            type_element = doc.GetElement(type_id)
    except Exception:
        type_element = None

    names.update(_collect_text_param_names(type_element))
    return names


def _get_common_text_parameters(elements):
    if not elements:
        return []

    common_names = None
    for element in elements:
        names = _get_text_param_names(element)
        if common_names is None:
            common_names = names
        else:
            common_names &= names
        if not common_names:
            break

    if not common_names:
        return []

    return sorted(list(common_names), key=lambda x: x.lower())


def _read_parameter_string(param):
    if not param:
        return u""
    try:
        if param.StorageType != DB.StorageType.String:
            return u""
    except Exception:
        return u""

    try:
        value = param.AsString()
    except Exception:
        value = None

    if not value:
        try:
            value = param.AsValueString()
        except Exception:
            value = None

    return _safe_text(value)


def _get_parameter_value(element, param_name):
    if not element or not param_name:
        return u""

    param = None
    try:
        param = element.LookupParameter(param_name)
    except Exception:
        param = None

    value = _read_parameter_string(param)
    if value:
        return value

    try:
        type_id = element.GetTypeId()
        if type_id and type_id != DB.ElementId.InvalidElementId:
            type_element = doc.GetElement(type_id)
            if type_element:
                type_param = type_element.LookupParameter(param_name)
                return _read_parameter_string(type_param)
    except Exception:
        pass

    return u""


def _resolve_single_name(elements, param_name):
    values = []
    seen = set()
    for element in elements or []:
        value = _get_parameter_value(element, param_name)
        if not value:
            continue
        if value in seen:
            continue
        seen.add(value)
        values.append(value)

    if not values:
        return None, "No non-empty value found for '{}' on connected selected devices.".format(param_name)

    if len(values) > 1:
        preview = ", ".join(values[:3])
        return None, "Multiple values found for '{}': {}".format(param_name, preview)

    return values[0], None


def _get_circuit_start_slot(circuit):
    start_slot = 0
    try:
        start_slot_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT)
        if start_slot_param and start_slot_param.HasValue:
            start_slot = start_slot_param.AsInteger()
    except Exception:
        pass
    return start_slot


def _format_circuit_label(circuit):
    panel_name = "No Panel"
    try:
        if circuit.BaseEquipment:
            panel_name = _safe_text(circuit.BaseEquipment.Name) or "No Panel"
    except Exception:
        panel_name = "No Panel"

    circuit_number = _safe_text(circuit.CircuitNumber) or "?"
    load_name = _safe_text(circuit.LoadName)
    if load_name:
        return "[{}] {} / {} - {}".format(circuit.Id.IntegerValue, panel_name, circuit_number, load_name)
    return "[{}] {} / {}".format(circuit.Id.IntegerValue, panel_name, circuit_number)


def _get_circuit_sort_key(circuit):
    panel_name = "No Panel"
    try:
        if circuit.BaseEquipment:
            panel_name = _safe_text(circuit.BaseEquipment.Name) or "No Panel"
    except Exception:
        panel_name = "No Panel"

    return (
        panel_name.lower(),
        _get_circuit_start_slot(circuit),
        _safe_text(circuit.CircuitNumber),
        circuit.Id.IntegerValue
    )


def _set_circuit_name(circuit, new_name):
    try:
        target_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
    except Exception:
        target_param = None

    if target_param and (not target_param.IsReadOnly):
        try:
            target_param.Set(new_name)
            return None
        except Exception as ex:
            return "Failed to set circuit name: {}".format(str(ex))

    try:
        lookup_param = circuit.LookupParameter("Load Name")
    except Exception:
        lookup_param = None

    if lookup_param and (not lookup_param.IsReadOnly):
        try:
            lookup_param.Set(new_name)
            return None
        except Exception as ex:
            return "Failed to set 'Load Name': {}".format(str(ex))

    return "Circuit name parameter is unavailable or read-only."


class CircuitParameterSelectionWindow(forms.WPFWindow):
    def __init__(self, xaml_path, row_data):
        forms.WPFWindow.__init__(self, xaml_path)
        self._row_data = row_data
        self._combos = {}
        self.selections = {}

        header_text = self.FindName("HeaderText")
        if header_text is not None:
            header_text.Text = "Select one text parameter for each circuit (default is No Change)."

        self._build_rows()

        apply_btn = self.FindName("ApplyButton")
        cancel_btn = self.FindName("CancelButton")
        if apply_btn is not None:
            apply_btn.Click += self._on_apply
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    def _build_rows(self):
        host = self.FindName("CircuitRowsPanel")
        if host is None:
            raise Exception("CircuitRowsPanel not found in XAML.")

        for row in self._row_data:
            panel = StackPanel()
            panel.Orientation = Orientation.Horizontal
            panel.Margin = Thickness(0, 0, 0, 6)

            label = TextBlock()
            label.Text = row["circuit_label"]
            label.Width = 440
            label.Margin = Thickness(0, 0, 10, 0)
            panel.Children.Add(label)

            combo = ComboBox()
            combo.Width = 240
            combo.Items.Add(NO_CHANGE_OPTION)

            if row["parameter_names"]:
                for pname in row["parameter_names"]:
                    combo.Items.Add(pname)
                combo.SelectedIndex = 0
            else:
                combo.SelectedIndex = 0

            self._combos[row["circuit_id"]] = combo
            panel.Children.Add(combo)
            host.Children.Add(panel)

    def _on_apply(self, sender, args):
        selections = {}

        for row in self._row_data:
            circuit_id = row["circuit_id"]
            combo = self._combos.get(circuit_id)
            if combo is None:
                selections[circuit_id] = None
                continue

            selected = combo.SelectedItem
            if selected is None:
                forms.alert("Select a parameter for every enabled circuit row.", exitscript=False)
                return

            selected_name = _safe_text(selected)
            if not selected_name:
                forms.alert("Select a parameter for every enabled circuit row.", exitscript=False)
                return

            selections[circuit_id] = selected_name

        self.selections = selections
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


def _build_row_data(circuit_map):
    sorted_circuits = sorted(circuit_map.values(), key=lambda x: _get_circuit_sort_key(x["circuit"]))
    row_data = []
    for data in sorted_circuits:
        circuit = data["circuit"]
        source_elements = data["source_elements"]
        parameter_names = _get_common_text_parameters(source_elements)

        row_data.append({
            "circuit_id": circuit.Id.IntegerValue,
            "circuit": circuit,
            "circuit_label": _format_circuit_label(circuit),
            "source_elements": source_elements,
            "parameter_names": parameter_names
        })
    return row_data


def _circuit_ref(circuit):
    panel_name = "No Panel"
    try:
        if circuit.BaseEquipment:
            panel_name = _safe_text(circuit.BaseEquipment.Name) or "No Panel"
    except Exception:
        panel_name = "No Panel"
    circuit_number = _safe_text(circuit.CircuitNumber) or "?"
    return "{} / {} ({})".format(panel_name, circuit_number, output.linkify(circuit.Id))


def main():
    selected_elements = _get_selected_elements()
    circuit_map = _build_circuit_map(selected_elements)
    row_data = _build_row_data(circuit_map)

    if not row_data:
        forms.alert("No valid circuits found from the current selection.", exitscript=True)

    ui_window = CircuitParameterSelectionWindow(XAML_PATH, row_data)
    result = ui_window.show_dialog()
    if not result:
        script.exit()

    selected_parameters = ui_window.selections or {}

    results = []
    renamed_count = 0
    skipped_count = 0

    with revit.Transaction("Rename Circuits by Device Parameter"):
        for row in row_data:
            circuit = row["circuit"]
            circuit_id = row["circuit_id"]
            selected_param = selected_parameters.get(circuit_id)
            previous_name = _safe_text(circuit.LoadName)

            if (not selected_param) or selected_param == NO_CHANGE_OPTION:
                skipped_count += 1
                results.append([
                    _circuit_ref(circuit),
                    selected_param or NO_CHANGE_OPTION,
                    previous_name or "-",
                    "-",
                    "No Change"
                ])
                continue

            new_name, resolve_error = _resolve_single_name(row["source_elements"], selected_param)
            if resolve_error:
                skipped_count += 1
                results.append([
                    _circuit_ref(circuit),
                    selected_param,
                    previous_name or "-",
                    "-",
                    "Skipped: {}".format(resolve_error)
                ])
                continue

            set_error = _set_circuit_name(circuit, new_name)
            if set_error:
                skipped_count += 1
                results.append([
                    _circuit_ref(circuit),
                    selected_param,
                    previous_name or "-",
                    new_name,
                    "Skipped: {}".format(set_error)
                ])
                continue

            renamed_count += 1
            results.append([
                _circuit_ref(circuit),
                selected_param,
                previous_name or "-",
                new_name,
                "Renamed"
            ])

    output.print_md("### Rename Circuits by Device Parameter")
    output.print_md("Processed {} circuit(s): {} renamed, {} skipped.".format(len(row_data), renamed_count, skipped_count))
    output.print_table(results, ["Circuit", "Parameter", "Previous Name", "New Name", "Status"])

    forms.alert(
        "Processed {} circuit(s).\nRenamed: {}\nSkipped: {}".format(len(row_data), renamed_count, skipped_count),
        title="Rename Circuits by Device Parameter",
        exitscript=False
    )


if __name__ == "__main__":
    main()

