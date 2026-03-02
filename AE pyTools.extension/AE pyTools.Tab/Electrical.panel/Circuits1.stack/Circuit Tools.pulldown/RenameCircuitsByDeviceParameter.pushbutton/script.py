# -*- coding: utf-8 -*-
import os
from collections import OrderedDict

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import revit, DB, forms, script
from Snippets import _elecutils as eu


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
output.close_others()

XAML_PATH = os.path.join(os.path.dirname(__file__), "ParameterSelectionWindow.xaml")
CONFIG_KEY = "rename_circuit_builder_config"

TOKEN_PARAM = "param"
TOKEN_SEPARATOR = "sep"
TOKEN_CUSTOM = "custom"


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


def _raw_text(value):
    if value is None:
        return u""
    try:
        return text_type(value)
    except Exception:
        return u""


def _clone_tokens(tokens):
    cloned = []
    for token in tokens or []:
        ttype = _safe_text(token.get("type"))
        tval = _raw_text(token.get("value"))
        if not ttype:
            continue
        cloned.append({"type": ttype, "value": tval})
    return cloned


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
                circuit_map[cid] = {"circuit": circuit, "source_elements": []}
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


def _resolve_preferred_param_value(elements, param_name):
    non_empty_counts = {}
    first_element_id_by_value = {}

    sortable = []
    for element in elements or []:
        try:
            sortable.append((element.Id.IntegerValue, element))
        except Exception:
            continue
    sortable.sort(key=lambda x: x[0])

    for eid, element in sortable:
        value = _get_parameter_value(element, param_name)
        if not value:
            continue
        if value not in non_empty_counts:
            non_empty_counts[value] = 0
            first_element_id_by_value[value] = eid
        non_empty_counts[value] += 1
        if eid < first_element_id_by_value[value]:
            first_element_id_by_value[value] = eid

    if not non_empty_counts:
        return u""

    best_value = None
    best_count = -1
    best_first_id = 2147483647
    for value, count in non_empty_counts.items():
        first_id = first_element_id_by_value.get(value, 2147483647)
        if count > best_count or (count == best_count and first_id < best_first_id):
            best_value = value
            best_count = count
            best_first_id = first_id

    return best_value or u""


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


def _build_row_data(circuit_map):
    sorted_circuits = sorted(circuit_map.values(), key=lambda x: _get_circuit_sort_key(x["circuit"]))
    row_data = []
    for data in sorted_circuits:
        circuit = data["circuit"]
        source_elements = data["source_elements"]
        parameter_names = _get_common_text_parameters(source_elements)
        panel_name = "No Panel"
        try:
            if circuit.BaseEquipment:
                panel_name = _safe_text(circuit.BaseEquipment.Name) or "No Panel"
        except Exception:
            panel_name = "No Panel"
        circuit_number = _safe_text(circuit.CircuitNumber) or "<unnamed>"
        current_load_name = _safe_text(circuit.LoadName) or "<unnamed>"
        selector_name = "{} | {} | {}".format(panel_name, circuit_number, current_load_name)

        resolved_values = {}
        for pname in parameter_names:
            resolved_values[pname] = _resolve_preferred_param_value(source_elements, pname)

        row_data.append({
            "circuit_id": circuit.Id.IntegerValue,
            "circuit": circuit,
            "circuit_label": _format_circuit_label(circuit),
            "selector_name": selector_name,
            "sort_key": _get_circuit_sort_key(circuit),
            "source_elements": source_elements,
            "parameter_names": parameter_names,
            "resolved_values": resolved_values,
            "existing_name": _safe_text(circuit.LoadName),
        })
    return row_data


def _token_to_label(token):
    ttype = _safe_text(token.get("type"))
    tval = _raw_text(token.get("value"))
    if ttype == TOKEN_PARAM:
        return "[Param] {}".format(tval)
    if ttype == TOKEN_SEPARATOR:
        return "[Sep] {}".format(tval.replace(" ", "<space>"))
    if ttype == TOKEN_CUSTOM:
        return "[Text] {}".format(tval)
    return "[?] {}".format(tval)


class _CircuitItem(object):
    def __init__(self, circuit_id, display):
        self.CircuitId = circuit_id
        self.Display = display

    def __str__(self):
        return self.Display

    def __unicode__(self):
        return self.Display


class _CircuitPreviewItem(object):
    def __init__(self, circuit_id, display):
        self.CircuitId = circuit_id
        self.Display = display

    def __str__(self):
        return self.Display

    def __unicode__(self):
        return self.Display


class CircuitStringBuilderWindow(forms.WPFWindow):
    def __init__(self, xaml_path, row_data):
        forms.WPFWindow.__init__(self, xaml_path)
        self._row_data = row_data or []
        self._rows_by_id = {}
        self._states = {}
        self._active_circuit_id = None
        self._is_programmatic_preview_update = False
        self._is_programmatic_selection_update = False
        self._is_programmatic_preview_selection_update = False

        self._config = script.get_config(CONFIG_KEY)
        self._templates = self._load_templates_from_config()

        self.result_names = {}
        self.result_target_ids = []

        for row in self._row_data:
            cid = row["circuit_id"]
            self._rows_by_id[cid] = row
            self._states[cid] = {
                "tokens": [],
                "preview_text": row.get("existing_name", u""),
                "manual_override": True,
            }

        self._bind_static_controls()
        self._populate_circuits()
        self._populate_templates_combo()
        self._populate_quick_add_defaults()
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def _bind_static_controls(self):
        self._circuit_list = self.FindName("CircuitList")
        self._active_circuit_text = self.FindName("ActiveCircuitText")
        self._available_params = self.FindName("AvailableParamsList")
        self._string_parts = self.FindName("StringPartsList")
        self._preview_textbox = self.FindName("PreviewTextBox")
        self._selected_preview_list = self.FindName("SelectedPreviewList")
        self._custom_field_combo = self.FindName("CustomFieldCombo")
        self._separator_combo = self.FindName("SeparatorCombo")
        self._saved_template_combo = self.FindName("SavedTemplateCombo")
        self._template_name_text = self.FindName("TemplateNameText")
        self._increment_checkbox = self.FindName("IncrementSelectedCheckBox")

    def _load_templates_from_config(self):
        templates = getattr(self._config, "saved_templates", {})
        if not isinstance(templates, dict):
            return {}
        cleaned = {}
        for name, tokens in templates.items():
            template_name = _safe_text(name)
            if not template_name:
                continue
            cleaned[template_name] = _clone_tokens(tokens)
        return cleaned

    def _save_templates_to_config(self):
        self._config.saved_templates = self._templates
        script.save_config()

    def _populate_quick_add_defaults(self):
        if self._separator_combo is not None:
            for item in [" ", "-", "_", "/", ".", ":"]:
                self._separator_combo.Items.Add(item)
            self._separator_combo.Text = "-"

        if self._custom_field_combo is not None:
            for item in ["", "EXISTING", "NEW", "SPARE"]:
                self._custom_field_combo.Items.Add(item)
            self._custom_field_combo.Text = ""

    def _populate_templates_combo(self):
        if self._saved_template_combo is None:
            return
        self._saved_template_combo.Items.Clear()
        for tname in sorted(self._templates.keys(), key=lambda x: x.lower()):
            self._saved_template_combo.Items.Add(tname)
        if self._saved_template_combo.Items.Count > 0:
            self._saved_template_combo.SelectedIndex = 0

    def _populate_circuits(self):
        if self._circuit_list is None:
            return

        self._is_programmatic_selection_update = True
        self._circuit_list.Items.Clear()

        for row in self._row_data:
            display = row.get("selector_name", "<unnamed>")
            self._circuit_list.Items.Add(_CircuitItem(row["circuit_id"], display))

        if self._circuit_list.Items.Count > 0:
            first_item = self._circuit_list.Items[0]
            self._circuit_list.SelectedItems.Add(first_item)
            self._active_circuit_id = getattr(first_item, "CircuitId", None)

        self._is_programmatic_selection_update = False

    def _get_selected_circuit_ids(self):
        selected_ids = []
        if self._circuit_list is None:
            return selected_ids
        for item in self._circuit_list.SelectedItems:
            cid = getattr(item, "CircuitId", None)
            if cid is not None:
                selected_ids.append(cid)
        return selected_ids

    def _get_active_row(self):
        if self._active_circuit_id is None:
            return None
        return self._rows_by_id.get(self._active_circuit_id)

    def _get_active_state(self):
        if self._active_circuit_id is None:
            return None
        return self._states.get(self._active_circuit_id)

    def _compute_preview_from_tokens(self, circuit_id, tokens):
        row = self._rows_by_id.get(circuit_id)
        if not row:
            return u""
        resolved_values = row.get("resolved_values", {})

        parts = []
        for token in tokens or []:
            ttype = _safe_text(token.get("type"))
            tval = _raw_text(token.get("value"))
            if ttype == TOKEN_PARAM:
                parts.append(_raw_text(resolved_values.get(tval, u"")))
            elif ttype in (TOKEN_SEPARATOR, TOKEN_CUSTOM):
                parts.append(tval)
        return u"".join(parts)

    def _get_preview_for_circuit(self, circuit_id):
        state = self._states.get(circuit_id)
        if not state:
            return u""
        if state.get("manual_override", False):
            return _raw_text(state.get("preview_text", u""))
        computed = self._compute_preview_from_tokens(circuit_id, state.get("tokens", []))
        state["preview_text"] = computed
        return computed

    def _set_preview_textbox(self, text_value):
        if self._preview_textbox is None:
            return
        self._is_programmatic_preview_update = True
        self._preview_textbox.Text = _raw_text(text_value)
        self._is_programmatic_preview_update = False

    def _refresh_active_circuit_views(self):
        row = self._get_active_row()
        state = self._get_active_state()
        if not row or not state:
            if self._active_circuit_text is not None:
                self._active_circuit_text.Text = "Active Circuit: (none)"
            if self._available_params is not None:
                self._available_params.Items.Clear()
            if self._string_parts is not None:
                self._string_parts.Items.Clear()
            self._set_preview_textbox(u"")
            return

        if self._active_circuit_text is not None:
            self._active_circuit_text.Text = "Active Circuit: {}".format(row["circuit_label"])

        if self._available_params is not None:
            self._available_params.Items.Clear()
            for pname in row.get("parameter_names", []):
                self._available_params.Items.Add(pname)

        if self._string_parts is not None:
            self._string_parts.Items.Clear()
            for token in state.get("tokens", []):
                self._string_parts.Items.Add(_token_to_label(token))

        self._set_preview_textbox(self._get_preview_for_circuit(self._active_circuit_id))

    def _refresh_selected_preview_list(self):
        if self._selected_preview_list is None:
            return
        self._is_programmatic_preview_selection_update = True
        self._selected_preview_list.Items.Clear()

        selected_ids = self._get_selected_circuit_ids()
        if not selected_ids and self._active_circuit_id is not None:
            selected_ids = [self._active_circuit_id]

        selected_rows = []
        for cid in selected_ids:
            row = self._rows_by_id.get(cid)
            if row:
                selected_rows.append(row)
        selected_rows.sort(key=lambda r: r["sort_key"])

        preview_names = {}
        for row in selected_rows:
            cid = row["circuit_id"]
            preview_names[cid] = _safe_text(self._get_preview_for_circuit(cid))

        if self._increment_checkbox is not None and bool(self._increment_checkbox.IsChecked):
            preview_names = self._apply_increment_suffix(preview_names)

        for row in selected_rows:
            cid = row["circuit_id"]
            preview_text = preview_names.get(cid, _safe_text(self._get_preview_for_circuit(cid)))
            item = _CircuitPreviewItem(cid, "{} -> {}".format(row["circuit_label"], preview_text or ""))
            self._selected_preview_list.Items.Add(item)
            if self._active_circuit_id is not None and cid == self._active_circuit_id:
                self._selected_preview_list.SelectedItem = item

        self._is_programmatic_preview_selection_update = False

    def _ensure_combo_has_value(self, combo, value):
        if combo is None:
            return
        raw = _raw_text(value)
        for item in combo.Items:
            if _raw_text(item) == raw:
                return
        combo.Items.Add(raw)

    def _append_token_to_active(self, token_type, token_value):
        state = self._get_active_state()
        if not state:
            return

        if token_type == TOKEN_PARAM:
            value = _safe_text(token_value)
            if not value:
                return
        else:
            value = _raw_text(token_value)
            if value == u"":
                return

        tokens = state.get("tokens", [])
        tokens.append({"type": token_type, "value": value})
        state["tokens"] = tokens
        state["manual_override"] = False
        state["preview_text"] = self._compute_preview_from_tokens(self._active_circuit_id, tokens)
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def _apply_tokens_to_targets(self, target_ids):
        state = self._get_active_state()
        if not state:
            return
        source_tokens = _clone_tokens(state.get("tokens", []))
        if not source_tokens:
            forms.alert("Active circuit has no string parts to apply.", exitscript=False)
            return

        for cid in target_ids:
            if cid not in self._states:
                continue
            target_state = self._states[cid]
            target_state["tokens"] = _clone_tokens(source_tokens)
            target_state["manual_override"] = False
            target_state["preview_text"] = self._compute_preview_from_tokens(cid, source_tokens)

        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def _get_selected_string_part_indices(self):
        indices = []
        if self._string_parts is None:
            return indices
        for item in self._string_parts.SelectedItems:
            try:
                idx = self._string_parts.Items.IndexOf(item)
            except Exception:
                idx = -1
            if idx >= 0:
                indices.append(idx)
        return sorted(indices)

    def _apply_increment_suffix(self, names_by_circuit):
        ordered_ids = sorted(names_by_circuit.keys(), key=lambda cid: self._rows_by_id[cid]["sort_key"])
        grouped = OrderedDict()
        for cid in ordered_ids:
            base_name = _safe_text(names_by_circuit.get(cid))
            if not base_name:
                continue
            if base_name not in grouped:
                grouped[base_name] = []
            grouped[base_name].append(cid)

        result = dict(names_by_circuit)
        for base_name, circuit_ids in grouped.items():
            if len(circuit_ids) <= 1:
                continue
            index = 1
            for cid in circuit_ids:
                if index == 1:
                    result[cid] = base_name
                else:
                    result[cid] = u"{} #{}".format(base_name, index)
                index += 1
        return result

    # --- UI Event Handlers ---
    def CircuitList_SelectionChanged(self, sender, args):
        if self._is_programmatic_selection_update:
            return

        selected_ids = self._get_selected_circuit_ids()
        if not selected_ids:
            if self._circuit_list is not None and self._circuit_list.Items.Count > 0:
                self._is_programmatic_selection_update = True
                fallback = self._circuit_list.Items[0]
                self._circuit_list.SelectedItems.Add(fallback)
                self._is_programmatic_selection_update = False
                fallback_id = getattr(fallback, "CircuitId", None)
                selected_ids = [fallback_id] if fallback_id is not None else []

        self._active_circuit_id = selected_ids[0] if selected_ids else None
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def SelectedPreviewList_SelectionChanged(self, sender, args):
        if self._is_programmatic_preview_selection_update:
            return
        if self._selected_preview_list is None:
            return

        selected_item = self._selected_preview_list.SelectedItem
        selected_circuit_id = getattr(selected_item, "CircuitId", None)
        if selected_circuit_id is None:
            return
        if selected_circuit_id == self._active_circuit_id:
            return

        self._active_circuit_id = selected_circuit_id
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def IncrementSelectedCheckBox_Changed(self, sender, args):
        self._refresh_selected_preview_list()

    def AddToStringButton_Click(self, sender, args):
        if self._available_params is None:
            return
        selected = self._available_params.SelectedItem
        if selected is None:
            forms.alert("Select a parameter first.", exitscript=False)
            return
        self._append_token_to_active(TOKEN_PARAM, selected)

    def AddCustomFieldButton_Click(self, sender, args):
        if self._custom_field_combo is None:
            return
        custom_value = _raw_text(self._custom_field_combo.Text)
        if custom_value == u"":
            forms.alert("Enter a custom field value first.", exitscript=False)
            return
        self._ensure_combo_has_value(self._custom_field_combo, custom_value)
        self._append_token_to_active(TOKEN_CUSTOM, custom_value)

    def AddSeparatorButton_Click(self, sender, args):
        if self._separator_combo is None:
            return
        separator_value = _raw_text(self._separator_combo.Text)
        if separator_value == u"":
            forms.alert("Enter a separator first.", exitscript=False)
            return
        self._ensure_combo_has_value(self._separator_combo, separator_value)
        self._append_token_to_active(TOKEN_SEPARATOR, separator_value)

    def RemoveFromStringButton_Click(self, sender, args):
        state = self._get_active_state()
        if not state:
            return
        indices = self._get_selected_string_part_indices()
        if not indices:
            forms.alert("Select one or more string parts to remove.", exitscript=False)
            return

        tokens = state.get("tokens", [])
        for idx in reversed(indices):
            if 0 <= idx < len(tokens):
                tokens.pop(idx)
        state["tokens"] = tokens
        state["manual_override"] = False
        state["preview_text"] = self._compute_preview_from_tokens(self._active_circuit_id, tokens)
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def MoveUpButton_Click(self, sender, args):
        state = self._get_active_state()
        if not state or self._string_parts is None:
            return
        selected_index = self._string_parts.SelectedIndex
        if selected_index <= 0:
            return

        tokens = state.get("tokens", [])
        tokens[selected_index - 1], tokens[selected_index] = tokens[selected_index], tokens[selected_index - 1]
        state["tokens"] = tokens
        state["manual_override"] = False
        state["preview_text"] = self._compute_preview_from_tokens(self._active_circuit_id, tokens)
        self._refresh_active_circuit_views()
        self._string_parts.SelectedIndex = selected_index - 1
        self._refresh_selected_preview_list()

    def MoveDownButton_Click(self, sender, args):
        state = self._get_active_state()
        if not state or self._string_parts is None:
            return
        selected_index = self._string_parts.SelectedIndex
        tokens = state.get("tokens", [])
        if selected_index < 0 or selected_index >= len(tokens) - 1:
            return

        tokens[selected_index + 1], tokens[selected_index] = tokens[selected_index], tokens[selected_index + 1]
        state["tokens"] = tokens
        state["manual_override"] = False
        state["preview_text"] = self._compute_preview_from_tokens(self._active_circuit_id, tokens)
        self._refresh_active_circuit_views()
        self._string_parts.SelectedIndex = selected_index + 1
        self._refresh_selected_preview_list()

    def ClearStringButton_Click(self, sender, args):
        state = self._get_active_state()
        if not state:
            return
        state["tokens"] = []
        state["manual_override"] = False
        state["preview_text"] = u""
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()

    def ApplyToSelectedButton_Click(self, sender, args):
        selected_ids = self._get_selected_circuit_ids()
        if not selected_ids:
            forms.alert("Select one or more circuits first.", exitscript=False)
            return
        self._apply_tokens_to_targets(selected_ids)

    def ApplyToAllButton_Click(self, sender, args):
        self._apply_tokens_to_targets([row["circuit_id"] for row in self._row_data])

    def SaveTemplateButton_Click(self, sender, args):
        state = self._get_active_state()
        if not state:
            return
        tokens = _clone_tokens(state.get("tokens", []))
        if not tokens:
            forms.alert("Active circuit has no string parts to save.", exitscript=False)
            return

        template_name = u""
        if self._template_name_text is not None:
            template_name = _safe_text(self._template_name_text.Text)
        if not template_name and self._saved_template_combo is not None:
            template_name = _safe_text(self._saved_template_combo.Text)

        if not template_name:
            forms.alert("Enter a template name first.", exitscript=False)
            return

        self._templates[template_name] = tokens
        self._save_templates_to_config()
        self._populate_templates_combo()
        if self._saved_template_combo is not None:
            self._saved_template_combo.SelectedItem = template_name
        if self._template_name_text is not None:
            self._template_name_text.Text = template_name

    def LoadTemplateButton_Click(self, sender, args):
        if self._saved_template_combo is None:
            return
        template_name = _safe_text(self._saved_template_combo.Text)
        if not template_name:
            forms.alert("Select a saved template first.", exitscript=False)
            return
        if template_name not in self._templates:
            forms.alert("Template '{}' was not found.".format(template_name), exitscript=False)
            return

        state = self._get_active_state()
        if not state:
            return
        tokens = _clone_tokens(self._templates.get(template_name, []))
        state["tokens"] = tokens
        state["manual_override"] = False
        state["preview_text"] = self._compute_preview_from_tokens(self._active_circuit_id, tokens)
        self._refresh_active_circuit_views()
        self._refresh_selected_preview_list()
        if self._template_name_text is not None:
            self._template_name_text.Text = template_name

    def DeleteTemplateButton_Click(self, sender, args):
        if self._saved_template_combo is None:
            return
        template_name = _safe_text(self._saved_template_combo.Text)
        if not template_name:
            forms.alert("Select a saved template first.", exitscript=False)
            return
        if template_name not in self._templates:
            forms.alert("Template '{}' was not found.".format(template_name), exitscript=False)
            return

        del self._templates[template_name]
        self._save_templates_to_config()
        self._populate_templates_combo()
        if self._template_name_text is not None:
            self._template_name_text.Text = ""

    def PreviewTextBox_TextChanged(self, sender, args):
        if self._is_programmatic_preview_update:
            return
        state = self._get_active_state()
        if not state:
            return
        state["preview_text"] = _raw_text(self._preview_textbox.Text if self._preview_textbox else u"")
        state["manual_override"] = True
        self._refresh_selected_preview_list()

    def RenameButton_Click(self, sender, args):
        target_ids = self._get_selected_circuit_ids()
        if not target_ids:
            target_ids = [row["circuit_id"] for row in self._row_data]

        names = {}
        for cid in target_ids:
            names[cid] = _safe_text(self._get_preview_for_circuit(cid))

        if self._increment_checkbox is not None and bool(self._increment_checkbox.IsChecked):
            names = self._apply_increment_suffix(names)

        self.result_target_ids = target_ids
        self.result_names = names
        self.DialogResult = True
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


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

    ui_window = CircuitStringBuilderWindow(XAML_PATH, row_data)
    result = ui_window.show_dialog()
    if not result:
        script.exit()

    rows_by_id = dict((row["circuit_id"], row) for row in row_data)
    target_ids = ui_window.result_target_ids or []
    names_by_id = ui_window.result_names or {}
    if not target_ids:
        forms.alert("No circuits selected for rename.", exitscript=True)

    results = []
    renamed_count = 0
    skipped_count = 0
    unchanged_count = 0

    with revit.Transaction("Rename Circuits by Device Parameter"):
        for cid in target_ids:
            row = rows_by_id.get(cid)
            if not row:
                continue

            circuit = row["circuit"]
            previous_name = _safe_text(circuit.LoadName)
            new_name = _safe_text(names_by_id.get(cid))

            if not new_name:
                skipped_count += 1
                results.append([_circuit_ref(circuit), previous_name or "-", "-", "Skipped: Resulting name is empty."])
                continue

            if new_name == previous_name:
                unchanged_count += 1
                results.append([_circuit_ref(circuit), previous_name or "-", new_name, "No Change"])
                continue

            set_error = _set_circuit_name(circuit, new_name)
            if set_error:
                skipped_count += 1
                results.append([_circuit_ref(circuit), previous_name or "-", new_name, "Skipped: {}".format(set_error)])
                continue

            renamed_count += 1
            results.append([_circuit_ref(circuit), previous_name or "-", new_name, "Renamed"])

    output.print_md("### Rename Circuits by Device Parameter")
    output.print_md(
        "Processed {} circuit(s): {} renamed, {} skipped, {} unchanged.".format(
            len(target_ids), renamed_count, skipped_count, unchanged_count
        )
    )
    output.print_table(results, ["Circuit", "Previous Name", "New Name", "Status"])

    forms.alert(
        "Processed {} circuit(s).\nRenamed: {}\nSkipped: {}\nNo Change: {}".format(
            len(target_ids), renamed_count, skipped_count, unchanged_count
        ),
        title="Rename Circuits by Device Parameter",
        exitscript=False
    )


if __name__ == "__main__":
    main()
