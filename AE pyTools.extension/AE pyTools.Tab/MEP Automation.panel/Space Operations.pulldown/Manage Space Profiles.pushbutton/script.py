
# -*- coding: utf-8 -*-
"""
Manage Space Profiles
---------------------
Manage per-space-type element templates and per-space overrides for classified spaces.
"""

import os
import sys
from collections import OrderedDict
from datetime import datetime

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    FamilyInstance,
    Group,
    GroupType,
)

output = script.get_output()
output.close_others()

TITLE = "Manage Space Profiles"
CLASSIFICATION_STORAGE_ID = "space_operations.classifications.v1"
SPACE_PROFILE_SCHEMA_VERSION = 1

KEY_TYPE_ELEMENTS = "space_type_elements"
KEY_SPACE_OVERRIDES = "space_overrides"

BUCKETS = [
    "Restrooms",
    "Offices",
    "Sales Floor",
    "Freezers",
    "Coolers",
    "Receiving",
    "Break",
    "Food Prep",
    "Utility",
    "Storage",
    "Other",
]


PLACEMENT_OPTIONS = [
    "Ceiling Corner Furthest from door",
    "One Foot off doorway wall",
    "Center of Furthest wall",
    "Center Ceiling",
    "Center Floor",
    "Center of Room",
    "Ceiling Corner Nearest Door",
]

DEFAULT_PLACEMENT_OPTION = "Center of Room"
def _resolve_lib_root():
    cursor = os.path.abspath(os.path.dirname(__file__))
    for _ in range(12):
        candidate = os.path.join(cursor, "CEDLib.lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if not parent or parent == cursor:
            break
        cursor = parent
    return None


LIB_ROOT = _resolve_lib_root()
if LIB_ROOT and LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

try:
    from ExtensibleStorage import ExtensibleStorage  # noqa: E402
except Exception:
    ExtensibleStorage = None


def _element_id_value(elem_id, default=""):
    if elem_id is None:
        return default
    for attr in ("IntegerValue", "Value"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return str(int(value))
        except Exception:
            try:
                return str(value)
            except Exception:
                continue
    return default


def _param_text(element, built_in_param):
    if element is None:
        return ""
    try:
        param = element.get_Parameter(built_in_param)
    except Exception:
        param = None
    if not param:
        return ""
    for getter_name in ("AsString", "AsValueString"):
        try:
            getter = getattr(param, getter_name)
            value = getter()
        except Exception:
            value = None
        if value is not None:
            text = str(value).strip()
            if text:
                return text
    return ""


def _space_name(space):
    name = _param_text(space, BuiltInParameter.ROOM_NAME)
    if name:
        return name
    try:
        value = getattr(space, "Name", None)
    except Exception:
        value = None
    return str(value).strip() if value else ""


def _space_number(space):
    return _param_text(space, BuiltInParameter.ROOM_NUMBER)


def _collect_spaces(doc):
    try:
        collector = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_MEPSpaces)
            .WhereElementIsNotElementType()
        )
        spaces = list(collector)
    except Exception:
        spaces = []
    return spaces


def _make_space_key(space_id, unique_id):
    return (unique_id or "").strip() or (space_id or "").strip()


def _lookup_assignment(assignments, space_id, unique_id):
    if not isinstance(assignments, dict):
        return None

    uid = (unique_id or "").strip()
    sid = (space_id or "").strip()

    if uid and uid in assignments and isinstance(assignments.get(uid), dict):
        return assignments.get(uid)
    if sid and sid in assignments and isinstance(assignments.get(sid), dict):
        return assignments.get(sid)

    for value in assignments.values():
        if not isinstance(value, dict):
            continue
        entry_uid = str(value.get("unique_id") or "").strip()
        entry_sid = str(value.get("space_id") or "").strip()
        if uid and entry_uid and uid == entry_uid:
            return value
        if sid and entry_sid and sid == entry_sid:
            return value

    return None


def _collect_classified_spaces(doc, assignments):
    rows = []
    for space in _collect_spaces(doc):
        space_id = _element_id_value(getattr(space, "Id", None), default="")
        unique_id = ""
        try:
            unique_id = str(getattr(space, "UniqueId", "") or "").strip()
        except Exception:
            unique_id = ""

        assignment = _lookup_assignment(assignments, space_id, unique_id)
        bucket = "Other"
        if isinstance(assignment, dict):
            candidate = str(assignment.get("bucket") or "").strip()
            if candidate in BUCKETS:
                bucket = candidate

        number = _space_number(space)
        name = _space_name(space) or "<Unnamed Space>"
        space_key = _make_space_key(space_id, unique_id)
        if not space_key:
            continue

        rows.append(
            {
                "space_key": space_key,
                "space_id": space_id,
                "unique_id": unique_id,
                "space_number": number,
                "space_name": name,
                "bucket": bucket,
            }
        )

    rows.sort(key=lambda row: ((row.get("space_number") or "").lower(), (row.get("space_name") or "").lower()))
    return rows


def _param_value_to_text(param):
    if param is None:
        return ""
    for getter_name in ("AsString", "AsValueString"):
        try:
            getter = getattr(param, getter_name)
            value = getter()
        except Exception:
            value = None
        if value is not None:
            text = str(value).strip()
            if text:
                return text

    storage_type = ""
    try:
        storage_type = str(param.StorageType)
    except Exception:
        storage_type = ""

    try:
        if "Integer" in storage_type:
            return str(param.AsInteger())
        if "Double" in storage_type:
            return str(param.AsDouble())
        if "ElementId" in storage_type:
            return _element_id_value(param.AsElementId(), default="")
    except Exception:
        pass
    return ""


def _collect_available_parameters(doc, kind, element_type_id):
    result = {}
    target_id = str(element_type_id or "").strip()
    if doc is None or not target_id:
        return OrderedDict()

    instances = []
    if kind == "family_type":
        try:
            elements = list(FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType())
        except Exception:
            elements = []
        for inst in elements:
            symbol = getattr(inst, "Symbol", None)
            symbol_id = _element_id_value(getattr(symbol, "Id", None), default="")
            if symbol_id == target_id:
                instances.append(inst)
    elif kind == "model_group":
        try:
            elements = list(FilteredElementCollector(doc).OfClass(Group).WhereElementIsNotElementType())
        except Exception:
            elements = []
        for grp in elements:
            grp_type = getattr(grp, "GroupType", None)
            grp_type_id = _element_id_value(getattr(grp_type, "Id", None), default="")
            if grp_type_id == target_id:
                instances.append(grp)

    for inst in instances:
        try:
            params = list(inst.Parameters)
        except Exception:
            params = []

        for param in params:
            if param is None:
                continue

            try:
                if param.IsReadOnly:
                    continue
            except Exception:
                continue

            try:
                definition = param.Definition
                name = getattr(definition, "Name", None)
            except Exception:
                name = None
            if not name:
                continue

            key = str(name).strip()
            if not key:
                continue

            try:
                storage_type = str(param.StorageType)
            except Exception:
                storage_type = "String"

            value = _param_value_to_text(param)
            existing = result.get(key)
            if not existing:
                result[key] = {
                    "storage_type": storage_type,
                    "current_value": value,
                    "read_only": False,
                }
            else:
                if (not existing.get("current_value")) and value:
                    existing["current_value"] = value

    ordered = OrderedDict()
    for key in sorted(result.keys(), key=lambda x: x.lower()):
        ordered[key] = result[key]
    return ordered


def _family_type_name(symbol):
    family_name = ""
    type_name = ""
    try:
        family_name = str(getattr(getattr(symbol, "Family", None), "Name", "") or "").strip()
    except Exception:
        family_name = ""

    try:
        type_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if type_param:
            type_name = str(type_param.AsString() or "").strip()
    except Exception:
        type_name = ""

    if not type_name:
        try:
            type_name = str(getattr(symbol, "Name", "") or "").strip()
        except Exception:
            type_name = ""

    if family_name and type_name:
        return "{} : {}".format(family_name, type_name)
    return family_name or type_name or "Family Type"

def _collect_family_symbol_choices(doc):
    options = OrderedDict()
    try:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol).WhereElementIsElementType()
        symbols = list(collector)
    except Exception:
        symbols = []

    for symbol in symbols:
        type_id = _element_id_value(getattr(symbol, "Id", None), default="")
        if not type_id:
            continue
        display = _family_type_name(symbol)
        label = "{} [{}]".format(display, type_id)
        options[label] = symbol

    ordered = OrderedDict()
    for label in sorted(options.keys(), key=lambda x: x.lower()):
        ordered[label] = options[label]
    return ordered



def _param_string(param):
    if not param:
        return ""
    for getter_name in ("AsString", "AsValueString"):
        try:
            getter = getattr(param, getter_name)
            value = getter()
        except Exception:
            value = None
        if value is not None:
            text = str(value).strip()
            if text:
                return text
    return ""


def _model_group_type_name(group_type):
    # GroupType.Name is often just "Model Group"; prefer explicit type-name params.
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            name = _param_string(group_type.get_Parameter(bip))
        except Exception:
            name = ""
        if name and name.lower() != "model group":
            return name

    try:
        name = str(getattr(group_type, "Name", "") or "").strip()
    except Exception:
        name = ""
    if name:
        return name

    return "Model Group"
def _collect_model_group_type_choices(doc):
    options = OrderedDict()
    try:
        model_group_cat = ElementId(BuiltInCategory.OST_IOSModelGroups).IntegerValue
    except Exception:
        model_group_cat = None

    try:
        collector = FilteredElementCollector(doc).OfClass(GroupType).WhereElementIsElementType()
        group_types = list(collector)
    except Exception:
        group_types = []

    for group_type in group_types:
        if model_group_cat is not None:
            try:
                cat = getattr(group_type, "Category", None)
                cat_id = getattr(getattr(cat, "Id", None), "IntegerValue", None)
                if cat_id != model_group_cat:
                    continue
            except Exception:
                continue

        type_id = _element_id_value(getattr(group_type, "Id", None), default="")
        if not type_id:
            continue
        name = _model_group_type_name(group_type)
        label = "{} [{}]".format(name, type_id)
        options[label] = group_type

    ordered = OrderedDict()
    for label in sorted(options.keys(), key=lambda x: x.lower()):
        ordered[label] = options[label]
    return ordered


def _sanitize_parameter_map(parameters):
    clean = OrderedDict()
    if not isinstance(parameters, dict):
        return clean

    for name, data in parameters.items():
        key = str(name or "").strip()
        if not key:
            continue
        if isinstance(data, dict):
            storage_type = str(data.get("storage_type") or "String")
            value = data.get("value")
            read_only = bool(data.get("read_only"))
        else:
            storage_type = "String"
            value = data
            read_only = False
        clean[key] = {
            "storage_type": storage_type,
            "value": "" if value is None else str(value),
            "read_only": read_only,
        }

    ordered = OrderedDict()
    for key in sorted(clean.keys(), key=lambda x: x.lower()):
        ordered[key] = clean[key]
    return ordered


def _sanitize_template_entry(entry):
    if not isinstance(entry, dict):
        return None

    entry_id = str(entry.get("id") or "").strip()
    kind = str(entry.get("kind") or "").strip().lower()
    if kind not in ("family_type", "model_group"):
        return None

    element_type_id = str(entry.get("element_type_id") or "").strip()
    if not element_type_id:
        element_type_id = entry_id.split(":", 1)[-1] if ":" in entry_id else ""

    if not entry_id and element_type_id:
        entry_id = "{}:{}".format(kind, element_type_id)

    if not entry_id:
        return None

    name = str(entry.get("name") or "").strip()
    if not name:
        name = "Family Type" if kind == "family_type" else "Model Group"

    placement_rule = str(entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION).strip()
    if placement_rule not in PLACEMENT_OPTIONS:
        placement_rule = DEFAULT_PLACEMENT_OPTION

    return {
        "id": entry_id,
        "kind": kind,
        "element_type_id": element_type_id,
        "name": name,
        "placement_rule": placement_rule,
        "parameters": _sanitize_parameter_map(entry.get("parameters") or {}),
    }


def _sanitize_template_list(raw_list):
    clean = []
    seen = set()
    for raw in raw_list or []:
        entry = _sanitize_template_entry(raw)
        if not entry:
            continue
        key = entry.get("id")
        if key in seen:
            continue
        seen.add(key)
        clean.append(entry)
    return clean


def _sanitize_type_elements(raw_map):
    data = {}
    if not isinstance(raw_map, dict):
        raw_map = {}
    for bucket in BUCKETS:
        data[bucket] = _sanitize_template_list(raw_map.get(bucket) or [])
    return data


def _sanitize_space_overrides(raw_map):
    data = {}
    if not isinstance(raw_map, dict):
        return data

    for space_key, entries in raw_map.items():
        key = str(space_key or "").strip()
        if not key:
            continue
        sanitized = _sanitize_template_list(entries or [])
        if sanitized:
            data[key] = sanitized
    return data


def _template_display(entry):
    if not isinstance(entry, dict):
        return "<Invalid>"
    name = entry.get("name") or "<Unnamed>"
    kind = entry.get("kind") or "unknown"
    kind_text = "Family Type" if kind == "family_type" else "Model Group"
    params = entry.get("parameters") or {}
    count = len(params)
    suffix = "{} param{}".format(count, "" if count == 1 else "s")
    placement = str(entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION)
    if placement not in PLACEMENT_OPTIONS:
        placement = DEFAULT_PLACEMENT_OPTION
    return "{} [{} | {} | Place: {}]".format(name, kind_text, suffix, placement)


def _clone_template_entry(entry):
    sanitized = _sanitize_template_entry(entry)
    if not sanitized:
        return None
    return {
        "id": sanitized.get("id"),
        "kind": sanitized.get("kind"),
        "element_type_id": sanitized.get("element_type_id"),
        "name": sanitized.get("name"),
        "placement_rule": sanitized.get("placement_rule") or DEFAULT_PLACEMENT_OPTION,
        "parameters": _sanitize_parameter_map(sanitized.get("parameters") or {}),
    }


def _upsert_template_entry(target_list, entry):
    if target_list is None:
        return "skipped"
    candidate = _clone_template_entry(entry)
    if not candidate:
        return "skipped"

    key = candidate.get("id")
    for idx, existing in enumerate(target_list):
        if isinstance(existing, dict) and existing.get("id") == key:
            target_list[idx] = candidate
            return "updated"

    target_list.append(candidate)
    return "added"


def _choose_single_option(options, title, button_name="Select"):
    labels = list(options.keys())
    if not labels:
        return None
    selection = forms.SelectFromList.show(
        labels,
        title=title,
        button_name=button_name,
        multiselect=False,
    )
    if not selection:
        return None
    chosen = selection[0] if isinstance(selection, list) else selection
    return options.get(chosen)


def _prompt_for_placement_rule(initial_rule=None, title="Select Placement Rule"):
    current = str(initial_rule or DEFAULT_PLACEMENT_OPTION).strip()
    if current not in PLACEMENT_OPTIONS:
        current = DEFAULT_PLACEMENT_OPTION

    options = list(PLACEMENT_OPTIONS)
    selected = forms.SelectFromList.show(
        options,
        title=title,
        button_name="Use",
        multiselect=False,
    )
    if not selected:
        return None

    choice = selected[0] if isinstance(selected, list) else selected
    choice = str(choice or "").strip()
    if choice not in PLACEMENT_OPTIONS:
        return None
    return choice


class ParameterEditorWindow(forms.WPFWindow):
    def __init__(self, xaml_path, element_name, available_params, initial_params=None):
        forms.WPFWindow.__init__(self, xaml_path)
        self.accepted = False

        self.available_params = OrderedDict(available_params or {})
        self.selected_params = _sanitize_parameter_map(initial_params or {})

        prompt = self.FindName("PromptText")
        if prompt is not None:
            prompt.Text = (
                "Choose parameter overrides for: {}\n"
                "Available editable instance parameters are listed."
            ).format(element_name or "Selected Element")

        self._available_combo = self.FindName("AvailableParamsCombo")
        self._selected_list = self.FindName("SelectedParamsList")
        self._value_text = self.FindName("ParamValueText")

        if self._available_combo is not None:
            self._available_combo.ItemsSource = list(self.available_params.keys())
            if self._available_combo.Items.Count > 0:
                self._available_combo.SelectedIndex = 0

        self._refresh_selected_params_list()

    def _selected_param_name(self):
        if self._selected_list is None:
            return None
        idx = int(getattr(self._selected_list, "SelectedIndex", -1))
        if idx < 0 or idx >= len(self._selected_param_names):
            return None
        return self._selected_param_names[idx]

    def _refresh_selected_params_list(self):
        self._selected_param_names = list(self.selected_params.keys())
        labels = []
        for name in self._selected_param_names:
            data = self.selected_params.get(name) or {}
            read_only = bool(data.get("read_only"))
            labels.append(
                "{} [{}{}] = {}".format(
                    name,
                    data.get("storage_type") or "String",
                    " | RO" if read_only else "",
                    data.get("value") or "",
                )
            )
        if self._selected_list is not None:
            self._selected_list.ItemsSource = labels
            try:
                self._selected_list.Items.Refresh()
            except Exception:
                pass

        if self._value_text is not None:
            selected_name = self._selected_param_name()
            if selected_name and selected_name in self.selected_params:
                self._value_text.Text = self.selected_params[selected_name].get("value") or ""
            else:
                self._value_text.Text = ""

    def OnAddParamClicked(self, sender, args):
        if self._available_combo is None:
            return
        name = getattr(self._available_combo, "SelectedItem", None)
        if not name:
            return

        key = str(name)
        if key in self.selected_params:
            self._refresh_selected_params_list()
            return

        available = self.available_params.get(key) or {}
        self.selected_params[key] = {
            "storage_type": str(available.get("storage_type") or "String"),
            "value": str(available.get("current_value") or ""),
            "read_only": bool(available.get("read_only")),
        }
        self._refresh_selected_params_list()

    def OnSelectedParamChanged(self, sender, args):
        if self._value_text is None:
            return
        key = self._selected_param_name()
        if key and key in self.selected_params:
            self._value_text.Text = self.selected_params[key].get("value") or ""
        else:
            self._value_text.Text = ""

    def OnApplyValueClicked(self, sender, args):
        key = self._selected_param_name()
        if not key or key not in self.selected_params:
            return
        value = ""
        if self._value_text is not None:
            value = getattr(self._value_text, "Text", "") or ""
        self.selected_params[key]["value"] = str(value)
        self._refresh_selected_params_list()

    def OnRemoveParamClicked(self, sender, args):
        key = self._selected_param_name()
        if not key or key not in self.selected_params:
            return
        self.selected_params.pop(key, None)
        self._refresh_selected_params_list()

    def OnSaveClicked(self, sender, args):
        self.OnApplyValueClicked(sender, args)
        self.accepted = True
        self.Close()

    def OnCancelClicked(self, sender, args):
        self.accepted = False
        self.Close()


def _build_template_entry_from_element(element, kind):
    type_id = _element_id_value(getattr(element, "Id", None), default="")
    if not type_id:
        return None

    if kind == "family_type":
        name = _family_type_name(element)
    else:
        name = "Model Group: {}".format(_model_group_type_name(element))

    return {
        "id": "{}:{}".format(kind, type_id),
        "kind": kind,
        "element_type_id": type_id,
        "name": name,
        "placement_rule": DEFAULT_PLACEMENT_OPTION,
        "parameters": OrderedDict(),
    }


def _prompt_for_template_entry(doc, param_editor_xaml_path):
    mode = forms.CommandSwitchWindow.show(
        ["Family Type", "Model Group"],
        message="Select what to add",
    )
    if not mode:
        return None

    if mode == "Family Type":
        options = _collect_family_symbol_choices(doc)
        selected_element = _choose_single_option(options, title="Select Family Type", button_name="Use")
        kind = "family_type"
    else:
        options = _collect_model_group_type_choices(doc)
        selected_element = _choose_single_option(options, title="Select Model Group Type", button_name="Use")
        kind = "model_group"

    if selected_element is None:
        return None

    entry = _build_template_entry_from_element(selected_element, kind)
    if not entry:
        return None

    placement_rule = _prompt_for_placement_rule(
        entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION,
        title="Select Placement Rule",
    )
    if not placement_rule:
        return None
    entry["placement_rule"] = placement_rule

    wants_params = forms.alert(
        "Are there any specific parameters you would like to set for this family : type?",
        title=TITLE,
        yes=True,
        no=True,
    )

    if wants_params:
        available_params = _collect_available_parameters(doc, kind, entry.get("element_type_id"))
        if not available_params:
            forms.alert(
                "No editable instance parameters were found for this selection.\nPlace an instance of this type/group in the model and try again.",
                title=TITLE,
            )
        else:
            editor = ParameterEditorWindow(
                param_editor_xaml_path,
                entry.get("name"),
                available_params,
                initial_params=entry.get("parameters") or {},
            )
            editor.ShowDialog()
            if not editor.accepted:
                return None
            entry["parameters"] = _sanitize_parameter_map(editor.selected_params)

    return _sanitize_template_entry(entry)


def _edit_template_entry_parameters(doc, entry, param_editor_xaml_path):
    if not isinstance(entry, dict):
        return False

    available_params = _collect_available_parameters(doc, entry.get("kind"), entry.get("element_type_id"))
    if not available_params:
        forms.alert(
            "No editable instance parameters were found for this selection.\n"
            "Place an instance of this type/group in the model and try again.",
            title=TITLE,
        )
        return False

    existing = _sanitize_parameter_map(entry.get("parameters") or {})
    initial = OrderedDict()
    for name, data in existing.items():
        if name in available_params:
            initial[name] = data

    editor = ParameterEditorWindow(
        param_editor_xaml_path,
        entry.get("name") or "Selected Element",
        available_params,
        initial_params=initial,
    )
    editor.ShowDialog()
    if not editor.accepted:
        return False

    entry["parameters"] = _sanitize_parameter_map(editor.selected_params)
    return True


def _edit_template_entry_placement(entry):
    if not isinstance(entry, dict):
        return False

    current = entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION
    selected = _prompt_for_placement_rule(current, title="Edit Placement Rule")
    if not selected:
        return False

    entry["placement_rule"] = selected
    return True


class ManageSpaceProfilesWindow(forms.WPFWindow):
    def __init__(self, xaml_path, doc, spaces, type_elements, space_overrides, param_editor_xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)
        self.doc = doc
        self.spaces = list(spaces or [])
        self.type_elements = _sanitize_type_elements(type_elements)
        self.space_overrides = _sanitize_space_overrides(space_overrides)
        self.param_editor_xaml_path = param_editor_xaml_path

        self.accepted = False

        self._space_type_combo = self.FindName("SpaceTypeCombo")
        self._spaces_list = self.FindName("SpacesList")
        self._type_elements_list = self.FindName("TypeElementsList")
        self._effective_elements_list = self.FindName("EffectiveElementsList")
        self._selected_space_text = self.FindName("SelectedSpaceText")
        self._summary_text = self.FindName("SummaryText")

        self._visible_spaces = []
        self._visible_type_entries = []
        self._visible_effective_entries = []

        if self._space_type_combo is not None:
            self._space_type_combo.ItemsSource = list(BUCKETS)

            default_bucket = BUCKETS[0]
            for bucket in BUCKETS:
                if any((row.get("bucket") == bucket) for row in self.spaces):
                    default_bucket = bucket
                    break

            self._space_type_combo.SelectedItem = default_bucket

        self._refresh_all()

    def _selected_bucket(self):
        if self._space_type_combo is None:
            return BUCKETS[0]
        selected = getattr(self._space_type_combo, "SelectedItem", None)
        if selected in BUCKETS:
            return selected
        return BUCKETS[0]

    def _selected_space(self):
        if self._spaces_list is None:
            return None
        idx = int(getattr(self._spaces_list, "SelectedIndex", -1))
        if idx < 0 or idx >= len(self._visible_spaces):
            return None
        return self._visible_spaces[idx]

    def _selected_type_entry(self):
        if self._type_elements_list is None:
            return None
        idx = int(getattr(self._type_elements_list, "SelectedIndex", -1))
        if idx < 0 or idx >= len(self._visible_type_entries):
            return None
        return self._visible_type_entries[idx]

    def _selected_effective_entry(self):
        if self._effective_elements_list is None:
            return None
        idx = int(getattr(self._effective_elements_list, "SelectedIndex", -1))
        if idx < 0 or idx >= len(self._visible_effective_entries):
            return None
        return self._visible_effective_entries[idx]

    @staticmethod
    def _space_label(row):
        if not isinstance(row, dict):
            return "<Invalid Space>"
        number = row.get("space_number") or "<No Number>"
        name = row.get("space_name") or "<Unnamed Space>"
        return "{} - {}".format(number, name)

    def _refresh_spaces_panel(self):
        bucket = self._selected_bucket()
        self._visible_spaces = [row for row in self.spaces if row.get("bucket") == bucket]
        labels = [self._space_label(row) for row in self._visible_spaces]

        if self._spaces_list is not None:
            self._spaces_list.ItemsSource = labels
            try:
                self._spaces_list.Items.Refresh()
            except Exception:
                pass
            if labels:
                self._spaces_list.SelectedIndex = 0

    def _refresh_type_elements_panel(self):
        bucket = self._selected_bucket()
        self._visible_type_entries = self.type_elements.get(bucket) or []
        labels = [_template_display(entry) for entry in self._visible_type_entries]

        if self._type_elements_list is not None:
            self._type_elements_list.ItemsSource = labels
            try:
                self._type_elements_list.Items.Refresh()
            except Exception:
                pass

    def _refresh_effective_panel(self):
        selected_space = self._selected_space()
        if not selected_space:
            self._visible_effective_entries = []
            if self._selected_space_text is not None:
                self._selected_space_text.Text = "Selected Space: <none>"
            if self._effective_elements_list is not None:
                self._effective_elements_list.ItemsSource = []
            return

        if self._selected_space_text is not None:
            self._selected_space_text.Text = "Selected Space: {}".format(self._space_label(selected_space))

        bucket = self._selected_bucket()
        type_entries = self.type_elements.get(bucket) or []

        space_key = selected_space.get("space_key")
        override_entries = self.space_overrides.get(space_key) or []

        effective = OrderedDict()
        for entry in type_entries:
            effective[entry.get("id")] = {
                "source": "type",
                "entry": entry,
            }
        for entry in override_entries:
            effective[entry.get("id")] = {
                "source": "override",
                "entry": entry,
            }

        self._visible_effective_entries = [value for value in effective.values() if value.get("entry")]
        labels = []
        for wrapped in self._visible_effective_entries:
            source = wrapped.get("source")
            prefix = "[Type]" if source == "type" else "[Space Override]"
            labels.append("{} {}".format(prefix, _template_display(wrapped.get("entry"))))

        if self._effective_elements_list is not None:
            self._effective_elements_list.ItemsSource = labels
            try:
                self._effective_elements_list.Items.Refresh()
            except Exception:
                pass

    def _refresh_summary(self):
        if self._summary_text is None:
            return

        counts = OrderedDict((bucket, 0) for bucket in BUCKETS)
        for row in self.spaces:
            bucket = row.get("bucket")
            if bucket in counts:
                counts[bucket] += 1
            else:
                counts["Other"] += 1

        type_total = sum(len(self.type_elements.get(bucket) or []) for bucket in BUCKETS)
        override_total = sum(len(entries or []) for entries in self.space_overrides.values())

        lines = [
            "Classified spaces: {}".format(len(self.spaces)),
            "Type templates: {} | Space overrides: {}".format(type_total, override_total),
            "Current type: {}".format(self._selected_bucket()),
        ]
        for bucket in BUCKETS:
            if counts.get(bucket, 0) <= 0:
                continue
            lines.append("{:<12} {}".format(bucket + ":", counts[bucket]))
        self._summary_text.Text = "\n".join(lines)

    def _refresh_all(self):
        self._refresh_spaces_panel()
        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def _add_template_to_type(self):
        bucket = self._selected_bucket()
        entry = _prompt_for_template_entry(self.doc, self.param_editor_xaml_path)
        if not entry:
            return

        target = self.type_elements.setdefault(bucket, [])
        _upsert_template_entry(target, entry)
        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def _remove_template_from_type(self):
        bucket = self._selected_bucket()
        selected = self._selected_type_entry()
        if not selected:
            return

        target = self.type_elements.setdefault(bucket, [])
        key = selected.get("id")
        self.type_elements[bucket] = [entry for entry in target if entry.get("id") != key]
        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def _add_template_to_space(self):
        selected_space = self._selected_space()
        if not selected_space:
            forms.alert("Select a space in panel 1 first.", title=TITLE)
            return

        entry = _prompt_for_template_entry(self.doc, self.param_editor_xaml_path)
        if not entry:
            return

        space_key = selected_space.get("space_key")
        target = self.space_overrides.setdefault(space_key, [])
        _upsert_template_entry(target, entry)
        self._refresh_effective_panel()
        self._refresh_summary()

    def _remove_template_from_space(self):
        selected_space = self._selected_space()
        if not selected_space:
            forms.alert("Select a space in panel 1 first.", title=TITLE)
            return

        wrapped = self._selected_effective_entry()
        if not wrapped:
            return

        source = wrapped.get("source")
        entry = wrapped.get("entry") or {}
        entry_id = entry.get("id")
        if source != "override":
            forms.alert(
                "Type-level inherited items cannot be removed from panel 3.\n"
                "Remove them in panel 2 instead.",
                title=TITLE,
            )
            return

        space_key = selected_space.get("space_key")
        existing = self.space_overrides.get(space_key) or []
        updated = [item for item in existing if item.get("id") != entry_id]
        if updated:
            self.space_overrides[space_key] = updated
        else:
            self.space_overrides.pop(space_key, None)

        self._refresh_effective_panel()
        self._refresh_summary()

    def _edit_template_params_in_type(self):
        selected = self._selected_type_entry()
        if not selected:
            forms.alert("Select an element in panel 2 first.", title=TITLE)
            return

        changed = _edit_template_entry_parameters(self.doc, selected, self.param_editor_xaml_path)
        if not changed:
            return

        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def _edit_template_placement_in_type(self):
        selected = self._selected_type_entry()
        if not selected:
            forms.alert("Select an element in panel 2 first.", title=TITLE)
            return

        changed = _edit_template_entry_placement(selected)
        if not changed:
            return

        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def _edit_template_params_in_space(self):
        wrapped = self._selected_effective_entry()
        if not wrapped:
            forms.alert("Select an element in panel 3 first.", title=TITLE)
            return

        entry = wrapped.get("entry")
        if not isinstance(entry, dict):
            return

        changed = _edit_template_entry_parameters(self.doc, entry, self.param_editor_xaml_path)
        if not changed:
            return

        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def _edit_template_placement_in_space(self):
        wrapped = self._selected_effective_entry()
        if not wrapped:
            forms.alert("Select an element in panel 3 first.", title=TITLE)
            return

        entry = wrapped.get("entry")
        if not isinstance(entry, dict):
            return

        changed = _edit_template_entry_placement(entry)
        if not changed:
            return

        self._refresh_type_elements_panel()
        self._refresh_effective_panel()
        self._refresh_summary()

    def OnSpaceTypeChanged(self, sender, args):
        self._refresh_all()

    def OnSpaceSelectionChanged(self, sender, args):
        self._refresh_effective_panel()
        self._refresh_summary()

    def OnTypeAddClicked(self, sender, args):
        self._add_template_to_type()

    def OnTypeEditParamsClicked(self, sender, args):
        self._edit_template_params_in_type()

    def OnTypeEditPlacementClicked(self, sender, args):
        self._edit_template_placement_in_type()

    def OnTypeRemoveClicked(self, sender, args):
        self._remove_template_from_type()

    def OnSpaceAddClicked(self, sender, args):
        self._add_template_to_space()

    def OnSpaceEditParamsClicked(self, sender, args):
        self._edit_template_params_in_space()

    def OnSpaceEditPlacementClicked(self, sender, args):
        self._edit_template_placement_in_space()

    def OnSpaceRemoveClicked(self, sender, args):
        self._remove_template_from_space()

    def OnSaveClicked(self, sender, args):
        self.accepted = True
        self.Close()

    def OnCancelClicked(self, sender, args):
        self.accepted = False
        self.Close()

def _plain_parameter_map(parameters):
    out = {}
    for name, data in (parameters or {}).items():
        key = str(name or "").strip()
        if not key:
            continue
        if isinstance(data, dict):
            storage_type = str(data.get("storage_type") or "String")
            value = data.get("value")
            read_only = bool(data.get("read_only"))
        else:
            storage_type = "String"
            value = data
            read_only = False
        out[key] = {
            "storage_type": storage_type,
            "value": "" if value is None else str(value),
            "read_only": read_only,
        }
    return out


def _plain_template_entry(entry):
    return {
        "id": str(entry.get("id") or ""),
        "kind": str(entry.get("kind") or ""),
        "element_type_id": str(entry.get("element_type_id") or ""),
        "name": str(entry.get("name") or ""),
        "placement_rule": str(entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION),
        "parameters": _plain_parameter_map(entry.get("parameters") or {}),
    }


def _plain_type_elements(type_elements):
    result = {}
    for bucket in BUCKETS:
        entries = type_elements.get(bucket) or []
        result[bucket] = [_plain_template_entry(entry) for entry in entries]
    return result


def _plain_space_overrides(space_overrides):
    result = {}
    for space_key, entries in (space_overrides or {}).items():
        key = str(space_key or "").strip()
        if not key:
            continue
        cleaned = [_plain_template_entry(entry) for entry in (entries or [])]
        if cleaned:
            result[key] = cleaned
    return result


def _save_summary_lines(type_elements, space_overrides):
    type_total = sum(len(type_elements.get(bucket) or []) for bucket in BUCKETS)
    override_total = sum(len(entries or []) for entries in space_overrides.values())
    return [
        "Saved Manage Space Profiles data.",
        "Storage ID: {}".format(CLASSIFICATION_STORAGE_ID),
        "",
        "Type templates: {}".format(type_total),
        "Space override entries: {}".format(override_total),
    ]


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    if ExtensibleStorage is None:
        forms.alert("Failed to load ExtensibleStorage library from CEDLib.lib.", title=TITLE)
        return

    payload = ExtensibleStorage.get_project_data(doc, CLASSIFICATION_STORAGE_ID, default=None)
    if not isinstance(payload, dict):
        forms.alert(
            "No saved space classification data found.\n\n"
            "Run Classify Spaces and save first.",
            title=TITLE,
        )
        return

    assignments = payload.get("space_assignments")
    if not isinstance(assignments, dict) or not assignments:
        forms.alert(
            "No space assignments found in saved classification data.\n\n"
            "Run Classify Spaces and save first.",
            title=TITLE,
        )
        return

    spaces = _collect_classified_spaces(doc, assignments)
    if not spaces:
        forms.alert("No MEP spaces found in this model.", title=TITLE)
        return

    type_elements = _sanitize_type_elements(payload.get(KEY_TYPE_ELEMENTS) or {})
    space_overrides = _sanitize_space_overrides(payload.get(KEY_SPACE_OVERRIDES) or {})

    xaml_path = os.path.join(os.path.dirname(__file__), "ManageSpaceProfilesWindow.xaml")
    param_editor_xaml_path = os.path.join(os.path.dirname(__file__), "ParameterEditorWindow.xaml")
    if not os.path.exists(xaml_path) or not os.path.exists(param_editor_xaml_path):
        forms.alert("Manage Space Profiles XAML files are missing.", title=TITLE)
        return

    window = ManageSpaceProfilesWindow(
        xaml_path,
        doc,
        spaces,
        type_elements,
        space_overrides,
        param_editor_xaml_path,
    )
    window.ShowDialog()
    if not window.accepted:
        return

    payload["space_profile_schema_version"] = SPACE_PROFILE_SCHEMA_VERSION
    payload["space_profile_saved_utc"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    payload[KEY_TYPE_ELEMENTS] = _plain_type_elements(window.type_elements)
    payload[KEY_SPACE_OVERRIDES] = _plain_space_overrides(window.space_overrides)

    try:
        saved = ExtensibleStorage.set_project_data(
            doc,
            CLASSIFICATION_STORAGE_ID,
            payload,
            transaction_name="{} Save".format(TITLE),
        )
    except Exception as exc:
        forms.alert("Failed to save space profiles:\n\n{}".format(exc), title=TITLE)
        return

    if not saved:
        forms.alert("Space profiles were not saved.", title=TITLE)
        return

    forms.alert("\n".join(_save_summary_lines(window.type_elements, window.space_overrides)), title=TITLE)


if __name__ == "__main__":
    main()





















