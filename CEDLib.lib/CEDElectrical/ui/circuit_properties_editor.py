# -*- coding: utf-8 -*-
"""Edit Circuit Properties window + view-model used by Circuit Browser actions."""

import os

import Autodesk.Revit.DB.Electrical as DBE
from System.Windows import Visibility
from pyrevit import DB, forms

from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Model.circuit_settings import IsolatedGroundBehavior, NeutralBehavior
from CEDElectrical.refdata.ampacity_table import WIRE_AMPACITY_TABLE
from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE
from CEDElectrical.refdata.conduit_area_table import CONDUIT_AREA_TABLE, CONDUIT_SIZE_INDEX
from Snippets import revit_helpers
from UIClasses import resource_loader


def _idval(item):
    return int(revit_helpers.get_elementid_value(item))


def _lookup_param_value(element, name):
    if element is None:
        return None
    try:
        param = element.LookupParameter(name)
    except Exception:
        param = None
    if not param:
        return None
    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString()
        if st == DB.StorageType.Integer:
            return param.AsInteger()
        if st == DB.StorageType.Double:
            return param.AsDouble()
        if st == DB.StorageType.ElementId:
            return param.AsElementId()
    except Exception:
        return None
    return None


def _lookup_param_text(element, name, default_value=""):
    value = _lookup_param_value(element, name)
    if value is None:
        return default_value
    return str(value or "")


def _as_int(value, default_value=0):
    try:
        return int(round(float(value or 0)))
    except Exception:
        return int(default_value)


def _as_float(value, default_value=0.0):
    try:
        return float(value)
    except Exception:
        return float(default_value)


def _fmt_number(value, digits=1):
    try:
        numeric = float(value)
    except Exception:
        return "-"
    if digits <= 0:
        return str(int(round(numeric, 0)))
    return str(round(numeric, digits))


def _fmt_amp(value, digits=0):
    text = _fmt_number(value, digits)
    if text == "-":
        return "-"
    return "{} A".format(text)


def _derive_branch_type(circuit):
    branch_type = _lookup_param_text(circuit, "CKT_Circuit Type_CEDT", "").strip().upper()
    if branch_type:
        return branch_type
    ctype = getattr(circuit, "CircuitType", None)
    if ctype == DBE.CircuitType.Space:
        return "SPACE"
    if ctype == DBE.CircuitType.Spare:
        return "SPARE"
    return "BRANCH"


def _build_wire_size_options():
    order = []
    seen = set()
    for material_table in list(WIRE_AMPACITY_TABLE.values()):
        if not isinstance(material_table, dict):
            continue
        for temp_rows in list(material_table.values()):
            for size_value, _ in list(temp_rows or []):
                text = str(size_value or "").strip()
                if not text or text in seen:
                    continue
                seen.add(text)
                order.append(text)
    return order


def _build_conduit_type_options():
    types = []
    for family_groups in list(CONDUIT_AREA_TABLE.values()):
        if not isinstance(family_groups, dict):
            continue
        for name in list(family_groups.keys()):
            text = str(name or "").strip()
            if text and text not in types:
                types.append(text)
    types.sort()
    return types


def _build_temperature_options():
    values = set()
    for material_table in list(WIRE_AMPACITY_TABLE.values()):
        if not isinstance(material_table, dict):
            continue
        for temp_value in list(material_table.keys()):
            try:
                temp_int = int(str(temp_value or "").replace("C", "").strip())
            except Exception:
                continue
            if temp_int > 0:
                values.add(temp_int)
    return ["{} C".format(x) for x in sorted(values)]


def _build_insulation_options():
    values = set()
    for row in list(CONDUCTOR_AREA_TABLE.values()):
        if not isinstance(row, dict):
            continue
        for name in list((row.get("area") or {}).keys()):
            text = str(name or "").strip().upper()
            if text:
                values.add(text)
    return sorted(list(values))


APPLY_PARAM_TYPES = {
    "CKT_User Override_CED": "int",
    "CKT_Rating_CED": "double",
    "CKT_Frame_CED": "double",
    "CKT_Length Makeup_CED": "double",
    "CKT_Number of Sets_CED": "int",
    "CKT_Include Neutral_CED": "int",
    "CKT_Include Isolated Ground_CED": "int",
    "CKT_Wire Hot Size_CEDT": "str",
    "CKT_Wire Neutral Size_CEDT": "str",
    "CKT_Wire Ground Size_CEDT": "str",
    "CKT_Wire Isolated Ground Size_CEDT": "str",
    "Conduit Size_CEDT": "str",
    "Conduit Type_CEDT": "str",
    "Wire Material_CEDT": "str",
    "Wire Temperature Rating_CEDT": "str",
    "Wire Insulation_CEDT": "str",
}


class CircuitPropertiesListItem(object):
    def __init__(self, target_item):
        circuit = getattr(target_item, "circuit", None)
        self.target_item = target_item
        self.circuit = circuit
        self.circuit_id = _idval(getattr(circuit, "Id", None)) if circuit is not None else 0
        self.panel = str(getattr(target_item, "panel", "-") or "-")
        self.circuit_number = str(getattr(target_item, "circuit_number", "-") or "-")
        self.load_name = str(getattr(target_item, "load_name", "") or "")
        self.panel_ckt_text = "{} / {}".format(self.panel or "-", self.circuit_number or "-")
        self.branch_type = str(getattr(target_item, "branch_type", "") or _derive_branch_type(circuit))
        self.status_text = ""
        self.status_visibility = Visibility.Collapsed

    def set_pending(self, is_pending):
        pending = bool(is_pending)
        self.status_text = "Pending change" if pending else ""
        self.status_visibility = Visibility.Visible if pending else Visibility.Collapsed


class CircuitPropertyEditorViewModel(object):
    def __init__(self, targets, settings):
        self.settings = settings
        self.rows = []
        self.rows_by_id = {}
        self.circuits_by_id = {}
        self.base_values = {}
        self.value_overrides = {}
        self.toggle_overrides = {}
        self.preview_rows = {}

        self.wire_size_options = _build_wire_size_options()
        self.conduit_size_options = list(CONDUIT_SIZE_INDEX or [])
        self.conduit_type_options = _build_conduit_type_options()
        self.temperature_options = _build_temperature_options()
        self.insulation_options = _build_insulation_options()
        self.material_options = ["CU", "AL"]

        for target in list(targets or []):
            row = CircuitPropertiesListItem(target)
            circuit = getattr(row, "circuit", None)
            if circuit is None or row.circuit_id <= 0:
                continue
            self.rows.append(row)
            self.rows_by_id[row.circuit_id] = row
            self.circuits_by_id[row.circuit_id] = circuit
            self.base_values[row.circuit_id] = self._collect_base_values(circuit)
            self.preview_rows[row.circuit_id] = self._build_preview_row(row.circuit_id)
            row.set_pending(False)

    def _collect_base_values(self, circuit):
        rating_value = _lookup_param_value(circuit, "CKT_Rating_CED")
        if rating_value is None:
            try:
                rating_value = float(circuit.Rating)
            except Exception:
                rating_value = 0.0
        frame_value = _lookup_param_value(circuit, "CKT_Frame_CED")
        if frame_value is None:
            try:
                frame_value = float(circuit.Frame)
            except Exception:
                frame_value = 0.0
        return {
            "CKT_User Override_CED": _as_int(_lookup_param_value(circuit, "CKT_User Override_CED"), 0),
            "CKT_Rating_CED": _as_float(rating_value, 0.0),
            "CKT_Frame_CED": _as_float(frame_value, 0.0),
            "CKT_Length Makeup_CED": _as_float(_lookup_param_value(circuit, "CKT_Length Makeup_CED"), 0.0),
            "CKT_Number of Sets_CED": _as_int(_lookup_param_value(circuit, "CKT_Number of Sets_CED"), 0),
            "CKT_Include Neutral_CED": _as_int(_lookup_param_value(circuit, "CKT_Include Neutral_CED"), 0),
            "CKT_Include Isolated Ground_CED": _as_int(_lookup_param_value(circuit, "CKT_Include Isolated Ground_CED"), 0),
            "CKT_Wire Hot Quantity_CED": _as_int(_lookup_param_value(circuit, "CKT_Wire Hot Quantity_CED"), 0),
            "CKT_Wire Hot Size_CEDT": _lookup_param_text(circuit, "CKT_Wire Hot Size_CEDT", ""),
            "CKT_Wire Neutral Quantity_CED": _as_int(_lookup_param_value(circuit, "CKT_Wire Neutral Quantity_CED"), 0),
            "CKT_Wire Neutral Size_CEDT": _lookup_param_text(circuit, "CKT_Wire Neutral Size_CEDT", ""),
            "CKT_Wire Ground Quantity_CED": _as_int(_lookup_param_value(circuit, "CKT_Wire Ground Quantity_CED"), 0),
            "CKT_Wire Ground Size_CEDT": _lookup_param_text(circuit, "CKT_Wire Ground Size_CEDT", ""),
            "CKT_Wire Isolated Ground Quantity_CED": _as_int(_lookup_param_value(circuit, "CKT_Wire Isolated Ground Quantity_CED"), 0),
            "CKT_Wire Isolated Ground Size_CEDT": _lookup_param_text(circuit, "CKT_Wire Isolated Ground Size_CEDT", ""),
            "Conduit Size_CEDT": _lookup_param_text(circuit, "Conduit Size_CEDT", ""),
            "Conduit Type_CEDT": _lookup_param_text(circuit, "Conduit Type_CEDT", ""),
            "Wire Material_CEDT": _lookup_param_text(circuit, "Wire Material_CEDT", ""),
            "Wire Temperature Rating_CEDT": _lookup_param_text(circuit, "Wire Temperature Rating_CEDT", ""),
            "Wire Insulation_CEDT": _lookup_param_text(circuit, "Wire Insulation_CEDT", ""),
        }

    def _default_toggle_state(self, circuit_id):
        base = dict(self.base_values.get(circuit_id, {}))
        hot_size = str(base.get("CKT_Wire Hot Size_CEDT", "") or "").strip()
        ground_size = str(base.get("CKT_Wire Ground Size_CEDT", "") or "").strip()
        conduit_size = str(base.get("Conduit Size_CEDT", "") or "").strip()
        allow_hot = hot_size != "-"
        allow_ground = allow_hot and ground_size != "-"
        allow_conduit = conduit_size != "-"
        include_neutral = bool(_as_int(base.get("CKT_Include Neutral_CED", 0), 0) == 1 or _as_int(base.get("CKT_Wire Neutral Quantity_CED", 0), 0) > 0)
        include_ig = bool(_as_int(base.get("CKT_Include Isolated Ground_CED", 0), 0) == 1 or _as_int(base.get("CKT_Wire Isolated Ground Quantity_CED", 0), 0) > 0)
        user_override = bool(_as_int(base.get("CKT_User Override_CED", 0), 0) == 1)
        return {
            "user_override": user_override,
            "allow_hot": allow_hot,
            "allow_ground": allow_ground,
            "allow_conduit": allow_conduit,
            "include_neutral": include_neutral,
            "include_ig": include_ig,
        }

    def _effective_toggles(self, circuit_id):
        state = dict(self._default_toggle_state(circuit_id))
        state.update(dict(self.toggle_overrides.get(circuit_id, {})))
        state["user_override"] = bool(state.get("user_override", False))
        state["allow_hot"] = bool(state.get("allow_hot", True))
        state["allow_ground"] = bool(state.get("allow_ground", True))
        state["allow_conduit"] = bool(state.get("allow_conduit", True))
        state["include_neutral"] = bool(state.get("include_neutral", True))
        state["include_ig"] = bool(state.get("include_ig", True))
        if not state["allow_hot"]:
            state["allow_ground"] = False
            state["include_neutral"] = False
            state["include_ig"] = False
        if not state["allow_ground"]:
            state["include_ig"] = False
        return state

    def _effective_values(self, circuit_id):
        values = dict(self.base_values.get(circuit_id, {}))
        values.update(dict(self.value_overrides.get(circuit_id, {})))
        toggles = self._effective_toggles(circuit_id)
        values["CKT_User Override_CED"] = 1 if toggles.get("user_override", False) else 0
        values["CKT_Include Neutral_CED"] = 1 if toggles.get("include_neutral", False) else 0
        values["CKT_Include Isolated Ground_CED"] = 1 if toggles.get("include_ig", False) else 0
        if toggles.get("user_override", False):
            if not toggles.get("allow_hot", True):
                values["CKT_Wire Hot Size_CEDT"] = "-"
                values["CKT_Wire Hot Quantity_CED"] = 0
                values["CKT_Number of Sets_CED"] = 0
                values["CKT_Wire Neutral Size_CEDT"] = "-"
                values["CKT_Wire Neutral Quantity_CED"] = 0
                values["CKT_Wire Ground Size_CEDT"] = "-"
                values["CKT_Wire Ground Quantity_CED"] = 0
                values["CKT_Wire Isolated Ground Size_CEDT"] = "-"
                values["CKT_Wire Isolated Ground Quantity_CED"] = 0
                values["CKT_Include Neutral_CED"] = 0
                values["CKT_Include Isolated Ground_CED"] = 0
            elif not toggles.get("allow_ground", True):
                values["CKT_Wire Ground Size_CEDT"] = "-"
                values["CKT_Wire Ground Quantity_CED"] = 0
                values["CKT_Wire Isolated Ground Size_CEDT"] = "-"
                values["CKT_Wire Isolated Ground Quantity_CED"] = 0
                values["CKT_Include Isolated Ground_CED"] = 0
            if not toggles.get("allow_conduit", True):
                values["Conduit Size_CEDT"] = "-"
        return values

    def _build_preview_inputs(self, circuit_id, toggles):
        preview_inputs = dict(self.value_overrides.get(circuit_id, {}))
        preview_inputs["CKT_User Override_CED"] = 1 if toggles.get("user_override", False) else 0
        preview_inputs["CKT_Include Neutral_CED"] = 1 if toggles.get("include_neutral", False) else 0
        preview_inputs["CKT_Include Isolated Ground_CED"] = 1 if toggles.get("include_ig", False) else 0

        if toggles.get("user_override", False):
            if not toggles.get("allow_hot", True):
                preview_inputs["CKT_Wire Hot Size_CEDT"] = "-"
                preview_inputs["CKT_Number of Sets_CED"] = 0
                preview_inputs["CKT_Wire Neutral Size_CEDT"] = "-"
                preview_inputs["CKT_Wire Ground Size_CEDT"] = "-"
                preview_inputs["CKT_Wire Isolated Ground Size_CEDT"] = "-"
                preview_inputs["CKT_Include Neutral_CED"] = 0
                preview_inputs["CKT_Include Isolated Ground_CED"] = 0
            elif not toggles.get("allow_ground", True):
                preview_inputs["CKT_Wire Ground Size_CEDT"] = "-"
                preview_inputs["CKT_Wire Isolated Ground Size_CEDT"] = "-"
                preview_inputs["CKT_Include Isolated Ground_CED"] = 0
            if not toggles.get("allow_conduit", True):
                preview_inputs["Conduit Size_CEDT"] = "-"

            explicit = set((self.value_overrides.get(circuit_id) or {}).keys())
            auto_recalc_params = (
                "CKT_Number of Sets_CED",
                "CKT_Wire Hot Size_CEDT",
                "CKT_Wire Neutral Size_CEDT",
                "CKT_Wire Ground Size_CEDT",
                "CKT_Wire Isolated Ground Size_CEDT",
                "Conduit Size_CEDT",
            )
            for param_name in auto_recalc_params:
                if param_name in explicit:
                    continue
                if param_name in preview_inputs and preview_inputs.get(param_name) not in (None, ""):
                    continue
                preview_inputs[param_name] = None

        return preview_inputs

    def _collect_notices(self, branch):
        notices = []
        collector = getattr(branch, "notices", None)
        if collector is None or not collector.has_items():
            return notices
        for _, severity, group, message in list(collector.items or []):
            label = "{} / {}".format(str(group or "Other"), str(severity or "NONE"))
            notices.append("{}: {}".format(label, str(message or "")))
        return notices

    def _build_preview_row(self, circuit_id):
        row = self.rows_by_id.get(circuit_id)
        circuit = self.circuits_by_id.get(circuit_id)
        if row is None or circuit is None:
            return None

        circuit_type = getattr(circuit, "CircuitType", None)
        system_type = getattr(circuit, "SystemType", None)
        is_supported = bool(
            system_type == DBE.ElectricalSystemType.PowerCircuit
            and circuit_type != DBE.CircuitType.Space
            and circuit_type != DBE.CircuitType.Spare
        )
        values = self._effective_values(circuit_id)
        toggles = self._effective_toggles(circuit_id)
        preview_inputs = self._build_preview_inputs(circuit_id, toggles)

        if not is_supported:
            return {
                "supported": False,
                "values": dict(values),
                "toggles": dict(toggles),
                "editability": {
                    "user_override": False,
                    "wire_manual_enabled": False,
                    "ground_manual_enabled": False,
                    "conduit_manual_enabled": False,
                    "include_neutral_enabled": False,
                    "include_ig_enabled": False,
                    "neutral_include_locked_by_type": False,
                    "neutral_size_manual_enabled": False,
                    "isolated_ground_size_manual_enabled": False,
                    "wire_specs_enabled": False,
                },
                "preview": {},
                "notices": ["Unsupported circuit type."],
            }

        branch = CircuitBranch(circuit, settings=self.settings, preview_values=preview_inputs)
        can_calculate = bool(branch.is_power_circuit and not branch.is_space and not branch.is_spare)
        if can_calculate:
            branch.calculate_hot_wire_size()
            branch.calculate_neutral_wire_size()
            branch.calculate_ground_wire_size()
            branch.calculate_isolated_ground_wire_size()
            branch.calculate_conduit_size()

        user_override = bool(getattr(branch, "_auto_calculate_override", False))
        hot_cleared = bool(getattr(branch, "_user_clear_hot", False))
        ground_cleared = bool(getattr(branch, "_user_clear_ground", False))
        conduit_cleared = bool(getattr(branch, "_user_clear_conduit", False))
        wire_manual_enabled = bool(user_override and not hot_cleared)
        ground_manual_enabled = bool(wire_manual_enabled and not ground_cleared)
        conduit_manual_enabled = bool(user_override and not conduit_cleared)
        hot_enabled = bool(toggles.get("allow_hot", False))
        ground_enabled = bool(hot_enabled and toggles.get("allow_ground", False))
        branch_type = str(getattr(branch, "branch_type", "") or "").upper()
        poles = _as_int(getattr(branch, "poles", 0), 0)
        neutral_locked_by_type = bool(
            poles == 1 or branch_type in ("FEEDER", "XFMR PRI", "XFMR SEC")
        )
        branch_neutral_qty = _as_int(getattr(branch, "neutral_wire_quantity", 0), 0)
        include_neutral_enabled = bool(hot_enabled and not neutral_locked_by_type)
        include_ig_enabled = bool(hot_enabled and ground_enabled)
        wire_specs_enabled = bool(hot_enabled)
        neutral_manual_enabled = bool(
            wire_manual_enabled
            and toggles.get("include_neutral", False)
            and include_neutral_enabled
            and self.settings.neutral_behavior == NeutralBehavior.MANUAL
        )
        ig_manual_enabled = bool(
            ground_manual_enabled
            and toggles.get("include_ig", False)
            and include_ig_enabled
            and self.settings.isolated_ground_behavior == IsolatedGroundBehavior.MANUAL
        )

        length_text = _fmt_number(getattr(branch, "length", None), 2)
        if length_text != "-":
            length_text = "{} ft".format(length_text)
        conduit_size_raw = str(getattr(getattr(branch, "conduit", None), "size", "") or "")
        if getattr(getattr(branch, "conduit", None), "cleared", False):
            conduit_size_raw = "-"
        hot_size_raw = str(getattr(getattr(branch, "cable", None), "hot_size", "") or "")
        neutral_size_raw = str(getattr(getattr(branch, "cable", None), "neutral_size", "") or "")
        ground_size_raw = str(getattr(getattr(branch, "cable", None), "ground_size", "") or "")
        ig_size_raw = str(getattr(getattr(branch, "cable", None), "ig_size", "") or "")
        if _as_int(getattr(branch, "neutral_wire_quantity", 0), 0) <= 0:
            neutral_size_raw = "-"
        if _as_int(getattr(branch, "ground_wire_quantity", 0), 0) <= 0:
            ground_size_raw = "-"
        if _as_int(getattr(branch, "isolated_ground_wire_quantity", 0), 0) <= 0:
            ig_size_raw = "-"

        voltage_drop_value = getattr(branch, "voltage_drop_percentage", None)
        if voltage_drop_value is None:
            voltage_drop_text = "-"
        else:
            voltage_drop_text = "{} %".format(_fmt_number(_as_float(voltage_drop_value, 0.0) * 100.0, 2))

        toggles_view = dict(toggles)
        if neutral_locked_by_type and branch_neutral_qty > 0:
            toggles_view["include_neutral"] = True

        preview_values = {
            "voltage_drop": voltage_drop_text,
            "load_current": _fmt_amp(getattr(branch, "circuit_load_current", None), 2),
            "ampacity": _fmt_amp(getattr(branch, "circuit_base_ampacity", None), 2),
            "total_length": length_text,
            "hot_qty": str(_as_int(getattr(branch, "hot_wire_quantity", 0), 0)),
            "neutral_qty": str(_as_int(getattr(branch, "neutral_wire_quantity", 0), 0)),
            "ground_qty": str(_as_int(getattr(branch, "ground_wire_quantity", 0), 0)),
            "ig_qty": str(_as_int(getattr(branch, "isolated_ground_wire_quantity", 0), 0)),
            "hot_size_raw": hot_size_raw,
            "neutral_size_raw": neutral_size_raw,
            "ground_size_raw": ground_size_raw,
            "ig_size_raw": ig_size_raw,
            "wire_material": str(getattr(branch, "wire_material", "") or ""),
            "wire_temp": str(getattr(branch, "wire_temp_rating", "") or ""),
            "wire_insulation": str(getattr(branch, "wire_insulation", "") or ""),
            "conduit_size_raw": conduit_size_raw,
            "conduit_type": str(getattr(branch, "conduit_type", "") or ""),
            "conduit_fill": "{} %".format(_fmt_number(_as_float(getattr(branch, "conduit_fill_percentage", 0.0), 0.0) * 100.0, 1)),
            "wire_summary": str(branch.get_wire_size_callout() or "-"),
            "conduit_summary": str(branch.get_conduit_and_wire_size() or "-"),
        }

        return {
            "supported": True,
            "values": dict(values),
            "toggles": dict(toggles_view),
            "editability": {
                "user_override": user_override,
                "wire_manual_enabled": wire_manual_enabled,
                "ground_manual_enabled": ground_manual_enabled,
                "conduit_manual_enabled": conduit_manual_enabled,
                "include_neutral_enabled": include_neutral_enabled,
                "include_ig_enabled": include_ig_enabled,
                "neutral_include_locked_by_type": neutral_locked_by_type,
                "neutral_size_manual_enabled": neutral_manual_enabled,
                "isolated_ground_size_manual_enabled": ig_manual_enabled,
                "wire_specs_enabled": wire_specs_enabled,
            },
            "preview": preview_values,
            "notices": self._collect_notices(branch),
        }

    def refresh_preview(self, circuit_id):
        self.preview_rows[circuit_id] = self._build_preview_row(circuit_id)
        row = self.rows_by_id.get(circuit_id)
        if row is not None:
            row.set_pending(self.has_pending_changes(circuit_id))

    def get_state(self, circuit_id):
        if circuit_id not in self.preview_rows:
            self.refresh_preview(circuit_id)
        return self.preview_rows.get(circuit_id)

    def set_toggle(self, circuit_id, toggle_name, value):
        state = dict(self.toggle_overrides.get(circuit_id, {}))
        state[str(toggle_name or "")] = bool(value)
        self.toggle_overrides[circuit_id] = state
        self.refresh_preview(circuit_id)

    def set_value(self, circuit_id, param_name, value):
        if circuit_id not in self.base_values:
            return
        key = str(param_name or "")
        if key not in APPLY_PARAM_TYPES:
            return
        value_type = APPLY_PARAM_TYPES[key]
        normalized = self._normalize_value(value, value_type)
        overrides = dict(self.value_overrides.get(circuit_id, {}))
        if self._values_equal(normalized, self.base_values[circuit_id].get(key), value_type):
            overrides.pop(key, None)
        else:
            overrides[key] = normalized
        if overrides:
            self.value_overrides[circuit_id] = overrides
        else:
            self.value_overrides.pop(circuit_id, None)
        self.refresh_preview(circuit_id)

    def apply_inputs(self, circuit_id, toggle_updates, value_updates):
        state = dict(self.toggle_overrides.get(circuit_id, {}))
        state.update(dict(toggle_updates or {}))
        self.toggle_overrides[circuit_id] = state

        overrides = dict(self.value_overrides.get(circuit_id, {}))
        for key, value in list(dict(value_updates or {}).items()):
            key = str(key or "")
            if key not in APPLY_PARAM_TYPES:
                continue
            value_type = APPLY_PARAM_TYPES[key]
            normalized = self._normalize_value(value, value_type)
            if self._values_equal(normalized, self.base_values[circuit_id].get(key), value_type):
                overrides.pop(key, None)
            else:
                overrides[key] = normalized
        if overrides:
            self.value_overrides[circuit_id] = overrides
        else:
            self.value_overrides.pop(circuit_id, None)
        self.refresh_preview(circuit_id)

    def _normalize_value(self, value, value_type):
        if value_type == "int":
            return _as_int(value, 0)
        if value_type == "double":
            return _as_float(value, 0.0)
        return str(value or "")

    def _values_equal(self, left, right, value_type):
        if value_type == "int":
            return _as_int(left, 0) == _as_int(right, 0)
        if value_type == "double":
            return abs(_as_float(left, 0.0) - _as_float(right, 0.0)) < 0.000001
        return str(left or "") == str(right or "")

    def _build_diff(self, circuit_id):
        base = dict(self.base_values.get(circuit_id, {}))
        effective = dict(self._effective_values(circuit_id))
        diff = {}
        for key, value_type in list(APPLY_PARAM_TYPES.items()):
            left = effective.get(key)
            right = base.get(key)
            if self._values_equal(left, right, value_type):
                continue
            diff[key] = self._normalize_value(left, value_type)
        return diff

    def has_pending_changes(self, circuit_id):
        return bool(self._build_diff(circuit_id))

    def pending_count(self):
        return len([x for x in list(self.rows or []) if self.has_pending_changes(x.circuit_id)])

    def build_apply_updates(self):
        updates = []
        for row in list(self.rows or []):
            circuit_id = row.circuit_id
            diff = self._build_diff(circuit_id)
            if not diff:
                row.set_pending(False)
                continue
            row.set_pending(True)
            updates.append(
                {
                    "circuit_id": int(circuit_id),
                    "param_values": diff,
                }
            )
        return updates


class CircuitPropertiesEditorWindow(forms.WPFWindow):
    def __init__(
        self,
        xaml_path,
        targets,
        settings,
        theme_mode="light",
        accent_mode="blue",
        resources_root=None,
    ):
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        self._resources_root = resources_root
        self._is_loading = False
        self.apply_requested = False
        self.apply_payload = {}
        self._vm = CircuitPropertyEditorViewModel(targets, settings=settings)

        forms.WPFWindow.__init__(self, os.path.abspath(xaml_path))
        self._apply_theme()

        self._list = self.FindName("CircuitList")
        self._panel = self.FindName("EditorPanelRoot")
        self._selected_text = self.FindName("SelectedCircuitText")
        self._pending_count = self.FindName("PendingCountText")
        self._status_text = self.FindName("StatusText")
        self._unsupported_text = self.FindName("UnsupportedText")

        self._user_override = self.FindName("UserOverrideCheck")
        self._hot_custom = self.FindName("CustomHotCheck")
        self._ground_custom = self.FindName("CustomGroundCheck")
        self._conduit_custom = self.FindName("CustomConduitCheck")
        self._include_neutral = self.FindName("IncludeNeutralCheck")
        self._include_ig = self.FindName("IncludeIGCheck")

        self._rating_tb = self.FindName("RatingTextBox")
        self._frame_tb = self.FindName("FrameTextBox")
        self._length_tb = self.FindName("LengthMakeupTextBox")
        self._total_length_text = self.FindName("TotalLengthText")
        self._num_sets_tb = self.FindName("NumberSetsTextBox")
        self._hot_qty_text = self.FindName("HotQtyValueText")
        self._neutral_qty_text = self.FindName("NeutralQtyValueText")
        self._ground_qty_text = self.FindName("GroundQtyValueText")
        self._ig_qty_text = self.FindName("IgQtyValueText")

        self._hot_size_cb = self.FindName("HotSizeCombo")
        self._neutral_size_cb = self.FindName("NeutralSizeCombo")
        self._ground_size_cb = self.FindName("GroundSizeCombo")
        self._ig_size_cb = self.FindName("IgSizeCombo")
        self._conduit_size_cb = self.FindName("ConduitSizeCombo")
        self._conduit_type_cb = self.FindName("ConduitTypeCombo")
        self._wire_temp_cb = self.FindName("WireTempCombo")
        self._wire_insulation_cb = self.FindName("WireInsulationCombo")
        self._wire_material_cu = self.FindName("WireMaterialCuRadio")
        self._wire_material_al = self.FindName("WireMaterialAlRadio")

        self._vd_text = self.FindName("PreviewVoltageDropText")
        self._load_text = self.FindName("PreviewLoadCurrentText")
        self._ampacity_text = self.FindName("PreviewAmpacityText")
        self._wire_summary_text = self.FindName("PreviewWireSummaryText")
        self._conduit_summary_text = self.FindName("PreviewConduitSummaryText")
        self._conduit_fill_text = self.FindName("ConduitFillText")
        self._notice_text = self.FindName("PreviewNoticeText")
        self._neutral_behavior_text = self.FindName("NeutralBehaviorText")
        self._ig_behavior_text = self.FindName("IgBehaviorText")

        if self._hot_size_cb is not None:
            self._hot_size_cb.ItemsSource = list(self._vm.wire_size_options or [])
        if self._neutral_size_cb is not None:
            self._neutral_size_cb.ItemsSource = list(self._vm.wire_size_options or [])
        if self._ground_size_cb is not None:
            self._ground_size_cb.ItemsSource = list(self._vm.wire_size_options or [])
        if self._ig_size_cb is not None:
            self._ig_size_cb.ItemsSource = list(self._vm.wire_size_options or [])
        if self._conduit_size_cb is not None:
            self._conduit_size_cb.ItemsSource = list(self._vm.conduit_size_options or [])
        if self._conduit_type_cb is not None:
            self._conduit_type_cb.ItemsSource = list(self._vm.conduit_type_options or [])
        if self._wire_temp_cb is not None:
            self._wire_temp_cb.ItemsSource = list(self._vm.temperature_options or [])
        if self._wire_insulation_cb is not None:
            self._wire_insulation_cb.ItemsSource = list(self._vm.insulation_options or [])

        if self._list is not None:
            self._list.ItemsSource = list(self._vm.rows or [])
            if self._vm.rows:
                self._list.SelectedItem = self._vm.rows[0]
        self._refresh_pending_count()
        self._refresh_behavior_text()
        self._load_selected_row()

    def _apply_theme(self):
        resource_loader.apply_theme(
            self,
            resources_root=self._resources_root,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )

    def _refresh_behavior_text(self):
        if self._neutral_behavior_text is not None:
            if self._vm.settings.neutral_behavior == NeutralBehavior.MANUAL:
                self._neutral_behavior_text.Text = "Neutral size: manual"
            else:
                self._neutral_behavior_text.Text = "Neutral size: matches hot (global setting)"
        if self._ig_behavior_text is not None:
            if self._vm.settings.isolated_ground_behavior == IsolatedGroundBehavior.MANUAL:
                self._ig_behavior_text.Text = "IG size: manual"
            else:
                self._ig_behavior_text.Text = "IG size: matches ground (global setting)"

    def _selected_row(self):
        if self._list is None:
            return None
        try:
            return getattr(self._list, "SelectedItem", None)
        except Exception:
            return None

    def _set_combo_value(self, combo, value):
        if combo is None:
            return
        target = str(value or "").strip()
        if not target:
            combo.SelectedIndex = -1
            return
        source = list(combo.ItemsSource or [])
        if target in source:
            combo.SelectedItem = target
            return

        norm_target = str(target or "").strip().upper()
        wire_prefix = str(getattr(self._vm.settings, "wire_size_prefix", "") or "").strip().upper()
        conduit_suffix = str(getattr(self._vm.settings, "conduit_size_suffix", "") or "").strip().upper()
        if wire_prefix and norm_target.startswith(wire_prefix):
            norm_target = norm_target[len(wire_prefix):].strip()
        if conduit_suffix and norm_target.endswith(conduit_suffix):
            norm_target = norm_target[:-len(conduit_suffix)].strip()
        has_temp_options = False
        for source_item in source:
            if str(source_item or "").strip().upper().endswith(" C"):
                has_temp_options = True
                break
        if has_temp_options and norm_target and norm_target.isdigit():
            norm_target = "{} C".format(norm_target)

        for source_item in source:
            item_text = str(source_item or "").strip()
            item_norm = item_text.upper()
            if wire_prefix and item_norm.startswith(wire_prefix):
                item_norm = item_norm[len(wire_prefix):].strip()
            if conduit_suffix and item_norm.endswith(conduit_suffix):
                item_norm = item_norm[:-len(conduit_suffix)].strip()
            if item_norm == norm_target:
                combo.SelectedItem = source_item
                return
        combo.SelectedIndex = -1

    def _parse_int_text(self, text, fallback):
        raw = str(text or "").strip()
        if not raw:
            return int(fallback)
        try:
            return int(round(float(raw)))
        except Exception:
            return int(fallback)

    def _parse_float_text(self, text, fallback):
        raw = str(text or "").strip()
        if not raw:
            return float(fallback)
        try:
            return float(raw)
        except Exception:
            return float(fallback)

    def _combo_text(self, combo, fallback):
        if combo is None:
            return str(fallback or "")
        selected = getattr(combo, "SelectedItem", None)
        if selected is None:
            return str(fallback or "")
        return str(selected or "")

    def _load_selected_row(self):
        row = self._selected_row()
        if row is None:
            if self._panel is not None:
                self._panel.IsEnabled = False
            if self._selected_text is not None:
                self._selected_text.Text = "Select a circuit"
            return

        self._is_loading = True
        try:
            state = self._vm.get_state(row.circuit_id) or {}
            values = dict(state.get("values") or {})
            toggles = dict(state.get("toggles") or {})
            editability = dict(state.get("editability") or {})
            preview = dict(state.get("preview") or {})
            notices = list(state.get("notices") or [])
            supported = bool(state.get("supported", False))

            if self._selected_text is not None:
                self._selected_text.Text = "{} - {}".format(row.panel_ckt_text, row.load_name or "-")
            if self._panel is not None:
                self._panel.IsEnabled = True
            if self._unsupported_text is not None:
                self._unsupported_text.Visibility = Visibility.Collapsed if supported else Visibility.Visible

            if self._user_override is not None:
                self._user_override.IsChecked = bool(toggles.get("user_override", False))
                self._user_override.IsEnabled = supported
            if self._hot_custom is not None:
                self._hot_custom.IsChecked = bool(toggles.get("allow_hot", True))
            if self._ground_custom is not None:
                self._ground_custom.IsChecked = bool(toggles.get("allow_ground", True))
            if self._conduit_custom is not None:
                self._conduit_custom.IsChecked = bool(toggles.get("allow_conduit", True))
            if self._include_neutral is not None:
                self._include_neutral.IsChecked = bool(toggles.get("include_neutral", False))
            if self._include_ig is not None:
                self._include_ig.IsChecked = bool(toggles.get("include_ig", False))

            self._rating_tb.Text = str(_as_int(values.get("CKT_Rating_CED", 0), 0))
            self._frame_tb.Text = str(_as_int(values.get("CKT_Frame_CED", 0), 0))
            self._length_tb.Text = _fmt_number(values.get("CKT_Length Makeup_CED", 0.0), 2)
            self._num_sets_tb.Text = str(_as_int(values.get("CKT_Number of Sets_CED", 0), 0))
            if self._total_length_text is not None:
                self._total_length_text.Text = str(preview.get("total_length", "-"))

            self._set_combo_value(self._hot_size_cb, preview.get("hot_size_raw", values.get("CKT_Wire Hot Size_CEDT", "")))
            self._set_combo_value(self._neutral_size_cb, preview.get("neutral_size_raw", values.get("CKT_Wire Neutral Size_CEDT", "")))
            self._set_combo_value(self._ground_size_cb, preview.get("ground_size_raw", values.get("CKT_Wire Ground Size_CEDT", "")))
            self._set_combo_value(self._ig_size_cb, preview.get("ig_size_raw", values.get("CKT_Wire Isolated Ground Size_CEDT", "")))
            self._set_combo_value(self._conduit_size_cb, preview.get("conduit_size_raw", values.get("Conduit Size_CEDT", "")))
            self._set_combo_value(self._conduit_type_cb, preview.get("conduit_type", values.get("Conduit Type_CEDT", "")))
            self._set_combo_value(self._wire_temp_cb, preview.get("wire_temp", values.get("Wire Temperature Rating_CEDT", "")))
            self._set_combo_value(self._wire_insulation_cb, preview.get("wire_insulation", values.get("Wire Insulation_CEDT", "")))

            wire_material_value = str(preview.get("wire_material", values.get("Wire Material_CEDT", "")) or "").strip().upper()
            if wire_material_value not in ("CU", "AL"):
                wire_material_value = str(values.get("Wire Material_CEDT", "CU") or "CU").strip().upper()
            if wire_material_value not in ("CU", "AL"):
                wire_material_value = "CU"
            if self._wire_material_cu is not None:
                self._wire_material_cu.IsChecked = bool(wire_material_value == "CU")
            if self._wire_material_al is not None:
                self._wire_material_al.IsChecked = bool(wire_material_value == "AL")

            wire_manual = bool(editability.get("wire_manual_enabled", False))
            ground_manual = bool(editability.get("ground_manual_enabled", False))
            conduit_manual = bool(editability.get("conduit_manual_enabled", False))
            include_neutral_enabled = bool(editability.get("include_neutral_enabled", False))
            include_ig_enabled = bool(editability.get("include_ig_enabled", False))
            neutral_locked_by_type = bool(editability.get("neutral_include_locked_by_type", False))
            neutral_manual = bool(editability.get("neutral_size_manual_enabled", False))
            ig_manual = bool(editability.get("isolated_ground_size_manual_enabled", False))
            wire_specs_enabled = bool(editability.get("wire_specs_enabled", False))

            if self._include_neutral is not None and neutral_locked_by_type:
                if _as_int(preview.get("neutral_qty", 0), 0) > 0:
                    self._include_neutral.IsChecked = True

            self._hot_custom.IsEnabled = bool(supported and self._user_override.IsChecked)
            self._ground_custom.IsEnabled = bool(supported and self._user_override.IsChecked and self._hot_custom.IsChecked)
            self._conduit_custom.IsEnabled = bool(supported and self._user_override.IsChecked)
            hot_checked = bool(self._hot_custom.IsChecked)
            ground_checked = bool(self._ground_custom.IsChecked)
            if self._include_neutral is not None:
                self._include_neutral.IsEnabled = bool(
                    supported
                    and include_neutral_enabled
                    and hot_checked
                    and not neutral_locked_by_type
                )
            if self._include_ig is not None:
                self._include_ig.IsEnabled = bool(
                    supported
                    and include_ig_enabled
                    and hot_checked
                    and ground_checked
                )

            self._rating_tb.IsEnabled = bool(supported)
            self._frame_tb.IsEnabled = bool(supported)
            self._length_tb.IsEnabled = bool(supported)

            self._num_sets_tb.IsEnabled = bool(supported and wire_manual and hot_checked)
            self._hot_size_cb.IsEnabled = bool(supported and wire_manual and hot_checked)
            neutral_checked = bool(self._include_neutral is not None and self._include_neutral.IsChecked)
            ig_checked = bool(self._include_ig is not None and self._include_ig.IsChecked)
            self._neutral_size_cb.IsEnabled = bool(
                supported
                and neutral_manual
                and hot_checked
                and neutral_checked
                and not neutral_locked_by_type
            )
            self._ground_size_cb.IsEnabled = bool(supported and ground_manual and hot_checked and ground_checked)
            self._ig_size_cb.IsEnabled = bool(supported and ig_manual and hot_checked and ground_checked and ig_checked)

            self._conduit_size_cb.IsEnabled = bool(supported and conduit_manual and self._conduit_custom.IsChecked)
            self._conduit_type_cb.IsEnabled = bool(supported and self._conduit_custom.IsChecked)
            if self._wire_material_cu is not None:
                self._wire_material_cu.IsEnabled = bool(supported and wire_specs_enabled)
            if self._wire_material_al is not None:
                self._wire_material_al.IsEnabled = bool(supported and wire_specs_enabled)
            if self._wire_temp_cb is not None:
                self._wire_temp_cb.IsEnabled = bool(supported and wire_specs_enabled)
            if self._wire_insulation_cb is not None:
                self._wire_insulation_cb.IsEnabled = bool(supported and wire_specs_enabled)

            if self._hot_qty_text is not None:
                self._hot_qty_text.Text = str(preview.get("hot_qty", "0"))
            if self._neutral_qty_text is not None:
                self._neutral_qty_text.Text = str(preview.get("neutral_qty", "0"))
            if self._ground_qty_text is not None:
                self._ground_qty_text.Text = str(preview.get("ground_qty", "0"))
            if self._ig_qty_text is not None:
                self._ig_qty_text.Text = str(preview.get("ig_qty", "0"))
            if self._conduit_fill_text is not None:
                self._conduit_fill_text.Text = str(preview.get("conduit_fill", "-"))
            self._vd_text.Text = str(preview.get("voltage_drop", "-"))
            self._load_text.Text = str(preview.get("load_current", "-"))
            self._ampacity_text.Text = str(preview.get("ampacity", "-"))
            self._wire_summary_text.Text = str(preview.get("wire_summary", "-"))
            self._conduit_summary_text.Text = str(preview.get("conduit_summary", "-"))
            self._notice_text.Text = "\n".join(notices) if notices else "No warnings."
        finally:
            self._is_loading = False
        self._refresh_pending_count()

    def _persist_selected_values(self, event_sender=None):
        if self._is_loading:
            return
        row = self._selected_row()
        if row is None:
            return
        state = self._vm.get_state(row.circuit_id) or {}
        current_values = dict(state.get("values") or {})
        cid = row.circuit_id
        existing_override_keys = set((self._vm.value_overrides.get(cid) or {}).keys())

        def _should_capture(param_name, control):
            if control is None:
                return False
            if not bool(control.IsEnabled):
                return False
            if event_sender is None:
                return True
            if event_sender is control:
                return True
            return param_name in existing_override_keys

        toggle_updates = {
            "user_override": bool(self._user_override.IsChecked),
            "allow_hot": bool(self._hot_custom.IsChecked),
            "allow_ground": bool(self._ground_custom.IsChecked),
            "allow_conduit": bool(self._conduit_custom.IsChecked),
            "include_neutral": bool(self._include_neutral.IsChecked) if self._include_neutral is not None else bool(current_values.get("CKT_Include Neutral_CED", 0)),
            "include_ig": bool(self._include_ig.IsChecked) if self._include_ig is not None else bool(current_values.get("CKT_Include Isolated Ground_CED", 0)),
        }

        value_updates = {}
        if bool(self._rating_tb.IsEnabled):
            value_updates["CKT_Rating_CED"] = self._parse_int_text(self._rating_tb.Text, current_values.get("CKT_Rating_CED", 0))
        if bool(self._frame_tb.IsEnabled):
            value_updates["CKT_Frame_CED"] = self._parse_int_text(self._frame_tb.Text, current_values.get("CKT_Frame_CED", 0))
        if bool(self._length_tb.IsEnabled):
            value_updates["CKT_Length Makeup_CED"] = self._parse_float_text(self._length_tb.Text, current_values.get("CKT_Length Makeup_CED", 0.0))
        if _should_capture("CKT_Number of Sets_CED", self._num_sets_tb):
            value_updates["CKT_Number of Sets_CED"] = self._parse_int_text(self._num_sets_tb.Text, current_values.get("CKT_Number of Sets_CED", 0))
        if _should_capture("CKT_Wire Hot Size_CEDT", self._hot_size_cb):
            value_updates["CKT_Wire Hot Size_CEDT"] = self._combo_text(self._hot_size_cb, current_values.get("CKT_Wire Hot Size_CEDT", ""))
        if _should_capture("CKT_Wire Neutral Size_CEDT", self._neutral_size_cb):
            value_updates["CKT_Wire Neutral Size_CEDT"] = self._combo_text(self._neutral_size_cb, current_values.get("CKT_Wire Neutral Size_CEDT", ""))
        if _should_capture("CKT_Wire Ground Size_CEDT", self._ground_size_cb):
            value_updates["CKT_Wire Ground Size_CEDT"] = self._combo_text(self._ground_size_cb, current_values.get("CKT_Wire Ground Size_CEDT", ""))
        if _should_capture("CKT_Wire Isolated Ground Size_CEDT", self._ig_size_cb):
            value_updates["CKT_Wire Isolated Ground Size_CEDT"] = self._combo_text(self._ig_size_cb, current_values.get("CKT_Wire Isolated Ground Size_CEDT", ""))
        if _should_capture("Conduit Size_CEDT", self._conduit_size_cb):
            value_updates["Conduit Size_CEDT"] = self._combo_text(self._conduit_size_cb, current_values.get("Conduit Size_CEDT", ""))
        if _should_capture("Conduit Type_CEDT", self._conduit_type_cb):
            value_updates["Conduit Type_CEDT"] = self._combo_text(self._conduit_type_cb, current_values.get("Conduit Type_CEDT", ""))
        if self._wire_material_cu is not None and self._wire_material_al is not None and bool(self._wire_material_cu.IsEnabled or self._wire_material_al.IsEnabled) and (
            event_sender is None or event_sender is self._wire_material_cu or event_sender is self._wire_material_al or "Wire Material_CEDT" in existing_override_keys
        ):
            value_updates["Wire Material_CEDT"] = "AL" if bool(self._wire_material_al.IsChecked) else "CU"
        if _should_capture("Wire Temperature Rating_CEDT", self._wire_temp_cb):
            value_updates["Wire Temperature Rating_CEDT"] = self._combo_text(
                self._wire_temp_cb, current_values.get("Wire Temperature Rating_CEDT", "")
            )
        if _should_capture("Wire Insulation_CEDT", self._wire_insulation_cb):
            value_updates["Wire Insulation_CEDT"] = self._combo_text(
                self._wire_insulation_cb, current_values.get("Wire Insulation_CEDT", "")
            )

        self._vm.apply_inputs(cid, toggle_updates, value_updates)

        try:
            if self._list is not None:
                self._list.Items.Refresh()
        except Exception:
            pass
        self._load_selected_row()

    def _refresh_pending_count(self):
        if self._pending_count is not None:
            self._pending_count.Text = "{} circuits pending".format(self._vm.pending_count())
        if self._status_text is not None:
            self._status_text.Text = "Ready"

    def circuit_selection_changed(self, sender, args):
        self._load_selected_row()

    def editable_value_changed(self, sender, args):
        self._persist_selected_values(event_sender=sender)

    def user_override_changed(self, sender, args):
        self._persist_selected_values(event_sender=sender)

    def hot_custom_changed(self, sender, args):
        if self._hot_custom is not None and bool(self._hot_custom.IsChecked):
            if self._ground_custom is not None and not bool(self._ground_custom.IsChecked):
                self._ground_custom.IsChecked = True
        self._persist_selected_values(event_sender=sender)

    def ground_custom_changed(self, sender, args):
        self._persist_selected_values(event_sender=sender)

    def conduit_custom_changed(self, sender, args):
        self._persist_selected_values(event_sender=sender)

    def include_neutral_changed(self, sender, args):
        self._persist_selected_values(event_sender=sender)

    def include_ig_changed(self, sender, args):
        self._persist_selected_values(event_sender=sender)

    def apply_clicked(self, sender, args):
        updates = list(self._vm.build_apply_updates() or [])
        if not updates:
            forms.alert("No staged circuit changes to apply.", title="Edit Circuit Properties")
            return
        self.apply_requested = True
        self.apply_payload = {"updates": updates}
        self.Close()

    def cancel_clicked(self, sender, args):
        self.apply_requested = False
        self.apply_payload = {}
        self.Close()
