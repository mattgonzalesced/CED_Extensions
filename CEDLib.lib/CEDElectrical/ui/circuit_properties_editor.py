# -*- coding: utf-8 -*-
"""Edit Circuit Properties window + view-model used by Circuit Browser actions."""

import os
import re

import Autodesk.Revit.DB.Electrical as DBE
from System.Windows import Visibility, Thickness
from System.Windows.Controls import Border, Control, ScrollViewer
from System.Windows.Media import VisualTreeHelper
from pyrevit import DB, forms

from CEDElectrical.Infrastructure.Revit.repositories.revit_circuit_repository import RevitCircuitRepository
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Model.circuit_settings import IsolatedGroundBehavior, NeutralBehavior
from CEDElectrical.refdata.ampacity_table import WIRE_AMPACITY_TABLE
from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE
from CEDElectrical.refdata.conduit_area_table import CONDUIT_AREA_TABLE, CONDUIT_SIZE_INDEX
from Snippets import revit_helpers
from Snippets.circuit_ui_actions import format_writeback_lock_reason
from UIClasses import resource_loader, ui_bases

CIRCUIT_NOTES_KEY = "__bip_circuit_notes__"
CIRCUIT_NAME_KEY = "__bip_circuit_name__"
UI_WIRE_SIZE_ORDER = [
    "12",
    "10",
    "8",
    "6",
    "4",
    "3",
    "2",
    "1",
    "1/0",
    "2/0",
    "3/0",
    "4/0",
    "250",
    "300",
    "350",
    "400",
    "500",
    "600",
    "700",
    "750",
    "800",
    "1000",
]

_WIRE_SIZE_ORDER_INDEX = dict((size, idx) for idx, size in enumerate(UI_WIRE_SIZE_ORDER))
_LOCK_REPOSITORY = RevitCircuitRepository()


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


def _lookup_builtin_text(element, bip, default_value=""):
    if element is None:
        return default_value
    try:
        param = element.get_Parameter(bip)
    except Exception:
        param = None
    if not param:
        return default_value
    try:
        return str(param.AsString() or "")
    except Exception:
        return default_value


def _lookup_builtin_numeric(element, bip):
    if element is None:
        return None
    try:
        param = element.get_Parameter(bip)
    except Exception:
        param = None
    if not param:
        return None
    try:
        st = param.StorageType
        if st == DB.StorageType.Double:
            return param.AsDouble()
        if st == DB.StorageType.Integer:
            return param.AsInteger()
    except Exception:
        return None
    return None


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
    found = set()
    extras = []
    seen = set()
    for material_table in list(WIRE_AMPACITY_TABLE.values()):
        if not isinstance(material_table, dict):
            continue
        for temp_rows in list(material_table.values()):
            for size_value, _ in list(temp_rows or []):
                text = str(size_value or "").strip()
                if text == "900":
                    continue
                if not text or text in seen:
                    continue
                seen.add(text)
                if text in UI_WIRE_SIZE_ORDER:
                    found.add(text)
                else:
                    extras.append(text)

    ordered = [x for x in UI_WIRE_SIZE_ORDER if x in found]
    if ordered:
        return ordered
    return extras


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


def _normalize_wire_size_text(value):
    text = str(value or "").strip().upper()
    if not text:
        return ""
    text = text.replace("#", "").replace(" AWG", "").strip()
    return text


def _wire_size_index(value):
    key = _normalize_wire_size_text(value)
    if not key:
        return -1
    return int(_WIRE_SIZE_ORDER_INDEX.get(key, -1))


def _is_wire_larger_than_limit(value, limit):
    value_idx = _wire_size_index(value)
    limit_idx = _wire_size_index(limit)
    if value_idx < 0 or limit_idx < 0:
        return False
    return bool(value_idx > limit_idx)


def _build_conduit_sizes_by_type():
    size_order = list(CONDUIT_SIZE_INDEX or [])
    rank = {}
    for idx, value in enumerate(size_order):
        rank[str(value or "").strip()] = int(idx)
    grouped = {}
    for family_groups in list(CONDUIT_AREA_TABLE.values()):
        if not isinstance(family_groups, dict):
            continue
        for conduit_type, sizes in list(family_groups.items()):
            key = str(conduit_type or "").strip()
            if not key or not isinstance(sizes, dict):
                continue
            bucket = grouped.setdefault(key, set())
            for conduit_size in list(sizes.keys()):
                size_text = str(conduit_size or "").strip()
                if size_text:
                    bucket.add(size_text)

    result = {}
    for key, values in list(grouped.items()):
        ordered = sorted(list(values), key=lambda x: rank.get(str(x or "").strip(), 9999))
        result[key] = ordered
    return result


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
    "Wire Temparature Rating_CEDT": "str",
    "Wire Insulation_CEDT": "str",
    CIRCUIT_NAME_KEY: "str",
    CIRCUIT_NOTES_KEY: "str",
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
        self.is_locked = False
        self.locked_owner = ""
        self.status_text = ""
        self.status_visibility = Visibility.Collapsed

    def set_locked(self, is_locked, owner=""):
        self.is_locked = bool(is_locked)
        self.locked_owner = str(owner or "")
        if self.is_locked:
            self.status_text = "Locked"
            self.status_visibility = Visibility.Visible
            return
        self.status_text = ""
        self.status_visibility = Visibility.Collapsed

    def set_pending(self, is_pending):
        if self.is_locked:
            self.status_text = "Locked"
            self.status_visibility = Visibility.Visible
            return
        pending = bool(is_pending)
        self.status_text = "Pending change" if pending else ""
        self.status_visibility = Visibility.Visible if pending else Visibility.Collapsed


class CircuitPropertyEditorViewModel(object):
    def __init__(self, targets, settings):
        self.settings = settings
        self.rows = []
        self.rows_by_id = {}
        self.circuits_by_id = {}
        self.locked_circuits = {}
        self.base_values = {}
        self.value_overrides = {}
        self.toggle_overrides = {}
        self.preview_rows = {}
        self.lock_rows_by_id = {}

        self.wire_size_options = _build_wire_size_options()
        self.conduit_size_options = list(CONDUIT_SIZE_INDEX or [])
        self.conduit_type_options = _build_conduit_type_options()
        self.conduit_sizes_by_type = _build_conduit_sizes_by_type()
        self.temperature_options = _build_temperature_options()
        self.insulation_options = _build_insulation_options()
        self.material_options = ["CU", "AL"]

        self._seed_lock_rows(targets)

        for target in list(targets or []):
            row = CircuitPropertiesListItem(target)
            circuit = getattr(row, "circuit", None)
            if circuit is None or row.circuit_id <= 0:
                continue
            self.rows.append(row)
            self.rows_by_id[row.circuit_id] = row
            self.circuits_by_id[row.circuit_id] = circuit
            is_locked, owner = self._worksharing_lock_info(circuit)
            if is_locked:
                self.locked_circuits[row.circuit_id] = str(owner or "")
            row.set_locked(is_locked, owner)
            self.base_values[row.circuit_id] = self._collect_base_values(circuit)
            self._seed_existing_override_inputs(row.circuit_id)
            self.preview_rows[row.circuit_id] = self._build_preview_row(row.circuit_id)
            row.set_pending(False)

    def _seed_lock_rows(self, targets):
        circuits = []
        for target in list(targets or []):
            circuit = getattr(target, "circuit", None)
            if circuit is not None:
                circuits.append(circuit)
        if not circuits:
            return
        try:
            doc = circuits[0].Document
        except Exception:
            doc = None
        if doc is None or not bool(getattr(doc, "IsWorkshared", False)):
            return
        try:
            # Match browser action windows lock evaluation behavior.
            _, _, locked_rows = _LOCK_REPOSITORY.partition_locked_elements(
                doc,
                circuits,
                self.settings,
                collect_all_device_owners=False,
            )
        except Exception:
            locked_rows = []
        lock_map = {}
        for row in list(locked_rows or []):
            try:
                cid = int((row or {}).get("circuit_id") or 0)
            except Exception:
                cid = 0
            if cid > 0:
                lock_map[cid] = dict(row or {})
        self.lock_rows_by_id = lock_map

    def _worksharing_lock_info(self, circuit):
        if circuit is None:
            return False, ""
        cid = 0
        try:
            cid = _idval(circuit.Id)
        except Exception:
            cid = 0
        lock_row = dict((self.lock_rows_by_id or {}).get(cid) or {})
        if lock_row:
            reason = format_writeback_lock_reason(lock_row)
            return True, "Blocked - {}".format(str(reason or "Blocked by ownership"))
        try:
            doc = circuit.Document
        except Exception:
            doc = None
        if doc is None or not bool(getattr(doc, "IsWorkshared", False)):
            return False, ""
        try:
            status = DB.WorksharingUtils.GetCheckoutStatus(doc, circuit.Id)
        except Exception:
            return False, ""
        if status != DB.CheckoutStatus.OwnedByOtherUser:
            return False, ""
        try:
            owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, circuit.Id).Owner or ""
        except Exception:
            owner = ""
        if owner:
            return True, "Blocked - Locked by {}".format(str(owner))
        return True, "Blocked - Locked by another user"

    def _collect_base_values(self, circuit):
        rating_value = _lookup_builtin_numeric(circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
        if rating_value is None:
            try:
                rating_value = float(circuit.Rating)
            except Exception:
                rating_value = 0.0
        frame_value = _lookup_builtin_numeric(circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM)
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
            "Wire Temparature Rating_CEDT": _lookup_param_text(circuit, "Wire Temparature Rating_CEDT", ""),
            "Wire Insulation_CEDT": _lookup_param_text(circuit, "Wire Insulation_CEDT", ""),
            CIRCUIT_NAME_KEY: _lookup_builtin_text(circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME, ""),
            CIRCUIT_NOTES_KEY: _lookup_builtin_text(circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM, ""),
            "__branch_type__": _derive_branch_type(circuit),
        }

    def _seed_existing_override_inputs(self, circuit_id):
        base = dict(self.base_values.get(circuit_id, {}))
        if _as_int(base.get("CKT_User Override_CED", 0), 0) != 1:
            return
        manual_seed_keys = (
            "CKT_Number of Sets_CED",
            "CKT_Wire Hot Size_CEDT",
            "CKT_Wire Neutral Size_CEDT",
            "CKT_Wire Ground Size_CEDT",
            "CKT_Wire Isolated Ground Size_CEDT",
            "Conduit Size_CEDT",
        )
        seeded = {}
        for key in manual_seed_keys:
            if key not in base:
                continue
            seeded[key] = base.get(key)
        if seeded:
            self.value_overrides[circuit_id] = seeded

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
        base = dict(self.base_values.get(circuit_id, {}))
        state = dict(self._default_toggle_state(circuit_id))
        state.update(dict(self.toggle_overrides.get(circuit_id, {})))
        state["user_override"] = bool(state.get("user_override", False))
        state["allow_hot"] = bool(state.get("allow_hot", True))
        state["allow_ground"] = bool(state.get("allow_ground", True))
        state["allow_conduit"] = bool(state.get("allow_conduit", True))
        state["include_neutral"] = bool(state.get("include_neutral", True))
        state["include_ig"] = bool(state.get("include_ig", True))
        branch_type = str(base.get("__branch_type__", "") or "").upper()
        if branch_type in ("FEEDER", "XFMR PRI", "XFMR SEC"):
            base_neutral = bool(
                _as_int(base.get("CKT_Include Neutral_CED", 0), 0) == 1
                or _as_int(base.get("CKT_Wire Neutral Quantity_CED", 0), 0) > 0
            )
            state["include_neutral"] = base_neutral
        if not state["allow_hot"]:
            state["allow_ground"] = False
            state["include_neutral"] = False
            state["include_ig"] = False
        if not state["allow_ground"]:
            state["include_ig"] = False
        return state

    def _effective_values(self, circuit_id):
        base_values = dict(self.base_values.get(circuit_id, {}))
        values = dict(base_values)
        values.update(dict(self.value_overrides.get(circuit_id, {})))
        toggles = self._effective_toggles(circuit_id)
        values["CKT_User Override_CED"] = 1 if toggles.get("user_override", False) else 0
        values["CKT_Include Neutral_CED"] = 1 if toggles.get("include_neutral", False) else 0
        values["CKT_Include Isolated Ground_CED"] = 1 if toggles.get("include_ig", False) else 0
        if not toggles.get("user_override", False):
            # Keep manual override values staged, but display calculated/base sets while override is off.
            values["CKT_Number of Sets_CED"] = base_values.get(
                "CKT_Number of Sets_CED",
                values.get("CKT_Number of Sets_CED", 0),
            )
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
        if not toggles.get("user_override", False):
            # Do not feed manual-set count into preview calculations while override is off.
            preview_inputs.pop("CKT_Number of Sets_CED", None)

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

    def _warning_flags_from_notices(self, notice_items, preview_values=None):
        items = list(notice_items or [])
        preview = dict(preview_values or {})
        lines = []
        alert_ids = set()
        lug_size_limit_value = ""

        for item in items:
            definition = None
            message = ""
            try:
                definition = item[0]
                message = str(item[3] or "")
            except Exception:
                definition = None
                message = ""
            lines.append(message.lower().strip())
            if definition is not None:
                try:
                    alert_ids.add(str(definition.GetId() or "").strip())
                except Exception:
                    pass
            if not lug_size_limit_value and message:
                match = re.search(r"\(([^)]+)\)", message)
                if match:
                    candidate = _normalize_wire_size_text(match.group(1))
                    if _wire_size_index(candidate) >= 0:
                        lug_size_limit_value = candidate

        if not lug_size_limit_value:
            lug_size_limit_value = _normalize_wire_size_text(preview.get("max_lug_size", ""))

        def _has_id(*keys):
            for key in list(keys or []):
                if str(key or "").strip() in alert_ids:
                    return True
            return False

        def _has_text(*tokens):
            for token in list(tokens or []):
                t = str(token or "").strip().lower()
                if not t:
                    continue
                for line in lines:
                    if t in line:
                        return True
            return False

        return {
            "length": _has_text("length", "wire length", "length makeup"),
            "insufficient_ampacity": _has_id("Design.InsufficientAmpacity") or _has_text("insufficient ampacity"),
            "insufficient_ampacity_breaker": _has_id("Design.InsufficientAmpacityBreaker") or _has_text("breaker ampacity"),
            "circuit_loads_null": _has_id("Design.CircuitLoadsNull") or _has_text("0 a load"),
            "circuit_panels_null": _has_id("Design.CircuitPanelsNull") or _has_text("not connected to a panel"),
            "lug_quantity_limit": _has_id(
                "Design.BreakerLugQuantityLimitOverride",
                "Calculations.BreakerLugQuantityLimit",
            ) or _has_text("lug quantity", "wire sets"),
            "excessive_voltage_drop": _has_id("Design.ExcessiveVoltDrop") or _has_text("voltage drop", "volt drop"),
            "excessive_conduit_fill": _has_id("Design.ExcessiveConduitFill") or _has_text("conduit fill"),
            "conduit_size_issue": _has_id(
                "Overrides.InvalidConduit",
                "Calculations.ConduitSizingFailed",
            ) or _has_text(
                "invalid conduit",
                "conduit sizing failed",
                "calculate conduit size",
            ),
            "lug_size_limit": _has_id(
                "Design.BreakerLugSizeLimitOverride",
                "Calculations.BreakerLugSizeLimit",
            ) or _has_text("lug size"),
            "lug_size_limit_value": lug_size_limit_value,
            "undersized_wire_egc": _has_id("Design.UndersizedWireEGC") or _has_text("undersized wire egc"),
            "undersized_wire_service_ground": _has_id("Design.UndersizedWireServiceGround") or _has_text("undersized wire service ground"),
            "rating": _has_id("Design.NonStandardOCPRating", "Design.UndersizedOCP") or _has_text("non-standard ampere rating", "undersized ocp"),
        }

    def _build_preview_row(self, circuit_id):
        row = self.rows_by_id.get(circuit_id)
        circuit = self.circuits_by_id.get(circuit_id)
        if row is None or circuit is None:
            return None
        if bool(row.is_locked):
            values = self._effective_values(circuit_id)
            return {
                "supported": True,
                "locked": True,
                "disabled_reason": str(row.locked_owner or "Blocked - Locked by another user"),
                "values": dict(values),
                "toggles": dict(self._effective_toggles(circuit_id)),
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
                "notices": [],
                "warning_flags": {},
            }

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
                "locked": False,
                "disabled_reason": "Unsupported circuit type.",
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
                "warning_flags": {
                    "length": False,
                    "insufficient_ampacity": False,
                    "insufficient_ampacity_breaker": False,
                    "circuit_loads_null": False,
                    "circuit_panels_null": False,
                    "lug_quantity_limit": False,
                    "excessive_voltage_drop": False,
                    "excessive_conduit_fill": False,
                    "conduit_size_issue": False,
                    "lug_size_limit": False,
                    "lug_size_limit_value": "",
                    "undersized_wire_egc": False,
                    "undersized_wire_service_ground": False,
                    "rating": False,
                },
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
        neutral_included = bool(
            toggles.get("include_neutral", False)
            or (neutral_locked_by_type and branch_neutral_qty > 0)
        )
        neutral_manual_enabled = bool(
            wire_manual_enabled
            and hot_enabled
            and neutral_included
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
        branch_voltage = _as_float(getattr(branch, "voltage", None), 0.0)
        if branch_voltage > 0 and poles > 0:
            volts_poles_text = "{} V / {}P".format(_fmt_number(branch_voltage, 0), poles)
        elif branch_voltage > 0:
            volts_poles_text = "{} V".format(_fmt_number(branch_voltage, 0))
        elif poles > 0:
            volts_poles_text = "{}P".format(poles)
        else:
            volts_poles_text = "-"

        toggles_view = dict(toggles)
        if neutral_locked_by_type and branch_neutral_qty > 0:
            toggles_view["include_neutral"] = True

        preview_values = {
            "circuit_type": str(getattr(branch, "branch_type", "") or "-"),
            "volts_poles": volts_poles_text,
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
            "max_lug_size": str(((getattr(branch, "_wire_info", {}) or {}).get("max_lug_size") or "")).strip(),
            "wire_summary": str(branch.get_wire_size_callout() or "-"),
            "conduit_summary": str(branch.get_conduit_and_wire_size() or "-"),
        }

        notice_items = list(getattr(getattr(branch, "notices", None), "items", []) or [])
        notices = self._collect_notices(branch)
        warnings = self._warning_flags_from_notices(notice_items, preview_values=preview_values)
        return {
            "supported": True,
            "locked": False,
            "disabled_reason": "",
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
            "notices": notices,
            "warning_flags": warnings,
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
        if circuit_id in self.locked_circuits:
            return
        state = dict(self.toggle_overrides.get(circuit_id, {}))
        state[str(toggle_name or "")] = bool(value)
        self.toggle_overrides[circuit_id] = state
        self.refresh_preview(circuit_id)

    def set_value(self, circuit_id, param_name, value):
        if circuit_id in self.locked_circuits:
            return
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
        if circuit_id in self.locked_circuits:
            return
        base = dict(self.base_values.get(circuit_id, {}))
        branch_type = str(base.get("__branch_type__", "") or "").upper()
        neutral_locked_by_type = bool(branch_type in ("FEEDER", "XFMR PRI", "XFMR SEC"))
        state = dict(self.toggle_overrides.get(circuit_id, {}))
        state.update(dict(toggle_updates or {}))
        if neutral_locked_by_type:
            base_neutral = bool(
                _as_int(base.get("CKT_Include Neutral_CED", 0), 0) == 1
                or _as_int(base.get("CKT_Wire Neutral Quantity_CED", 0), 0) > 0
            )
            state["include_neutral"] = base_neutral
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
        if neutral_locked_by_type:
            overrides.pop("CKT_Include Neutral_CED", None)
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
        if circuit_id in self.locked_circuits:
            return {}
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
            if circuit_id in self.locked_circuits:
                row.set_pending(False)
                continue
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

    def reset_all(self):
        self.value_overrides = {}
        self.toggle_overrides = {}
        for row in list(self.rows or []):
            self.refresh_preview(row.circuit_id)
            row.set_pending(False)

    def reset_circuit(self, circuit_id):
        cid = int(circuit_id or 0)
        if cid <= 0:
            return
        self.value_overrides.pop(cid, None)
        self.toggle_overrides.pop(cid, None)
        self.refresh_preview(cid)


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
        self._blocked_info_tb = self.FindName("BlockedInfoTextBox")
        self._reset_button = self.FindName("ResetButton")

        self._user_override = self.FindName("UserOverrideCheck")
        self._hot_custom = self.FindName("CustomHotCheck")
        self._ground_custom = self.FindName("CustomGroundCheck")
        self._conduit_custom = self.FindName("CustomConduitCheck")
        self._include_neutral = self.FindName("IncludeNeutralCheck")
        self._include_ig = self.FindName("IncludeIGCheck")

        self._rating_tb = self.FindName("RatingTextBox")
        self._frame_tb = self.FindName("FrameTextBox")
        self._length_tb = self.FindName("LengthMakeupTextBox")
        self._load_name_tb = self.FindName("LoadNameTextBox")
        self._volts_text = self.FindName("VoltsText")
        self._total_length_text = self.FindName("TotalLengthText")
        self._circuit_type_text = self.FindName("CircuitTypeText")
        self._sched_notes_tb = self.FindName("ScheduleNotesTextBox")
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
        self._vd_box = self.FindName("PreviewVoltageDropBox")
        self._load_text = self.FindName("PreviewLoadCurrentText")
        self._load_box = self.FindName("PreviewLoadCurrentBox")
        self._ampacity_box = self.FindName("PreviewAmpacityBox")
        self._ampacity_text = self.FindName("PreviewAmpacityText")
        self._wire_summary_text = self.FindName("PreviewWireSummaryText")
        self._conduit_summary_text = self.FindName("PreviewConduitSummaryText")
        self._conduit_fill_text = self.FindName("ConduitFillText")
        self._conduit_fill_box = self.FindName("ConduitFillBox")
        self._notice_text = self.FindName("PreviewNoticeText")
        self._notice_box = self.FindName("PreviewNoticeBox")
        self._neutral_behavior_text = self.FindName("NeutralBehaviorText")
        self._ig_behavior_text = self.FindName("IgBehaviorText")
        self._hot_include_state_text = self.FindName("HotIncludeStateText")
        self._ground_include_state_text = self.FindName("GroundIncludeStateText")

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
        self._bind_size_combo_wheel()

        if self._list is not None:
            self._list.ItemsSource = list(self._vm.rows or [])
            if self._vm.rows:
                self._list.SelectedItem = self._vm.rows[0]
        try:
            self.Loaded += self._on_loaded_bind_textbox_behaviors
        except Exception:
            pass
        self._on_loaded_bind_textbox_behaviors(None, None)
        self._refresh_pending_count()
        self._refresh_behavior_text()
        self._load_selected_row()

    def _sync_reset_button_state(self, row):
        if self._reset_button is None:
            return
        if row is None:
            self._reset_button.IsEnabled = False
            return
        cid = int(getattr(row, "circuit_id", 0) or 0)
        is_pending = bool(cid > 0 and self._vm.has_pending_changes(cid))
        self._reset_button.IsEnabled = bool(is_pending and not bool(getattr(row, "is_locked", False)))

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
        # Keep existing selection if no match is found; avoid clearing user-picked values.
        try:
            if getattr(combo, "SelectedItem", None) is None:
                combo.Text = target
        except Exception:
            pass

    def _conduit_size_options_for_type(self, conduit_type):
        text = str(conduit_type or "").strip()
        if text:
            values = list((self._vm.conduit_sizes_by_type or {}).get(text) or [])
            if values:
                return values
        return list(self._vm.conduit_size_options or [])

    def _refresh_conduit_size_items(self, conduit_type, preferred_size):
        if self._conduit_size_cb is None:
            return
        options = self._conduit_size_options_for_type(conduit_type)
        self._conduit_size_cb.ItemsSource = options
        target = str(preferred_size or "").strip()
        if target and target in options:
            self._conduit_size_cb.SelectedItem = target
            return
        if options:
            self._conduit_size_cb.SelectedItem = options[0]
            return
        self._conduit_size_cb.SelectedIndex = -1

    def _bind_size_combo_wheel(self):
        size_combos = (
            self._hot_size_cb,
            self._neutral_size_cb,
            self._ground_size_cb,
            self._ig_size_cb,
            self._conduit_size_cb,
        )
        for combo in size_combos:
            if combo is None:
                continue
            try:
                combo.PreviewMouseWheel += self._size_combo_preview_mouse_wheel
            except Exception:
                pass

    def _size_combo_preview_mouse_wheel(self, sender, args):
        combo = sender
        if combo is None or not bool(getattr(combo, "IsEnabled", False)):
            return
        if bool(getattr(combo, "IsDropDownOpen", False)):
            self._scroll_open_combo_popup(combo, getattr(args, "Delta", 0.0))
            args.Handled = True
            return
        items = list(getattr(combo, "ItemsSource", None) or [])
        if not items:
            args.Handled = True
            return
        try:
            idx = int(getattr(combo, "SelectedIndex", -1))
        except Exception:
            idx = -1
        if idx < 0:
            idx = 0

        step = -1 if float(getattr(args, "Delta", 0.0)) > 0 else 1
        new_idx = idx + step
        if new_idx < 0:
            new_idx = 0
        if new_idx >= len(items):
            new_idx = len(items) - 1
        if new_idx != idx:
            combo.SelectedIndex = int(new_idx)
        args.Handled = True

    def _scroll_open_combo_popup(self, combo, delta):
        popup = None
        try:
            template = getattr(combo, "Template", None)
            if template is not None:
                popup = template.FindName("PART_Popup", combo)
        except Exception:
            popup = None
        popup_child = getattr(popup, "Child", None) if popup is not None else None
        viewer = self._find_visual_descendant(popup_child, ScrollViewer)
        if viewer is None:
            return
        try:
            if float(delta or 0.0) > 0:
                viewer.LineUp()
            else:
                viewer.LineDown()
        except Exception:
            pass

    def _find_visual_descendant(self, root, target_type):
        if root is None:
            return None
        try:
            if isinstance(root, target_type):
                return root
        except Exception:
            pass
        try:
            child_count = VisualTreeHelper.GetChildrenCount(root)
        except Exception:
            return None
        for idx in range(int(child_count or 0)):
            try:
                child = VisualTreeHelper.GetChild(root, idx)
            except Exception:
                continue
            match = self._find_visual_descendant(child, target_type)
            if match is not None:
                return match
        return None

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
            try:
                text_value = str(getattr(combo, "Text", "") or "").strip()
            except Exception:
                text_value = ""
            if text_value:
                return text_value
            return str(fallback or "")
        try:
            return str(selected.Content or "")
        except Exception:
            return str(selected or "")

    def _set_control_style(self, control, style_key):
        if control is None or not style_key:
            return
        try:
            style = self.FindResource(style_key)
        except Exception:
            style = None
        if style is None:
            return
        try:
            control.Style = style
        except Exception:
            pass

    def _apply_override_input_styles(self, circuit_id, toggles):
        manual_keys = set((self._vm.value_overrides.get(circuit_id) or {}).keys())
        user_override_on = bool((toggles or {}).get("user_override", False))

        def _style_for(param_name, control, is_combo):
            enabled = bool(control is not None and bool(getattr(control, "IsEnabled", False)))
            if (not user_override_on) or (not enabled):
                return "EditorState.Combo.Readonly" if is_combo else "EditorState.Text.Readonly"
            if param_name in manual_keys:
                return "EditorState.Combo.Manual" if is_combo else "EditorState.Text.Manual"
            return "EditorState.Combo.Auto" if is_combo else "EditorState.Text.Auto"

        def _style_for_always_editable(param_name, control, is_combo):
            enabled = bool(control is not None and bool(getattr(control, "IsEnabled", False)))
            if not enabled:
                return "EditorState.Combo.Readonly" if is_combo else "EditorState.Text.Readonly"
            if param_name in manual_keys:
                return "EditorState.Combo.Manual" if is_combo else "EditorState.Text.Manual"
            return "EditorState.Combo.Auto" if is_combo else "EditorState.Text.Auto"

        self._set_control_style(
            self._num_sets_tb,
            _style_for("CKT_Number of Sets_CED", self._num_sets_tb, False),
        )

        always_numeric = (
            self._rating_tb,
            self._frame_tb,
            self._length_tb,
        )
        for control in always_numeric:
            enabled = bool(control is not None and bool(getattr(control, "IsEnabled", False)))
            self._set_control_style(
                control,
                "EditorState.Text.Auto" if enabled else "EditorState.Text.Readonly",
            )

        always_text = (
            self._load_name_tb,
            self._sched_notes_tb,
        )
        for control in always_text:
            enabled = bool(control is not None and bool(getattr(control, "IsEnabled", False)))
            self._set_control_style(
                control,
                "EditorState.Input.Auto" if enabled else "EditorState.Input.Readonly",
            )

        combo_targets = (
            (self._hot_size_cb, "CKT_Wire Hot Size_CEDT"),
            (self._neutral_size_cb, "CKT_Wire Neutral Size_CEDT"),
            (self._ground_size_cb, "CKT_Wire Ground Size_CEDT"),
            (self._ig_size_cb, "CKT_Wire Isolated Ground Size_CEDT"),
            (self._conduit_size_cb, "Conduit Size_CEDT"),
        )
        for control, param_name in combo_targets:
            style_key = _style_for(param_name, control, True)
            if style_key.startswith("EditorState.Combo."):
                style_key = "EditorState.Combo.Center." + style_key.split(".")[-1]
            self._set_control_style(control, style_key)

        always_combo_targets = (
            (self._conduit_type_cb, "Conduit Type_CEDT"),
            (self._wire_temp_cb, "Wire Temparature Rating_CEDT"),
            (self._wire_insulation_cb, "Wire Insulation_CEDT"),
        )
        for control, param_name in always_combo_targets:
            self._set_control_style(control, _style_for_always_editable(param_name, control, True))

    def _set_warning_border(self, element, is_warning):
        if element is None:
            return
        try:
            is_border = isinstance(element, Border)
        except Exception:
            is_border = False
        if is_border:
            brush_prop = Border.BorderBrushProperty
            thick_prop = Border.BorderThicknessProperty
        else:
            brush_prop = Control.BorderBrushProperty
            thick_prop = Control.BorderThicknessProperty
        if bool(is_warning):
            try:
                red_brush = self.FindResource("CED.Brush.AccentRed")
            except Exception:
                red_brush = None
            if red_brush is not None:
                try:
                    element.SetValue(brush_prop, red_brush)
                except Exception:
                    pass
            try:
                element.SetValue(thick_prop, Thickness(2))
            except Exception:
                pass
            return
        try:
            element.ClearValue(brush_prop)
        except Exception:
            pass
        try:
            element.ClearValue(thick_prop)
        except Exception:
            pass

    def _apply_warning_styles(self, warning_flags):
        flags = dict(warning_flags or {})
        length_warn = bool(flags.get("length", False))
        insufficient_ampacity_warn = bool(flags.get("insufficient_ampacity", False))
        insufficient_ampacity_breaker_warn = bool(flags.get("insufficient_ampacity_breaker", False))
        circuit_loads_null_warn = bool(flags.get("circuit_loads_null", False))
        circuit_panels_null_warn = bool(flags.get("circuit_panels_null", False))
        lug_qty_warn = bool(flags.get("lug_quantity_limit", False))
        conduit_fill_warn = bool(flags.get("excessive_conduit_fill", False))
        conduit_size_warn = bool(flags.get("conduit_size_issue", False))
        vd_warn = bool(flags.get("excessive_voltage_drop", False))
        lug_size_warn = bool(flags.get("lug_size_limit", False))
        lug_size_limit = _normalize_wire_size_text(flags.get("lug_size_limit_value", ""))
        undersized_ground_warn = bool(flags.get("undersized_wire_egc", False) or flags.get("undersized_wire_service_ground", False))
        rating_warn = bool(flags.get("rating", False))

        self._set_warning_border(self._length_tb, length_warn)
        self._set_warning_border(self._conduit_size_cb, conduit_size_warn or conduit_fill_warn)
        self._set_warning_border(self._conduit_fill_box, conduit_fill_warn)
        self._set_warning_border(self._vd_box, vd_warn)
        self._set_warning_border(self._rating_tb, rating_warn or insufficient_ampacity_breaker_warn)
        self._set_warning_border(self._frame_tb, False)
        self._set_warning_border(self._ampacity_box, False)
        self._set_warning_border(self._num_sets_tb, lug_qty_warn)
        self._set_warning_border(self._load_box, insufficient_ampacity_warn or circuit_loads_null_warn or circuit_panels_null_warn)

        hot_warn = bool(insufficient_ampacity_warn or insufficient_ampacity_breaker_warn)
        neutral_warn = False
        ground_warn = bool(undersized_ground_warn)
        ig_included = bool(self._include_ig is not None and bool(self._include_ig.IsChecked))
        ig_warn = bool(undersized_ground_warn and ig_included)

        if lug_size_warn:
            hot_size = _normalize_wire_size_text(self._combo_text(self._hot_size_cb, ""))
            neutral_size = _normalize_wire_size_text(self._combo_text(self._neutral_size_cb, ""))
            if lug_size_limit and _is_wire_larger_than_limit(hot_size, lug_size_limit):
                hot_warn = True
            if lug_size_limit and _is_wire_larger_than_limit(neutral_size, lug_size_limit):
                neutral_warn = True
            if not lug_size_limit:
                hot_warn = True

        self._set_warning_border(self._hot_size_cb, hot_warn)
        self._set_warning_border(self._neutral_size_cb, neutral_warn)
        self._set_warning_border(self._ground_size_cb, ground_warn)
        self._set_warning_border(self._ig_size_cb, ig_warn)

    def _load_selected_row(self):
        row = self._selected_row()
        if row is None:
            if self._panel is not None:
                self._panel.IsEnabled = False
            if self._selected_text is not None:
                self._selected_text.Text = "Select a circuit"
            if self._unsupported_text is not None:
                self._unsupported_text.Visibility = Visibility.Collapsed
            if self._blocked_info_tb is not None:
                self._blocked_info_tb.Visibility = Visibility.Collapsed
                self._blocked_info_tb.Text = ""
            if self._notice_text is not None:
                self._notice_text.Text = "No warnings."
            self._set_notice_warning_state(False)
            self._sync_reset_button_state(None)
            return

        self._is_loading = True
        try:
            state = self._vm.get_state(row.circuit_id) or {}
            values = dict(state.get("values") or {})
            toggles = dict(state.get("toggles") or {})
            editability = dict(state.get("editability") or {})
            preview = dict(state.get("preview") or {})
            notices = list(state.get("notices") or [])
            warning_flags = dict(state.get("warning_flags") or {})
            supported = bool(state.get("supported", False))
            locked = bool(state.get("locked", False))
            disabled_reason = str(state.get("disabled_reason", "") or "")

            if self._selected_text is not None:
                self._selected_text.Text = str(row.panel_ckt_text or "-")
            if self._panel is not None:
                # Keep panel/scroll active for locked rows; individual controls handle readonly state.
                self._panel.IsEnabled = bool(supported)
            if self._unsupported_text is not None:
                if not supported and not locked:
                    self._unsupported_text.Visibility = Visibility.Visible
                    self._unsupported_text.Text = disabled_reason or "Editing is disabled."
                else:
                    self._unsupported_text.Visibility = Visibility.Collapsed
                    self._unsupported_text.Text = "Unsupported circuit type. Editing is disabled."
            if self._blocked_info_tb is not None:
                if locked:
                    self._blocked_info_tb.Visibility = Visibility.Visible
                    self._blocked_info_tb.Text = disabled_reason or "Blocked - Locked by another user"
                else:
                    self._blocked_info_tb.Visibility = Visibility.Collapsed
                    self._blocked_info_tb.Text = ""

            if self._user_override is not None:
                self._user_override.IsChecked = bool(toggles.get("user_override", False))
                self._user_override.IsEnabled = bool(supported and not locked)
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
            if self._hot_include_state_text is not None:
                self._hot_include_state_text.Text = "Yes" if bool(self._hot_custom is not None and self._hot_custom.IsChecked) else "No"
            if self._ground_include_state_text is not None:
                self._ground_include_state_text.Text = "Yes" if bool(self._ground_custom is not None and self._ground_custom.IsChecked) else "No"

            self._rating_tb.Text = str(_as_int(values.get("CKT_Rating_CED", 0), 0))
            self._frame_tb.Text = str(_as_int(values.get("CKT_Frame_CED", 0), 0))
            self._length_tb.Text = _fmt_number(values.get("CKT_Length Makeup_CED", 0.0), 2)
            if self._load_name_tb is not None:
                self._load_name_tb.Text = str(values.get(CIRCUIT_NAME_KEY, row.load_name or "") or "")
            if self._sched_notes_tb is not None:
                self._sched_notes_tb.Text = str(values.get(CIRCUIT_NOTES_KEY, "") or "")
            self._num_sets_tb.Text = str(_as_int(values.get("CKT_Number of Sets_CED", 0), 0))
            if self._total_length_text is not None:
                self._total_length_text.Text = str(preview.get("total_length", "-"))
            if self._circuit_type_text is not None:
                self._circuit_type_text.Text = str(preview.get("circuit_type", row.branch_type or "-") or "-")
            if self._volts_text is not None:
                self._volts_text.Text = str(preview.get("volts_poles", "-") or "-")

            self._set_combo_value(self._hot_size_cb, preview.get("hot_size_raw", values.get("CKT_Wire Hot Size_CEDT", "")))
            self._set_combo_value(self._neutral_size_cb, preview.get("neutral_size_raw", values.get("CKT_Wire Neutral Size_CEDT", "")))
            self._set_combo_value(self._ground_size_cb, preview.get("ground_size_raw", values.get("CKT_Wire Ground Size_CEDT", "")))
            self._set_combo_value(self._ig_size_cb, preview.get("ig_size_raw", values.get("CKT_Wire Isolated Ground Size_CEDT", "")))
            conduit_type_value = preview.get("conduit_type", values.get("Conduit Type_CEDT", ""))
            conduit_size_value = preview.get("conduit_size_raw", values.get("Conduit Size_CEDT", ""))
            self._set_combo_value(self._conduit_type_cb, conduit_type_value)
            self._refresh_conduit_size_items(conduit_type_value, conduit_size_value)
            self._set_combo_value(self._wire_temp_cb, preview.get("wire_temp", values.get("Wire Temparature Rating_CEDT", "")))
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

            self._hot_custom.IsEnabled = bool(supported and not locked and self._user_override.IsChecked)
            self._ground_custom.IsEnabled = bool(supported and not locked and self._user_override.IsChecked and self._hot_custom.IsChecked)
            self._conduit_custom.IsEnabled = bool(supported and not locked and self._user_override.IsChecked)
            hot_checked = bool(self._hot_custom.IsChecked)
            ground_checked = bool(self._ground_custom.IsChecked)
            if self._include_neutral is not None:
                self._include_neutral.IsEnabled = bool(
                    supported
                    and not locked
                    and include_neutral_enabled
                    and hot_checked
                    and not neutral_locked_by_type
                )
            if self._include_ig is not None:
                self._include_ig.IsEnabled = bool(
                    supported
                    and not locked
                    and include_ig_enabled
                    and hot_checked
                    and ground_checked
                )

            self._rating_tb.IsEnabled = bool(supported and not locked)
            self._frame_tb.IsEnabled = bool(supported and not locked)
            self._length_tb.IsEnabled = bool(supported and not locked)
            if self._load_name_tb is not None:
                self._load_name_tb.IsEnabled = bool(supported and not locked)
            if self._sched_notes_tb is not None:
                self._sched_notes_tb.IsEnabled = bool(supported and not locked)

            self._num_sets_tb.IsEnabled = bool(supported and not locked and hot_checked)
            self._num_sets_tb.IsReadOnly = not bool(supported and not locked and wire_manual and hot_checked)
            self._hot_size_cb.IsEnabled = bool(supported and not locked and wire_manual and hot_checked)
            neutral_checked = bool(self._include_neutral is not None and self._include_neutral.IsChecked)
            ig_checked = bool(self._include_ig is not None and self._include_ig.IsChecked)
            self._neutral_size_cb.IsEnabled = bool(
                supported
                and not locked
                and neutral_manual
                and hot_checked
                and neutral_checked
            )
            self._ground_size_cb.IsEnabled = bool(supported and not locked and ground_manual and hot_checked and ground_checked)
            self._ig_size_cb.IsEnabled = bool(supported and not locked and ig_manual and hot_checked and ground_checked and ig_checked)

            self._conduit_size_cb.IsEnabled = bool(supported and not locked and conduit_manual and self._conduit_custom.IsChecked)
            self._conduit_type_cb.IsEnabled = bool(supported and not locked and self._conduit_custom.IsChecked)
            if self._wire_material_cu is not None:
                self._wire_material_cu.IsEnabled = bool(supported and not locked and wire_specs_enabled)
            if self._wire_material_al is not None:
                self._wire_material_al.IsEnabled = bool(supported and not locked and wire_specs_enabled)
            if self._wire_temp_cb is not None:
                self._wire_temp_cb.IsEnabled = bool(supported and not locked and wire_specs_enabled)
            if self._wire_insulation_cb is not None:
                self._wire_insulation_cb.IsEnabled = bool(supported and not locked and wire_specs_enabled)

            self._apply_override_input_styles(row.circuit_id, toggles)

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
            self._set_notice_warning_state(bool(notices))
            self._apply_warning_styles(warning_flags)
        finally:
            self._is_loading = False
        self._refresh_pending_count()
        self._sync_reset_button_state(row)

    def _set_notice_warning_state(self, has_notices):
        if self._notice_box is None:
            return
        if bool(has_notices):
            try:
                warning_brush = self.FindResource("CED.Brush.DataGridWarningBackground")
            except Exception:
                warning_brush = None
            if warning_brush is not None:
                try:
                    self._notice_box.Background = warning_brush
                except Exception:
                    pass
            return
        try:
            self._notice_box.ClearValue(Border.BackgroundProperty)
        except Exception:
            pass

    def _persist_selected_values(self, event_sender=None):
        if self._is_loading:
            return
        row = self._selected_row()
        if row is None:
            return
        if bool(getattr(row, "is_locked", False)):
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
        if _should_capture(CIRCUIT_NAME_KEY, self._load_name_tb):
            value_updates[CIRCUIT_NAME_KEY] = str(self._load_name_tb.Text or "")
        if _should_capture(CIRCUIT_NOTES_KEY, self._sched_notes_tb):
            value_updates[CIRCUIT_NOTES_KEY] = str(self._sched_notes_tb.Text or "")
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
        if _should_capture("Wire Temparature Rating_CEDT", self._wire_temp_cb):
            value_updates["Wire Temparature Rating_CEDT"] = self._combo_text(
                self._wire_temp_cb, current_values.get("Wire Temparature Rating_CEDT", "")
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
        if sender is not None and getattr(sender, "SelectedItem", None) is None:
            try:
                added_items = list(getattr(args, "AddedItems", None) or [])
            except Exception:
                added_items = []
            if added_items:
                try:
                    sender.SelectedItem = added_items[0]
                except Exception:
                    pass
        if sender is self._conduit_type_cb and self._conduit_type_cb is not None:
            conduit_type = self._combo_text(self._conduit_type_cb, "")
            current_size = self._combo_text(self._conduit_size_cb, "")
            self._refresh_conduit_size_items(conduit_type, current_size)
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

    def _on_loaded_bind_textbox_behaviors(self, sender, args):
        mode_map = {
            "LoadNameTextBox": ui_bases.TEXTBOX_MODE_DEFAULT,
            "RatingTextBox": ui_bases.TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK,
            "FrameTextBox": ui_bases.TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK,
            "LengthMakeupTextBox": ui_bases.TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK,
            "ScheduleNotesTextBox": ui_bases.TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK,
            "NumberSetsTextBox": ui_bases.TEXTBOX_MODE_SELECT_ALL_ON_FIRST_CLICK,
        }
        try:
            ui_bases.wire_textbox_interaction_modes(
                self,
                textbox_mode_map=mode_map,
                default_mode=ui_bases.TEXTBOX_MODE_DEFAULT,
            )
        except Exception:
            pass

    def _textbox_preview_mouse_down(self, sender, args):
        ui_bases._handle_textbox_preview_mouse_down_for_instance(self, sender, args)

    def _textbox_got_keyboard_focus(self, sender, args):
        ui_bases._handle_textbox_got_keyboard_focus_for_instance(self, sender, args)

    def apply_clicked(self, sender, args):
        updates = list(self._vm.build_apply_updates() or [])
        if not updates:
            forms.alert("No staged circuit changes to apply.", title="Edit Circuit Properties")
            return
        self.apply_requested = True
        self.apply_payload = {"updates": updates}
        self.Close()

    def reset_clicked(self, sender, args):
        selected = self._selected_row()
        if selected is None:
            return
        self._vm.reset_circuit(selected.circuit_id)
        try:
            if self._list is not None:
                self._list.Items.Refresh()
        except Exception:
            pass
        try:
            self._list.SelectedItem = selected
        except Exception:
            pass
        self._load_selected_row()
        if self._status_text is not None:
            self._status_text.Text = "Selected circuit reset."

    def cancel_clicked(self, sender, args):
        self.apply_requested = False
        self.apply_payload = {}
        self.Close()
