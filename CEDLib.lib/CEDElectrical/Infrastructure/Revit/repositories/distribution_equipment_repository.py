# -*- coding: utf-8 -*-
"""Revit mappers for DistributionEquipment domain models."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Model.distribution_equipment import DistributionEquipment, PowerBus, Transformer
from CEDElectrical.part_types import (
    PART_TYPE_MAP,
    PART_TYPE_OTHER_PANEL,
    PART_TYPE_PANELBOARD,
    PART_TYPE_SWITCHBOARD,
    PART_TYPE_TRANSFORMER,
)
from Snippets import revit_helpers

PSTYPE_UNKNOWN = DBE.PanelScheduleType.Unknown
PSTYPE_BRANCH = DBE.PanelScheduleType.Branch
PSTYPE_SWITCHBOARD = DBE.PanelScheduleType.Switchboard
PSTYPE_DATA = DBE.PanelScheduleType.Data

PART_TYPE_TO_PANEL_SCHEDULE_TYPE = {
    PART_TYPE_PANELBOARD: PSTYPE_BRANCH,
    PART_TYPE_SWITCHBOARD: PSTYPE_SWITCHBOARD,
    PART_TYPE_OTHER_PANEL: PSTYPE_DATA,
}


def _idval(item):
    """Return numeric value for ElementId-like objects."""
    return revit_helpers.get_elementid_value(item)


def _to_text(value, fallback=""):
    """Return safe string conversion."""
    if value is None:
        return fallback
    try:
        return str(value)
    except Exception:
        return fallback


BIP_ELEC_PANEL_CONFIGURATION = DB.BuiltInParameter.RBS_ELEC_PANEL_CONFIGURATION_PARAM
BIP_FAMILY_CONTENT_PART_TYPE = DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE
BIP_FAMILY_DIST_SYSTEM = DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM
BIP_FAMILY_SECONDARY_DIST_SYSTEM = DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS
BIP_SYMBOL_NAME = DB.BuiltInParameter.SYMBOL_NAME_PARAM
BIP_VOLTAGE_TYPE_VOLTAGE = DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM
BIP_PANEL_TOTAL_LOAD = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM
BIP_PANEL_TOTAL_CONNECTED_LOAD = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM
BIP_PANEL_TOTAL_CURRENT = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM
BIP_PANEL_TOTAL_CONNECTED_CURRENT = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM
BIP_PANEL_TOTAL_DEMAND_LOAD = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALESTLOAD_PARAM
BIP_PANEL_TOTAL_DEMAND_CURRENT = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM
BIP_PANEL_CONNECTED_CURRENT_PHASEA = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_CURRENT_PHASEA
BIP_PANEL_CONNECTED_CURRENT_PHASEB = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_CURRENT_PHASEB
BIP_PANEL_CONNECTED_CURRENT_PHASEC = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_CURRENT_PHASEC
BIP_PANEL_CONNECTED_LOAD_PHASEA = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_APPARENT_LOAD_PHASEA
BIP_PANEL_CONNECTED_LOAD_PHASEB = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_APPARENT_LOAD_PHASEB
BIP_PANEL_CONNECTED_LOAD_PHASEC = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_APPARENT_LOAD_PHASEC
BIP_PANEL_NAME =DB.BuiltInParameter.RBS_ELEC_PANEL_NAME
BIP_PANEL_MCB_RATING = DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM
BIP_PANEL_MAINS_RATING = DB.BuiltInParameter.RBS_ELEC_MAINS
BIP_PANEL_MODIFICATIONS = DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS
BIP_PANEL_MOUNTING = DB.BuiltInParameter.RBS_ELEC_MOUNTING
BIP_PANEL_ENCLOSURE = DB.BuiltInParameter.RBS_ELEC_ENCLOSURE
BIP_PANEL_FEED_THRU_LUGS = DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_THRU_LUGS_PARAM
BIP_PANEL_FEED = DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_PARAM
BIP_PANEL_MAX_BREAKERS = DB.BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS
BIP_PANEL_MAX_CIRCUITS = DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS
BIP_PANEL_SHORT_CIRCUIT_RATING = DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING


def _param_from_names(element, names, include_type=True):
    """Return first matching parameter by name from instance/type."""
    for name in list(names or []):
        param = revit_helpers.get_parameter(
            element,
            name,
            include_type=bool(include_type),
            case_insensitive=False,
        )
        if param is not None:
            return param
    return None


def _param_value(param, default=None):
    """Return native parameter value."""
    return revit_helpers.get_parameter_value(param, default=default)


def _param_from_bips(element, bips):
    """Return first non-empty built-in parameter value."""
    for bip in list(bips or []):
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
        if param is None or not param.HasValue:
            continue
        value = _param_value(param, default=None)
        if value is not None:
            return value
    return None


def _enum_equals(value, target):
    """Return True when two enum-like values represent the same member."""
    try:
        if value == target:
            return True
    except Exception:
        pass
    try:
        return int(value) == int(target)
    except Exception:
        pass
    return False


PCFG_ONE_COLUMN = DBE.PanelConfiguration.OneColumn
PCFG_TWO_COLUMNS_ACROSS = DBE.PanelConfiguration.TwoColumnsCircuitsAcross
PCFG_TWO_COLUMNS_DOWN = DBE.PanelConfiguration.TwoColumnsCircuitsDown


def _normalize_panel_configuration(value):
    """Return canonical DBE.PanelConfiguration member when possible."""
    if value is None:
        return None
    for candidate in (PCFG_ONE_COLUMN, PCFG_TWO_COLUMNS_ACROSS, PCFG_TWO_COLUMNS_DOWN):
        if candidate is None:
            continue
        if _enum_equals(value, candidate):
            return candidate
    return None


def _family_parameter_from_bip(equipment, bip):
    """Return built-in parameter from Family definition of an instance."""
    if equipment is None:
        return None
    if not isinstance(equipment, DB.FamilyInstance):
        return None
    try:
        symbol = equipment.Symbol
    except Exception:
        symbol = None
    try:
        family = symbol.Family if symbol is not None else None
    except Exception:
        family = None
    if family is None:
        return None
    try:
        param = family.get_Parameter(bip)
    except Exception:
        param = None
    if param is None:
        return None
    try:
        if not bool(param.HasValue):
            return None
    except Exception:
        pass
    return param


def _panel_configuration_for_equipment(equipment, part_type):
    """Return panel configuration using part-type rules and family fallback."""
    if part_type in (PART_TYPE_SWITCHBOARD, PART_TYPE_OTHER_PANEL):
        return PCFG_ONE_COLUMN
    if part_type != PART_TYPE_PANELBOARD:
        return PCFG_ONE_COLUMN
    param = _family_parameter_from_bip(equipment, BIP_ELEC_PANEL_CONFIGURATION)
    value = _param_value(param, default=None)
    config = _normalize_panel_configuration(value)
    if config is not None:
        return config
    return PCFG_ONE_COLUMN


def _param_bool(value):
    """Normalize mixed parameter value types into bool/None."""
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    try:
        return bool(int(value))
    except Exception:
        pass
    return None


def _volts_from_internal(value):
    """Convert Revit internal electrical value to volts when possible."""
    if value is None:
        return None
    try:
        return DB.UnitUtils.ConvertFromInternalUnits(float(value), DB.UnitTypeId.Volts)
    except Exception:
        return value


def get_family_part_type(equipment):
    """Return FAMILY_CONTENT_PART_TYPE integer from family definition."""
    if equipment is None or not isinstance(equipment, DB.FamilyInstance):
        return None
    try:
        symbol = equipment.Symbol
        family = symbol.Family if symbol else None
        param = family.get_Parameter(BIP_FAMILY_CONTENT_PART_TYPE) if family else None
        if param and param.HasValue and param.StorageType == DB.StorageType.Integer:
            return int(param.AsInteger())
    except Exception:
        pass
    return None


def equipment_type_from_part_type(part_type):
    """Return equipment type label from family part type."""
    return PART_TYPE_MAP.get(part_type, "Unknown")


def expected_panel_schedule_type_for_equipment(equipment):
    """Return expected DBE.PanelScheduleType from equipment part type."""
    part_type = get_family_part_type(equipment)
    if part_type in PART_TYPE_TO_PANEL_SCHEDULE_TYPE:
        return PART_TYPE_TO_PANEL_SCHEDULE_TYPE.get(part_type, PSTYPE_BRANCH)
    return PSTYPE_UNKNOWN


def _distribution_system_snapshot(doc, dist_system_id):
    """Return distribution system snapshot from DistributionSysType element id."""
    result = {
        "id": 0,
        "name": "",
        "phase": None,
        "wire_count": None,
        "lg_voltage": None,
        "ll_voltage": None,
    }
    if dist_system_id is None:
        return result
    try:
        dist_id_val = _idval(dist_system_id)
    except Exception:
        dist_id_val = 0
    if dist_id_val <= 0:
        return result
    result["id"] = dist_id_val
    dist = doc.GetElement(dist_system_id)
    if dist is None:
        return result
    try:
        name_param = dist.get_Parameter(BIP_SYMBOL_NAME)
        if name_param and name_param.HasValue:
            result["name"] = _to_text(name_param.AsString(), "")
    except Exception:
        pass
    try:
        result["phase"] = dist.ElectricalPhase
    except Exception:
        pass
    try:
        value = dist.NumWires
        if value is not None:
            wires = int(value)
            if wires > 0:
                result["wire_count"] = wires
    except Exception:
        pass
    try:
        lg = dist.VoltageLineToGround
        if lg is not None:
            lg_param = lg.get_Parameter(BIP_VOLTAGE_TYPE_VOLTAGE)
            if lg_param and lg_param.HasValue:
                result["lg_voltage"] = _volts_from_internal(lg_param.AsDouble())
    except Exception:
        pass
    try:
        ll = dist.VoltageLineToLine
        if ll is not None:
            ll_param = ll.get_Parameter(BIP_VOLTAGE_TYPE_VOLTAGE)
            if ll_param and ll_param.HasValue:
                result["ll_voltage"] = _volts_from_internal(ll_param.AsDouble())
    except Exception:
        pass
    return result


def _branch_circuit_options(primary_profile, secondary_profile=None):
    """Build branch-circuit voltage/pole options from distribution profiles."""
    options = []

    def _append_options(profile, source):
        if not profile:
            return
        lg = profile.get("lg_voltage")
        ll = profile.get("ll_voltage")
        phase = profile.get("phase")
        is_single_phase = False
        try:
            if phase == DBE.ElectricalPhase.SinglePhase:
                is_single_phase = True
        except Exception:
            pass
        if lg is not None:
            options.append({"source": source, "poles": 1, "voltage": lg})
        if ll is not None:
            options.append({"source": source, "poles": 2, "voltage": ll})
            if not is_single_phase:
                options.append({"source": source, "poles": 3, "voltage": ll})

    _append_options(primary_profile, "primary")
    _append_options(secondary_profile, "secondary")
    deduped = []
    seen = set()
    for item in options:
        key = (int(item.get("poles", 0) or 0), _to_text(item.get("voltage"), ""))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
    return deduped


def _system_ids(systems):
    """Return sorted circuit id values from ElectricalSystem collections."""
    ids = []
    for system in list(systems or []):
        try:
            cid = _idval(system.Id)
        except Exception:
            cid = 0
        if cid > 0:
            ids.append(cid)
    ids.sort()
    return ids


def _schedule_slot_count(schedule_view):
    """Return schedule slot count from PanelScheduleView."""
    if schedule_view is None:
        return 0
    try:
        table = schedule_view.GetTableData()
        return int(table.NumberOfSlots or 0)
    except Exception:
        return 0


def _total_power_current_snapshot(equipment):
    """Return connected/demand panel load totals when available."""
    return {
        "power_connected_total": _param_from_bips(
            equipment,
            [
                BIP_PANEL_TOTAL_LOAD,
                BIP_PANEL_TOTAL_CONNECTED_LOAD,
            ],
        ),
        "current_connected_total": _param_from_bips(
            equipment,
            [
                BIP_PANEL_TOTAL_CURRENT,
                BIP_PANEL_TOTAL_CONNECTED_CURRENT,
            ],
        ),
        "power_demand_total": _param_from_bips(
            equipment,
            [BIP_PANEL_TOTAL_DEMAND_LOAD],
        ),
        "current_demand_total": _param_from_bips(
            equipment,
            [BIP_PANEL_TOTAL_DEMAND_CURRENT],
        ),
        "branch_current_phase_a": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_CURRENT_PHASEA],
        ),
        "branch_current_phase_b": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_CURRENT_PHASEB],
        ),
        "branch_current_phase_c": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_CURRENT_PHASEC],
        ),
        "branch_load_phase_a": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_LOAD_PHASEA],
        ),
        "branch_load_phase_b": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_LOAD_PHASEB],
        ),
        "branch_load_phase_c": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_LOAD_PHASEC],
        ),
    }


def build_distribution_equipment(doc, equipment, schedule_view=None):
    """Map a Revit electrical equipment instance into a domain model object."""
    if equipment is None:
        return None

    part_type = get_family_part_type(equipment)
    equipment_type = equipment_type_from_part_type(part_type)

    primary_dist_id = _param_from_bips(
        equipment,
        [BIP_FAMILY_DIST_SYSTEM],
    )
    secondary_dist_id = _param_from_bips(
        equipment,
        [BIP_FAMILY_SECONDARY_DIST_SYSTEM],
    )
    primary_profile = _distribution_system_snapshot(doc, primary_dist_id)
    secondary_profile = _distribution_system_snapshot(doc, secondary_dist_id)

    mep = None
    try:
        mep = equipment.MEPModel
    except Exception:
        mep = None

    all_systems = []
    assigned_systems = []
    if mep is not None:
        try:
            all_systems = list(mep.GetElectricalSystems() or [])
        except Exception:
            all_systems = []
        try:
            assigned_systems = list(mep.GetAssignedElectricalSystems() or [])
        except Exception:
            assigned_systems = []
    assigned_ids = set(_system_ids(assigned_systems))
    supply_systems = []
    for system in list(all_systems or []):
        try:
            sid = _idval(system.Id)
        except Exception:
            sid = 0
        if sid > 0 and sid not in assigned_ids:
            supply_systems.append(system)

    mains_rating = _param_value(
        _param_from_names(equipment, ["Mains Rating_CED", "Mains Rating"], include_type=True),
        default=None,
    )
    mains_type = _param_value(
        _param_from_names(equipment, ["Mains Type_CEDT", "Mains Type"], include_type=True),
        default=None,
    )
    ocp_rating = _param_value(
        _param_from_names(equipment, ["Main Breaker Rating_CED", "Main Breaker Rating"], include_type=True),
        default=None,
    )
    short_circuit_rating = _param_value(
        _param_from_names(equipment, ["Short Circuit Rating_CEDT", "Short Circuit Rating"], include_type=True),
        default=None,
    )

    has_feed_thru_lugs = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Feed Thru Lugs_CED"], include_type=True),
            default=None,
        )
    )
    has_neutral_bus = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Neutral Bus_CED"], include_type=True),
            default=None,
        )
    )
    has_ground_bus = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Ground Bus_CED"], include_type=True),
            default=None,
        )
    )
    has_isolated_ground_bus = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Isolated Ground Bus_CED"], include_type=True),
            default=None,
        )
    )

    totals = _total_power_current_snapshot(equipment)
    options = _branch_circuit_options(primary_profile, secondary_profile)

    voltage = primary_profile.get("ll_voltage") or primary_profile.get("lg_voltage")
    poles = None
    if options:
        poles = max([int(x.get("poles", 0) or 0) for x in options if int(x.get("poles", 0) or 0) > 0] or [None])

    max_poles = None
    if part_type in (PART_TYPE_PANELBOARD, PART_TYPE_TRANSFORMER, PART_TYPE_OTHER_PANEL):
        max_poles = _param_from_bips(equipment, [BIP_PANEL_MAX_BREAKERS])
        try:
            max_poles = int(max_poles or 0)
        except Exception:
            max_poles = 0
        if max_poles <= 0:
            max_poles = _param_value(
                _param_from_names(
                    equipment,
                    ["Max Number of Single Pole Breakers_CED", "Max Number of Single Pole Breakers"],
                    include_type=True,
                ),
                default=None,
            )
        try:
            max_poles = int(max_poles or 0)
        except Exception:
            max_poles = 0
        if max_poles <= 0:
            max_poles = _schedule_slot_count(schedule_view)
    elif part_type == PART_TYPE_SWITCHBOARD:
        max_poles = 0
        mep_model = getattr(equipment, "MEPModel", None)
        if mep_model is not None:
            for attr in ("MaxNumberOfCircuits", "maxNumberOfCircuits"):
                try:
                    value = int(getattr(mep_model, attr, 0) or 0)
                except Exception:
                    value = 0
                if value > 0:
                    max_poles = int(value)
                    break
        if max_poles <= 0:
            value = _param_from_bips(equipment, [BIP_PANEL_MAX_CIRCUITS])
            try:
                max_poles = int(value or 0)
            except Exception:
                max_poles = 0
        if max_poles <= 0:
            max_poles = _param_value(
                _param_from_names(
                    equipment,
                    ["Max Number of Circuits_CED", "Max Number of Circuits"],
                    include_type=True,
                ),
                default=None,
            )
            try:
                max_poles = int(max_poles or 0)
            except Exception:
                max_poles = 0
        if max_poles <= 0:
            max_poles = _schedule_slot_count(schedule_view)

    equipment_name = None
    try:
        equipment_name = equipment.Name
    except Exception:
        equipment_name = None

    base_kwargs = {
        "id": _idval(equipment.Id),
        "name": _to_text(equipment_name, None),
        "part_type": part_type,
        "equipment_type": equipment_type,
        "voltage": voltage,
        "poles": poles,
        "distribution_system": primary_profile,
        "distribution_system_secondary": secondary_profile,
        "supply_circuits": _system_ids(supply_systems),
        "branch_circuits": _system_ids(assigned_systems),
        "branch_circuit_options": options,
        "mains_rating": mains_rating,
        "mains_type": mains_type,
        "has_ocp": bool(ocp_rating not in (None, "", 0)),
        "ocp_type": _to_text(mains_type, None),
        "ocp_rating": ocp_rating,
        "has_feed_thru_lugs": has_feed_thru_lugs,
        "has_neutral_bus": has_neutral_bus,
        "has_ground_bus": has_ground_bus,
        "has_isolated_ground_bus": has_isolated_ground_bus,
        "max_poles": max_poles,
        "short_circuit_rating": short_circuit_rating,
        "power_connected_total": totals.get("power_connected_total"),
        "current_connected_total": totals.get("current_connected_total"),
        "power_demand_total": totals.get("power_demand_total"),
        "current_demand_total": totals.get("current_demand_total"),
        "branch_current_phase_a": totals.get("branch_current_phase_a"),
        "branch_current_phase_b": totals.get("branch_current_phase_b"),
        "branch_current_phase_c": totals.get("branch_current_phase_c"),
        "branch_load_phase_a": totals.get("branch_load_phase_a"),
        "branch_load_phase_b": totals.get("branch_load_phase_b"),
        "branch_load_phase_c": totals.get("branch_load_phase_c"),
    }

    if part_type == PART_TYPE_TRANSFORMER:
        base_kwargs.update(
            {
                "xfmr_rating": _param_value(
                    _param_from_names(equipment, ["Transformer Rating_CED"], include_type=True),
                    default=None,
                ),
                "xfmr_impedance": _param_value(
                    _param_from_names(equipment, ["Transformer %Z_CED"], include_type=True),
                    default=None,
                ),
                "xfmr_kfactor": _param_value(
                    _param_from_names(equipment, ["Transformer K-Factor_CEDT"], include_type=True),
                    default=None,
                ),
            }
        )
        return Transformer(**base_kwargs)

    if part_type in (PART_TYPE_PANELBOARD, PART_TYPE_SWITCHBOARD, PART_TYPE_OTHER_PANEL):
        panel_configuration = _panel_configuration_for_equipment(equipment, part_type)
        base_kwargs.update(
            {
                "has_panel_schedule": bool(schedule_view is not None),
                "panel_configuration": panel_configuration,
            }
        )
        return PowerBus(**base_kwargs)

    return DistributionEquipment(**base_kwargs)
