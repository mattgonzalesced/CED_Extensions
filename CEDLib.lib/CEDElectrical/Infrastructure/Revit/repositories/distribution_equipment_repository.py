# -*- coding: utf-8 -*-
"""Revit mappers for DistributionEquipment domain models."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Model.distribution_equipment import DistributionEquipment, PowerBus, Transformer
from Snippets import revit_helpers

try:
    _INTEGER_TYPES = (int, long)
except Exception:
    _INTEGER_TYPES = (int,)

PART_TYPE_MAP = {
    14: "Panelboard",
    15: "Transformer",
    16: "Switchboard",
    17: "Other Panel",
    18: "Equipment Switch",
}


def _panel_schedule_type(name, fallback=None):
    """Return DBE.PanelScheduleType member by name."""
    try:
        return getattr(DBE.PanelScheduleType, name)
    except Exception:
        return fallback


PSTYPE_UNKNOWN = _panel_schedule_type("Unknown", None)
PSTYPE_BRANCH = _panel_schedule_type("Branch", PSTYPE_UNKNOWN)
PSTYPE_SWITCHBOARD = _panel_schedule_type("Switchboard", PSTYPE_BRANCH)
PSTYPE_DATA = _panel_schedule_type("Data", PSTYPE_BRANCH)

PART_TYPE_TO_PANEL_SCHEDULE_TYPE = {
    14: PSTYPE_BRANCH,
    16: PSTYPE_SWITCHBOARD,
    17: PSTYPE_DATA,
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


def _bip(name):
    """Return BuiltInParameter member by name."""
    try:
        return getattr(DB.BuiltInParameter, name)
    except Exception:
        return None


def _param_from_names(element, names, include_type=True):
    """Return first matching parameter by name from instance/type."""
    for name in list(names or []):
        param = revit_helpers.get_parameter(
            element,
            name,
            include_type=bool(include_type),
            case_insensitive=True,
        )
        if param is not None:
            return param
    return None


def _param_value(param, default=None):
    """Return native parameter value."""
    return revit_helpers.get_parameter_value(param, default=default)


def _param_from_bips(element, bip_names):
    """Return first non-empty built-in parameter value."""
    for bip_name in list(bip_names or []):
        bip = _bip(bip_name)
        if bip is None:
            continue
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
    return _to_text(value, "").strip().lower() == _to_text(target, "").strip().lower()


def _panel_configuration_member(name, fallback=None):
    """Return DBE.PanelConfiguration member by name."""
    try:
        return getattr(DBE.PanelConfiguration, name)
    except Exception:
        return fallback


PCFG_ONE_COLUMN = _panel_configuration_member("OneColumn", None)
PCFG_TWO_COLUMNS_ACROSS = _panel_configuration_member("TwoColumnsCircuitsAcross", None)
PCFG_TWO_COLUMNS_DOWN = _panel_configuration_member("TwoColumnsCircuitsDown", None)


def _normalize_panel_configuration(value):
    """Return canonical DBE.PanelConfiguration member when possible."""
    if value is None:
        return None
    for candidate in (PCFG_ONE_COLUMN, PCFG_TWO_COLUMNS_ACROSS, PCFG_TWO_COLUMNS_DOWN):
        if candidate is None:
            continue
        if _enum_equals(value, candidate):
            return candidate
    text = _to_text(value, "").strip().lower()
    if text in ("onecolumn", "one column"):
        return PCFG_ONE_COLUMN
    if text in ("twocolumnscircuitsacross", "two columns circuits across", "across"):
        return PCFG_TWO_COLUMNS_ACROSS
    if text in ("twocolumnscircuitsdown", "two columns circuits down", "down"):
        return PCFG_TWO_COLUMNS_DOWN
    return None


def _family_parameter_from_bip(equipment, bip_name):
    """Return built-in parameter from Family definition of an instance."""
    bip = _bip(bip_name)
    if bip is None or equipment is None:
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
    if int(part_type or 0) in (16, 17):
        return PCFG_ONE_COLUMN
    if int(part_type or 0) != 14:
        return PCFG_ONE_COLUMN
    param = _family_parameter_from_bip(equipment, "RBS_ELEC_PANEL_CONFIGURATION_PARAM")
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
        if isinstance(value, _INTEGER_TYPES):
            return bool(int(value))
    except Exception:
        pass
    try:
        if isinstance(value, float):
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
        param = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE) if family else None
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
        name_param = dist.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
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
            lg_param = lg.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            if lg_param and lg_param.HasValue:
                result["lg_voltage"] = _volts_from_internal(lg_param.AsDouble())
    except Exception:
        pass
    try:
        ll = dist.VoltageLineToLine
        if ll is not None:
            ll_param = ll.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
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
                "RBS_ELEC_PANEL_TOTALLOAD_PARAM",
                "RBS_ELEC_PANEL_TOTAL_CONNECTED_LOAD",
            ],
        ),
        "current_connected_total": _param_from_bips(
            equipment,
            [
                "RBS_ELEC_PANEL_TOTALLOAD_CURRENT_PARAM",
                "RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT",
            ],
        ),
        "power_demand_total": _param_from_bips(
            equipment,
            ["RBS_ELEC_PANEL_TOTALESTLOAD_PARAM"],
        ),
        "current_demand_total": _param_from_bips(
            equipment,
            ["RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM"],
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
        ["RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM"],
    )
    secondary_dist_id = _param_from_bips(
        equipment,
        ["RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS"],
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
            _param_from_names(equipment, ["Feed Thru Lugs", "Has Feed Thru Lugs"], include_type=True),
            default=None,
        )
    )
    has_neutral_bus = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Neutral Bus", "Has Neutral Bus"], include_type=True),
            default=None,
        )
    )
    has_ground_bus = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Ground Bus", "Has Ground Bus"], include_type=True),
            default=None,
        )
    )
    has_isolated_ground_bus = _param_bool(
        _param_value(
            _param_from_names(equipment, ["Isolated Ground Bus", "Has Isolated Ground Bus"], include_type=True),
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
    if part_type in (14, 17):
        max_poles = _schedule_slot_count(schedule_view)
        if max_poles <= 0:
            max_poles = _param_value(
                _param_from_names(
                    equipment,
                    ["Max Number of Single Pole Breakers_CED", "Max Number of Single Pole Breakers"],
                    include_type=True,
                ),
                default=None,
            )
    elif part_type in (15, 16):
        max_poles = len(_system_ids(assigned_systems))
        if max_poles <= 0:
            max_poles = _param_value(
                _param_from_names(
                    equipment,
                    ["Max Number of Circuits_CED", "Max Number of Circuits"],
                    include_type=True,
                ),
                default=None,
            )

    base_kwargs = {
        "id": _idval(equipment.Id),
        "name": _to_text(getattr(equipment, "Name", None), None),
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
    }

    if part_type == 15:
        base_kwargs.update(
            {
                "xfmr_rating": _param_value(
                    _param_from_names(equipment, ["Transformer Rating", "XFMR Rating"], include_type=True),
                    default=None,
                ),
                "xfmr_impedance": _param_value(
                    _param_from_names(equipment, ["Transformer Impedance", "XFMR Impedance"], include_type=True),
                    default=None,
                ),
                "xfmr_kfactor": _param_value(
                    _param_from_names(equipment, ["K-Factor", "Transformer K-Factor"], include_type=True),
                    default=None,
                ),
            }
        )
        return Transformer(**base_kwargs)

    if part_type in (14, 16, 17):
        panel_configuration = _panel_configuration_for_equipment(equipment, part_type)
        base_kwargs.update(
            {
                "has_panel_schedule": bool(schedule_view is not None),
                "panel_configuration": panel_configuration,
            }
        )
        return PowerBus(**base_kwargs)

    return DistributionEquipment(**base_kwargs)
