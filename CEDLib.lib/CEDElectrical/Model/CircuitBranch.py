# -*- coding: utf-8 -*-

import Autodesk.Revit.DB.Electrical as DBE
from System import Guid
from pyrevit import DB, script, revit

from CEDElectrical.Model.circuit_settings import (
    CircuitSettings,
    FeederVDMethod,
    NeutralBehavior,
)
from CEDElectrical.refdata.ampacity_table import WIRE_AMPACITY_TABLE
from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE
from CEDElectrical.refdata.conduit_area_table import CONDUIT_AREA_TABLE, CONDUIT_SIZE_INDEX
from CEDElectrical.refdata.egc_table import EGC_TABLE
from CEDElectrical.refdata.impedance_table import WIRE_IMPEDANCE_TABLE
from CEDElectrical.refdata.ocp_cable_defaults import OCP_CABLE_DEFAULTS
from CEDElectrical.refdata.shared_params_table import SHARED_PARAMS
from CEDElectrical.refdata.standard_ocp_table import BREAKER_FRAME_SWITCH_TABLE

console = script.get_output()
logger = script.get_logger()


class NoticeCollector(object):
    """Collects warnings/errors for a branch so they can be summarized later."""

    def __init__(self, circuit_name):
        self.circuit_name = circuit_name
        self.items = []  # list of (level, message)

    def add(self, level, message):
        self.items.append((level.upper(), message))

    def has_items(self):
        return bool(self.items)

    def grouped(self):
        grouped = {"WARNING": [], "ERROR": []}
        for level, msg in self.items:
            if level not in grouped:
                grouped[level] = []
            grouped[level].append(msg)
        return grouped

PART_TYPE_MAP = {
    14: "Panelboard",
    15: "Transformer",
    16: "Switchboard",
    17: "Other Panel",
    18: "Equipment Switch"
}

ALLOWED_WIRE_SIZES = [
    "12", "10", "8", "6", "4", "3", "2", "1",
    "1/0", "2/0", "3/0", "4/0",
    "250", "300", "350", "400",
    "500", "600", "700", "750", "800", "1000"
]




# ---------------------------------------------------------------------
# Cable / Conduit helper models
# ---------------------------------------------------------------------
class CableSet(object):
    """Pure data for conductors on one circuit."""
    def __init__(self):
        # sizes are normalized strings (no #, no C)
        self.hot_size = None
        self.neutral_size = None
        self.ground_size = None
        self.ig_size = None

        self.hot_qty = 0
        self.neutral_qty = 0
        self.ground_qty = 0
        self.ig_qty = 0

        self.sets = 1

        self.material = None          # CU / AL
        self.temp_c = None           # 60/75/90
        self.insulation = None       # THWN, etc.

        self.base_ampacity = None    # per conductor
        self.total_ampacity = None   # ampacity * sets
        self.voltage_drop = None     # decimal fraction

        self.cleared = False         # user explicitly blanked cable set
        self.calc_failed = False     # calculator could not find solution

    def get_total_area(self):
        total = 0.0
        table = CONDUCTOR_AREA_TABLE
        ins = self.insulation

        items = [
            (self.hot_size, self.hot_qty),
            (self.neutral_size or self.hot_size, self.neutral_qty),
            (self.ground_size, self.ground_qty),
            (self.ig_size or self.ground_size, self.ig_qty),
        ]

        for size, qty in items:
            if not size or not qty:
                continue
            if size not in table:
                continue
            areas = table[size]['area']
            if ins not in areas:
                continue
            total += qty * areas[ins]

        return total

    def clear(self):
        """Reset all cable-related data (but do not touch flags)."""
        self.hot_size = None
        self.neutral_size = None
        self.ground_size = None
        self.ig_size = None

        self.hot_qty = 0
        self.neutral_qty = 0
        self.ground_qty = 0
        self.ig_qty = 0

        self.sets = None
        self.material = None
        self.temp_c = None
        self.insulation = None

        self.base_ampacity = None
        self.total_ampacity = None
        self.voltage_drop = None


class ConduitRun(object):
    """Pure data for the conduit that carries the CableSet."""
    def __init__(self):
        self.conduit_type = None      # EMT, PVC, etc.
        self.material_type = None     # "Magnetic" / "Non-Magnetic"
        self.size = None              # normalized (no C suffix)
        self.fill_ratio = None        # decimal 0-1

        self.cleared = False
        self.calc_failed = False

    def clear(self):
        """Reset conduit geometry (but not flags)."""
        self.conduit_type = None
        self.material_type = None
        self.size = None
        self.fill_ratio = None

    def set_type_from_value(self, conduit_type):
        """Resolve conduit_type into type + material_type based on tables."""
        for material, type_dict in CONDUIT_AREA_TABLE.items():
            if conduit_type in type_dict:
                self.conduit_type = conduit_type
                self.material_type = material
                return True
        return False

    def apply_override_size(self, size_norm, total_area):
        """
        Attempt to apply a user-override size (normalized, no suffix).
        Returns True if applied, False if invalid for this type/material.
        """
        table = CONDUIT_AREA_TABLE.get(self.material_type, {}).get(self.conduit_type, {})
        if not table:
            return False

        if size_norm not in table:
            return False

        area = table[size_norm]
        fill = total_area / float(area)

        self.size = size_norm
        self.fill_ratio = round(fill, 5)

        return True

    def pick_size(self, total_area, settings):
        """
        Auto-pick first conduit size that keeps fill <= max_conduit_fill,
        respecting settings.min_conduit_size.
        """
        table = CONDUIT_AREA_TABLE.get(self.material_type, {}).get(self.conduit_type, {})
        if not table:
            return None

        enum = CONDUIT_SIZE_INDEX
        if settings.min_conduit_size not in enum:
            return None

        start_index = enum.index(settings.min_conduit_size)

        chosen_size = None
        chosen_fill = None

        for size in enum[start_index:]:
            if size not in table:
                continue
            area = table[size]
            fill_ratio = total_area / float(area)
            if fill_ratio <= settings.max_conduit_fill:
                chosen_size = size
                chosen_fill = round(fill_ratio, 5)
                break

        if not chosen_size:
            return None

        self.size = chosen_size
        self.fill_ratio = chosen_fill
        return chosen_size


# ---------------------------------------------------------------------
# CircuitBranch main class
# ---------------------------------------------------------------------
class CircuitBranch(object):
    def __init__(self, circuit, settings=None):
        self.circuit = circuit
        self.settings = settings if settings else CircuitSettings()

        self.circuit_id = circuit.Id.Value
        self.panel = getattr(circuit.BaseEquipment, "Name", None) if circuit.BaseEquipment else ""
        self.circuit_number = circuit.CircuitNumber
        self.name = "{}-{}".format(self.panel, self.circuit_number)

        # feeder/transformer flags
        self._is_transformer_primary = False
        self._is_feeder = self._detect_feeder()

        # wire length (Revit length + makeup)
        self._wire_length = None
        self._wire_length_makeup = 0.0
        self._wire_info = None  # dict from OCP_CABLE_DEFAULTS

        # override flags
        self._auto_calculate_override = False
        self._include_neutral = False
        self._include_isolated_ground = False

        # overrides (raw values from Revit)
        self._breaker_override = None
        self._wire_sets_override = None
        self._wire_material_override = None
        self._wire_temp_rating_override = None
        self._wire_insulation_override = None
        self._wire_hot_size_override = None
        self._wire_neutral_size_override = None
        self._wire_ground_size_override = None
        self._conduit_type_override = None
        self._conduit_size_override = None
        self._user_clear_hot = False
        self._user_clear_conduit = False

        # calculation results
        self._calculated_breaker = None

        # models for wires & conduit
        self.cable = CableSet()
        self.conduit = ConduitRun()

        # warnings/errors for summary output
        self.notices = NoticeCollector(self.name)

        # overall failure flag (if true, all output strings should blank)
        self.calc_failed = False

        # prep pipeline
        self._load_core_inputs()
        self._load_overrides()
        self._validate_overrides()
        self._setup_structural_quantities()

    # -----------------------------------------------------------------
    # Logging helpers
    # -----------------------------------------------------------------
    def log_info(self, msg, *args):
        logger.info("{}: {}".format(self.name, msg), *args)

    def log_warning(self, msg, *args):
        formatted = msg.format(*args) if args else msg
        self.notices.add("WARNING", formatted)
        logger.warning("{}: {}".format(self.name, formatted))

    def log_error(self, msg, *args):
        formatted = msg.format(*args) if args else msg
        self.notices.add("ERROR", formatted)
        logger.error("{}: {}".format(self.name, formatted))

    def log_debug(self, msg, *args):
        logger.debug("{}: {}".format(self.name, msg), *args)

    def _warn_if_overloaded(self):
        try:
            load = self.circuit_load_current
            rating = self.rating
            if load is not None and rating is not None and load > rating:
                self.log_warning(
                    "Circuit load {:.2f}A exceeds breaker rating {}A.".format(load, rating)
                )
        except Exception:
            pass

    # -----------------------------------------------------------------
    # Basic classification
    # -----------------------------------------------------------------
    @property
    def branch_type(self):
        if self.cable.cleared and self.conduit.cleared:
            return "N/A"
        if self.cable.cleared:
            return "CONDUIT ONLY"
        if self._is_feeder:
            return "FEEDER"
        if self.is_space:
            return "SPACE"
        if self.is_spare:
            return "SPARE"
        return "BRANCH"

    @property
    def is_power_circuit(self):
        return self.circuit.SystemType == DBE.ElectricalSystemType.PowerCircuit

    def _detect_feeder(self):
        """Looks at connected elements' PART_TYPE to decide if feeder."""
        try:
            for el in self.circuit.Elements:
                if isinstance(el, DB.FamilyInstance):
                    family = el.Symbol.Family
                    param = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
                    if not param or param.StorageType != DB.StorageType.Integer:
                        continue
                    part_value = param.AsInteger()

                    if part_value == 15:
                        self._is_transformer_primary = True

                    if part_value in [14, 15, 16, 17]:
                        return True
        except Exception as e:
            logger.debug("is_feeder detection failed on {}: {}".format(self.name, e))
        return False

    @property
    def is_feeder(self):
        return self._is_feeder

    @property
    def is_spare(self):
        return self.circuit.CircuitType == DBE.CircuitType.Spare

    @property
    def is_space(self):
        return self.circuit.CircuitType == DBE.CircuitType.Space

    @property
    def max_voltage_drop(self):
        if self._is_feeder:
            return self.settings.max_feeder_voltage_drop
        return self.settings.max_branch_voltage_drop

    # -----------------------------------------------------------------
    # Core Revit inputs, wire info, overrides
    # -----------------------------------------------------------------
    def _load_core_inputs(self):
        self._wire_length = None
        self._wire_length_makeup = 0.0

        if self.is_power_circuit and not self.is_spare and not self.is_space:
            try:
                rvt_length = self.circuit.Length
                makeup = self._get_param_value(SHARED_PARAMS['CKT_Length Makeup_CED']['GUID'])
                if makeup is None:
                    makeup = 0.0

                self._wire_length_makeup = makeup
                final_length = rvt_length + makeup
                if final_length <= 0:
                    self.log_warning(
                        "Wire makeup length results in a total length <= 0. Using Revit Length only."
                    )
                    final_length = rvt_length
                self._wire_length = final_length
            except Exception as e:
                logger.debug("Failed to compute wire length for {}: {}".format(self.name, e))

        # override flags (yes/no)
        try:
            self._include_neutral = self._get_yesno(SHARED_PARAMS['CKT_Include Neutral_CED']['GUID'])
            self._include_isolated_ground = self._get_yesno(
                SHARED_PARAMS['CKT_Include Isolated Ground_CED']['GUID']
            )
            self._auto_calculate_override = self._get_yesno(
                SHARED_PARAMS['CKT_User Override_CED']['GUID']
            )
        except Exception as e:
            logger.debug("_load_core_inputs flags failed for {}: {}".format(self.name, e))

        # wire info defaults
        self._wire_info = self._get_wire_info_for_rating()

    def _get_wire_info_for_rating(self):
        if not self.is_power_circuit:
            return {}

        rating = self.rating
        if rating is None:
            self.log_debug("No Revit rating; wire_info empty.")
            return {}

        rating_key = int(rating)
        table = OCP_CABLE_DEFAULTS

        if rating_key in table:
            return table[rating_key]

        sorted_keys = sorted(table.keys())
        for key in sorted_keys:
            if key >= rating_key:
                self.log_warning(
                    "No exact wire info match for breaker {}; using next available {}.".format(
                        rating_key, key
                    )
                )
                return table[key]

        fallback = sorted_keys[-1]
        self.log_warning(
            "Breaker rating {} exceeds defaults; using max available {}.".format(
                rating_key, fallback
            )
        )
        return table[fallback]

    def _load_overrides(self):
        try:
            # These may be provided in both auto and manual modes
            self._wire_material_override = self._get_param_value(
                SHARED_PARAMS['Wire Material_CEDT']['GUID']
            )
            self._wire_temp_rating_override = self._get_param_value(
                SHARED_PARAMS['Wire Temperature Rating_CEDT']['GUID']
            )
            self._wire_insulation_override = self._get_param_value(
                SHARED_PARAMS['Wire Insulation_CEDT']['GUID']
            )
            self._conduit_type_override = self._get_param_value(
                SHARED_PARAMS['Conduit Type_CEDT']['GUID']
            )
            self._conduit_size_override = self._get_param_value(
                SHARED_PARAMS['Conduit Size_CEDT']['GUID']
            )

            if self._auto_calculate_override:
                self._breaker_override = self._get_param_value(
                    SHARED_PARAMS['CKT_Rating_CED']['GUID']
                )
                self._wire_sets_override = self._get_param_value(
                    SHARED_PARAMS['CKT_Number of Sets_CED']['GUID']
                )
                self._wire_hot_size_override = self._get_param_value(
                    SHARED_PARAMS['CKT_Wire Hot Size_CEDT']['GUID']
                )
                self._wire_neutral_size_override = self._get_param_value(
                    SHARED_PARAMS['CKT_Wire Neutral Size_CEDT']['GUID']
                )
                self._wire_ground_size_override = self._get_param_value(
                    SHARED_PARAMS['CKT_Wire Ground Size_CEDT']['GUID']
                )
        except Exception as e:
            logger.debug("_load_overrides failed for {}: {}".format(self.name, e))

    def _validate_overrides(self):
        valid_insulations = set()
        for v in CONDUCTOR_AREA_TABLE.values():
            valid_insulations.update(v.get("area", {}).keys())

        # --- material ---
        if self._wire_material_override:
            norm = str(self._wire_material_override).upper().strip()
            if norm in WIRE_AMPACITY_TABLE:
                self._wire_material_override = norm
            else:
                self.log_warning(
                    "Wire material '{}' not recognized; using defaults.".format(
                        self._wire_material_override
                    )
                )
                self._wire_material_override = None

        # --- temp rating ---
        if self._wire_temp_rating_override:
            try:
                t = int(str(self._wire_temp_rating_override).replace("C", "").strip())
                if t not in (60, 75, 90):
                    raise ValueError()
                self._wire_temp_rating_override = t
            except Exception:
                self.log_warning(
                    "Wire temp '{}' invalid; reverting to defaults.".format(
                        self._wire_temp_rating_override
                    )
                )
                self._wire_temp_rating_override = None

        # --- insulation ---
        if self._wire_insulation_override:
            if not isinstance(self._wire_insulation_override, str) or not self._wire_insulation_override.strip():
                self.log_warning(
                    "Wire insulation '{}' invalid; using defaults.".format(
                        self._wire_insulation_override
                    )
                )
                self._wire_insulation_override = None
            else:
                norm_ins = self._wire_insulation_override.strip().upper()
                if valid_insulations and norm_ins not in valid_insulations:
                    self.log_warning(
                        "Wire insulation '{}' not found in tables; using defaults.".format(
                            self._wire_insulation_override
                        )
                    )
                    self._wire_insulation_override = None
                else:
                    self._wire_insulation_override = norm_ins

        # --- conduit type ---
        if self._conduit_type_override:
            raw = self._conduit_type_override
            valid = any(raw in types for _, types in CONDUIT_AREA_TABLE.items())
            if not valid:
                self.log_warning(
                    "Conduit type '{}' invalid; using defaults.".format(raw)
                )
                self._conduit_type_override = None

        # --- conduit size ---
        if self._conduit_size_override:
            raw = self._conduit_size_override
            norm = self._normalize_conduit_type(raw)
            if norm not in CONDUIT_SIZE_INDEX:
                self.log_warning(
                    "Conduit size '{}' invalid; using calculated size.".format(raw)
                )
                self._conduit_size_override = None

        # Manual-only validations below
        if not self._auto_calculate_override:
            return

        # --- wire sets ---
        if not (isinstance(self._wire_sets_override, int) and self._wire_sets_override > 0):
            try:
                parsed_sets = int(str(self._wire_sets_override).strip())
            except Exception:
                parsed_sets = None

            if parsed_sets is None or parsed_sets <= 0:
                if self._wire_sets_override not in (None, ""):
                    self.log_warning("Wire sets override '{}' is invalid. Ignoring.".format(self._wire_sets_override))
                self._wire_sets_override = None
            else:
                self._wire_sets_override = parsed_sets

        if self._wire_sets_override:
            max_sets = self._wire_info.get("max_lug_qty", 1) or 1
            if self._wire_sets_override > max_sets:
                self.log_warning(
                    "Wire sets override {} exceeds lug capacity of {} set(s); keeping user override per request.".format(
                        self._wire_sets_override, max_sets
                    )
                )

            rating = self.rating or 0
            poles = self.poles or 0
            if ((rating and rating < 100) or poles < 2) and self._wire_sets_override != 1:
                self.log_warning(
                    "Parallel sets not allowed for {}P breaker {}A; keeping {} set(s) as requested.".format(
                        poles or 0, rating, self._wire_sets_override
                    )
                )

        # --- wire sizes ---
        def _check_size(name, attr):
            raw = getattr(self, attr)
            if not raw:
                return
            if isinstance(raw, str) and raw.strip() == "-":
                setattr(self, "_user_clear_{}".format(name), True)
                setattr(self, attr, None)
                return
            norm = self._normalize_wire_size(raw)
            if norm not in CONDUCTOR_AREA_TABLE:
                self.log_warning("{} size override '{}' invalid; will auto size.".format(name.capitalize(), raw))
                setattr(self, attr, None)

        self._user_clear_hot = False
        self._user_clear_conduit = False

        _check_size("hot", "_wire_hot_size_override")
        _check_size("neutral", "_wire_neutral_size_override")
        _check_size("ground", "_wire_ground_size_override")

        if self._conduit_size_override and isinstance(self._conduit_size_override, str):
            if self._conduit_size_override.strip() == "-":
                self._user_clear_conduit = True
                self._conduit_size_override = None

        if self._wire_sets_override and self._wire_sets_override > 1 and self._is_feeder:
            hot_norm = self._normalize_wire_size(self._wire_hot_size_override) or ""
            if hot_norm and self._is_wire_below_one_aught(hot_norm):
                self.log_warning(
                    "Feeders smaller than 1/0 are typically not paralleled; keeping {} set(s) as requested.".format(
                        self._wire_sets_override
                    )
                )

        max_hot_size = self._wire_info.get("max_lug_size")
        if self._wire_hot_size_override and max_hot_size:
            hot_norm = self._normalize_wire_size(self._wire_hot_size_override)
            if self._is_wire_larger_than_limit(hot_norm, max_hot_size):
                self.log_warning(
                    "Hot size override {} exceeds lug size block {}; keeping user override per request.".format(
                        self._wire_hot_size_override, max_hot_size
                    )
                )

    def _setup_structural_quantities(self):
        """Establish default quantities based on poles, flags, feeder logic."""
        # hot qty = poles (or 0)
        self.cable.hot_qty = self.poles or 0

        # neutral qty:
        if self.poles == 1:
            neut_qty = 1
        elif self._is_feeder:
            neut_qty = self._has_feeder_ln_voltage()
        elif self._include_neutral:
            neut_qty = 1
        else:
            neut_qty = 0
        self.cable.neutral_qty = neut_qty

        # ground qty = 1 for load circuits
        if self.circuit.CircuitType == DBE.CircuitType.Circuit:
            self.cable.ground_qty = 1
        else:
            self.cable.ground_qty = 0

        # isolated ground
        self.cable.ig_qty = 1 if self._include_isolated_ground else 0

        # sets default from wire_info / 1
        base_sets = self._wire_info.get("number_of_parallel_sets", 1) or 1
        self.cable.sets = self._apply_set_constraints(base_sets, source="defaults")
        if self._user_clear_hot:
            self.cable.hot_qty = 0
            self.cable.neutral_qty = 0
            self.cable.ground_qty = 0
            self.cable.ig_qty = 0
            self.cable.cleared = True
            self.cable.sets = self._wire_sets_override or self.cable.sets or 1
            self.cable.material = None
            self.cable.temp_c = None
            self.cable.insulation = None
            return

        material_value, temp_c, insulation_value = self._resolve_wire_specs()
        self.cable.material = material_value
        self.cable.temp_c = temp_c
        self.cable.insulation = insulation_value

    def _has_feeder_ln_voltage(self):
        doc = revit.doc
        try:
            for el in self.circuit.Elements:
                if isinstance(el, DB.FamilyInstance):
                    ds_param = el.get_Parameter(DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
                    if not ds_param or not ds_param.HasValue:
                        continue
                    ds_elem = doc.GetElement(ds_param.AsElementId())
                    if isinstance(ds_elem, DBE.DistributionSysType):
                        if ds_elem.VoltageLineToGround:
                            return 1
        except Exception as e:
            logger.debug("Feeder neutral check failed on {}: {}".format(self.name, e))
        return 0

    def _apply_set_constraints(self, sets_value, source="override", enforce_design=True):
        """Clamp number of sets to breaker/pole/lug limits when requested."""
        if sets_value is None:
            return None

        try:
            sets = int(sets_value)
        except Exception:
            self.log_warning("{} set value '{}' is invalid; defaulting to 1 set.".format(source.capitalize(), sets_value))
            return 1

        if sets < 1:
            self.log_warning("{} set value '{}' is invalid; defaulting to 1 set.".format(source.capitalize(), sets_value))
            sets = 1

        rating = self.rating or 0
        poles = self.poles or 0
        max_sets = self._wire_info.get("max_lug_qty", 1) or 1

        if enforce_design:
            if (rating and rating < 100) or poles < 2:
                if sets != 1:
                    self.log_warning(
                        "Parallel sets not allowed for {}P breaker {}A. Resetting to 1 set.".format(poles or 0, rating)
                    )
                    sets = 1

            if sets > max_sets:
                self.log_warning(
                    "Requested {} sets exceeds lug capacity of {} set(s); clamping to {}.".format(sets, max_sets, max_sets)
                )
                sets = max_sets

        return sets

    def _resolve_wire_specs(self):
        if self._user_clear_hot:
            return None, None, None

        material_default = self._wire_info.get("wire_material", "CU")
        material = self._wire_material_override or material_default
        try:
            material = str(material).strip().upper()
        except Exception:
            material = material_default

        temp_default = self._wire_info.get("wire_temperature_rating", "75 C")
        try:
            temp_c = int(str(self._wire_temp_rating_override or temp_default).replace("C", "").strip())
        except Exception:
            temp_c = 75

        insulation_default = self._wire_info.get("wire_insulation")
        insulation = self._wire_insulation_override or insulation_default
        if insulation:
            try:
                insulation = str(insulation).strip().upper()
            except Exception:
                pass

        return material, temp_c, insulation

    def _resolve_conduit_type(self):
        if self._user_clear_conduit:
            return None
        return self._conduit_type_override or self._wire_info.get("conduit_type")

    # -----------------------------------------------------------------
    # Core circuit properties (Revit getters)
    # -----------------------------------------------------------------
    @property
    def load_name(self):
        try:
            return self.circuit.LoadName
        except Exception:
            return None

    @property
    def rating(self):
        try:
            if self.is_power_circuit and not self.is_space:
                return self.circuit.Rating
        except Exception:
            return None

    @property
    def frame(self):
        try:
            return self.circuit.Frame
        except Exception:
            return None

    @property
    def circuit_notes(self):
        try:
            param = self.circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
            if param and param.StorageType == DB.StorageType.String:
                return param.AsString()
        except Exception:
            pass
        return ""

    @property
    def length(self):
        return self._wire_length

    @property
    def wire_length_makeup(self):
        return self._wire_length_makeup

    @property
    def voltage(self):
        try:
            param = self.circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
            if param and param.HasValue:
                raw = param.AsDouble()
                return DB.UnitUtils.ConvertFromInternalUnits(raw, DB.UnitTypeId.Volts)
        except Exception:
            pass
        return None

    @property
    def apparent_power(self):
        try:
            return DBE.ElectricalSystem.ApparentLoad.__get__(self.circuit)
        except Exception:
            return None

    @property
    def apparent_current(self):
        try:
            return DBE.ElectricalSystem.ApparentCurrent.__get__(self.circuit)
        except Exception:
            return None

    @property
    def circuit_load_current(self):
        if self.circuit.CircuitType != DBE.CircuitType.Circuit:
            return None
        if self._is_feeder:
            return self.get_downstream_demand_current()
        return self.apparent_current

    def _get_voltage_drop_current(self):
        """Resolve feeder voltage-drop current based on settings."""
        if not self._is_feeder:
            return self.apparent_current

        method = getattr(self.settings, "feeder_vd_method", FeederVDMethod.DEMAND)
        demand_current = self.get_downstream_demand_current()
        connected_current = self.apparent_current
        base_demand = demand_current if demand_current is not None else connected_current

        if method == FeederVDMethod.CONNECTED:
            return connected_current if connected_current is not None else base_demand

        if method == FeederVDMethod.EIGHTY_PERCENT:
            eighty_current = None
            try:
                rating = self.rating
                if rating is not None:
                    eighty_current = 0.8 * float(rating)
            except Exception:
                eighty_current = None

            if eighty_current is None:
                return base_demand

            if base_demand is None:
                return eighty_current

            return base_demand if eighty_current < base_demand else eighty_current

        # Default: demand load basis
        return base_demand

    @property
    def poles(self):
        try:
            return DBE.ElectricalSystem.PolesNumber.__get__(self.circuit)
        except Exception:
            return None

    @property
    def phase(self):
        if not self.poles:
            return 0
        if self.poles == 3:
            return 3
        return 1

    @property
    def power_factor(self):
        try:
            return DBE.ElectricalSystem.PowerFactor.__get__(self.circuit)
        except Exception:
            return None

    # -----------------------------------------------------------------
    # Public "resolved" properties for writing back to Revit
    # -----------------------------------------------------------------
    @property
    def breaker_rating(self):
        if self.is_space:
            return None
        if not self.settings.auto_calculate_breaker:
            return self.rating
        if self._auto_calculate_override and self._breaker_override:
            return self._breaker_override
        return self._calculated_breaker

    @property
    def wire_material(self):
        if self.cable.cleared or self.calc_failed:
            return ""
        return self.cable.material

    @property
    def wire_temp_rating(self):
        if self.cable.cleared or self.calc_failed:
            return ""
        if self.cable.temp_c is None:
            return ""
        return "{} C".format(self.cable.temp_c)

    @property
    def wire_insulation(self):
        if self.cable.cleared or self.calc_failed:
            return ""
        return self.cable.insulation or ""

    @property
    def hot_wire_quantity(self):
        if self.cable.cleared or self.calc_failed:
            return 0
        return self.cable.hot_qty or 0

    @property
    def neutral_wire_quantity(self):
        if self.cable.cleared or self.calc_failed:
            return 0
        return self.cable.neutral_qty or 0

    @property
    def ground_wire_quantity(self):
        if self.cable.cleared or self.calc_failed:
            return 0
        return self.cable.ground_qty or 0

    @property
    def isolated_ground_wire_quantity(self):
        if self.cable.cleared or self.calc_failed:
            return 0
        return self.cable.ig_qty or 0

    def _format_wire_size(self, normalized):
        if not normalized or self.cable.cleared or self.calc_failed:
            return ""
        prefix = self.settings.wire_size_prefix or ""
        return "{}{}".format(prefix, normalized)

    @property
    def hot_wire_size(self):
        return self._format_wire_size(self.cable.hot_size)

    @property
    def neutral_wire_size(self):
        if self.neutral_wire_quantity == 0:
            return ""
        if self.cable.neutral_size:
            return self._format_wire_size(self.cable.neutral_size)
        return self._format_wire_size(self.cable.hot_size)

    @property
    def ground_wire_size(self):
        return self._format_wire_size(self.cable.ground_size)

    @property
    def isolated_ground_wire_size(self):
        if self.isolated_ground_wire_quantity == 0:
            return ""
        return self._format_wire_size(self.cable.ig_size)

    @property
    def number_of_sets(self):
        if self.calc_failed:
            return None
        if self.cable.cleared:
            return self.cable.sets or 1
        return self.cable.sets or 1

    @property
    def number_of_wires(self):
        if self.cable.cleared or self.calc_failed:
            return 0
        return self.hot_wire_quantity + self.neutral_wire_quantity

    @property
    def circuit_base_ampacity(self):
        if self.cable.cleared or self.calc_failed:
            return None
        return self.cable.total_ampacity

    @property
    def voltage_drop_percentage(self):
        if self.cable.cleared or self.calc_failed:
            return None
        return self.cable.voltage_drop

    @property
    def conduit_material_type(self):
        if self.conduit.cleared or self.calc_failed:
            return ""
        return self.conduit.material_type or ""

    @property
    def conduit_type(self):
        if self.conduit.cleared or self.calc_failed:
            return ""
        return self.conduit.conduit_type or ""

    @property
    def conduit_size(self):
        if self.conduit.cleared or self.calc_failed:
            return ""
        size = self.conduit.size
        if not size:
            return ""
        suffix = self.settings.conduit_size_suffix or ""
        return "{}{}".format(size, suffix)

    @property
    def conduit_fill_percentage(self):
        if self.conduit.cleared or self.calc_failed:
            return None
        return self.conduit.fill_ratio

    # -----------------------------------------------------------------
    # Calculations
    # -----------------------------------------------------------------
    def calculate_breaker_size(self):
        """Only used if settings.auto_calculate_breaker is True."""
        try:
            amps = self.apparent_current
            if not amps:
                self._calculated_breaker = None
                return
            amps = amps * 1.25
            if amps < self.settings.min_breaker_size:
                amps = self.settings.min_breaker_size
            for b in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
                if b >= amps:
                    self._calculated_breaker = b
                    return
            self._calculated_breaker = None
        except Exception:
            self._calculated_breaker = None

    def calculate_hot_wire_size(self):
        """Fill self.cable.hot_size, sets, total_ampacity, voltage_drop.

        Uses overrides first (if valid), then automatic sizing.
        """
        if self._user_clear_hot and self.cable.cleared:
            self.cable.voltage_drop = None
            return

        rating = self.breaker_rating
        if rating is None:
            self._fail_cable_sizing("No breaker rating.")
            return

        self._warn_if_overloaded()

        if self._auto_calculate_override and self._wire_hot_size_override:
            if self._try_override_hot_size(rating):
                return
            self.log_info("Override hot size rejected; falling back to automatic sizing.")

        self._auto_hot_sizing(rating)

    def calculate_neutral_wire_size(self):
        if self.cable.cleared or self.calc_failed:
            self.cable.neutral_size = None
            return

        behavior = getattr(self.settings, "neutral_behavior", NeutralBehavior.MATCH_HOT)

        if self._auto_calculate_override:
            if behavior == NeutralBehavior.MATCH_HOT:
                self.cable.neutral_size = self.cable.hot_size if self.cable.neutral_qty else None
                return
            if behavior == NeutralBehavior.MANUAL:
                if self._try_override_neutral_size():
                    return

        # USER OVERRIDE
        if self._auto_calculate_override and self._wire_neutral_size_override:
            if self._try_override_neutral_size():
                return  # done, override accepted

        # DEFAULT: neutral follows hot unless qty = 0
        if self.cable.neutral_qty == 0:
            self.cable.neutral_size = None
        else:
            self.cable.neutral_size = self.cable.neutral_size or self.cable.hot_size

    def calculate_ground_wire_size(self):
        """EGC sizing with fallback to table-lookup scaling."""
        if self.cable.cleared or self.calc_failed:
            self.cable.ground_size = None
            return

        # USER OVERRIDE
        if self._auto_calculate_override and self._wire_ground_size_override:
            if self._try_override_ground_size():
                return

        amps = self.breaker_rating
        if amps is None:
            self.cable.ground_size = None
            return

        material = self.cable.material or self._wire_info.get("wire_material", "CU")
        wire_info = self._wire_info or {}

        base_ground = wire_info.get("wire_ground_size")
        base_hot = wire_info.get("wire_hot_size")
        base_sets = wire_info.get("number_of_parallel_sets", 1) or 1

        calc_hot = self.cable.hot_size
        calc_sets = self.cable.sets or 1

        if not base_ground:
            egc_list = EGC_TABLE.get(material)
            if not egc_list:
                self.log_warning("EGC table missing for material {}.".format(material))
                self.cable.ground_size = None
                return
            for threshold, size in egc_list:
                if amps <= threshold:
                    self.cable.ground_size = size
                    return
            self.log_warning(
                "Breaker {}A exceeds EGC table; using largest EGC size {}.".format(
                    amps, egc_list[-1][1]
                )
            )
            self.cable.ground_size = egc_list[-1][1]
            return

        if not (base_hot and base_ground and calc_hot):
            self.log_warning(
                "Unable to scale ground size (missing base sizes); leaving ground blank."
            )
            self.cable.ground_size = None
            return

        base_hot_cmil = CONDUCTOR_AREA_TABLE.get(base_hot, {}).get("cmil")
        calc_hot_cmil = CONDUCTOR_AREA_TABLE.get(calc_hot, {}).get("cmil")
        base_ground_cmil = CONDUCTOR_AREA_TABLE.get(base_ground, {}).get("cmil")

        if not (base_hot_cmil and calc_hot_cmil and base_ground_cmil):
            self.log_warning("Ground size lookup failed for provided wire sizes; leaving blank.")
            self.cable.ground_size = None
            return

        total_base = base_sets * base_hot_cmil
        total_calc = calc_sets * calc_hot_cmil
        if total_base <= 0:
            self.cable.ground_size = None
            return

        new_ground_cmil = base_ground_cmil * (float(total_calc) / total_base)

        candidates = sorted(
            CONDUCTOR_AREA_TABLE.items(), key=lambda kv: kv[1].get("cmil", 0)
        )
        for wire, data in candidates:
            cmil = data.get("cmil")
            if cmil and cmil >= new_ground_cmil:
                self.cable.ground_size = wire
                return

        self.cable.ground_size = None

    def calculate_conduit_size(self):
        """Size conduit (or apply override) using CableSet + ConduitRun."""
        if (self.cable.cleared and not self._user_clear_hot) or self.calc_failed:
            self._clear_conduit_data()
            return

        if self._user_clear_conduit:
            self._clear_conduit_data()
            self.conduit.cleared = True
            return

        self.conduit.cleared = False
        conduit_type = self._resolve_conduit_type()
        if not conduit_type:
            self._clear_conduit_data()
            return

        if not self.conduit.set_type_from_value(conduit_type):
            self._clear_conduit_data()
            return

        total_area = self.cable.get_total_area()

        # override path first
        if self._auto_calculate_override and self._conduit_size_override:
            size_norm = self._normalize_conduit_type(self._conduit_size_override)
            if self.conduit.apply_override_size(size_norm, total_area):
                if self.conduit.fill_ratio and self.conduit.fill_ratio > self.settings.max_conduit_fill:
                    self.log_warning(
                        "Override conduit size {} exceeds max fill ({} > {}).".format(
                            size_norm, round(self.conduit.fill_ratio, 3), self.settings.max_conduit_fill
                        )
                    )
                return
            else:
                self.log_warning(
                    "Invalid conduit size override '{}'; calculating instead.".format(
                        self._conduit_size_override
                    )
                )

        # auto sizing path
        if not self.conduit.pick_size(total_area, self.settings):
            self.log_warning(
                "{}: No conduit size fits total area {:.4f} at max fill {}.".format(
                    self.name, total_area, self.settings.max_conduit_fill
                )
            )
            self.conduit.calc_failed = True
            self.calc_failed = True
            self._clear_conduit_data()
    def _try_override_hot_size(self, rating):
        override = self._normalize_wire_size(self._wire_hot_size_override)
        if not override:
            return False

        if override not in ALLOWED_WIRE_SIZES:
            self.log_warning(
                "Hot size override {} is invalid; using auto-sizing while keeping your material/insulation overrides."
                .format(self._wire_hot_size_override)
            )
            return False

        material = self._wire_material_override or self.cable.material or "CU"
        temp_c = self.cable.temp_c or 75
        wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp_c, [])

        sets = self._wire_sets_override or self.cable.sets or 1

        for w, ampacity in wire_set:
            if w != override:
                continue

            total_amp = ampacity * sets

            # Accept override even if it fails VD or breaker, but warn
            if not self._is_ampacity_acceptable(rating, total_amp, self.circuit_load_current):
                self.log_warning("Override {} set(s) x #{}  fails ampacity ({} A). Saving anyway."
                                 .format( sets, w,total_amp))

            vd = self._safe_voltage_drop_calc(w, sets)
            if vd is not None and vd > self.max_voltage_drop:
                self.log_warning("Override {} set(s) x #{}  fails volt drop check ({}%). Saving anyway."
                                 .format( sets,w, round(100*vd,2)))

            # ACCEPT regardless (but with warnings)
            self.cable.hot_size = w
            self.cable.sets = sets
            self.cable.base_ampacity = ampacity
            self.cable.total_ampacity = total_amp
            self.cable.voltage_drop = vd
            return True

        self.log_warning("Hot override {} not found in ampacity table."
                         .format(self._wire_hot_size_override))
        return False

    def _try_override_neutral_size(self):
        override = self._normalize_wire_size(self._wire_neutral_size_override)
        if not override:
            return False

        # allow neutral to differ from hot
        if override in ALLOWED_WIRE_SIZES:
            self.cable.neutral_size = override
            return True

        self.log_warning("Neutral size override '{}' invalid; using calculated neutral size."
                         .format(self._wire_neutral_size_override))
        return False

    def _try_override_ground_size(self):
        override = self._normalize_wire_size(self._wire_ground_size_override)
        if not override:
            return False

        if override in ALLOWED_WIRE_SIZES:
            self.cable.ground_size = override
            return True

        self.log_warning(
            "Ground size override '{}' invalid; auto-calculating based on breaker/material.".format(
                self._wire_ground_size_override
            )
        )
        return False

    def _auto_hot_sizing(self, rating):
        """Automatic hot conductor sizing with allowed sizes + VD check."""
        wire_info = self._wire_info or {}
        if not wire_info:
            self._fail_cable_sizing("No wire_info defaults.")
            return

        temp_str = wire_info.get("wire_temperature_rating", "75 C")
        try:
            temp_c = int(str(temp_str).replace("C", "").strip())
        except Exception:
            temp_c = 75

        # Prefer user-provided material/insulation overrides even when hot size override
        # is rejected so ampacity and fill are based on user intent.
        material = self.cable.material or wire_info.get("wire_material", "CU")
        base_wire = wire_info.get("wire_hot_size")
        base_sets = wire_info.get("number_of_parallel_sets", 1) or 1
        max_size = wire_info.get("max_lug_size")
        max_sets = wire_info.get("max_lug_qty", 1) or 1
        max_feeder_size = wire_info.get("max_feeder_size")

        wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(self.cable.temp_c or temp_c, [])
        if not wire_set:
            self._fail_cable_sizing(
                "No ampacity table for {} at {} C.".format(material, temp_c)
            )
            return

        sets = base_sets
        any_lug_limit_hit = False
        solution_found = False

        while sets <= max_sets and not solution_found:
            start_index = 0
            if base_wire:
                for i, (w, _) in enumerate(wire_set):
                    if w == base_wire:
                        start_index = i
                        break

            reached_max_lug_size = False
            selected_over_soft_limit = False

            for wire, ampacity in wire_set[start_index:]:
                if wire not in ALLOWED_WIRE_SIZES:
                    continue

                over_soft_limit = False
                if max_feeder_size and self._is_wire_larger_than_limit(wire, max_feeder_size):
                    over_soft_limit = True

                if max_size and self._is_wire_larger_than_limit(wire, max_size):
                    reached_max_lug_size = True
                    any_lug_limit_hit = True
                    if sets < max_sets:
                        continue
                    else:
                        self.log_warning(
                            "Exceeded lug size block {} at max sets {}; selecting {} with warning.".format(
                                max_size, max_sets, wire
                            )
                        )

                if self._is_feeder and sets > 1 and self._is_wire_below_one_aught(wire):
                    self.log_warning(
                        "Skipping parallel set attempt with {} on feeder; conductors smaller than 1/0 cannot be paralleled.".format(
                            wire
                        )
                    )
                    continue

                total_amp = ampacity * sets

                if not self._is_ampacity_acceptable(rating, total_amp, self.circuit_load_current):
                    if max_size and wire == max_size:
                        reached_max_lug_size = True
                        any_lug_limit_hit = True
                    continue

                vd = self._safe_voltage_drop_calc(wire, sets)
                if vd is not None and vd > self.max_voltage_drop:
                    if max_size and wire == max_size:
                        reached_max_lug_size = True
                        any_lug_limit_hit = True
                    continue

                self.cable.hot_size = wire
                self.cable.sets = sets
                self.cable.base_ampacity = ampacity
                self.cable.total_ampacity = total_amp
                self.cable.voltage_drop = vd
                selected_over_soft_limit = over_soft_limit
                solution_found = True
                break

            sets += 1

        if selected_over_soft_limit and max_feeder_size:
            self.log_warning(
                "Exceeded max feeder size {} to satisfy ampacity/VD; selected {} instead.".format(
                    max_feeder_size, self.cable.hot_size
                )
            )

        if reached_max_lug_size and solution_found:
            self.log_warning(
                "Reached lug size block {} when sizing hots; continuing with allowable configuration.".format(max_size)
            )

        if not solution_found:
            msg = "Reached max lug qty {} and could not size hot conductor for breaker {} A.".format(
                max_sets, rating
            )
            if any_lug_limit_hit and max_size:
                msg += " Lug size block at {} prevented further upsizing.".format(max_size)
            self._fail_cable_sizing(msg)

    def _safe_voltage_drop_calc(self, wire_size, sets):
        try:
            return self.calculate_voltage_drop(wire_size, sets)
        except Exception as e:
            logger.debug(
                "Voltage drop calc failed for {} x {} sets: {}".format(
                    wire_size, sets, e
                )
            )
            return None

    def _fail_cable_sizing(self, msg):
        self.log_error(
            "{} Wire sizing failed. All cable outputs will be cleared.".format(msg)
        )
        self.cable.calc_failed = True
        self.calc_failed = True
        self._clear_cable_data()

    def _clear_cable_data(self):
        self.cable.clear()


    def _is_ampacity_acceptable(self, breaker_rating, ampacity, circuit_amps):
        """NEC 240.4(B) logic; same as before."""
        if circuit_amps is None:
            return False

        if ampacity < circuit_amps:
            return False

        if ampacity >= breaker_rating:
            return True

        if breaker_rating > 800:
            return False

        for std in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
            if std >= ampacity:
                return std >= breaker_rating
        return False

    def calculate_voltage_drop(self, wire_size_formatted, sets):
        if self.cable.cleared or self.calc_failed:
            return None
        try:
            length = self.length
            volts = self.voltage
            pf = self.power_factor or 0.9
            phase = self.phase
            amps = self._get_voltage_drop_current()

            if not amps or not length or not volts:
                return 0

            material = self.cable.material or self._wire_info.get("wire_material", "CU")
            conduit_material = self.conduit_material_type or self._wire_info.get(
                "conduit_material_type"
            )
            wire_size = self._normalize_wire_size(wire_size_formatted)

            impedance = WIRE_IMPEDANCE_TABLE.get(wire_size)
            if not impedance:
                logger.debug(
                    "{}: no impedance found for wire size {}".format(self.name, wire_size)
                )
                return None

            R = impedance['R'].get(material, {}).get(conduit_material)
            X = impedance['X'].get(conduit_material)
            if R is None or X is None:
                return None

            R = R / float(sets)
            X = X / float(sets)
            sin_phi = (1 - pf ** 2) ** 0.5

            if phase == 3:
                drop = (1.732 * amps * (R * pf + X * sin_phi) * length) / 1000.0
            else:
                drop = (2 * amps * (R * pf + X * sin_phi) * length) / 1000.0

            return drop / volts
        except Exception:
            return 0

    def get_downstream_demand_current(self):
        try:
            for el in self.circuit.Elements:
                if self._is_transformer_primary:
                    va_param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALESTLOAD_PARAM)
                    if va_param and va_param.HasValue:
                        raw_va = va_param.AsDouble()
                        demand_va = DB.UnitUtils.ConvertFromInternalUnits(
                            raw_va, DB.UnitTypeId.VoltAmperes
                        )
                        volts = self.voltage
                        phase = self.phase
                        if volts:
                            divisor = volts if phase == 1 else volts * 3 ** 0.5
                            return demand_va / divisor

                param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM)
                if param and param.StorageType == DB.StorageType.Double:
                    return param.AsDouble()
        except Exception:
            pass
        return None



    def _clear_conduit_data(self):
        self.conduit.clear()
        self.conduit.cleared = True

    def calculate_conduit_fill_percentage(self):
        """Kept for compatibility; all work is done in calculate_conduit_size()."""
        return self.conduit_fill_percentage

    # -----------------------------------------------------------------
    # Wire / conduit callout strings
    # -----------------------------------------------------------------
    def get_wire_set_string(self):
        if self.cable.cleared or self.calc_failed:
            return "-"

        wp = self.settings.wire_size_prefix or ""

        hot_size = self.cable.hot_size
        neut_size = self.cable.neutral_size or hot_size
        gnd_size = self.cable.ground_size
        ig_size = self.cable.ig_size or gnd_size

        hot_qty = self.cable.hot_qty or 0
        neut_qty = self.cable.neutral_qty or 0
        gnd_qty = self.cable.ground_qty or 0
        ig_qty = self.cable.ig_qty or 0

        parts = []

        if neut_qty and hot_size and neut_size:
            if hot_size == neut_size:
                combined = hot_qty + neut_qty
                if combined:
                    parts.append("{}{}{}".format(combined, wp, hot_size))
            else:
                if hot_qty:
                    parts.append("{}{}{}H".format(hot_qty, wp, hot_size))
                if neut_qty:
                    parts.append("{}{}{}N".format(neut_qty, wp, neut_size))
        else:
            if hot_qty and hot_size:
                parts.append("{}{}{}".format(hot_qty, wp, hot_size))

        if gnd_qty and gnd_size:
            parts.append("{}{}{}G".format(gnd_qty, wp, gnd_size))

        if ig_qty and ig_size:
            parts.append("{}{}{}IG".format(ig_qty, wp, ig_size))

        material = self.wire_material or ""
        suffix = material if material.upper() != "CU" else ""

        final = " + ".join(parts)
        if not final:
            return "-"
        return "{} {}".format(final, suffix) if suffix else final

    def get_wire_size_callout(self):
        if self.cable.cleared or self.calc_failed:
            return "-"

        sets = self.number_of_sets or 1
        wire_str = self.get_wire_set_string()
        if wire_str == "-":
            return "-"
        if sets > 1:
            return "({}) {}".format(sets, wire_str)
        return wire_str

    def get_conduit_and_wire_size(self):
        if (self.conduit.cleared or self.calc_failed) and (self.cable.cleared or self.calc_failed):
            return "-"

        sets = self.number_of_sets or 1
        prefix = "({}) ".format(sets) if sets > 1 else ""

        conduit_norm = self.conduit.size
        if not conduit_norm or self.conduit.cleared or self.calc_failed:
            return self.get_wire_size_callout()

        conduit_str = "{}{}".format(
            conduit_norm, self.settings.conduit_size_suffix or ""
        )

        wire_callout = self.get_wire_set_string()
        if wire_callout == "-":
            return "{}{}".format(prefix, conduit_str)

        return "{}{}-({})".format(prefix, conduit_str, wire_callout)

    # -----------------------------------------------------------------
    # Utility helpers
    # -----------------------------------------------------------------
    def _get_yesno(self, guid):
        try:
            param = self.circuit.get_Parameter(Guid(guid))
            if not param:
                return False
            return bool(param.AsInteger())
        except Exception as e:
            logger.debug("Failed to read yes/no {} on {}: {}".format(guid, self.name, e))
            return False

    def _get_param_value(self, guid):
        try:
            param = self.circuit.get_Parameter(Guid(guid))
            if not param:
                return None
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            if param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
            if param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            if param.StorageType == DB.StorageType.ElementId:
                return param.AsElementId()
        except Exception as e:
            logger.debug("Failed to read param {} on {}: {}".format(guid, self.name, e))
        return None

    def _normalize_wire_size(self, val):
        if not val:
            return None
        prefix = self.settings.wire_size_prefix or ""
        return str(val).replace(prefix, "").strip()

    def _normalize_conduit_type(self, val):
        if not val:
            return None
        suffix = self.settings.conduit_size_suffix or ""
        return str(val).replace(suffix, "").strip()

    def _wire_index(self, wire):
        try:
            return ALLOWED_WIRE_SIZES.index(wire)
        except ValueError:
            return -1

    def _is_wire_below_one_aught(self, wire):
        idx = self._wire_index(wire)
        threshold = self._wire_index("1/0")
        return idx != -1 and threshold != -1 and idx < threshold

    def _is_wire_larger_than_limit(self, wire, limit_wire):
        idx = self._wire_index(wire)
        limit_idx = self._wire_index(limit_wire)
        if idx == -1 or limit_idx == -1:
            return False
        return idx > limit_idx
