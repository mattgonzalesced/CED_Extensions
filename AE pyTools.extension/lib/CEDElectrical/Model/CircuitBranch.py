# -*- coding: utf-8 -*-

import Autodesk.Revit.DB.Electrical as DBE
from System import Guid
from pyrevit import DB, script, revit

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

PART_TYPE_MAP = {
    14: "Panelboard",
    15: "Transformer",
    16: "Switchboard",
    17: "Other Panel",
    18: "Equipment Switch"
}

# TODO: make sure circuit settings is set up correctly for use in the circuit branch class. these settings
#  should apply to every instance of circuit branch, so we need to make sure we're efficient and not recreating
#  data unnecessarily.

class CircuitSettings(object):
    def __init__(self):
        # User-adjustable settings (can be loaded from file or UI later)
        self.min_wire_size = '12'
        self.max_wire_size = '600'
        self.min_breaker_size = 20
        self.auto_calculate_breaker = False
        self.min_conduit_size = '3/4"'
        self.max_conduit_fill = .36
        self.max_branch_voltage_drop = .03
        self.max_feeder_voltage_drop = .02
        self.wire_size_prefix = '#'
        self.conduit_size_suffix = 'C'

    def to_dict(self):
        return {
            'min_wire_size': self.min_wire_size,
            'max_wire_size': self.max_wire_size,
            'min_breaker_size': self.min_breaker_size,
            'auto_calculate_breaker': self.auto_calculate_breaker
        }

    def load_from_dict(self, data):
        self.min_wire_size = data.get('min_wire_size', self.min_wire_size)
        self.max_wire_size = data.get('max_wire_size', self.max_wire_size)
        self.min_breaker_size = data.get('min_breaker_size', self.min_breaker_size)
        self.auto_calculate_breaker = data.get('auto_calculate_breaker', self.auto_calculate_breaker)


class CircuitBranch(object):
    def __init__(self, circuit, settings=None):

        self.circuit = circuit
        self.settings = settings if settings else CircuitSettings()
        self.circuit_id = circuit.Id.Value
        self.panel = getattr(circuit.BaseEquipment, 'Name', None) if circuit.BaseEquipment else ""
        self.circuit_number = circuit.CircuitNumber
        self.name = "{}-{}".format(self.panel, self.circuit_number)
        self._wire_info = self.wire_info  # Lazy-loaded wire info dictionary
        self._is_transformer_primary = False
        self._is_feeder = self.is_feeder

        # User overrides (None = no override)
        self._auto_calculate_override = False
        self._include_neutral = False
        self._include_isolated_ground = False
        self._get_override_flags()

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
        self._get_user_overrides()

        # Calculated values (set by calculation methods)
        self._calculated_breaker = None
        self._calculated_hot_wire = None
        self._calculated_wire_sets = None
        self._calculated_hot_ampacity = None
        self._calculated_ground_wire = None
        self._calculated_conduit_size = None
        self._calculated_conduit_fill = None



    # ----------- Classification -----------

    def log_info(self, msg, *args):
        logger.info("{}: {}".format(self.name, msg), *args)
   
    def log_warning(self, msg, *args):
        logger.warning("{}: {}".format(self.name, msg), *args)

    def log_debug(self, msg, *args):
        logger.debug("{}: {}".format(self.name, msg), *args)




    #------------Prevent Ridiculous User Overrides------------
    @property
    def reasonable_user_overrides(self):
        reasonable = True

        if self._wire_sets_override and isinstance(self._wire_sets_override, int):
            if self._wire_sets_override < 1 or self._wire_sets_override > 7:
                self.log_warning("Unreasonable wire sets override: {}. Resetting to None.".format(
                    self._wire_sets_override))
                self._wire_sets_override = None
                reasonable = False
        else:
            self._wire_sets_override = None
            reasonable = False

        if self._wire_material_override:
            if self._wire_material_override not in WIRE_AMPACITY_TABLE.keys():
                self.log_warning("Unreasonable wire material override: {}. Resetting to None.".format(
                    self._wire_material_override))
                self._wire_material_override = None
                reasonable = False

        if self._wire_temp_rating_override:
            try:
                temp = int(str(self._wire_temp_rating_override).replace('C', '').strip())
                if temp not in [60, 75, 90]:
                    self.log_warning("Unreasonable wire temperature rating override: {}. Resetting to None.".format(
                        self._wire_temp_rating_override))
                    self._wire_temp_rating_override = None
                    reasonable = False
            except Exception as e:
                self.log_warning("Invalid wire temperature rating override: {}. Resetting to None.".format(
                    self._wire_temp_rating_override))
                self._wire_temp_rating_override = None
                reasonable = False

        if self._wire_insulation_override:
            # For simplicity, just check if it's a non-empty string
            if not isinstance(self._wire_insulation_override, str) or not self._wire_insulation_override.strip():
                self.log_warning("Unreasonable wire insulation override: {}. Resetting to None.".format(
                    self._wire_insulation_override))
                self._wire_insulation_override = None
                reasonable = False

        if self._wire_hot_size_override:
            normalized = self._normalize_wire_size(self._wire_hot_size_override)
            if normalized not in CONDUCTOR_AREA_TABLE.keys():
                self.log_warning("Unreasonable hot wire size override: {}. Resetting to None.".format(
                    self._wire_hot_size_override))
                self._wire_hot_size_override = None
                reasonable = False 

        if self._wire_neutral_size_override:
            normalized = self._normalize_wire_size(self._wire_neutral_size_override)
            if normalized not in CONDUCTOR_AREA_TABLE.keys():
                self.log_warning("Unreasonable neutral wire size override: {}. Resetting to None.".format(
                    self._wire_neutral_size_override))
                self._wire_neutral_size_override = None
                reasonable = False
        
        if self._wire_ground_size_override:
            normalized = self._normalize_wire_size(self._wire_ground_size_override)
            if normalized not in CONDUCTOR_AREA_TABLE.keys():
                self.log_warning("Unreasonable ground wire size override: {}. Resetting to None.".format(
                    self._wire_ground_size_override))
                self._wire_ground_size_override = None
                reasonable = False
        
        if self._conduit_size_override:
            size = self._conduit_size_override
            if size not in CONDUIT_AREA_TABLE.keys():
                self.log_warning("Unreasonable conduit size override: {}. Resetting to None.".format(
                    self._conduit_size_override))
                self._conduit_size_override = None
                reasonable = False

        if self._conduit_type_override:
            if self._conduit_type_override not in CONDUIT_SIZE_INDEX.keys():
                self.log_warning("Unreasonable conduit type override: {}. Resetting to None.".format(
                    self._conduit_type_override))
                self._conduit_type_override = None
                reasonable = False

        
        return reasonable

    @property
    def branch_type(self):
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

    @property
    def is_feeder(self):
        self._is_feeder = False
        try:
            elements = list(self.circuit.Elements)
            logger.debug("🔍 Checking is_feeder for circuit: {}".format(self.name))

            for el in elements:
                if isinstance(el, DB.FamilyInstance):
                    family = el.Symbol.Family
                    part_type = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)

                    if part_type and part_type.StorageType == DB.StorageType.Integer:
                        part_value = part_type.AsInteger()
                        logger.debug("➡️ Element: {} (ID: {}), PART_TYPE: {}".format(el.Name, el.Id, part_value))

                        if part_value == 15:
                            self._is_transformer_primary = True
                            logger.debug("⚡ Transformer primary detected on circuit {}".format(self.name))

                        if part_value in [14, 15, 16, 17]:
                            self._is_feeder = True
                            logger.debug(
                                "✅ Marked as feeder (PART_TYPE: {}) for circuit {}".format(part_value, self.name))
                            return self._is_feeder

            logger.debug("❌ No feeder-type load found for circuit {}".format(self.name))
        except Exception as e:
            logger.debug("🚨 Error in is_feeder for circuit {}: {}".format(self.name, str(e)))

        return self._is_feeder

    @property
    def is_spare(self):
        return self.circuit.CircuitType == DBE.CircuitType.Spare

    @property
    def is_space(self):
        return self.circuit.CircuitType == DBE.CircuitType.Space

    # ----------- Wire Info Dictionary -----------

    @property
    def max_voltage_drop(self):
        if self._is_feeder:
            return self.settings.max_feeder_voltage_drop
        else:
            return self.settings.max_branch_voltage_drop


    @property
    def wire_info(self):
        if not self.is_power_circuit:
            return {}

        rating = self.rating
        if rating is None:
            logger.debug("⚠️ No rating found for circuit {}, wire info is empty.".format(self.name))
            self._wire_info = {}
            return self._wire_info

        rating_key = int(rating)
        table = OCP_CABLE_DEFAULTS

        if rating_key in table:
            self._wire_info = table[rating_key]
            return self._wire_info

        sorted_keys = sorted(table.keys())
        for key in sorted_keys:
            if key >= rating_key:
                self._wire_info = table[key]
                logger.debug("⚠️ No exact wire info match for {}, using next available: {}".format(rating_key, key))
                return self._wire_info

        # If no match or next larger found, use largest available
        fallback_key = sorted_keys[-1]
        self._wire_info = table[fallback_key]
        logger.debug("⚠️ Rating {} exceeds all defaults. Using max available: {}".format(rating_key, fallback_key))
        return self._wire_info

    # ----------- Circuit Properties -----------
    # TODO: need debug loggers in these try except blocks. this applies to many methods and properties
    #  we need clear understanding of what fails and we need to determine how to handle failures
    @property
    def load_name(self):
        try:
            return self.circuit.LoadName
        except:
            return None

    @property
    def rating(self):
        try:
            if self.is_power_circuit and not self.is_space:
                return self.circuit.Rating
        except:
            return None

    @property
    def frame(self):
        try:
            return self.circuit.Frame
        except:
            return None

    @property
    def circuit_notes(self):
        try:
            param = self.circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
            if param and param.StorageType == DB.StorageType.String:
                return param.AsString()
        except Exception as e:
            logger.debug("circuit_notes: {}".format(e))
        return ""

    @property
    def length(self):
        try:
            if self.is_power_circuit and not self.is_spare and not self.is_space:
                return self.circuit.Length
        except:
            return None

    @property
    def voltage(self):
        """Returns voltage in Volts, converted from internal units (kV)."""
        try:
            param = self.circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
            if param and param.HasValue:
                raw_volt = param.AsDouble()  # internal unit = kV
                volts = DB.UnitUtils.ConvertFromInternalUnits(raw_volt, DB.UnitTypeId.Volts)
                return volts
        except Exception as e:
            logger.debug("voltage conversion error: {}".format(e))
        return None

    @property
    def apparent_power(self):
        try:
            return DBE.ElectricalSystem.ApparentLoad.__get__(self.circuit)
        except:
            return None

    @property
    def apparent_current(self):
        try:
            return DBE.ElectricalSystem.ApparentCurrent.__get__(self.circuit)
        except:
            return None

    @property
    def circuit_load_current(self):
        if self.circuit.CircuitType == DBE.CircuitType.Circuit:
            if self._is_feeder:
                return self.get_downstream_demand_current()
            else:
                return self.apparent_current

    @property
    def poles(self):
        try:
            return DBE.ElectricalSystem.PolesNumber.__get__(self.circuit)
        except:
            return None

    @property
    def phase(self):
        if not self.poles:
            return 0

        if self.poles == 3:
            return 3
        else:
            return 1

    @property
    def power_factor(self):
        try:
            return DBE.ElectricalSystem.PowerFactor.__get__(self.circuit)
        except:
            return None

    # ----------- Override Setters -----------
    def _get_override_flags(self):
        """
        Reads shared Yes/No override flags from the circuit.
        This should be called once during init.
        """
        try:
            self._include_neutral = self._get_yesno(SHARED_PARAMS['CKT_Include Neutral_CED']['GUID'])
            self._include_isolated_ground = self._get_yesno(SHARED_PARAMS['CKT_Include Isolated Ground_CED']['GUID'])
            self._auto_calculate_override = self._get_yesno(SHARED_PARAMS['CKT_User Override_CED']['GUID'])
        except Exception as e:
            logger.debug("_get_override_flags: {}".format(e))

    def _get_user_overrides(self):
        """
        Pulls user-entered override values from Revit shared parameters
        if the auto-calculate flag is enabled.
        """
        if not self._auto_calculate_override:
            return

        try:
            self._breaker_override = self._get_param_value(SHARED_PARAMS['CKT_Rating_CED']['GUID'])
            self._wire_sets_override = self._get_param_value(SHARED_PARAMS['CKT_Number of Sets_CED']['GUID'])
            self._wire_hot_size_override = self._get_param_value(SHARED_PARAMS['CKT_Wire Hot Size_CEDT']['GUID'])
            self._wire_neutral_size_override = self._get_param_value(
                SHARED_PARAMS['CKT_Wire Neutral Size_CEDT']['GUID'])
            self._wire_ground_size_override = self._get_param_value(SHARED_PARAMS['CKT_Wire Ground Size_CEDT']['GUID'])
            self._conduit_type_override = self._get_param_value(SHARED_PARAMS['Conduit Type_CEDT']['GUID'])
            self._conduit_size_override = self._get_param_value(SHARED_PARAMS['Conduit Size_CEDT']['GUID'])
            self._wire_material_override = self._get_param_value(SHARED_PARAMS['Wire Material_CEDT']['GUID'])
            self._wire_temp_rating_override = self._get_param_value(
                SHARED_PARAMS['Wire Temperature Rating_CEDT']['GUID'])
            self._wire_insulation_override = self._get_param_value(SHARED_PARAMS['Wire Insulation_CEDT']['GUID'])
            logger.debug("got overrides")
        except Exception as e:
            logger.debug("_get_user_overrides failed: {}".format(e))

    # ----------- Public Access Properties -----------

    @property
    def breaker_rating(self):
        if self.is_space:
            return None

        if not self.settings.auto_calculate_breaker:
            return self.rating

        if self._breaker_override:
            return self._breaker_override

        else:
            return self._calculated_breaker

    @property
    def wire_material(self):
        if self._wire_material_override:
            return self._wire_material_override
        return self._wire_info.get('wire_material')

    @property
    def wire_temp_rating(self):
        if self._wire_temp_rating_override:
            return self._wire_temp_rating_override
        return self._wire_info.get('wire_temperature_rating')

    @property
    def wire_insulation(self):
        if self._wire_insulation_override:
            return self._wire_insulation_override
        return self._wire_info.get('wire_insulation')

    # @property
    # def wire_hot_size(self):
    #     if self._wire_hot_size_override:
    #         return self._wire_hot_size_override
    #     return self._wire_info.get('wire_hot_size')

    @property
    def hot_wire_quantity(self):
        if self.circuit.CircuitType == DBE.CircuitType.Circuit:
            return self.poles
        else:
            return 0

    @property
    def hot_wire_size(self):
        raw = self._wire_hot_size_override if self._auto_calculate_override else self._calculated_hot_wire
        size = self._normalize_wire_size(raw)
        if size and self.settings.wire_size_prefix:
            return "{}{}".format(self.settings.wire_size_prefix, size)
        return size

    @property
    def ground_wire_quantity(self):
        if self.circuit.CircuitType == DBE.CircuitType.Circuit:
            return 1
        else:
            return 0

    @property
    def ground_wire_size(self):
        raw = self._wire_ground_size_override if self._auto_calculate_override else self._calculated_ground_wire
        size = self._normalize_wire_size(raw)
        if size and self.settings.wire_size_prefix:
            return "{}{}".format(self.settings.wire_size_prefix, size)
        return size

    @property
    def neutral_wire_quantity(self):
        # Case 1: Always 1 for single-pole circuits
        if self.poles == 1:
            return 1

        # Case 3: Automatic check if it's a feeder with L-N voltage
        if self._is_feeder:
            doc = revit.doc
            try:
                for el in self.circuit.Elements:
                    if isinstance(el, DB.FamilyInstance):
                        ds_param = el.get_Parameter(DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
                        if ds_param and ds_param.HasValue:
                            ds_elem = doc.GetElement(ds_param.AsElementId())
                            if isinstance(ds_elem, DBE.DistributionSysType):
                                l_n_voltage = ds_elem.VoltageLineToGround
                                if l_n_voltage:
                                    return 1
                return 0  # No L-N voltage found
            except Exception as e:
                logger.debug("Feeder neutral check failed: {}".format(e))
                return 0

        # Case 2: Explicit override
        if self._include_neutral:
            return 1

        # Case 4: Default to no neutral
        return 0

    @property
    def neutral_wire_size(self):
        if self.neutral_wire_quantity == 0:
            return ""

        if self._auto_calculate_override:
            if self._wire_neutral_size_override is not None:
                raw = self._wire_neutral_size_override
            else:
                return self.hot_wire_size

            size = self._normalize_wire_size(raw)
            if size and self.settings.wire_size_prefix:
                return "{}{}".format(self.settings.wire_size_prefix, size)
        return self.hot_wire_size

    @property
    def isolated_ground_wire_quantity(self):
        if self._include_isolated_ground:
            return 1
        return 0

    @property
    def isolated_ground_wire_size(self):
        if self.isolated_ground_wire_quantity == 0:
            return ""
        return self.ground_wire_size

    @property
    def number_of_sets(self):
        if self._auto_calculate_override and self._wire_sets_override:
            return self._wire_sets_override
        return self._calculated_wire_sets

    @property
    def number_of_wires(self):
        return self.hot_wire_quantity + self.neutral_wire_quantity

    @property
    def circuit_base_ampacity(self):
        return self._calculated_hot_ampacity

    @property
    def voltage_drop_percentage(self):
        vdp = self.calculate_voltage_drop(self.hot_wire_size, self.number_of_sets)
        return vdp

    @property
    def conduit_material_type(self):
        if self._auto_calculate_override and self._conduit_type_override:
            return self.get_conduit_material_from_type()
        return self._wire_info.get('conduit_material_type')

    @property
    def conduit_type(self):
        if self._auto_calculate_override and self._conduit_type_override:
            return self._conduit_type_override
        return self._wire_info.get('conduit_type', "EMT")

    @property
    def conduit_size(self):
        raw = self._conduit_size_override if self._auto_calculate_override else self._calculated_conduit_size
        size = self._normalize_conduit_type(raw)
        if size and self.settings.conduit_size_suffix:
            return "{}{}".format(size, self.settings.conduit_size_suffix)
        return size

    @property
    def conduit_fill_percentage(self):
        return self.calculate_conduit_fill_percentage()

    # ----------- Calculations -----------

    def calculate_breaker_size(self):
        try:
            amps = self.apparent_current
            if amps:
                amps *= 1.25
                if amps < self.settings.min_breaker_size:
                    amps = self.settings.min_breaker_size

                for b in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
                    if b >= amps:
                        self._calculated_breaker = b
                        break
        except:
            self._calculated_breaker = None

    def calculate_hot_wire_size(self):
        rating = self.breaker_rating
        if rating is None:
            return

        # --- Handle user overrides directly ---
        if self._auto_calculate_override and self._wire_hot_size_override:
            self.log_info("overrides set for calc hot wire size.")
            try:
                material = self._wire_material_override
                temp = int(str(self._wire_temp_rating_override).replace('C', '').strip())
                sets = self._wire_sets_override or 1
                wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp, [])

                for wire, ampacity in wire_set:
                    if wire == self._normalize_wire_size(self._wire_hot_size_override):
                        self._calculated_hot_wire = wire
                        self._calculated_wire_sets = sets
                        self._calculated_hot_ampacity = ampacity * sets
                        logger.debug(
                            "Ampacity with override: wire={}, sets={}, ampacity per wire={}, total ampacity={}".format(
                                wire,
                                sets,
                                ampacity,
                                ampacity * sets
                            ))

                        return
            except Exception as e:
                logger.debug("Override ampacity calc failed: {}".format(e))
                return

        wire_info = self._wire_info
        try:
            temp = int(wire_info.get('wire_temperature_rating', '75').replace('C', '').strip())
        except Exception as e:
            logger.debug("Invalid wire temperature: {}".format(e))
            return

        material = wire_info.get('wire_material', 'CU')
        base_wire = wire_info.get('wire_hot_size')
        base_sets = wire_info.get('number_of_parallel_sets')
        max_size = wire_info.get('max_lug_size')
        max_sets = wire_info.get('max_lug_qty', 1)

        wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp, [])
        sets = base_sets or 1
        start_index = 0

        for i, (wire, _) in enumerate(wire_set):
            if wire == base_wire:
                start_index = i
                break

        while sets <= max_sets:
            reached_max_size = False
            for wire, ampacity in wire_set[start_index:]:
                total_ampacity = ampacity * sets

                # NEC 240.4(B) check here
                if not self._is_ampacity_acceptable(rating, total_ampacity, self.circuit_load_current):
                    continue

                try:
                    vd = self.calculate_voltage_drop(wire, sets)
                except Exception as e:
                    logger.debug("Voltage drop failed for {}x{}: {}".format(wire, sets, e))
                    vd = None

                if vd is None or vd <= self.max_voltage_drop:
                    self._calculated_hot_wire = wire
                    self._calculated_wire_sets = sets
                    self._calculated_hot_ampacity = total_ampacity
                    return

                if wire == max_size:
                    self._calculated_hot_wire = wire
                    self._calculated_wire_sets = sets
                    self._calculated_hot_ampacity = total_ampacity
                    reached_max_size = True
                    break

            if reached_max_size:
                logger.warning("{}: wire reached max size for breaker rating.".format(self.name))
                break
            sets += 1

            # fallback: use base wire size directly if no match
            if base_wire:
                for wire, ampacity in wire_set:
                    if wire == base_wire:
                        self._calculated_hot_wire = wire
                        self._calculated_wire_sets = base_sets
                        self._calculated_hot_ampacity = base_sets * ampacity
                        return

            self._calculated_hot_wire = None
            self._calculated_wire_sets = None
            self._calculated_hot_ampacity = None

    def calculate_ground_wire_size(self):
        try:
            amps = self.breaker_rating
            if amps is None:
                return

            wire_info = self._wire_info
            base_ground = wire_info.get('wire_ground_size')
            base_hot = wire_info.get('wire_hot_size')
            base_sets = wire_info.get('number_of_parallel_sets', 1)

            calc_hot = self._calculated_hot_wire
            calc_sets = self._calculated_wire_sets
            material = self.wire_material
            logger.debug("hot: {}, gnd: {}, calc hot: {}".format(base_hot, base_ground, calc_hot))
            # If base_ground is missing, fallback to EGC lookup
            if not base_ground:
                egc_list = EGC_TABLE.get(material)
                if egc_list:
                    for threshold, size in egc_list:
                        if amps <= threshold:
                            logger.debug("EGC lookup: material={}, rating={}, selected={}".format(material, amps, size))
                            self._calculated_ground_wire = size
                            return
                    # amps > all table entries
                    fallback = egc_list[-1][1]
                    logger.warning(
                        "{}: Breaker rating {}A exceeds EGC table. Using max ground size: {}".format(self.name, amps,
                                                                                                     fallback))
                    self._calculated_ground_wire = fallback
                    return
                else:
                    logger.warning("EGC table missing for material: {}".format(material))
                    self._calculated_ground_wire = None
                    return

            if not (base_ground and base_hot and calc_hot and calc_sets):
                self._calculated_ground_wire = None
                return

            # Get circular mils
            base_hot_cmil = CONDUCTOR_AREA_TABLE.get(base_hot, {}).get('cmil')
            calc_hot_cmil = CONDUCTOR_AREA_TABLE.get(calc_hot, {}).get('cmil')
            base_ground_cmil = CONDUCTOR_AREA_TABLE.get(base_ground, {}).get('cmil')

            if not all([base_hot_cmil, calc_hot_cmil, base_ground_cmil]):
                self._calculated_ground_wire = None
                return

            # Compute new ground CMIL
            total_base_hot_cmil = base_sets * base_hot_cmil
            total_calc_hot_cmil = calc_sets * calc_hot_cmil

            new_ground_cmil = base_ground_cmil * (float(total_calc_hot_cmil) / total_base_hot_cmil)

            # Find matching or next-larger ground wire
            for wire, data in sorted(CONDUCTOR_AREA_TABLE.items(), key=lambda x: x[1]['cmil']):
                cmil = data['cmil']
                if cmil >= new_ground_cmil:
                    self._calculated_ground_wire = wire
                    return

            self._calculated_ground_wire = None

        except Exception:
            self._calculated_ground_wire = None

    def _is_ampacity_acceptable(self, breaker_rating, ampacity, circuit_amps):
        """
        Returns True if ampacity is acceptable per NEC 240.4(B) exception.
        Allows undersized ampacity for breakers <= 800A if:
        - Ampacity is >= load
        - No smaller standard breaker fits above ampacity
        """
        # Basic check: Ampacity must cover the load regardless
        if ampacity < circuit_amps:
            return False

        # Normal compliance
        if ampacity >= breaker_rating:
            return True

        # Exception rule applies only to breakers 800A and below
        if breaker_rating > 800:
            return False

        # Check if breaker is the next standard size above ampacity
        for std_rating in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
            if std_rating >= ampacity:
                return std_rating >= breaker_rating  # True if no smaller frame fits
        return False

    def calculate_voltage_drop(self, wire_size_formatted, sets=1):
        try:
            length = self.length
            voltage = self.voltage
            pf = self.power_factor or 0.9
            phase = self.phase
            amps = self.circuit_load_current

            if not amps or not length or not voltage:
                return 0

            wire_info = self._wire_info
            material = wire_info.get('wire_material', 'CU')
            conduit_material = self.conduit_material_type
            wire_size = self._normalize_wire_size(wire_size_formatted)

            impedance = WIRE_IMPEDANCE_TABLE.get(wire_size)
            if not impedance:
                logger.debug("{}: no impedance found for wire size {}".format(self.name, wire_size))
                return None

            R = impedance['R'].get(material, {}).get(conduit_material)
            X = impedance['X'].get(conduit_material)
            if R is None or X is None:
                return None

            R = R / sets
            X = X / sets
            sin_phi = (1 - pf ** 2) ** 0.5

            if phase == 3:
                drop = (1.732 * amps * (R * pf + X * sin_phi) * length) / 1000.0
            else:
                drop = (2 * amps * (R * pf + X * sin_phi) * length) / 1000.0

            return (drop / voltage)
        except Exception:
            return 0

    def get_downstream_demand_current(self):
        try:
            logger.debug("🔍 Checking downstream demand current for circuit: {}".format(self.name))

            for el in self.circuit.Elements:
                logger.debug("➡️ Inspecting element: {} (ID: {})".format(el.Name, el.Id))

                # Transformer-specific calculation
                if self._is_transformer_primary:
                    logger.debug("⚡ Detected transformer primary on circuit: {}".format(self.name))
                    va_param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALESTLOAD_PARAM)
                    if va_param and va_param.HasValue:
                        raw_va = va_param.AsDouble()
                        logger.debug("✅ Found TOTALESTLOAD_PARAM: {} kVA (internal)".format(raw_va))

                        demand_va = DB.UnitUtils.ConvertFromInternalUnits(raw_va, DB.UnitTypeId.VoltAmperes)
                        logger.debug("🔧 Converted VA: {} VA".format(demand_va))

                        voltage = self.voltage
                        phase = self.phase
                        logger.debug("🔌 Voltage: {} V, Phase: {}".format(voltage, phase))

                        if voltage and demand_va:
                            divisor = voltage if phase == 1 else voltage * 3 ** 0.5
                            self._demand_current = demand_va / divisor
                            logger.debug(
                                "✅ Calculated transformer primary current: {:.2f} A".format(self._demand_current))
                            return self._demand_current
                        else:
                            logger.debug("⚠️ Missing voltage or demand_va for current calculation.")

                    else:
                        logger.debug("❌ Missing or invalid TOTALESTLOAD_PARAM on transformer element.")

                # Default panel total demand fallback
                param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM)
                if param and param.StorageType == DB.StorageType.Double:
                    self._demand_current = param.AsDouble()
                    logger.debug("✅ Found panel TOTAL_DEMAND_CURRENT: {:.2f} A".format(self._demand_current))
                    return self._demand_current
                else:
                    logger.debug("❌ No TOTAL_DEMAND_CURRENT_PARAM or invalid storage type on element.")

        except Exception as e:
            logger.debug("🚨 Exception in get_downstream_demand_current: {}".format(str(e)))

        self._demand_current = None
        logger.debug("❌ No valid demand current found for circuit: {}".format(self.name))
        return None

    def get_conduit_material_from_type(self):
        conduit_type = self.conduit_type  # use property
        for material, type_dict in CONDUIT_AREA_TABLE.items():
            if conduit_type in type_dict:
                return material
        return None

    def calculate_conduit_size(self):
        self._calculated_conduit_size = None
        self._calculated_conduit_fill = None

        wire_info = self._wire_info
        insulation = self.wire_insulation
        conduit_material = self.conduit_material_type
        conduit_type = self.conduit_type  # already exists

        if not insulation or not conduit_material or not conduit_type:
            logger.debug("Missing wire_info insulation or conduit data.")
            return

        sizes_and_qtys = [
            (self._normalize_wire_size(self.hot_wire_size), self.hot_wire_quantity),
            (self._normalize_wire_size(self.neutral_wire_size), self.neutral_wire_quantity),
            (self._normalize_wire_size(self.ground_wire_size), self.ground_wire_quantity),
            (self._normalize_wire_size(self.isolated_ground_wire_size), self.isolated_ground_wire_quantity)
        ]

        total_area = sum(
            CONDUCTOR_AREA_TABLE[size]['area'][insulation] * qty
            for size, qty in sizes_and_qtys
            if size and qty and size in CONDUCTOR_AREA_TABLE and insulation in CONDUCTOR_AREA_TABLE[size]['area']
        )

        conduit_table = CONDUIT_AREA_TABLE.get(conduit_material, {}).get(conduit_type, {})
        if not conduit_table:
            logger.debug("No conduit table found for {} / {}".format(conduit_material, conduit_type))
            return

        enum = CONDUIT_SIZE_INDEX
        if self.settings.min_conduit_size not in enum:
            logger.warning("Invalid min conduit size setting: {}".format(self.settings.min_conduit_size))
            return

        min_index = enum.index(self.settings.min_conduit_size)

        for size in enum[min_index:]:
            area = conduit_table.get(size)
            if area and area * self.settings.max_conduit_fill >= total_area:
                self._calculated_conduit_size = size
                self._calculated_conduit_fill = round(total_area / area, 5)  # ⚠️ decimal, not percent
                return

        logger.warning("{}: No conduit size found that fits total area {:.4f}".format(self.name, total_area))

    def calculate_conduit_fill_percentage(self):
        conduit_formatted = self._conduit_size_override if self._auto_calculate_override else self._calculated_conduit_size
        conduit_size = self._normalize_conduit_type(conduit_formatted)
        wire_info = self._wire_info
        insulation = self.wire_insulation
        conduit_material = self.conduit_material_type
        conduit_type = self.conduit_type

        if not conduit_size or not conduit_material or not conduit_type or not insulation:
            return None

        conduit_area = CONDUIT_AREA_TABLE.get(conduit_material, {}).get(conduit_type, {}).get(conduit_size)
        if not conduit_area:
            return None

        sizes_and_qtys = [
            (self._normalize_wire_size(self.hot_wire_size), self.hot_wire_quantity),
            (self._normalize_wire_size(self.neutral_wire_size), self.neutral_wire_quantity),
            (self._normalize_wire_size(self.ground_wire_size), self.ground_wire_quantity),
            (self._normalize_wire_size(self.isolated_ground_wire_size), self.isolated_ground_wire_quantity)
        ]
        # TODO: need to build protections in for sizes that are NOT found! need to decide if this should happen here
        #  or in the wire size attributes
        total_area = sum(
            CONDUCTOR_AREA_TABLE[size]['area'][insulation] * qty
            for size, qty in sizes_and_qtys
            if size and qty and size in CONDUCTOR_AREA_TABLE and insulation in CONDUCTOR_AREA_TABLE[size]['area']
        )

        return round(total_area / conduit_area, 5)  # ⚠️ keep as decimal

    # TODO: need to handle cases where the Neutral does NOT equal hot. it needs to be
    #  a separate chunk, cant just be number of wires.
    #  example if 3 hots and 1 neutral, and hot_size = neutral_size: 4#12
    #  example if 3 hots and 1 neutral, and hot_size not = neutral_size: 3#12H + 1#10N
    def get_wire_set_string(self):
        wp = self.settings.wire_size_prefix or ''
        parts = []

        total_hn_qty = self.number_of_wires  # hot + neutral quantity
        hot_size = self._normalize_wire_size(self.hot_wire_size)

        if total_hn_qty and hot_size:
            parts.append("{}{}{}".format(total_hn_qty, wp, hot_size))

        if self.ground_wire_quantity:
            grd_size = self._normalize_wire_size(self.ground_wire_size)
            parts.append("{}{}{}G".format(self.ground_wire_quantity, wp, grd_size))

        if self.isolated_ground_wire_quantity:
            ig_size = self._normalize_wire_size(self.isolated_ground_wire_size)
            parts.append("{}{}{}IG".format(self.isolated_ground_wire_quantity, wp, ig_size))

        material = self.wire_material
        suffix = material if material != "CU" else ""

        return "{} {}".format(" + ".join(parts), suffix) if suffix else " + ".join(parts)

    def get_wire_size_callout(self):
        sets = self.number_of_sets or 1
        wire_set_string = self.get_wire_set_string()

        if sets > 1:
            return "({}) {}".format(sets, wire_set_string)
        else:
            return wire_set_string

    def get_conduit_and_wire_size(self):
        sets = self.number_of_sets or 1
        prefix = "{} SETS - ".format(sets) if sets > 1 else ""

        conduit = self._normalize_conduit_type(self.conduit_size)
        if not conduit:
            return None

        conduit = '{}{}'.format(conduit, self.settings.conduit_size_suffix or '')

        wire_callout = self.get_wire_set_string()
        return "{}{}-({})".format(prefix, conduit, wire_callout)

    def print_info(self, include_wire_info=True, include_all_properties=False):
        print("\n=== CircuitBranch: {} (ID: {}) ===".format(self.name, self.circuit_id))

        # Wire Info
        if include_wire_info:
            print("\nWire Info:")
            if self.wire_info:
                for key, val in self.wire_info.items():
                    print("    {}: {}".format(key, val if val else "N/A"))
            else:
                print("    No wire info available.")

        print("Feeder: {}".format(self._is_feeder))
        if self._is_feeder:
            current_source = self.get_downstream_demand_current()
            print("Current used for voltage drop: {:.2f} A (Feeder demand)".format(current_source or 0.0))
        else:
            print("Current used for voltage drop: {:.2f} A (Circuit apparent)".format(self.apparent_current or 0.0))

        # Circuit Info
        print("\nCircuit Info:")
        info_fields = [
            'rating', 'voltage', 'length',
            'apparent_power', 'apparent_current',
            'phase', 'power_factor'
        ]
        for attr in info_fields:
            try:
                value = getattr(self, attr)
                print("    {}: {}".format(attr, value if value is not None else "N/A"))
            except:
                print("    {}: [Error]".format(attr))

        if self._calculated_hot_wire and self._calculated_wire_sets:
            # First wire tried: Revit rating, min wire size

            base_wire = self._wire_info.get('wire_hot_size')
            base_sets = self._wire_info.get('number_of_parallel_sets', 1)
            initial_vd = self.calculate_voltage_drop(base_wire, base_sets)
            final_vd = self.calculate_voltage_drop(self._calculated_hot_wire, self._calculated_wire_sets)

            print("\nVoltage Drop Calculation:")
            print("    Base breaker size (from Revit): {} A".format(self.rating))
            print("    Initial VD ({} x {} set(s)): {:.2f}%".format(base_wire, base_sets, initial_vd or 0))
            print("    Final wire: {} x {} set(s)".format(self._calculated_hot_wire, self._calculated_wire_sets))
            print("    Final voltage drop: {:.2f}% (Max allowed: {}%)".format(final_vd or 0, self.max_voltage_drop))
        else:
            print("\nVoltage Drop Calculation: [⚠️ Sizing failed or skipped]")

        # Calculation Results
        print("\nCalculated/Resolved Values:")
        print("    Breaker Rating: {}".format(self.breaker_rating or "N/A"))
        print("    Hot Wire Size: {}".format(self.hot_wire_size or "N/A"))
        print("    Number of Sets: {}".format(self.number_of_sets or "N/A"))
        print("    Circuit Base Ampacity: {}".format(self.circuit_base_ampacity or "N/A"))
        print("    Ground Wire Size: {}".format(self.ground_wire_size or "N/A"))

        # Quantities
        print("\nWire Quantities:")
        print("    Hot Conductors: {}".format(self.hot_wire_quantity))
        print("    Ground Conductors: {}".format(self.ground_wire_quantity))
        print("    Neutral Conductors: {}".format(self.neutral_wire_quantity))
        print("    Isolated Ground Conductors: {}".format(self.isolated_ground_wire_quantity))

    def _get_yesno(self, guid):
        try:
            param = self.circuit.get_Parameter(Guid(guid))
            if param:
                logger.debug(("✅ Param found:", param.Definition.Name, "=", param.AsInteger()))

            return bool(param.AsInteger()) if param else False
        except Exception as e:
            logger.debug("Failed to read param {}: {}".format(guid, e))
            return False

    def _get_param_value(self, guid):
        try:
            param = self.circuit.get_Parameter(Guid(guid))

            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.ElementId:
                return param.AsElementId()
        except Exception as e:
            logger.debug("Failed to read param GUID {}: {}".format(guid, e))
            return None

    def _normalize_wire_size(self, val):
        prefix = self.settings.wire_size_prefix
        return val.replace(prefix, '').strip() if val else None

    def _normalize_conduit_type(self, val):
        suffix = self.settings.conduit_size_suffix
        return val.replace(suffix, '').strip() if val else None
