<<<<<<< ours
from __future__ import annotations

=======
>>>>>>> theirs
import Autodesk.Revit.DB.Electrical as DBE
from System import Guid
from pyrevit import DB, revit, script

from CEDElectrical.circuit_sizing.models.circuit_branch import (
    CircuitBranchModel,
    CircuitOverrides,
    CircuitSettings,
)
from CEDElectrical.refdata.ocp_cable_defaults import OCP_CABLE_DEFAULTS
from CEDElectrical.refdata.shared_params_table import SHARED_PARAMS

logger = script.get_logger()


class RevitCircuitReader:
    """Extracts Revit electrical system data into a pure model object."""

    def __init__(self, circuit, settings=None):
        self.circuit = circuit
        self.settings = settings or CircuitSettings()

    # --- Public API ---
<<<<<<< ours
    def to_model(self) -> CircuitBranchModel:
=======
    def to_model(self):
>>>>>>> theirs
        is_feeder, is_transformer_primary = self._is_feeder()
        is_transformer_secondary = self._is_transformer_secondary()
        is_space = self.circuit.CircuitType == DBE.CircuitType.Space
        is_spare = self.circuit.CircuitType == DBE.CircuitType.Spare

        branch_type = "BRANCH"
        if is_transformer_primary:
            branch_type = "TRANSFORMER_PRIMARY"
        elif is_transformer_secondary:
            branch_type = "TRANSFORMER_SECONDARY"
        elif is_feeder:
            branch_type = "FEEDER"
        elif is_space:
            branch_type = "SPACE"
        elif is_spare:
            branch_type = "SPARE"

        overrides = self._read_overrides()
        wire_info = self._wire_info()

        model = CircuitBranchModel(
            circuit_id=self.circuit.Id.IntegerValue,
            panel=getattr(self.circuit.BaseEquipment, "Name", "") if self.circuit.BaseEquipment else "",
            circuit_number=self.circuit.CircuitNumber,
            name="{}-{}".format(
                getattr(self.circuit.BaseEquipment, "Name", "") if self.circuit.BaseEquipment else "",
                self.circuit.CircuitNumber,
            ),
            branch_type=branch_type,
            rating=self._rating(),
            frame=self._frame(),
            length=self._length(),
            voltage=self._voltage(),
            apparent_power=self._apparent_power(),
            apparent_current=self._apparent_current(),
            circuit_load_current=self._circuit_load_current(is_feeder, is_transformer_primary),
            poles=self._poles(),
            phase=self._phase(),
            power_factor=self._power_factor(),
            load_name=self._load_name(),
            circuit_notes=self._circuit_notes(),
            wire_info=wire_info,
            overrides=overrides,
            settings=self.settings,
            is_feeder=is_feeder,
            is_spare=is_spare,
            is_space=is_space,
            is_transformer_primary=is_transformer_primary,
            is_transformer_secondary=is_transformer_secondary,
        )
        return model

    # --- Base properties ---
    def _load_name(self):
        try:
            return self.circuit.LoadName
        except Exception:
            return None

    def _rating(self):
        try:
            if self.circuit.SystemType == DBE.ElectricalSystemType.PowerCircuit and not (
                self.circuit.CircuitType == DBE.CircuitType.Space
            ):
                return self.circuit.Rating
        except Exception:
            return None

    def _frame(self):
        try:
            return self.circuit.Frame
        except Exception:
            return None

    def _length(self):
        try:
            if self.circuit.SystemType == DBE.ElectricalSystemType.PowerCircuit and self.circuit.CircuitType == DBE.CircuitType.Circuit:
                return self.circuit.Length
        except Exception:
            return None

    def _voltage(self):
        try:
            param = self.circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
            if param and param.HasValue:
                raw_volt = param.AsDouble()
                return DB.UnitUtils.ConvertFromInternalUnits(raw_volt, DB.UnitTypeId.Volts)
        except Exception:
            return None

    def _apparent_power(self):
        try:
            return DBE.ElectricalSystem.ApparentLoad.__get__(self.circuit)
        except Exception:
            return None

    def _apparent_current(self):
        try:
            return DBE.ElectricalSystem.ApparentCurrent.__get__(self.circuit)
        except Exception:
            return None

    def _poles(self):
        try:
            return DBE.ElectricalSystem.PolesNumber.__get__(self.circuit)
        except Exception:
            return None

    def _phase(self):
        poles = self._poles()
        if not poles:
            return 0
        return 3 if poles == 3 else 1

    def _power_factor(self):
        try:
            return DBE.ElectricalSystem.PowerFactor.__get__(self.circuit)
        except Exception:
            return None

    def _circuit_notes(self):
        try:
            param = self.circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
            if param and param.StorageType == DB.StorageType.String:
                return param.AsString() or ""
        except Exception as e:
            logger.debug("circuit_notes: {}".format(e))
        return ""

    def _circuit_load_current(self, is_feeder: bool, is_transformer_primary: bool):
        if self.circuit.CircuitType != DBE.CircuitType.Circuit:
            return None
        if is_feeder:
            demand = self._get_downstream_demand_current(is_transformer_primary)
            return demand
        return self._apparent_current()

    # --- Override handling ---
    def _get_yesno(self, guid):
        try:
            param = self.circuit.get_Parameter(Guid(guid))
            return bool(param.AsInteger()) if param else False
        except Exception:
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
        except Exception:
            return None

<<<<<<< ours
    def _read_overrides(self) -> CircuitOverrides:
=======
    def _read_overrides(self):
>>>>>>> theirs
        overrides = CircuitOverrides()
        try:
            overrides.include_neutral = self._get_yesno(SHARED_PARAMS['CKT_Include Neutral_CED']['GUID'])
            overrides.include_isolated_ground = self._get_yesno(SHARED_PARAMS['CKT_Include Isolated Ground_CED']['GUID'])
            overrides.auto_calculate = self._get_yesno(SHARED_PARAMS['CKT_User Override_CED']['GUID'])
        except Exception as e:
            logger.debug("override flags: {}".format(e))
            return overrides

        if not overrides.auto_calculate:
            return overrides

        try:
            overrides.breaker_override = self._get_param_value(SHARED_PARAMS['CKT_Rating_CED']['GUID'])
            overrides.wire_sets_override = self._get_param_value(SHARED_PARAMS['CKT_Number of Sets_CED']['GUID'])
            overrides.wire_hot_size_override = self._get_param_value(SHARED_PARAMS['CKT_Wire Hot Size_CEDT']['GUID'])
            overrides.wire_neutral_size_override = self._get_param_value(SHARED_PARAMS['CKT_Wire Neutral Size_CEDT']['GUID'])
            overrides.wire_ground_size_override = self._get_param_value(SHARED_PARAMS['CKT_Wire Ground Size_CEDT']['GUID'])
            overrides.conduit_type_override = self._get_param_value(SHARED_PARAMS['Conduit Type_CEDT']['GUID'])
            overrides.conduit_size_override = self._get_param_value(SHARED_PARAMS['Conduit Size_CEDT']['GUID'])
            overrides.wire_material_override = self._get_param_value(SHARED_PARAMS['Wire Material_CEDT']['GUID'])
            overrides.wire_temp_rating_override = self._get_param_value(SHARED_PARAMS['Wire Temperature Rating_CEDT']['GUID'])
            overrides.wire_insulation_override = self._get_param_value(SHARED_PARAMS['Wire Insulation_CEDT']['GUID'])
        except Exception as e:
            logger.debug("override params: {}".format(e))
        return overrides

    # --- Wire info ---
    def _wire_info(self):
        if self.circuit.SystemType != DBE.ElectricalSystemType.PowerCircuit:
            return {}

        rating = self._rating()
        if rating is None:
            return {}
        rating_key = int(rating)
        table = OCP_CABLE_DEFAULTS
        if rating_key in table:
            return table[rating_key]

        sorted_keys = sorted(table.keys())
        for key in sorted_keys:
            if key >= rating_key:
                return table[key]

        fallback_key = sorted_keys[-1]
        return table[fallback_key]

    # --- Feeder helpers ---
    def _is_feeder(self):
        is_feeder = False
        is_transformer_primary = False
        try:
            for el in list(self.circuit.Elements):
                if isinstance(el, DB.FamilyInstance):
                    family = el.Symbol.Family
                    part_type = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
                    if part_type and part_type.StorageType == DB.StorageType.Integer:
                        part_value = part_type.AsInteger()
                        if part_value == 15:
                            is_transformer_primary = True
                        if part_value in [14, 15, 16, 17]:
                            is_feeder = True
                            return is_feeder, is_transformer_primary
        except Exception as e:
            logger.debug("is_feeder: {}".format(e))
        return is_feeder, is_transformer_primary

    def _is_transformer_secondary(self):
        try:
            base_equipment = getattr(self.circuit, "BaseEquipment", None)
            if not base_equipment or not isinstance(base_equipment, DB.FamilyInstance):
                return False
            family = base_equipment.Symbol.Family
            part_type = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
            if part_type and part_type.StorageType == DB.StorageType.Integer:
                return part_type.AsInteger() == 15
        except Exception as e:
            logger.debug("is_transformer_secondary: {}".format(e))
        return False

    def _get_downstream_demand_current(self, is_transformer_primary):
        try:
            for el in self.circuit.Elements:
                if is_transformer_primary:
                    va_param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALESTLOAD_PARAM)
                    if va_param and va_param.HasValue:
                        raw_va = va_param.AsDouble()
                        demand_va = DB.UnitUtils.ConvertFromInternalUnits(raw_va, DB.UnitTypeId.VoltAmperes)
                        voltage = self._voltage()
                        if voltage:
                            divisor = voltage if self._phase() == 1 else voltage * 3 ** 0.5
                            return demand_va / divisor
                param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM)
                if param and param.StorageType == DB.StorageType.Double:
                    return param.AsDouble()
        except Exception as e:
            logger.debug("downstream demand current: {}".format(e))
        return None
