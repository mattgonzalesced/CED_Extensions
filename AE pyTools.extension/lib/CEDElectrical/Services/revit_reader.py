# -*- coding: utf-8 -*-
"""Adapters that extract Revit data into domain models."""

from System import Guid
import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Model.CircuitBranch import CircuitBranchModel, CircuitSettings
from CEDElectrical.refdata.shared_params_table import SHARED_PARAMS
from CEDElectrical.refdata.ocp_cable_defaults import OCP_CABLE_DEFAULTS


class RevitCircuitReader(object):
    def __init__(self, doc, settings=None):
        self.doc = doc
        self.settings = settings or CircuitSettings()

    def _get_parameter_value(self, element, param_name):
        try:
            param = element.LookupParameter(param_name)
            if param:
                if param.StorageType == DB.StorageType.String:
                    return param.AsString()
                if param.StorageType == DB.StorageType.Integer:
                    return param.AsInteger()
                if param.StorageType == DB.StorageType.Double:
                    return param.AsDouble()
        except Exception:
            pass
        return None

    def _get_length_in_feet(self, circuit):
        try:
            length_internal = circuit.Length
        except Exception:
            length_internal = None
        if length_internal is None:
            try:
                param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM)
                if param:
                    length_internal = param.AsDouble()
            except Exception:
                length_internal = None
        if length_internal is None:
            return None
        try:
            return DB.UnitUtils.ConvertFromInternalUnits(length_internal, DB.DisplayUnitType.DUT_FEET)
        except Exception:
            return length_internal

    def _base_wire_info(self, rating):
        if rating is None:
            return {}
        try:
            rating_key = int(rating)
        except Exception:
            return {}

        table = OCP_CABLE_DEFAULTS
        if rating_key in table:
            return table[rating_key]

        sorted_keys = sorted(table.keys())
        for key in sorted_keys:
            if key >= rating_key:
                return table[key]
        if sorted_keys:
            return table[sorted_keys[-1]]
        return {}

    def _get_yes_no(self, element, guid):
        try:
            param = element.get_Parameter(Guid(guid))
            if param and param.StorageType == DB.StorageType.Integer:
                return bool(param.AsInteger())
        except Exception:
            pass
        return None

    def _collect_overrides(self, circuit):
        overrides = {}
        flags = {
            'auto_calculate_override': 'CKT_User Override_CED',
            'include_neutral': 'CKT_Include Neutral_CED',
            'include_isolated_ground': 'CKT_Include Isolated Ground_CED'
        }
        for key, param_name in flags.items():
            if param_name in SHARED_PARAMS:
                guid = SHARED_PARAMS[param_name].get('GUID')
                if guid:
                    overrides[key] = self._get_yes_no(circuit, guid)
        return overrides

    def read(self, circuit):
        overrides = self._collect_overrides(circuit)
        base_equipment = getattr(circuit, 'BaseEquipment', None)
        panel = getattr(base_equipment, 'Name', '') if base_equipment else ''

        branch_type = 'Unknown'
        try:
            if circuit.SystemType:
                branch_type = str(circuit.SystemType)
        except Exception:
            pass

        rating = getattr(circuit, 'Rating', None)
        frame = getattr(circuit, 'Frame', None)
        length = self._get_length_in_feet(circuit)
        circuit_notes = self._get_parameter_value(circuit, 'CKT_Schedule Notes_CEDT')
        voltage = getattr(circuit, 'Voltage', None)
        apparent_current = getattr(circuit, 'ApparentCurrent', None)
        apparent_power = getattr(circuit, 'ApparentLoad', None)
        power_factor = getattr(circuit, 'PowerFactor', None)
        poles = getattr(circuit, 'PolesNumber', None) or 0

        is_power_circuit = False
        is_feeder = False
        try:
            is_power_circuit = circuit.CircuitType == DBE.CircuitType.Circuit
            is_feeder = bool(circuit.IsFeedToPanel)
        except Exception:
            pass

        base_wire_info = self._base_wire_info(rating)

        return CircuitBranchModel(
            circuit_id=circuit.Id.IntegerValue if hasattr(circuit, 'Id') else None,
            panel=panel,
            circuit_number=getattr(circuit, 'CircuitNumber', None),
            branch_type=branch_type,
            load_name=getattr(circuit, 'LoadName', None),
            rating=rating,
            frame=frame,
            length=length,
            circuit_notes=circuit_notes,
            voltage=voltage,
            apparent_current=apparent_current,
            apparent_power=apparent_power,
            power_factor=power_factor,
            poles=poles,
            include_neutral=overrides.get('include_neutral'),
            include_isolated_ground=overrides.get('include_isolated_ground'),
            auto_calculate_override=overrides.get('auto_calculate_override', False),
            base_wire_info=base_wire_info,
            overrides=overrides,
            is_power_circuit=is_power_circuit,
            is_feeder=is_feeder,
            settings=self.settings
        )

    def get_selected_circuits(self, selection, picker):
        circuits = []
        if not selection:
            return picker()

        for el in selection:
            if isinstance(el, DBE.ElectricalSystem):
                circuits.append(el)
        if circuits:
            return circuits
        return picker()
