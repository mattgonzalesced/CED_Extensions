# -*- coding: utf-8 -*-
"""Adapters that extract Revit data into domain models."""

from System import Guid
import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Model.CircuitBranch import CircuitBranchModel, CircuitSettings
from CEDElectrical.refdata.shared_params_table import SHARED_PARAMS


class RevitCircuitReader(object):
    def __init__(self, doc, settings=None):
        self.doc = doc
        self.settings = settings or CircuitSettings()

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
        return CircuitBranchModel(circuit, self.settings, overrides)

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
