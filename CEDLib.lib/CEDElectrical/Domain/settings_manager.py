# -*- coding: utf-8 -*-
from pyrevit import DB

from CEDElectrical.Model.circuit_settings import CircuitSettings

GP_NAME = "CED_Circuit_Settings"


# ---------------------------
# INTERNAL HELPERS
# ---------------------------

def _get_global_param(doc):
    """Return existing global parameter Element or None."""
    gp_id = DB.GlobalParametersManager.FindByName(doc, GP_NAME)
    if gp_id:
        return doc.GetElement(gp_id)
    return None


def _create_global_param(doc):
    """Create a new global text parameter and return it."""
    spec = DB.SpecTypeId.String.Text  # text parameter spec
    gp = DB.GlobalParameter.Create(doc, GP_NAME, spec)
    return gp


def _get_or_create_global_param(doc):
    gp = _get_global_param(doc)
    if gp:
        return gp
    return _create_global_param(doc)


# ---------------------------
# PUBLIC API
# ---------------------------

def load_circuit_settings(doc):
    """Return a CircuitSettings instance using stored GP JSON (or defaults)."""
    gp = _get_or_create_global_param(doc)

    value_obj = gp.GetValue()
    if value_obj and isinstance(value_obj, DB.StringParameterValue):
        json_text = value_obj.Value
    else:
        json_text = None

    return CircuitSettings.from_json(json_text)


def save_circuit_settings(doc, settings):
    """Write settings JSON back into the global parameter."""
    gp = _get_or_create_global_param(doc)
    json_text = settings.to_json()

    spv = DB.StringParameterValue(json_text)
    gp.SetValue(spv)
