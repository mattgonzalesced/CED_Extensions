# -*- coding: utf-8 -*-
from pyrevit import DB, script

from CEDElectrical.Model.circuit_settings import CircuitSettings

GP_NAME = "CED_Circuit_Settings"

RESULT_PARAM_NAMES = [
    'CKT_Circuit Type_CEDT',
    'CKT_Panel_CEDT',
    'CKT_Circuit Number_CEDT',
    'CKT_Load Name_CEDT',
    'CKT_Rating_CED',
    'CKT_Frame_CED',
    'CKT_Length_CED',
    'CKT_Schedule Notes_CEDT',
    'Voltage Drop Percentage_CED',
    'CKT_Wire Hot Size_CEDT',
    'CKT_Number of Wires_CED',
    'CKT_Number of Sets_CED',
    'CKT_Wire Hot Quantity_CED',
    'CKT_Wire Ground Size_CEDT',
    'CKT_Wire Ground Quantity_CED',
    'CKT_Wire Neutral Size_CEDT',
    'CKT_Wire Neutral Quantity_CED',
    'CKT_Wire Isolated Ground Size_CEDT',
    'CKT_Wire Isolated Ground Quantity_CED',
    'Wire Material_CEDT',
    'Wire Temparature Rating_CEDT',
    'Wire Insulation_CEDT',
    'Conduit Size_CEDT',
    'Conduit Type_CEDT',
    'Conduit Fill Percentage_CED',
    'Wire Size_CEDT',
    'Conduit and Wire Size_CEDT',
    'Circuit Load Current_CED',
    'Circuit Ampacity_CED',
    'CKT_Length Makeup_CED',
]


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
    t = DB.Transaction(doc, "Create {}".format(GP_NAME))
    t.Start()
    gp = DB.GlobalParameter.Create(doc, GP_NAME, spec)
    t.Commit()
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
    t = DB.Transaction(doc, "Save {}".format(GP_NAME))
    t.Start()
    gp.SetValue(spv)
    t.Commit()


def _clear_param(param):
    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            param.Set("")
        elif st == DB.StorageType.Integer:
            param.Set(0)
        elif st == DB.StorageType.Double:
            param.Set(0.0)
        elif st == DB.StorageType.ElementId:
            param.Set(DB.ElementId.InvalidElementId)
        return True
    except Exception:
        return False


def clear_downstream_results(doc, clear_equipment=False, clear_fixtures=False, logger=None):
    """Blank all stored circuit data on downstream elements after toggles are disabled."""
    if not (clear_equipment or clear_fixtures):
        return 0, 0

    logger = logger or script.get_logger()
    cleared_equipment = 0
    cleared_fixtures = 0

    t = DB.Transaction(doc, "Clear downstream circuit data")
    t.Start()
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance)
        for el in collector:
            cat = el.Category
            if not cat:
                continue

            cat_id = cat.Id
            is_fixture = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures)
            is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)

            if (is_fixture and not clear_fixtures) or (is_equipment and not clear_equipment):
                continue

            changed = False
            for param_name in RESULT_PARAM_NAMES:
                param = el.LookupParameter(param_name)
                if not param:
                    continue
                if _clear_param(param):
                    changed = True

            if changed:
                if is_fixture:
                    cleared_fixtures += 1
                elif is_equipment:
                    cleared_equipment += 1
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    if cleared_equipment or cleared_fixtures:
        logger.error(
            "Cleared stored circuit data on {} equipment and {} fixtures after write toggles were disabled.".format(
                cleared_equipment, cleared_fixtures
            )
        )

    return cleared_equipment, cleared_fixtures
