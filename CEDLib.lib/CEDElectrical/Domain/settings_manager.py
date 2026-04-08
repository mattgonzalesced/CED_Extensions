# -*- coding: utf-8 -*-
import os

from System.Collections.Generic import List
from pyrevit import DB, script

from CEDElectrical.Model.circuit_settings import CircuitSettings

GP_NAME = "CED_Circuit_Settings"
AUTO_PARAM_PROBE_NAME = "Circuit Data_CED"

LOAD_PARAMS_COLUMNS = (
    "GUID",
    "UniqueId",
    "Parameter Name",
    "Discipline",
    "Type of Parameter",
    "Group Under",
    "Instance/Type",
    "Categories",
    "Groups",
)
LOAD_PARAMS_GROUP_MAP = {
    "Electrical": DB.GroupTypeId.Electrical,
    "Identity Data": DB.GroupTypeId.IdentityData,
    "Electrical - Circuiting": DB.GroupTypeId.ElectricalCircuiting,
    "Other": DB.GroupTypeId.Data,
}

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

FIXTURE_CATEGORY_IDS = [
    DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures),
    DB.ElementId(DB.BuiltInCategory.OST_LightingDevices),
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures),
    DB.ElementId(DB.BuiltInCategory.OST_SecurityDevices),
    DB.ElementId(DB.BuiltInCategory.OST_FireAlarmDevices),
    DB.ElementId(DB.BuiltInCategory.OST_DataDevices),
    DB.ElementId(DB.BuiltInCategory.OST_MechanicalControlDevices),
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


def has_project_parameter_binding(doc, parameter_name):
    """Return True when a project parameter binding exists by name."""
    name_text = str(parameter_name or "").strip()
    if not name_text:
        return False
    iterator = doc.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        key = iterator.Key
        if key and str(getattr(key, "Name", "") or "") == name_text:
            return True
    return False


def ensure_electrical_parameters_for_calculate(doc, logger=None):
    """
    Ensure required electrical shared parameters are available before calculate.
    This is silent (no alerts/output) and runs once per calculate call.
    """
    if has_project_parameter_binding(doc, AUTO_PARAM_PROBE_NAME):
        return {"status": "present", "updated": 0, "unchanged": 0, "skipped": 0}

    logger = logger or script.get_logger()
    app = doc.Application
    shared_txt, table_xlsx = _resolve_load_params_files()
    if not shared_txt or not table_xlsx:
        return {
            "status": "failed",
            "reason": "missing_files",
            "updated": 0,
            "unchanged": 0,
            "skipped": 0,
        }

    rows = _load_parameter_rows(table_xlsx)
    updated = 0
    unchanged = 0
    skipped = 0

    original_shared_file = app.SharedParametersFilename
    shared_param_file = None
    tx = None
    try:
        app.SharedParametersFilename = shared_txt
        shared_param_file = app.OpenSharedParameterFile()
        if not shared_param_file:
            return {
                "status": "failed",
                "reason": "shared_file_open_failed",
                "updated": 0,
                "unchanged": 0,
                "skipped": len(rows),
            }

        tx = DB.Transaction(doc, "Auto-Load Electrical Parameters")
        tx.Start()
        bindmap = doc.ParameterBindings

        for row in rows:
            name = row.get("Parameter Name")
            if not name:
                skipped += 1
                continue

            definition = _get_shared_definition(shared_param_file, name)
            if definition is None:
                skipped += 1
                continue

            category_names = [c.strip() for c in str(row.get("Categories") or "").split(",") if c and c.strip()]
            category_set, inserted = _category_set_from_names(doc, category_names)
            if inserted <= 0:
                skipped += 1
                continue

            is_instance = str(row.get("Instance/Type") or "").strip().lower() == "instance"
            group_label = str(row.get("Group Under") or "").strip()
            group_id = LOAD_PARAMS_GROUP_MAP.get(group_label, DB.GroupTypeId.ElectricalCircuiting)

            existing_binding = _get_existing_binding(doc, definition.Name)
            if existing_binding:
                current_is_instance = isinstance(existing_binding, DB.InstanceBinding)
                current_categories = _category_id_set(existing_binding.Categories)
                target_categories = _category_id_set(category_set)
                needs_update = not (current_is_instance == is_instance and current_categories == target_categories)
            else:
                needs_update = True

            if not needs_update:
                unchanged += 1
                continue

            binding = _create_binding(app, category_set, is_instance)
            if existing_binding:
                bindmap.ReInsert(definition, binding, group_id)
            else:
                bindmap.Insert(definition, binding, group_id)
            updated += 1

        tx.Commit()
        return {
            "status": "loaded",
            "updated": updated,
            "unchanged": unchanged,
            "skipped": skipped,
        }
    except Exception as ex:
        try:
            if tx and tx.HasStarted():
                tx.RollBack()
        except Exception:
            pass
        try:
            logger.warning("Auto-load electrical parameters failed: {}".format(ex))
        except Exception:
            pass
        return {
            "status": "failed",
            "reason": str(ex),
            "updated": 0,
            "unchanged": 0,
            "skipped": len(rows) if rows else 0,
        }
    finally:
        try:
            app.SharedParametersFilename = original_shared_file or ""
        except Exception:
            pass


def _extensions_root():
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))


def _candidate_load_params_content_dirs():
    root = _extensions_root()
    suffix = (
        "Electrical.panel",
        "Circuits1.stack",
        "Circuit Tools.pulldown",
        "Load Electrical Parameters.pushbutton",
        "Content",
    )
    candidates = []
    try:
        for entry in os.listdir(root):
            ext_path = os.path.join(root, entry)
            if not os.path.isdir(ext_path):
                continue
            if not str(entry or "").lower().endswith(".extension"):
                continue
            for tab_name in ("AE pyTools.tab", "AE pyTools.Tab"):
                candidates.append(os.path.join(ext_path, tab_name, *suffix))
    except Exception:
        pass

    return candidates


def _resolve_load_params_files():
    for content_dir in _candidate_load_params_content_dirs():
        shared_txt = os.path.join(content_dir, "ELEC SHARED PARAMS.txt")
        table_xlsx = os.path.join(content_dir, "ELEC SHARED PARAM TABLE.xlsx")
        if os.path.exists(shared_txt) and os.path.exists(table_xlsx):
            return shared_txt, table_xlsx

    root = _extensions_root()
    try:
        for current_root, dirs, files in os.walk(root):
            if os.path.basename(current_root) != "Content":
                continue
            parent = os.path.basename(os.path.dirname(current_root))
            if parent != "Load Electrical Parameters.pushbutton":
                continue
            candidate_txt = os.path.join(current_root, "ELEC SHARED PARAMS.txt")
            candidate_xlsx = os.path.join(current_root, "ELEC SHARED PARAM TABLE.xlsx")
            if os.path.exists(candidate_txt) and os.path.exists(candidate_xlsx):
                return candidate_txt, candidate_xlsx
    except Exception:
        pass
    return None, None


def _load_parameter_rows(config_path):
    from pyrevit.interop import xl as pyxl

    xldata = pyxl.load(config_path, headers=False)
    sheet = xldata.get("Parameter List")
    if not sheet:
        raise Exception("Sheet 'Parameter List' not found in ELEC SHARED PARAM TABLE.xlsx.")
    rows = [dict(zip(LOAD_PARAMS_COLUMNS, row)) for row in sheet["rows"][1:] if len(row) >= len(LOAD_PARAMS_COLUMNS)]
    return sorted(rows, key=lambda row: row.get("UniqueId", ""))


def _get_shared_definition(shared_param_file, name):
    for group in shared_param_file.Groups:
        definition = group.Definitions.get_Item(name)
        if definition:
            return definition
    return None


def _get_existing_binding(doc, definition_name):
    iterator = doc.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        key = iterator.Key
        if key and str(getattr(key, "Name", "") or "") == str(definition_name or ""):
            return iterator.Current
    return None


def _category_set_from_names(doc, names):
    category_set = DB.CategorySet()
    category_map = {cat.Name: cat for cat in doc.Settings.Categories}
    inserted = 0
    for name in list(names or []):
        cat = category_map.get(name)
        if cat is None:
            continue
        category_set.Insert(cat)
        inserted += 1
    return category_set, inserted


def _category_id_set(categories):
    return set([int(getattr(cat.Id, "IntegerValue", -1)) for cat in list(categories or []) if cat is not None])


def _create_binding(app, category_set, is_instance):
    try:
        return DB.InstanceBinding(category_set) if is_instance else DB.TypeBinding(category_set)
    except Exception:
        creator = getattr(app, "Create", None)
        if creator is None:
            raise
        return creator.NewInstanceBinding(category_set) if is_instance else creator.NewTypeBinding(category_set)


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


def clear_downstream_results(doc, clear_equipment=False, clear_fixtures=False, logger=None, check_ownership=True):
    """Blank stored circuit data on downstream elements after toggles are disabled."""
    if not (clear_equipment or clear_fixtures):
        return 0, 0, []

    logger = logger or script.get_logger()
    cleared_equipment = 0
    cleared_fixtures = 0
    locked = []

    def _is_locked(eid):
        if not getattr(doc, "IsWorkshared", False):
            return False
        try:
            status = DB.WorksharingUtils.GetCheckoutStatus(doc, eid)
            return status == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _owner_name(eid):
        try:
            info = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, eid)
            return info.Owner
        except Exception:
            return None

    # Filter to only electrical fixtures/equipment that have an MEP model to avoid
    # grouped annotation and other non-relevant family instances.
    category_ids = []
    if clear_equipment:
        category_ids.append(DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment))
    if clear_fixtures:
        category_ids.extend(FIXTURE_CATEGORY_IDS)

    if not category_ids:
        return 0, 0, []

    multi_filter = DB.ElementMulticategoryFilter(List[DB.ElementId](category_ids))
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)

    t = DB.Transaction(doc, "Clear downstream circuit data")
    t.Start()
    try:
        collector = (
            DB.FilteredElementCollector(doc)
            .WherePasses(multi_filter)
            .WherePasses(option_filter)
            .OfClass(DB.FamilyInstance)
        )

        for el in collector:
            try:
                # Skip non-MEP model family instances which are typically annotation or nested items.
                if getattr(el, "MEPModel", None) is None:
                    continue
            except Exception:
                continue

            cat = el.Category
            if not cat:
                continue

            cat_id = cat.Id
            is_fixture = cat_id in FIXTURE_CATEGORY_IDS
            is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)

            if (is_fixture and not clear_fixtures) or (is_equipment and not clear_equipment):
                continue

            if check_ownership and _is_locked(el.Id):
                locked.append(
                    {
                        "element_id": el.Id,
                        "owner": _owner_name(el.Id) or "",
                        "category": cat.Name if cat else "",
                    }
                )
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
        logger.info(
            "Cleared stored circuit data on {} equipment and {} fixtures after write toggles were disabled.".format(
                cleared_equipment, cleared_fixtures
            )
        )

    return cleared_equipment, cleared_fixtures, locked
