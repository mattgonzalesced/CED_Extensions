# -*- coding: utf-8 -*-
import os

from System.Collections.Generic import List
from pyrevit import DB, script

from CEDElectrical.Model.circuit_settings import CircuitSettings
from Snippets import categories as category_utils
from Snippets import revit_helpers

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


def _elementid_value(item, default=0):
    return revit_helpers.get_elementid_value(item, default=default)


def _is_valid_element_id(item):
    return _elementid_value(item, default=-1) > 0


def _definition_element_id(doc, definition):
    if definition is None:
        return DB.ElementId.InvalidElementId

    try:
        definition_id = getattr(definition, "Id", None)
    except Exception:
        definition_id = None
    if _is_valid_element_id(definition_id):
        return definition_id

    try:
        definition_guid = getattr(definition, "GUID", None)
    except Exception:
        definition_guid = None
    if definition_guid:
        try:
            shared_param_element = DB.SharedParameterElement.Lookup(doc, definition_guid)
        except Exception:
            shared_param_element = None
        if shared_param_element is not None and _is_valid_element_id(getattr(shared_param_element, "Id", None)):
            return shared_param_element.Id

    return DB.ElementId.InvalidElementId


def _parameter_definition_owned_by_other(doc, definition):
    if not getattr(doc, "IsWorkshared", False):
        return False, ""

    definition_id = _definition_element_id(doc, definition)
    if not _is_valid_element_id(definition_id):
        return False, ""

    try:
        status = DB.WorksharingUtils.GetCheckoutStatus(doc, definition_id)
    except Exception:
        status = None
    if status != DB.CheckoutStatus.OwnedByOtherUser:
        return False, ""

    owner = ""
    try:
        tooltip = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, definition_id)
        owner = str(getattr(tooltip, "Owner", "") or "")
    except Exception:
        owner = ""
    return True, owner


def _resolve_target_category_set(doc, row, settings):
    category_text = str((row or {}).get("Categories") or "").strip()
    category_ids, missing_tokens = category_utils.resolve_binding_category_ids(doc, category_text)

    writeback_flags = {}
    try:
        writeback_flags = dict(getattr(settings, "get_binding_writeback_flags")() or {})
    except Exception:
        writeback_flags = {}

    write_equipment = bool(writeback_flags.get("write_equipment_results", getattr(settings, "write_equipment_results", True)))
    write_fixtures = bool(writeback_flags.get("write_fixture_results", getattr(settings, "write_fixture_results", False)))
    filtered_ids = category_utils.apply_writeback_filter(
        doc,
        category_ids,
        write_equipment_results=write_equipment,
        write_fixture_results=write_fixtures,
    )

    category_set, inserted, missing_ids = category_utils.build_category_set(doc, filtered_ids)
    return category_set, inserted, missing_tokens, missing_ids


def _binding_needs_update(existing_binding, target_category_set, is_instance):
    if existing_binding is None:
        return True, set()
    current_is_instance = isinstance(existing_binding, DB.InstanceBinding)
    current_categories = category_utils.category_id_values_from_categories(getattr(existing_binding, "Categories", []))
    target_categories = category_utils.category_id_values_from_categories(target_category_set)
    # Conservative mode: never unbind existing categories automatically.
    removed = set()
    needs_update = not (current_is_instance == is_instance and target_categories.issubset(current_categories))
    return bool(needs_update), removed


def _sync_parameter_bindings(doc, app, rows, shared_param_file, settings, logger, check_ownership=True):
    bindmap = doc.ParameterBindings
    updated = 0
    unchanged = 0
    skipped = 0
    locked = []
    warnings = []
    errors = []
    unbound = 0

    for row in list(rows or []):
        name = str((row or {}).get("Parameter Name") or "").strip()
        if not name:
            skipped += 1
            continue

        definition = _get_shared_definition(shared_param_file, name)
        if definition is None:
            warnings.append("Missing shared definition: {}".format(name))
            skipped += 1
            continue

        category_set, inserted, missing_tokens, missing_ids = _resolve_target_category_set(doc, row, settings)
        if missing_tokens:
            warnings.append("{}: unresolved category token(s): {}".format(name, ", ".join(list(missing_tokens))))
        if missing_ids:
            warnings.append("{}: category id(s) not found in project: {}".format(name, ", ".join(list(missing_ids))))
        if inserted <= 0:
            warnings.append("{}: no valid categories after writeback filtering; skipped.".format(name))
            skipped += 1
            continue

        is_instance = str((row or {}).get("Instance/Type") or "").strip().lower() == "instance"
        group_label = str((row or {}).get("Group Under") or "").strip()
        group_id = LOAD_PARAMS_GROUP_MAP.get(group_label, DB.GroupTypeId.ElectricalCircuiting)

        existing_binding = _get_existing_binding(doc, definition.Name)
        if existing_binding is not None:
            merged_set, merged_inserted, merged_missing_ids = category_utils.merge_category_sets(
                doc,
                category_set,
                getattr(existing_binding, "Categories", []),
            )
            if merged_missing_ids:
                warnings.append("{}: category id(s) not found in project: {}".format(name, ", ".join(list(merged_missing_ids))))
            if merged_inserted > 0:
                category_set = merged_set
                inserted = int(merged_inserted)

        needs_update, removed_categories = _binding_needs_update(existing_binding, category_set, is_instance)
        if not needs_update:
            unchanged += 1
            continue

        if removed_categories:
            unbound += 1

        if check_ownership and existing_binding is not None:
            owned_by_other, owner_name = _parameter_definition_owned_by_other(doc, definition)
            if owned_by_other:
                skipped += 1
                locked.append({"parameter": name, "owner": owner_name})
                warnings.append(
                    "{}: skipped binding update because parameter definition is owned by '{}'.".format(
                        name, owner_name or "another user"
                    )
                )
                continue

        try:
            binding = _create_binding(app, category_set, is_instance)
            if existing_binding is not None:
                success = bindmap.ReInsert(definition, binding, group_id)
            else:
                success = bindmap.Insert(definition, binding, group_id)
            if isinstance(success, bool) and not success:
                errors.append("{}: Revit returned False when updating binding.".format(name))
                skipped += 1
                continue
            updated += 1
        except Exception as ex:
            errors.append("{}: {}".format(name, ex))
            skipped += 1

    return {
        "updated": int(updated),
        "unchanged": int(unchanged),
        "skipped": int(skipped),
        "locked": list(locked),
        "warnings": list(warnings),
        "errors": list(errors),
        "unbound": int(unbound),
    }


def sync_electrical_parameter_bindings(
    doc,
    logger=None,
    settings=None,
    check_ownership=True,
    transaction_name="Load Electrical Parameters",
):
    """Load and reconcile electrical shared parameter bindings using current writeback settings."""
    logger = logger or script.get_logger()
    app = doc.Application
    settings = settings or load_circuit_settings(doc)

    shared_txt, table_xlsx = _resolve_load_params_files()
    if not shared_txt or not table_xlsx:
        return {
            "status": "failed",
            "reason": "missing_files",
            "updated": 0,
            "unchanged": 0,
            "skipped": 0,
            "warnings": [],
            "errors": [],
            "locked": [],
            "unbound": 0,
            "total": 0,
        }

    rows = _load_parameter_rows(table_xlsx)
    original_shared_file = app.SharedParametersFilename
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
                "warnings": [],
                "errors": ["Failed to open shared parameter file."],
                "locked": [],
                "unbound": 0,
                "total": len(rows),
            }

        tx = DB.Transaction(doc, str(transaction_name or "Load Electrical Parameters"))
        tx.Start()
        summary = _sync_parameter_bindings(
            doc,
            app,
            rows,
            shared_param_file,
            settings=settings,
            logger=logger,
            check_ownership=check_ownership,
        )
        tx.Commit()

        status = "loaded" if int(summary.get("updated", 0)) > 0 else "present"
        summary["status"] = status
        summary["reason"] = ""
        summary["total"] = int(len(rows))
        return summary
    except Exception as ex:
        try:
            if tx and tx.HasStarted():
                tx.RollBack()
        except Exception:
            pass
        try:
            logger.warning("Load electrical parameters failed: {}".format(ex))
        except Exception:
            pass
        return {
            "status": "failed",
            "reason": str(ex),
            "updated": 0,
            "unchanged": 0,
            "skipped": len(rows) if rows else 0,
            "warnings": [],
            "errors": [str(ex)],
            "locked": [],
            "unbound": 0,
            "total": len(rows) if rows else 0,
        }
    finally:
        try:
            app.SharedParametersFilename = original_shared_file or ""
        except Exception:
            pass


def ensure_electrical_parameters_for_calculate(doc, logger=None):
    """
    Ensure required electrical shared parameters are available before calculate.
    This silently reconciles category bindings against current writeback settings.
    """
    settings = load_circuit_settings(doc)
    return sync_electrical_parameter_bindings(
        doc,
        logger=logger,
        settings=settings,
        check_ownership=True,
        transaction_name="Auto-Load Electrical Parameters",
    )


def unbind_disabled_writeback_categories(doc, logger=None, settings=None, check_ownership=True):
    """
    Compatibility wrapper for callers expecting category reconciliation.
    Current conservative mode does not unbind categories automatically.
    """
    active_settings = settings or load_circuit_settings(doc)
    return sync_electrical_parameter_bindings(
        doc,
        logger=logger,
        settings=active_settings,
        check_ownership=check_ownership,
        transaction_name="Update Electrical Parameter Categories",
    )


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
    fixture_category_ids = list(category_utils.get_fixture_category_ids(doc) or [])
    fixture_category_values = set(
        [revit_helpers.get_elementid_value(x, default=-1) for x in list(fixture_category_ids or [])]
    )
    equipment_category_ids = list(category_utils.get_equipment_category_ids(doc) or [])
    equipment_category_id = equipment_category_ids[0] if equipment_category_ids else None
    equipment_category_value = revit_helpers.get_elementid_value(equipment_category_id, default=-1)

    if clear_equipment:
        if equipment_category_id is not None:
            category_ids.append(equipment_category_id)
    if clear_fixtures:
        category_ids.extend(fixture_category_ids)

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
            cat_id_value = revit_helpers.get_elementid_value(cat_id, default=-1)
            is_fixture = cat_id_value in fixture_category_values
            is_equipment = cat_id_value == equipment_category_value

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
