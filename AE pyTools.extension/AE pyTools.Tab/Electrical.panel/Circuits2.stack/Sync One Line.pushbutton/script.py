# -*- coding: utf-8 -*-
# IRONPYTHON 2.7 COMPATIBLE (no f-strings, no .format usage)
from pyrevit import revit, DB, script, forms

logger = script.get_logger()

# ----------------------------------------------------------------------
# SOURCE: The circuit/panel data is read from Revit elements:
#   Circuits (ElectricalSystem) -> read built-in param: RBS_ELEC_CIRCUIT_PANEL_PARAM, RBS_ELEC_CIRCUIT_NUMBER, etc.
#   Panels  (ElectricalEquipment) -> read built-in param: RBS_ELEC_PANEL_NAME, etc.
#
# DESTINATION: The detail items hold custom "which circuit/panel?" params:
#   e.g. "CKT_Panel_CEDT", "CKT_Circuit Number_CEDT", "Panel Name_CEDT"
# so we can figure out which circuit/panel data to copy in.
#
# Then we apply the circuit/panel values into the detail items' "CKT_Rating_CED" etc.

# These are the custom param names on detail items that hold the "which circuit/panel?" data:
DETAIL_PARAM_CKT_PANEL = "CKT_Panel_CEDT"
DETAIL_PARAM_CKT_NUMBER = "CKT_Circuit Number_CEDT"
DETAIL_PARAM_CKT_LOAD_NAME = "CKT_Load Name_CEDT"
DETAIL_PARAM_PANEL_NAME = "Panel Name_CEDT"

DEVICE_CATEGORY_IDS = [
    DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures),
    DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment),
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures),
    DB.ElementId(DB.BuiltInCategory.OST_DataDevices)
]

# Circuit built-in param -> we store them in a dictionary so we can quickly read
# Then we map them to detail param names that we’ll set.
# Key here is "DetailParamName" : "BuiltInParamOnCircuit"
CIRCUIT_VALUE_MAP = {
    "x VD Schedule": "x VD Schedule",
    "Circuit Tree Sort_CED": "Circuit Tree Sort_CED",
    "CKT_Circuit Type_CEDT": "CKT_Circuit Type_CEDT",
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "CKT_Schedule Notes_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM,
    "CKT_Length_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "Voltage_CED": DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Wire Material_CEDT": "Wire Material_CEDT",
    "Wire Insulation_CEDT": "Wire Insulation_CEDT",
    "Wire Temparature Rating_CEDT": "Wire Temparature Rating_CEDT",
    "Wire Size_CEDT": "Wire Size_CEDT",
    "Conduit and Wire Size_CEDT": "Conduit and Wire Size_CEDT",
    "Conduit Type_CEDT": "Conduit Type_CEDT",
    "Conduit Size_CEDT": "Conduit Size_CEDT",
    "Conduit Fill Percentage_CED": "Conduit Fill Percentage_CED",
    "Voltage Drop Percentage_CED": "Voltage Drop Percentage_CED",
    "Circuit Load Current_CED": "Circuit Load Current_CED"
}

# Panel built-in param -> detail param name map
PANEL_VALUE_MAP = {
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Mains Rating_CED": "Mains Rating_CED",
    "Mains Type_CEDT": "Mains Type_CEDT",
    "Phase_CED": "Phase_CED",
    "Main Breaker Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Short Circuit Rating_CEDT": DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING,
    "Mounting_CEDT": DB.BuiltInParameter.RBS_ELEC_MOUNTING,
    "Panel Modifications_CEDT": DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS,
    "Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM,
    "Secondary Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS,
    "Total Connected Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM,
    "Total Demand Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Total Connected Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM,
    "Total Demand Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Max Number of Single Pole Breakers_CED": DB.BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS,
    "Max Number of Circuits_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS,
    "Transformer Rating_CEDT": "Transformer Rating_CEDT",
    "Transformer Rating_CED": "Transformer Rating_CEDT",
    "Transformer Primary Description_CEDT": "Transformer Primary Description_CEDT",
    "Transformer Secondary Description_CEDT": "Transformer Secondary Description_CEDT",
    "Transformer %Z_CED": "Transformer %Z_CED",
    "Panel Feed_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_PARAM,
}


def get_model_param_value(elem, param_key, allow_type_fallback=True):
    """
    Reads either a BuiltInParameter or a shared parameter (by name).
    Checks instance first, then type if not found and allowed.
    Returns string/int/double or None.
    """
    param = None

    # Try instance-level parameter
    if isinstance(param_key, DB.BuiltInParameter):
        param = elem.get_Parameter(param_key)
    elif isinstance(param_key, str):
        param = elem.LookupParameter(param_key)
    else:
        logger.debug("get_model_param_value: Invalid param key type: " + str(param_key))
        return None

    # Try type-level parameter if instance param is missing and fallback is allowed
    if not param and allow_type_fallback:
        try:
            type_elem = elem.Document.GetElement(elem.GetTypeId())
            if type_elem:
                if isinstance(param_key, DB.BuiltInParameter):
                    param = type_elem.get_Parameter(param_key)
                elif isinstance(param_key, str):
                    param = type_elem.LookupParameter(param_key)
        except Exception as e:
            logger.debug("get_model_param_value: Error accessing type element for " + str(elem.Id) + ": " + str(e))

    # Final check
    if not param:
        logger.debug("get_model_param_value: Param '{}' not found on element {}".format(param_key, elem.Id))
        return None

    st = param.StorageType
    if st == DB.StorageType.String:
        return param.AsString()
    elif st == DB.StorageType.Integer:
        return param.AsInteger()
    elif st == DB.StorageType.Double:
        return param.AsDouble()
    elif st == DB.StorageType.ElementId:
        return param.AsValueString()

    return None


def get_detail_param_value(elem, param_name):
    """
    Read a string/int/double from the detail item’s custom param param_name. Return None if not found or empty.
    """
    p = elem.LookupParameter(param_name)
    if not p:
        logger.debug("    get_detail_param_value: Param '" + param_name + "' not found on detail item " + str(elem.Id))
        return None

    st = p.StorageType
    if st == DB.StorageType.String:
        return p.AsString()
    elif st == DB.StorageType.Integer:
        return p.AsInteger()
    elif st == DB.StorageType.Double:
        return p.AsDouble()
    elif st == DB.StorageType.ElementId:
        return p.AsValueString()
    return None


def set_detail_param_value(elem, param_name, new_value):
    """
    Sets the detail item’s param_name to str(new_value).
    """
    p = elem.LookupParameter(param_name)
    if not p:
        return
    if p.IsReadOnly:
        return
    try:
        if new_value is None:
            new_value = ""
        p.Set(new_value)
    except:
        logger.debug("      set_detail_param_value: FAILED setting '" + param_name + "' on " + str(elem.Id))


def is_not_in_group(element):
    return element.GroupId == DB.ElementId.InvalidElementId


def get_detail_type_label(doc, detail_elem):
    try:
        type_elem = doc.GetElement(detail_elem.GetTypeId())
        if type_elem:
            family_name = getattr(type_elem, "FamilyName", "") or ""
            type_name = getattr(type_elem, "Name", "") or ""
            if family_name and type_name:
                return family_name + ": " + type_name
            return type_name or family_name
    except Exception:
        pass
    return ""


def element_category_in_targets(elem):
    """
    Returns True if the element's category is one of the device categories we want to update.
    """
    cat = elem.Category
    if not cat:
        return False
    for target_cat_id in DEVICE_CATEGORY_IDS:
        if cat.Id == target_cat_id:
            return True
    return False


def _linkify_id(output, id_value, link_text=None):
    if not id_value:
        return ""
    try:
        return output.linkify(DB.ElementId(int(id_value)), link_text)
    except Exception:
        return ""


def _linkify_ids(output, id_values):
    if not id_values:
        return ""
    unique_ids = sorted(set([str(val) for val in id_values if val]))
    if not unique_ids:
        return ""
    if len(unique_ids) == 1:
        return _linkify_id(output, unique_ids[0], unique_ids[0])
    element_ids = [DB.ElementId(int(val)) for val in unique_ids]
    return output.linkify(element_ids, "Select")


def _ensure_drafting_view(doc):
    active_view = doc.ActiveView
    if not active_view or active_view.ViewType != DB.ViewType.DraftingView:
        forms.alert("Sync One Line only runs in a Drafting View. Please open a drafting view and try again.")
        return None
    return active_view


def collect_all_circuits(doc, option_filter):
    circuits = []

    ckt_collector = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Electrical.ElectricalSystem) \
        .WherePasses(option_filter) \
        .ToElements()

    for ckt in ckt_collector:

        panel_name = get_model_param_value(
            ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM
        )
        cnum = get_model_param_value(
            ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER
        )

        if not panel_name or not cnum:
            continue

        circuits.append({
            "panel_name": str(panel_name),
            "ckt_number": str(cnum),
            "element": ckt,
            "id": str(ckt.Id.IntegerValue)
        })

    return circuits


def build_circuits_by_panel(resolved_panels, circuits):
    result = {}

    for panel_name, pdata in resolved_panels.items():
        pid = pdata["panel_id"]
        panel_elem = pdata["panel"]

        result[pid] = {}

        try:
            systems = panel_elem.MEPModel.GetAssignedElectricalSystems()
        except:
            systems = []

        for sys in systems:
            cnum = get_model_param_value(
                sys, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER
            )
            if not cnum:
                continue

            cdata = {}
            for dp, bip in CIRCUIT_VALUE_MAP.items():
                cdata[dp] = get_model_param_value(sys, bip)

            cdata["circuit_id"] = str(sys.Id.IntegerValue)
            result[pid][str(cnum)] = cdata

    return result


def _collect_circuits(doc, option_filter):
    circuit_map = {}
    circuited_panel_names = set()

    ckt_collector = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Electrical.ElectricalSystem) \
        .WherePasses(option_filter) \
        .ToElements()

    for ckt in ckt_collector:
        pval = get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        cnum = get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)

        if pval:
            circuited_panel_names.add(str(pval))

        if pval and cnum:
            key = (str(pval), str(cnum))
            cdata = {}
            for detail_param_name, bip in CIRCUIT_VALUE_MAP.items():
                cdata[detail_param_name] = get_model_param_value(ckt, bip)
            cdata["circuit_id"] = str(ckt.Id.IntegerValue)
            circuit_map[key] = cdata

    return ckt_collector, circuit_map, circuited_panel_names


def _is_spare_or_space_circuit(electrical_system):
    """
    Returns True if ElectricalSystem.CircuitType is Spare or Space.
    Safe for IronPython / enum weirdness.
    """
    if not electrical_system:
        return False

    try:
        ctype = electrical_system.CircuitType
    except:
        return False

    # Prefer enum compare if available
    try:
        if ctype == DB.Electrical.CircuitType.Spare:
            return True
        if ctype == DB.Electrical.CircuitType.Space:
            return True
    except:
        pass

    # Fallback to string compare (for API/enum binding edge cases)
    try:
        ctype_str = str(ctype).upper()
        if "SPARE" in ctype_str:
            return True
        if "SPACE" in ctype_str:
            return True
    except:
        pass

    return False


def _get_fed_from_label(equipment):
    """
    Returns 'PANEL / CIRCUIT' for the primary feeder supplying this equipment,
    or empty string if not found.
    """
    try:
        mep_model = equipment.MEPModel
        if not mep_model:
            return ""

        conn_mgr = mep_model.ConnectorManager
        if not conn_mgr:
            return ""

        connectors = conn_mgr.Connectors
        if not connectors:
            return ""

        primary_connector = None

        # Find primary connector on the equipment
        for conn in connectors:
            try:
                info = conn.GetMEPConnectorInfo()
                if info and info.IsPrimary:
                    primary_connector = conn
                    break
            except:
                continue

        if not primary_connector:
            return ""

        # Traverse references to find connected ElectricalSystem
        refs = primary_connector.AllRefs
        if not refs:
            return ""

        for ref_conn in refs:
            try:
                owner = ref_conn.Owner
                if isinstance(owner, DB.Electrical.ElectricalSystem):
                    panel = get_model_param_value(
                        owner,
                        DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM
                    )
                    number = get_model_param_value(
                        owner,
                        DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER
                    )

                    if panel and number:
                        return str(panel) + " / " + str(number)
            except:
                continue

    except:
        pass

    return ""


def _collect_panels(doc, option_filter):
    panel_map = {}  # name -> [pdata, pdata, ...]
    panel_map_by_id = {}

    pnl_collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType() \
        .WherePasses(option_filter) \
        .ToElements()

    for pnl in pnl_collector:
        pname = get_model_param_value(pnl, DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        if not pname:
            continue

        pdata = {}
        for detail_param_name, bip in PANEL_VALUE_MAP.items():
            pdata[detail_param_name] = get_model_param_value(pnl, bip)

        pdata["panel_id"] = str(pnl.Id.IntegerValue)
        pdata["_element"] = pnl  # keep element for sorting/debug

        key = str(pname)
        panel_map.setdefault(key, []).append(pdata)
        panel_map_by_id[pdata["panel_id"]] = pdata

    return panel_map, panel_map_by_id


def resolve_panels(panel_map, circuited_panel_names):
    resolved = {}  # panel_name -> {panel, panel_id}
    failed = {}  # panel_name -> {used, rejected}

    for panel_name, candidates in panel_map.items():

        # Single panel — always valid
        if len(candidates) == 1:
            pdata = candidates[0]
            resolved[panel_name] = {
                "panel": pdata["_element"],
                "panel_id": pdata["panel_id"]
            }
            continue

        # Prefer circuited panels
        circuited = []
        if panel_name in circuited_panel_names:
            for pdata in candidates:
                pnl = pdata["_element"]
                try:
                    if pnl.MEPModel and pnl.MEPModel.GetElectricalSystems():
                        circuited.append(pdata)
                except:
                    pass

        pool = circuited if circuited else candidates

        winner = sorted(pool, key=lambda d: int(d["panel_id"]))[0]
        losers = [d for d in candidates if d != winner]

        resolved[panel_name] = {
            "panel": winner["_element"],
            "panel_id": winner["panel_id"]
        }

        failed[panel_name] = {
            "used": winner["panel_id"],
            "rejected": [d["panel_id"] for d in losers]
        }

    return resolved, failed




def _get_supplied_panel_id_from_circuit(circuit):
    """
    Try to find a single ElectricalEquipment element that this circuit supplies.
    Returns panel_id string or None.
    """
    try:
        elems = circuit.Elements  # ElementSet-like
    except:
        elems = None

    if not elems:
        return None

    supplied_ids = []

    try:
        for el in elems:
            try:
                if el and el.Category and el.Category.Id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment):
                    supplied_ids.append(str(el.Id.IntegerValue))
            except:
                continue
    except:
        return None

    supplied_ids = sorted(set([x for x in supplied_ids if x]))
    if len(supplied_ids) == 1:
        return supplied_ids[0]

    return None


def reconcile_panel_identity_from_circuit(
        ditem,
        resolved_panels,
        circuits_by_panel,
        auto_panel_updates,
        auto_panel_warnings
):
    # -------------------------
    # Guardrail 1: must have all 3 params
    # -------------------------
    pname_val = get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)
    cpanel_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
    cnum_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)

    if not (pname_val and cpanel_val and cnum_val):
        return False

    identity_name = str(pname_val)
    circuit_panel_name = str(cpanel_val)
    ckt_number = str(cnum_val)

    # -------------------------
    # Guardrail 2: identity panel must be resolved
    # -------------------------
    identity_pdata = resolved_panels.get(identity_name)
    if not identity_pdata:
        return False

    current_panel_id = identity_pdata.get("panel_id")
    if not current_panel_id:
        return False

    # -------------------------
    # Guardrail 3: resolve circuit via circuit's panel
    # -------------------------
    circuit_panel_pdata = resolved_panels.get(circuit_panel_name)
    if not circuit_panel_pdata:
        return False

    owner_panel_id = circuit_panel_pdata.get("panel_id")
    if not owner_panel_id:
        return False

    cdict = circuits_by_panel.get(owner_panel_id, {}).get(ckt_number)
    if not cdict:
        return False

    cid = cdict.get("circuit_id")
    if not cid:
        return False

    try:
        circuit_elem = revit.doc.GetElement(DB.ElementId(int(cid)))
    except:
        return False

    if not circuit_elem:
        return False

    # -------------------------
    # ✅ NEW GUARDRAIL: ignore Spare / Space circuits
    # -------------------------
    if _is_spare_or_space_circuit(circuit_elem):
        return False

    # -------------------------
    # Guardrail 4: resolve supplied panel
    # -------------------------
    supplied_panel_id = _get_supplied_panel_id_from_circuit(circuit_elem)
    if not supplied_panel_id:
        return False

    supplied_panel_name = None
    supplied_panel_elem = None

    for name, pdata in resolved_panels.items():
        if pdata.get("panel_id") == supplied_panel_id:
            supplied_panel_name = name
            supplied_panel_elem = pdata.get("panel")
            break

    if not supplied_panel_name or not supplied_panel_elem:
        return False

    # -------------------------
    # Already correct
    # -------------------------
    if supplied_panel_id == current_panel_id:
        return False

    # -------------------------
    # APPLY CORRECTION
    # -------------------------
    set_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME, supplied_panel_name)

    for detail_param_name, bip in PANEL_VALUE_MAP.items():
        set_detail_param_value(
            ditem,
            detail_param_name,
            get_model_param_value(supplied_panel_elem, bip)
        )

    auto_panel_updates.setdefault(
        "The following Equipment Symbols were updated automatically based on new supply circuit number.",
        set()
    ).add(str(ditem.Id.IntegerValue))

    return True


def _collect_detail_items(doc, option_filter, active_view):
    detail_item_collector = DB.FilteredElementCollector(doc, active_view.Id) \
        .OfCategory(DB.BuiltInCategory.OST_DetailComponents) \
        .WhereElementIsNotElementType() \
        .WherePasses(option_filter) \
        .ToElements()

    detail_items = [
        el for el in detail_item_collector
        if is_not_in_group(el)]
    logger.debug("Collected " + str(len(detail_items)) + " detail item(s).")
    return detail_items


def _get_circuit_sort_key(circuit_label):
    """
    Returns an integer circuit number if possible, otherwise a high fallback
    so non-numeric circuits sort last.
    Expected format: 'DP1 / 12 - LOAD'
    """
    try:
        if "/" in circuit_label:
            right = circuit_label.split("/", 1)[1]
            num_part = right.split("-", 1)[0].strip()
            return int(num_part)
    except:
        pass
    return 999999


def _build_output_summary(detail_items, circuit_map, panel_map, panel_map_by_id, resolved_panels, failed_panels,
                          auto_panel_updates, auto_panel_warnings):
    equipment_rows = {}
    circuit_rows = {}
    unmapped_details = []
    mapped_panel_names = set()

    for ditem in detail_items:
        detail_id = str(ditem.Id.IntegerValue)

        pname_val = get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)
        cpanel_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
        cnum_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)

        had_mapping = False

        if pname_val:
            panel_name = str(pname_val)

            row = equipment_rows.get(panel_name)
            if not row:
                row = {"panel_ids": set(), "detail_ids": set()}

                pdict = resolved_panels.get(panel_name)
                if pdict:
                    pid = pdict.get("panel_id")
                    if pid:
                        row["panel_ids"].add(pid)

                equipment_rows[panel_name] = row

            row["detail_ids"].add(detail_id)
            mapped_panel_names.add(panel_name)
            had_mapping = True

        if cpanel_val and cnum_val:
            ckt_panel = str(cpanel_val)
            ckt_number = str(cnum_val)
            ckey = (ckt_panel, ckt_number)

            row = circuit_rows.get(ckey)
            if not row:
                row = {
                    "ckt_panel": ckt_panel,
                    "ckt_number": ckt_number,
                    "load_name": "",
                    "circuit_ids": set(),
                    "detail_ids": set()
                }

                cdict = circuit_map.get(ckey)
                if cdict:
                    cid = cdict.get("circuit_id")
                    if cid:
                        row["circuit_ids"].add(cid)
                    row["load_name"] = cdict.get(DETAIL_PARAM_CKT_LOAD_NAME) or ""

                circuit_rows[ckey] = row

            row["detail_ids"].add(detail_id)
            had_mapping = True

        if not had_mapping:
            label = get_detail_type_label(revit.doc, ditem)
            unmapped_details.append((detail_id, label))

    unmapped_panels = []

    # (1) Never-mapped panels
    for panel_id, pdata in panel_map_by_id.items():
        panel_name = pdata.get(DETAIL_PARAM_PANEL_NAME) or pdata.get(DETAIL_PARAM_CKT_PANEL, "") or "(Unnamed Panel)"
        if panel_name not in mapped_panel_names:
            unmapped_panels.append({"name": panel_name, "ids": [panel_id]})

    # (2) Duplicate-name rejected panels
    if failed_panels:
        for panel_name, dupdata in failed_panels.items():
            rejected_ids = dupdata.get("rejected", [])
            if rejected_ids:
                unmapped_panels.append({
                    "name": panel_name + " (Duplicate Name – Not Used)",
                    "ids": rejected_ids
                })

    return equipment_rows, circuit_rows, unmapped_details, unmapped_panels, auto_panel_updates, auto_panel_warnings


def _render_summary(equipment_rows, circuit_rows, unmapped_details, unmapped_panels, failed_panels,
                    auto_panel_updates, auto_panel_warnings):
    output = script.get_output()
    output.close_others()
    output.print_md("## Sync One Line Results")

    headers = ["Element", "Category", "Name", "Detail Items", "Detail Count"]

    # Duplicate-name conflicts (your "bigger problem" text stays)
    if failed_panels:
        output.print_md("### ⚠ Duplicate Panel Name Conflicts")
        for panel_name in sorted(failed_panels.keys(), key=lambda x: x.upper()):
            data = failed_panels.get(panel_name, {})
            used_id = data.get("used")
            rejected_ids = data.get("rejected", [])

            used_link = _linkify_id(output, used_id, "element " + str(used_id)) if used_id else "(unknown)"
            rejected_link = _linkify_ids(output, rejected_ids) if rejected_ids else ""

            output.print_md(
                "- Multiple panels named **'" + panel_name + "'** detected. "
                "Using " + used_link + " for mapping panel and associated circuits on one-line diagram."
            )
            if rejected_link:
                output.print_md("  - Please give elements " + rejected_link + " unique names.")

    # Auto panel identity updates (grouped)
    if auto_panel_updates:
        output.print_md("### Notices")
        for msg in sorted(auto_panel_updates.keys()):
            output.print_md("- " + msg + " " + _linkify_ids(output, auto_panel_updates.get(msg, set())))

    # Auto panel identity warnings (grouped)
    if auto_panel_warnings:
        output.print_md("### Warnings")
        for msg in sorted(auto_panel_warnings.keys()):
            output.print_md("- " + msg + " " + _linkify_ids(output, auto_panel_warnings.get(msg, set())))

    # -------------------------
    # Build flattened panel order
    # -------------------------
    panel_names = set()
    panel_names.update(equipment_rows.keys())
    panel_names.update([row["ckt_panel"] for row in circuit_rows.values()])

    def _ckt_sort_key(row):
        try:
            return int(row.get("ckt_number"))
        except:
            return 999999

    table_data = []

    for panel_name in sorted(panel_names, key=lambda x: x.upper()):

        # ---- EE row (if exists) ----
        if panel_name in equipment_rows:
            erow = equipment_rows[panel_name]
            table_data.append([
                _linkify_ids(output, erow.get("panel_ids", set())),
                "EE",
                panel_name,
                _linkify_ids(output, erow.get("detail_ids", set())),
                str(len(erow.get("detail_ids", set())))
            ])

        # ---- EC rows for this panel ----
        panel_circuits = [
            row for row in circuit_rows.values()
            if row.get("ckt_panel") == panel_name
        ]

        for crow in sorted(panel_circuits, key=_ckt_sort_key):
            label = (
                    panel_name + " / " +
                    str(crow.get("ckt_number")) +
                    (" - " + crow.get("load_name") if crow.get("load_name") else "")
            )

            table_data.append([
                _linkify_ids(output, crow.get("circuit_ids", set())),
                "EC",
                label,
                _linkify_ids(output, crow.get("detail_ids", set())),
                str(len(crow.get("detail_ids", set())))
            ])

    if not table_data:
        table_data.append(["", "", "None", "", "0"])

    output.print_table(
        table_data=table_data,
        title="Panels & Circuits",
        columns=headers,
        formats=["", "", "", "", ""]
    )

    # -------------------------
    # Unmapped Detail Items
    # -------------------------
    detail_table_data = []
    if unmapped_details:
        for detail_id, detail_label in sorted(unmapped_details, key=lambda d: (d[1] or "", d[0])):
            detail_table_data.append([
                _linkify_ids(output, [detail_id]),
                detail_label or "(Unnamed Detail Item)"
            ])
    else:
        detail_table_data.append(["", "None"])

    output.print_table(
        table_data=detail_table_data,
        title="Unmapped Detail Items",
        columns=["Element", "Name"],
        formats=["", ""]
    )

    # -------------------------
    # Unmapped Model Equipment
    # -------------------------
    equipment_table_data = []
    if unmapped_panels:
        for panel in sorted(unmapped_panels, key=lambda p: p.get("name") or ""):
            equipment_table_data.append([
                _linkify_ids(output, panel.get("ids", [])),
                "EE",
                panel.get("name") or "(Unnamed Panel)",
                "",
                "0"
            ])
    else:
        equipment_table_data.append(["", "", "None", "", "0"])

    output.print_table(
        table_data=equipment_table_data,
        title="Unmapped Model Equipment",
        columns=headers,
        formats=["", "", "", "", ""]
    )


def main():
    doc = revit.doc
    logger.info("=== Syncing Circuit/Panel param values to detail items ===")

    active_view = _ensure_drafting_view(doc)
    if not active_view:
        return

    auto_panel_updates = {}  # message -> set(detail_ids)
    auto_panel_warnings = {}  # message -> set(detail_ids)

    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)

    logger.debug("Collecting circuits...")
    ckt_collector, circuit_map, circuited_panel_names = _collect_circuits(doc, option_filter)

    logger.debug("Collecting panels...")
    panel_map, panel_map_by_id = _collect_panels(doc, option_filter)

    detail_items = _collect_detail_items(doc, option_filter, active_view)

    # Resolve duplicate panels ONCE
    resolved_panels, failed_panels = resolve_panels(panel_map, circuited_panel_names)

    # Build circuits scoped to resolved panels
    all_circuits = collect_all_circuits(doc, option_filter)
    circuits_by_panel = build_circuits_by_panel(resolved_panels, all_circuits)

    # Apply ALL model changes (ONE TRANSACTION)
    t = DB.Transaction(doc, "Sync Circuits/Panels to Detail Items")
    t.Start()

    update_count = 0

    for ditem in detail_items:
        logger.debug("Detail item " + str(ditem.Id) + ":")

        cpanel_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
        cnum_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)
        pname_val = get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)

        changed = False

        panel_identity_name = None
        circuit_panel_name = None

        if pname_val:
            panel_identity_name = str(pname_val)

        if cpanel_val:
            circuit_panel_name = str(cpanel_val)

        # -------------------------
        # PANEL SYNC (identity-based)
        # -------------------------
        if panel_identity_name and panel_identity_name in resolved_panels:
            pdata = resolved_panels[panel_identity_name]
            panel_elem = pdata["panel"]

            for detail_param_name, bip in PANEL_VALUE_MAP.items():
                set_detail_param_value(
                    ditem,
                    detail_param_name,
                    get_model_param_value(panel_elem, bip)
                )
            changed = True
        # -------------------------
        # CIRCUIT SYNC (ownership-based)
        # -------------------------
        if circuit_panel_name and cnum_val:
            cpdata = resolved_panels.get(circuit_panel_name)
            if cpdata:
                cp_id = cpdata["panel_id"]
                cdict = circuits_by_panel.get(cp_id, {}).get(str(cnum_val))
                if cdict:
                    for detail_pname, ckt_val in cdict.items():
                        set_detail_param_value(ditem, detail_pname, ckt_val)
                    changed = True

        did_reconcile = reconcile_panel_identity_from_circuit(
            ditem,
            resolved_panels,
            circuits_by_panel,
            auto_panel_updates,
            auto_panel_warnings
        )

        if changed or did_reconcile:
            update_count += 1

    t.Commit()

    logger.info("Sync finished. Updated " + str(update_count) + " detail item(s).")

    # Reporting (NO TRANSACTION)
    equipment_rows, circuit_rows, unmapped_details, unmapped_panels, auto_panel_updates, auto_panel_warnings = _build_output_summary(
        detail_items,
        circuit_map,
        panel_map,
        panel_map_by_id,
        resolved_panels,
        failed_panels,
        auto_panel_updates,
        auto_panel_warnings
    )

    choice = forms.alert(
        "Data sync complete.\n\nPrint output report?",
        ok=False,
        yes=True,
        no=True
    )

    if choice:
        _render_summary(
            equipment_rows,
            circuit_rows,
            unmapped_details,
            unmapped_panels,
            failed_panels,
            auto_panel_updates,
            auto_panel_warnings
        )


if __name__ == "__main__":
    main()
