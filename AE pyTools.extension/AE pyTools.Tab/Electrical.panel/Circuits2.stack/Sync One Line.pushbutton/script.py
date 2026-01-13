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
DETAIL_PARAM_SC_PANEL_ID = "SC_Panel ElementId"
DETAIL_PARAM_SC_CIRCUIT_ID = "SC_Circuit ElementId"

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
        logger.debug("      set_detail_param_value: Param '" + param_name + "' not found on " + str(elem.Id))
        return
    if p.IsReadOnly:
        logger.debug("      set_detail_param_value: Param '" + param_name + "' is read-only on " + str(elem.Id))
        return
    try:
        if new_value is None:
            new_value = ""
        p.Set(new_value)
        logger.debug(
            "      set_detail_param_value: Set '" + param_name + "' to '" + str(new_value) + "' on " + str(elem.Id))
    except:
        logger.debug("      set_detail_param_value: FAILED setting '" + param_name + "' on " + str(elem.Id))


def is_not_in_group(element):
    return element.GroupId == DB.ElementId.InvalidElementId


def element_has_sc_params(elem):
    """
    Returns True if the element has any parameters whose names start with 'SC_'.
    Used so we only flag missing references on elements expecting SC_ data.
    """
    try:
        for param in elem.Parameters:
            try:
                definition = param.Definition
                if definition:
                    pname = definition.Name
                    if pname and pname.startswith("SC_"):
                        return True
            except Exception as exc:
                logger.debug("element_has_sc_params: Failed checking param on {}: {}".format(elem.Id, exc))
                continue
    except Exception as exc:
        logger.debug("element_has_sc_params: Could not iterate params on {}: {}".format(elem.Id, exc))
    return False


def describe_detail_item(doc, detail_elem):
    """
    Returns (name, owner_view_name) tuple for reporting back to the user.
    """
    try:
        name = detail_elem.Name
    except Exception:
        name = ""

    view_name = ""
    try:
        owner_view_id = getattr(detail_elem, "OwnerViewId", None)
        if owner_view_id and owner_view_id != DB.ElementId.InvalidElementId:
            owner_view = doc.GetElement(owner_view_id)
            if owner_view:
                view_name = owner_view.Name
    except Exception as exc:
        logger.debug("describe_detail_item: Failed getting view for {}: {}".format(detail_elem.Id, exc))

    return name, view_name


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


def update_devices_with_sc_ids(circuit):
    """
    Writes SC element ids to family instances connected to the provided circuit.
    """
    if not circuit:
        return 0
    panel_elem = None
    try:
        panel_elem = circuit.BaseEquipment
    except Exception as panel_exc:
        logger.debug("update_devices_with_sc_ids: Failed to get panel for {}: {}".format(circuit.Id, panel_exc))

    panel_id_str = ""
    if panel_elem:
        try:
            panel_id_str = str(panel_elem.Id.IntegerValue)
        except Exception:
            panel_id_str = ""

    circuit_id_str = str(circuit.Id.IntegerValue)
    updated = 0
    for elem in circuit.Elements:
        if not isinstance(elem, DB.FamilyInstance):
            continue
        if not element_category_in_targets(elem):
            continue
        set_detail_param_value(elem, DETAIL_PARAM_SC_CIRCUIT_ID, circuit_id_str)
        if panel_id_str:
            set_detail_param_value(elem, DETAIL_PARAM_SC_PANEL_ID, panel_id_str)
        updated += 1
    return updated


def _ensure_drafting_view(doc):
    active_view = doc.ActiveView
    if not active_view or active_view.ViewType != DB.ViewType.DraftingView:
        forms.alert("Sync One Line only runs in a Drafting View. Please open a drafting view and try again.")
        return None
    return active_view


def _collect_circuits(doc, option_filter):
    circuit_map = {}
    circuit_map_by_id = {}

    ckt_collector = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Electrical.ElectricalSystem) \
        .WherePasses(option_filter) \
        .ToElements()

    for ckt in ckt_collector:
        pval = get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        cnum = get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        if pval and cnum:
            key = (str(pval), str(cnum))
            cdata = {}
            for detail_param_name, bip in CIRCUIT_VALUE_MAP.items():
                cdata[detail_param_name] = get_model_param_value(ckt, bip)
            cdata[DETAIL_PARAM_SC_CIRCUIT_ID] = str(ckt.Id.IntegerValue)
            circuit_map[key] = cdata
            circuit_map_by_id[str(ckt.Id.IntegerValue)] = cdata
        else:
            logger.debug("  Circuit " + str(ckt.Id) + " missing panel or circuit # => skipping")

    return ckt_collector, circuit_map, circuit_map_by_id


def _collect_panels(doc, option_filter):
    panel_map = {}
    panel_map_by_id = {}

    pnl_collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType() \
        .WherePasses(option_filter) \
        .ToElements()

    for pnl in pnl_collector:
        pname = get_model_param_value(pnl, DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        try:
            mep_model = getattr(pnl, "MEPModel", None)
            connector_manager = mep_model.ConnectorManager if mep_model else None
        except Exception:
            connector_manager = None
        if pname:
            if connector_manager:
                pdata = {}
                for detail_param_name, bip in PANEL_VALUE_MAP.items():
                    pdata[detail_param_name] = get_model_param_value(pnl, bip)
                pdata[DETAIL_PARAM_SC_PANEL_ID] = str(pnl.Id.IntegerValue)
                panel_map[str(pname)] = pdata
                panel_map_by_id[str(pnl.Id.IntegerValue)] = pdata
        else:
            logger.debug("  Panel " + str(pnl.Id) + " has no name => skipping")

    return panel_map, panel_map_by_id


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


def _build_output_summary(detail_items, circuit_map, circuit_map_by_id, panel_map, panel_map_by_id):
    panel_summary = {}
    unmapped_details = []
    mapped_panel_names = set()
    sc_missing_reference_records = []

    for ditem in detail_items:
        cpanel_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
        cnum_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)
        pname_val = get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)

        missing_contexts = []
        sc_circuit_id_val = get_detail_param_value(ditem, DETAIL_PARAM_SC_CIRCUIT_ID)
        sc_panel_id_val = get_detail_param_value(ditem, DETAIL_PARAM_SC_PANEL_ID)

        circuit_id_lookup = None
        if sc_circuit_id_val not in (None, "", 0):
            circuit_id_lookup = str(sc_circuit_id_val).strip()
            if not circuit_id_lookup or circuit_id_lookup == "0":
                circuit_id_lookup = None

        panel_id_lookup = None
        if sc_panel_id_val not in (None, "", 0):
            panel_id_lookup = str(sc_panel_id_val).strip()
            if not panel_id_lookup or panel_id_lookup == "0":
                panel_id_lookup = None

        has_circuit_reference = bool(circuit_id_lookup) or (cpanel_val and cnum_val)
        if not has_circuit_reference:
            missing_fields = []
            if not circuit_id_lookup:
                missing_fields.append(DETAIL_PARAM_SC_CIRCUIT_ID)
            if not cpanel_val:
                missing_fields.append(DETAIL_PARAM_CKT_PANEL)
            if not cnum_val:
                missing_fields.append(DETAIL_PARAM_CKT_NUMBER)
            missing_contexts.append({
                "context": "Circuit reference",
                "fields": missing_fields
            })

        has_panel_reference = bool(panel_id_lookup) or bool(pname_val) or bool(cpanel_val)
        if not has_panel_reference:
            missing_fields = []
            if not panel_id_lookup:
                missing_fields.append(DETAIL_PARAM_SC_PANEL_ID)
            if not pname_val:
                missing_fields.append(DETAIL_PARAM_PANEL_NAME)
            if not cpanel_val:
                missing_fields.append(DETAIL_PARAM_CKT_PANEL)
            missing_contexts.append({
                "context": "Panel reference",
                "fields": missing_fields
            })

        cdict = None
        if circuit_id_lookup and circuit_id_lookup in circuit_map_by_id:
            cdict = circuit_map_by_id[circuit_id_lookup]
        if not cdict and cpanel_val and cnum_val:
            ckey = (str(cpanel_val), str(cnum_val))
            cdict = circuit_map.get(ckey)

        pdict = None
        if pname_val:
            pdict = panel_map.get(str(pname_val))

        detail_id = str(ditem.Id.IntegerValue)

        panel_name_key = None
        if cdict:
            panel_name_key = cdict.get(DETAIL_PARAM_CKT_PANEL, "")
        if not panel_name_key and pdict:
            panel_name_key = pdict.get(DETAIL_PARAM_PANEL_NAME) or pdict.get(DETAIL_PARAM_CKT_PANEL, "")

        if panel_name_key:
            if panel_name_key not in panel_summary:
                panel_summary[panel_name_key] = {
                    "name": panel_name_key or "(Unnamed Panel)",
                    "ids": set(),
                    "detail_ids": set(),
                    "circuits": {}
                }
            panel_lookup = panel_map.get(panel_name_key)
            if panel_lookup:
                panel_id = panel_lookup.get(DETAIL_PARAM_SC_PANEL_ID)
                if panel_id:
                    panel_summary[panel_name_key]["ids"].add(panel_id)

        if cdict and panel_name_key:
            circuit_id = cdict.get(DETAIL_PARAM_SC_CIRCUIT_ID)
            cpanel_name = cdict.get(DETAIL_PARAM_CKT_PANEL, "")
            cnum_name = cdict.get(DETAIL_PARAM_CKT_NUMBER, "")
            load_name = cdict.get(DETAIL_PARAM_CKT_LOAD_NAME, "")
            circuit_label = "{} / {} - {}".format(cpanel_name or "", cnum_name or "", load_name or "").strip()
            circuit_key = circuit_label or "(Unnamed Circuit)"
            circuits = panel_summary[panel_name_key]["circuits"]
            if circuit_key not in circuits:
                circuits[circuit_key] = {
                    "label": circuit_label or "(Unnamed Circuit)",
                    "ids": set(),
                    "detail_ids": set()
                }
            if circuit_id:
                circuits[circuit_key]["ids"].add(circuit_id)
            circuits[circuit_key]["detail_ids"].add(detail_id)
        elif pdict and panel_name_key:
            panel_summary[panel_name_key]["detail_ids"].add(detail_id)
            mapped_panel_names.add(panel_name_key)

        if not cdict and not pdict:
            detail_label = get_detail_type_label(revit.doc, ditem)
            unmapped_details.append((detail_id, detail_label))

        if missing_contexts and element_has_sc_params(ditem):
            name, view_name = describe_detail_item(revit.doc, ditem)
            sc_missing_reference_records.append({
                "id": ditem.Id.IntegerValue,
                "name": name,
                "view": view_name,
                "contexts": missing_contexts
            })

    unmapped_panels = []
    for panel_id, pdata in panel_map_by_id.items():
        panel_name = pdata.get(DETAIL_PARAM_PANEL_NAME) or pdata.get(DETAIL_PARAM_CKT_PANEL, "") or "(Unnamed Panel)"
        if panel_name not in mapped_panel_names:
            unmapped_panels.append({
                "name": panel_name,
                "ids": [panel_id]
            })

    return panel_summary, unmapped_details, unmapped_panels, sc_missing_reference_records


def _render_summary(panel_summary, unmapped_details, unmapped_panels, sc_missing_reference_records):
    output = script.get_output()
    output.close_others()
    output.print_md("## Sync One Line Results")

    headers = ["Element", "Category", "Name", "Detail Items", "Detail Count"]
    table_data = []
    if panel_summary:
        for panel_key in sorted(panel_summary.keys()):
            panel = panel_summary[panel_key]
            panel_ids = panel.get("ids", set())
            panel_link = _linkify_ids(output, panel_ids)
            panel_details = panel.get("detail_ids", set())
            detail_link = _linkify_ids(output, panel_details)
            table_data.append([
                panel_link,
                "EE",
                panel.get("name") or "(Unnamed Panel)",
                detail_link,
                str(len(panel_details))
            ])
            circuits = panel.get("circuits", {})
            for circuit_key in sorted(circuits.keys()):
                circuit = circuits[circuit_key]
                circuit_ids = circuit.get("ids", set())
                circuit_link = _linkify_ids(output, circuit_ids)
                circuit_details = circuit.get("detail_ids", set())
                circuit_detail_link = _linkify_ids(output, circuit_details)
                table_data.append([
                    circuit_link,
                    "EC",
                    circuit.get("label") or "(Unnamed Circuit)",
                    circuit_detail_link,
                    str(len(circuit_details))
                ])
    else:
        table_data.append(["", "", "None", "", "0"])
    output.print_table(
        table_data=table_data,
        title="Panels & Circuits",
        columns=headers,
        formats=["", "", "", "", ""]
    )

    detail_table_data = []
    if unmapped_details:
        for detail_id, detail_label in sorted(unmapped_details, key=lambda d: (d[1] or "", d[0])):
            detail_link = _linkify_ids(output, [detail_id])
            detail_table_data.append([
                detail_link,
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

    equipment_table_data = []
    if unmapped_panels:
        for panel in sorted(unmapped_panels, key=lambda p: p.get("name") or ""):
            panel_link = _linkify_ids(output, panel.get("ids", []))
            equipment_table_data.append([
                panel_link,
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

    if sc_missing_reference_records:
        output.print_md("## SC Detail Items Skipped")
        for record in sc_missing_reference_records:
            label_name = record["name"] if record["name"] else "(Unnamed Detail Item)"
            base_label = "`{}` (Id {})".format(label_name, record["id"])
            if record["view"]:
                base_label += " in view '{}'".format(record["view"])
            context_msgs = []
            for ctx in record["contexts"]:
                context_msgs.append(ctx["context"] + " missing: " + ", ".join(ctx["fields"]))
            output.print_md("* {} – {}".format(base_label, "; ".join(context_msgs)))


def main():
    doc = revit.doc
    logger.info("=== Syncing Circuit/Panel param values to detail items ===")

    active_view = _ensure_drafting_view(doc)
    if not active_view:
        return

    # 1) Build circuit dictionary keyed by (panel_name, circuit_number)
    logger.debug("Collecting circuits...")

    # design option filter
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)

    ckt_collector, circuit_map, circuit_map_by_id = _collect_circuits(doc, option_filter)

    # 2) Build panel dictionary keyed by panel_name
    logger.debug("Collecting panels...")

    panel_map, panel_map_by_id = _collect_panels(doc, option_filter)

    # 3) Collect detail items in the active view
    detail_items = _collect_detail_items(doc, option_filter, active_view)

    t = DB.Transaction(doc, "Sync Circuits/Panels to Detail Items")
    t.Start()
    update_count = 0
    sc_device_update_count = 0
    sc_missing_reference_records = []

    for ckt in ckt_collector:
        sc_device_update_count += update_devices_with_sc_ids(ckt)

    for ditem in detail_items:
        logger.debug("Detail item " + str(ditem.Id) + ":")
        # read which circuit/panel from the detail item’s custom parameters
        cpanel_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
        cnum_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)
        pname_val = get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)

        logger.debug("    cpanel_val='" + str(cpanel_val) + "' cnum_val='" + str(cnum_val) + "' pname_val='" + str(
            pname_val) + "'")
        changed = False
        missing_contexts = []
        sc_circuit_id_val = get_detail_param_value(ditem, DETAIL_PARAM_SC_CIRCUIT_ID)
        sc_panel_id_val = get_detail_param_value(ditem, DETAIL_PARAM_SC_PANEL_ID)

        circuit_id_lookup = None
        if sc_circuit_id_val not in (None, "", 0):
            circuit_id_lookup = str(sc_circuit_id_val).strip()
            if not circuit_id_lookup or circuit_id_lookup == "0":
                circuit_id_lookup = None

        panel_id_lookup = None
        if sc_panel_id_val not in (None, "", 0):
            panel_id_lookup = str(sc_panel_id_val).strip()
            if not panel_id_lookup or panel_id_lookup == "0":
                panel_id_lookup = None

        has_circuit_reference = bool(circuit_id_lookup) or (cpanel_val and cnum_val)
        if not has_circuit_reference:
            missing_fields = []
            if not circuit_id_lookup:
                missing_fields.append(DETAIL_PARAM_SC_CIRCUIT_ID)
            if not cpanel_val:
                missing_fields.append(DETAIL_PARAM_CKT_PANEL)
            if not cnum_val:
                missing_fields.append(DETAIL_PARAM_CKT_NUMBER)
            missing_contexts.append({
                "context": "Circuit reference",
                "fields": missing_fields
            })

        has_panel_reference = bool(panel_id_lookup) or bool(pname_val) or bool(cpanel_val)
        if not has_panel_reference:
            missing_fields = []
            if not panel_id_lookup:
                missing_fields.append(DETAIL_PARAM_SC_PANEL_ID)
            if not pname_val:
                missing_fields.append(DETAIL_PARAM_PANEL_NAME)
            if not cpanel_val:
                missing_fields.append(DETAIL_PARAM_CKT_PANEL)
            missing_contexts.append({
                "context": "Panel reference",
                "fields": missing_fields
            })

        # Resolve circuit data by ElementId first, then by (panel, number)
        cdict = None
        if circuit_id_lookup:
            if circuit_id_lookup in circuit_map_by_id:
                cdict = circuit_map_by_id[circuit_id_lookup]
                logger.debug("    Found circuit data for ElementId " + circuit_id_lookup)
            else:
                logger.debug("    No circuit data for ElementId " + circuit_id_lookup)
        if not cdict and cpanel_val and cnum_val:
            ckey = (str(cpanel_val), str(cnum_val))
            if ckey in circuit_map:
                cdict = circuit_map[ckey]
                logger.debug("    Found circuit data for key " + str(ckey))
            else:
                logger.debug("    No circuit data in circuit_map for key " + str(ckey))

        if cdict:
            for detail_pname, ckt_val in cdict.items():
                set_detail_param_value(ditem, detail_pname, ckt_val)
            changed = True

        # Resolve panel data: ElementId -> Panel Name -> CKT_Panel fallback
        pdict = None
        if panel_id_lookup:
            if panel_id_lookup in panel_map_by_id:
                pdict = panel_map_by_id[panel_id_lookup]
                logger.debug("    Found panel data for ElementId " + panel_id_lookup)
            else:
                logger.debug("    No panel data for ElementId " + panel_id_lookup)
        if not pdict and pname_val:
            panel_lookup_name = str(pname_val)
            if panel_lookup_name in panel_map:
                pdict = panel_map[panel_lookup_name]
                logger.debug("    Found panel data for '" + panel_lookup_name + "' (from " + DETAIL_PARAM_PANEL_NAME + ")")
            else:
                logger.debug("    No panel data in panel_map for '" + panel_lookup_name + "' (from " + DETAIL_PARAM_PANEL_NAME + ")")
        if not pdict and cpanel_val:
            panel_lookup_name = str(cpanel_val)
            if not pname_val:
                logger.debug("    Panel reference missing '" + DETAIL_PARAM_PANEL_NAME + "' so inferring from '" + DETAIL_PARAM_CKT_PANEL + "' value '" + panel_lookup_name + "'")
            else:
                logger.debug("    Panel name '" + str(pname_val) + "' not found; trying '" + DETAIL_PARAM_CKT_PANEL + "' value '" + panel_lookup_name + "' instead")
            if panel_lookup_name in panel_map:
                pdict = panel_map[panel_lookup_name]
                logger.debug("    Found panel data for '" + panel_lookup_name + "' (from " + DETAIL_PARAM_CKT_PANEL + ")")
            else:
                logger.debug("    No panel data in panel_map for '" + panel_lookup_name + "' (from " + DETAIL_PARAM_CKT_PANEL + ")")

        if pdict:
            for detail_pname, pval in pdict.items():
                set_detail_param_value(ditem, detail_pname, pval)
            changed = True

        if changed:
            logger.debug("    => Updated this detail item.")
            update_count += 1
        else:
            logger.debug("    => No changes for this detail item.")

        if missing_contexts and element_has_sc_params(ditem):
            name, view_name = describe_detail_item(doc, ditem)
            sc_missing_reference_records.append({
                "id": ditem.Id.IntegerValue,
                "name": name,
                "view": view_name,
                "contexts": missing_contexts
            })

    t.Commit()
    logger.info(
        "Sync finished. Updated " + str(update_count) + " detail item(s) and propagated SC ids to " + str(
            sc_device_update_count) + " device(s).")

    panel_summary, unmapped_details, unmapped_panels, sc_missing_reference_records = _build_output_summary(
        detail_items,
        circuit_map,
        circuit_map_by_id,
        panel_map,
        panel_map_by_id
    )

    choice = forms.alert(
        "Data sync complete.\n\nPrint output report?",
        ok=False,
        yes=True,
        no=True
    )
    if choice:
        _render_summary(
            panel_summary,
            unmapped_details,
            unmapped_panels,
            sc_missing_reference_records
        )


if __name__ == "__main__":
    main()
