# -*- coding: utf-8 -*-
# IRONPYTHON 2.7 COMPATIBLE (no f-strings, no .format usage)
from pyrevit import revit, DB, script

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
DETAIL_PARAM_PANEL_NAME = "Panel Name_CEDT"

# Circuit built-in param -> we store them in a dictionary so we can quickly read
# Then we map them to detail param names that we’ll set.
# Key here is "DetailParamName" : "BuiltInParamOnCircuit"
CIRCUIT_VALUE_MAP = {
    "x VD Schedule": "x VD Schedule",
    "Circuit Tree Sort_CED":"Circuit Tree Sort_CED",
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "CKT_Schedule Notes_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM,
    "CKT_Length_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "Voltage_CED":DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Wire Material_CEDT":"Wire Material_CEDT",
    "Wire Insulation_CEDT":"Wire Insulation_CEDT",
    "Wire Temperature Rating_CEDT" : "Wire Temperature Rating_CEDT",
    "Wire Size_CEDT":"Wire Size_CEDT",
    "Conduit and Wire Size_CEDT":"Conduit and Wire Size_CEDT",
    "Conduit Type_CEDT":"Conduit Type_CEDT",
    "Conduit Size_CEDT":"Conduit Size_CEDT",
    "Conduit Fill Percentage_CEDT":"Conduit Fill Percentage_CED",
    "Voltage Drop Percentage_CEDT":"Voltage Drop Percentage_CED",
}

# Panel built-in param -> detail param name map
PANEL_VALUE_MAP = {
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Mains Rating_CED": "Mains Rating_CED",
    "Mains Type_CEDT" : "Mains Type_CEDT",
    "Phase_CED":"Phase_CED",
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
    "Max Number of Circuits_CED":DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS,
    "Transformer Rating_CEDT": "Transformer Rating_CEDT",
    "Transformer Rating_CED": "Transformer Rating_CEDT",
    "Transformer Primary Description_CEDT":"Transformer Primary Description_CEDT",
    "Transformer Secondary Description_CEDT":"Transformer Secondary Description_CEDT",
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
        logger.debug("      set_detail_param_value: Set '" + param_name + "' to '" + str(new_value) + "' on " + str(elem.Id))
    except:
        logger.debug("      set_detail_param_value: FAILED setting '" + param_name + "' on " + str(elem.Id))

def is_not_in_group(element):
    return element.GroupId == DB.ElementId.InvalidElementId

def main():
    doc = revit.doc
    logger.info("=== Syncing Circuit/Panel param values to detail items ===")

    # 1) Build circuit dictionary keyed by (panel_name, circuit_number)
    circuit_map = {}
    logger.debug("Collecting circuits...")

    ckt_collector = DB.FilteredElementCollector(doc)\
                      .OfClass(DB.Electrical.ElectricalSystem)\
                      .ToElements()

    for ckt in ckt_collector:
        pval = get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        cnum = get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        if pval and cnum:
            key = (str(pval), str(cnum))
            # gather all relevant param values into a sub-dict
            cdata = {}
            for detail_param_name, bip in CIRCUIT_VALUE_MAP.items():
                cdata[detail_param_name] = get_model_param_value(ckt, bip)
            circuit_map[key] = cdata
            logger.debug("  Circuit " + str(ckt.Id) + " => key " + str(key) + " stored")
        else:
            logger.debug("  Circuit " + str(ckt.Id) + " missing panel or circuit # => skipping")

    # 2) Build panel dictionary keyed by panel_name
    panel_map = {}
    logger.debug("Collecting panels...")

    pnl_collector = DB.FilteredElementCollector(doc)\
                      .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)\
                      .WhereElementIsNotElementType()\
                      .ToElements()

    for pnl in pnl_collector:
        pname = get_model_param_value(pnl, DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        if pname:
            pdata = {}
            for detail_param_name, bip in PANEL_VALUE_MAP.items():
                pdata[detail_param_name] = get_model_param_value(pnl, bip)
            panel_map[str(pname)] = pdata
            logger.debug("  Panel " + str(pnl.Id) + " => name '" + pname + "' stored")
        else:
            logger.debug("  Panel " + str(pnl.Id) + " has no name => skipping")

    # 3) Collect detail items
    detail_item_collector = DB.FilteredElementCollector(doc)\
                     .OfCategory(DB.BuiltInCategory.OST_DetailComponents)\
                     .WhereElementIsNotElementType()\
                     .ToElements()

    detail_items = [el for el in detail_item_collector if is_not_in_group(el)]

    logger.debug("Collected " + str(len(detail_items)) + " detail item(s).")

    t = DB.Transaction(doc, "Sync Circuits/Panels to Detail Items")
    t.Start()
    update_count = 0

    for ditem in detail_items:
        logger.debug("Detail item " + str(ditem.Id) + ":")
        # read which circuit/panel from the detail item’s custom parameters
        cpanel_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
        cnum_val = get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)
        pname_val = get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)

        logger.debug("    cpanel_val='" + str(cpanel_val) + "' cnum_val='" + str(cnum_val) + "' pname_val='" + str(pname_val) + "'")
        changed = False

        # If circuit reference is found, see if we can retrieve that from circuit_map
        if cpanel_val and cnum_val:
            ckey = (str(cpanel_val), str(cnum_val))
            if ckey in circuit_map:
                cdict = circuit_map[ckey]
                logger.debug("    Found circuit data for key " + str(ckey))
                for detail_pname, ckt_val in cdict.items():
                    set_detail_param_value(ditem, detail_pname, ckt_val)
                changed = True
            else:
                logger.debug("    No circuit data in circuit_map for key " + str(ckey))

        # If panel reference is found, see if we have that in panel_map
        if pname_val:
            if pname_val in panel_map:
                pdict = panel_map[pname_val]
                logger.debug("    Found panel data for '" + pname_val + "'")
                for detail_pname, pval in pdict.items():
                    set_detail_param_value(ditem, detail_pname, pval)
                changed = True
            else:
                logger.debug("    No panel data in panel_map for '" + pname_val + "'")

        if changed:
            logger.debug("    => Updated this detail item.")
            update_count += 1
        else:
            logger.debug("    => No changes for this detail item.")

    t.Commit()
    logger.info("Sync finished. Updated " + str(update_count) + " detail item(s).")

if __name__ == "__main__":
    main()
