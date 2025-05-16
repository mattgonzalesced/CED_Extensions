# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script, HOST_APP
from pyrevit.revit import query, ui
import pyrevit.extensions as exts

doc = HOST_APP.doc
uidoc = HOST_APP.uidoc
logger = script.get_logger()

FAMILY_NAME = "EF-U_Refrig Power Connector-Balanced_CED-WM"
PARAM_NAME = "Symbol Visible_CED"

def get_toggle_state():
    """Returns 0 or 1 depending on the first type's current state, or None if indeterminate"""
    fam_types = DB.FilteredElementCollector(doc) \
        .OfClass(DB.FamilySymbol) \
        .WhereElementIsElementType()

    for symbol in fam_types:
        if symbol.Family.Name != FAMILY_NAME:
            continue
        param = symbol.LookupParameter(PARAM_NAME)
        if param and param.HasValue:
            return param.AsInteger()
    return None


def main():
    fam_types = DB.FilteredElementCollector(doc) \
        .OfClass(DB.FamilySymbol) \
        .WhereElementIsElementType()

    targets = [s for s in fam_types if s.Family.Name == FAMILY_NAME]
    if not targets:
        logger.warning("No types found for family '{}'.".format(FAMILY_NAME))
        return None

    initial_state = get_toggle_state()
    if initial_state is None:
        logger.warning("Could not determine initial toggle state.")
        return None

    new_state = 0 if initial_state else 1
    passed_items = []

    with revit.Transaction("Toggle Connector Symbols"):
        for symbol in targets:
            param = symbol.LookupParameter(PARAM_NAME)
            if not param or param.IsReadOnly:
                continue
            if param.StorageType == DB.StorageType.Integer:
                param.Set(new_state)
                passed_items.append(query.get_name(symbol))

    state_text = "ON" if new_state == 1 else "OFF"
    logger.success("Case Power Symbols toggled {} for {} Family Types".format(state_text, len(passed_items)))
    return state_text


# Smart button self-initialization
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    # Get state by reading it from a type
    state = get_toggle_state()

    off_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_OFF_ICON_FILE)
    on_icon = ui.resolve_icon_file(script_cmp.directory, "on.png")

    icon_path = on_icon if state == 1 else off_icon
    ui_button_cmp.set_icon(icon_path)


if __name__ == "__main__":
    main()
