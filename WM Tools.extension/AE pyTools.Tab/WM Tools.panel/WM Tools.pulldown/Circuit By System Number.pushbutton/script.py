# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script
from pyrevit.revit import query
from wmlib import extract_system_id, ChildGroup

doc = revit.doc
logger = script.get_logger()


def collect_groups_with_circuit_param():
    selected_ids = revit.get_selection().element_ids
    if selected_ids:
        elements = [doc.GetElement(eid) for eid in selected_ids]
        logger.info("Using selection: {} elements.".format(len(elements)))
    else:
        elements = DB.FilteredElementCollector(doc, doc.ActiveView.Id) \
            .OfCategory(DB.BuiltInCategory.OST_IOSModelGroups) \
            .WhereElementIsNotElementType() \
            .ToElements()
        logger.info("No selection. Scanning all groups in view: {} elements.".format(len(elements)))

    valid_groups = []
    for g in elements:
        if not isinstance(g, DB.Group):
            continue
        param = g.LookupParameter(CIRCUIT_PARAM)
        if param and param.HasValue and param.AsString():
            logger.debug("Valid group: {} (ID {}) with circuit '{}'".format(g.Name, g.Id, param.AsString()))
            valid_groups.append(g)
        else:
            logger.debug("Skipping group without '{}' or value: {} (ID {})".format(CIRCUIT_PARAM, g.Name, g.Id))

    return valid_groups

def main():

if __name__ == "__main__":
    main()
