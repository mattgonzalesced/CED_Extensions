# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script
from pyrevit.revit import query
from wmlib import extract_system_id, ChildGroup

doc = revit.doc
logger = script.get_logger()

CIRCUIT_PARAM = "Refrigeration Circuit Number_CEDT"
SYSTEM_PARAM = "System Number_CEDT"

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
    target_groups = collect_groups_with_circuit_param()
    if not target_groups:
        logger.warning("No model groups with '{}' parameter found.".format(CIRCUIT_PARAM))
        return

    with DB.Transaction(doc, "Ungroup Power Groups & Propagate Circuit Info") as t:
        t.Start()
        for group in target_groups:
            try:
                instance = ChildGroup.from_existing_group(group, CIRCUIT_PARAM,doc)
                if not instance:
                    logger.warning("Could not create ChildGroup from group ID {}".format(group.Id))
                    continue
                instance.ungroup_and_propagate(CIRCUIT_PARAM, SYSTEM_PARAM)
            except Exception as e:
                logger.error("Error ungrouping group ID {}: {}".format(group.Id, e))
        t.Commit()

if __name__ == "__main__":
    main()
