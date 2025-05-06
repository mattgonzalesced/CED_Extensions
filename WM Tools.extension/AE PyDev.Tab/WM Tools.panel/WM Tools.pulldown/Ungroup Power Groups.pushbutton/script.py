# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script
from pyrevit.revit import query
from wmlib import *

PARAM_NAME = "Refrigeration Circuit Number_CEDT"
doc = revit.doc
logger = script.get_logger()




def main():
    groups = ChildGroup.collect_target_groups()
    if not groups:
        logger.warning("No matching model groups found to ungroup.")
        return

    with DB.Transaction(doc, "Ungroup Case Power Connections") as t:
        t.Start()
        for g in groups:
            cg = ChildGroup.from_existing_group(g)
            if cg:
                cg.ungroup_and_propagate()
        t.Commit()


if __name__ == "__main__":
    main()