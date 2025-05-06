# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script
from pyrevit.revit import query
import re
from collections import defaultdict
from wmlib import *


doc = revit.doc
logger = script.get_logger()


BOX_ONLY_NAME = "Case Power Tags - BOX ONLY"

def main():
    parameter_mapping = {
        "Circuit #": "Refrigeration Circuit Number_CEDT"
    }

    tags = collect_reference_tags()
    parents = [ParentElement.from_family_instance(t) for t in tags]
    model_type = get_model_group_type("Case Power - 1 Case, 3 Ckts")
    children = [ChildGroup(p, model_type) for p in parents if p]
    detail_types = get_attached_detail_types(model_type)

    box_only = next((dt for dt in detail_types if query.get_name(dt) == BOX_ONLY_NAME), None)

    if not box_only:
        logger.error("'{}' detail group is missing.".format(BOX_ONLY_NAME))
        script.exit()

    systems = defaultdict(list)
    for c in children:
        sys_id = extract_system_id(c.parent.circuit_number)
        if sys_id:
            systems[sys_id].append(c)

    with DB.Transaction(doc, "Place Case Power Groups & Write Circuit Info") as trans:
        trans.Start()
        for sys_id, group_members in systems.items():
            for c in group_members:
                c.place()
                c.rotate_to_match_parent()
                c.copy_parameters(parameter_mapping)
                c.attach_detail_group_by_type(box_only)
        trans.Commit()


if __name__ == "__main__":
    main()
