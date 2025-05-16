# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script, forms
from pyrevit.revit import query
from wmlib import *
from collections import defaultdict

import re

doc = revit.doc
logger = script.get_logger()

def pick_model_group():
    group_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_IOSModelGroups).WhereElementIsElementType()
    group_options = group_collector

    sorted_labels = sorted(["{} (ID: {})".format(DB.Element.Name.__get__(g), g.Id.IntegerValue) for g in group_options])

    selected = forms.SelectFromList.show(
        sorted_labels,
        title="Select a Model Group Type",
        multiselect=False
    )

    if not selected:
        logger.info("No selection made.")
        script.exit()

    for g in group_options:
        label = "{} (ID: {})".format(DB.Element.Name.__get__(g), g.Id.IntegerValue)
        if selected == label:
            return g

    logger.error("No matching group found.")
    script.exit()

def pick_detail_group(attached_detail_types):
    if not attached_detail_types:
        logger.warning("No attached detail groups available. Will place model group without detail..")
        return None

    labels = ["{} (ID: {})".format(query.get_name(dt), dt.Id.IntegerValue) for dt in attached_detail_types]
    selected = forms.SelectFromList.show(
        sorted(labels),
        title="Select a Detail Group to Attach",
        multiselect=False
    )
    if not selected:
        logger.warning("No detail group selected.")
        return None

    for dt in attached_detail_types:
        if selected.startswith(query.get_name(dt)):
            return dt

    logger.error("No matching detail group found.")
    return None

def main():
    parameter_mapping = {
        "Circuit #": "Refrigeration Circuit Number_CEDT"
    }

    model_type = pick_model_group()
    detail_types = get_attached_detail_types(model_type)
    default_detail = pick_detail_group(detail_types)

    selected_ids = revit.get_selection().element_ids
    tags = []

    if selected_ids:
        selected_elements = [doc.GetElement(eid) for eid in selected_ids]
        for inst in selected_elements:
            tags.append(inst)

        logger.info("Using {} selected element(s) as parent references.".format(len(tags)))
    else:
        tags = [
            inst for inst in DB.FilteredElementCollector(doc, doc.ActiveView.Id)
            .OfClass(DB.FamilyInstance)
            if inst.Symbol.Family.Name == "Refrigeration Case Tag - EMS"
            and query.get_name(inst.Symbol) == "EMS Circuit Label"
        ]
        logger.info("Using {} EMS tags from active view.".format(len(tags)))

    if not tags:
        logger.error("No valid parent elements found.")
        script.exit()

    parents = [ParentElement.from_family_instance(t) for t in tags]
    children = [ChildGroup(p, model_type) for p in parents if p]

    systems = defaultdict(list)
    for c in children:
        sys_id = extract_system_id(c.parent.circuit_number)
        if not sys_id:
            sys_id = "no_system"
        systems[sys_id].append(c)

    with DB.Transaction(doc, "Place Case Power Groups & Write Circuit Info") as trans:
        trans.Start()
        for sys_id, group_members in systems.items():
            for c in group_members:
                c.place()
                c.rotate_to_match_parent()

                if c.parent.circuit_number:
                    c.copy_parameters(parameter_mapping)
                else:
                    group_element = doc.GetElement(c.child_id)
                    param = group_element.LookupParameter("Refrigeration Circuit Number_CEDT")
                    if param and not param.IsReadOnly:
                        param.Set("ckt # not found")

                if default_detail:
                    c.attach_detail_group_by_type(default_detail)
        trans.Commit()

if __name__ == "__main__":
    main()
