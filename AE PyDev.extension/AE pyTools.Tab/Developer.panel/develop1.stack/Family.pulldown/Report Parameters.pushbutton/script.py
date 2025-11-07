# -*- coding: utf-8 -*-
from pyrevit import DB, script
from pyrevit import HOST_APP


doc = HOST_APP.doc
output = script.get_output()
logger = script.get_logger()

def get_parameter_group_name(group_id):
    """Return readable group name from BuiltInParameterGroup enum."""
    try:
        return DB.LabelUtils.GetLabelFor(group_id)
    except:
        return str(group_id)

def describe_parameter(param):
    """Return info tuple for a FamilyParameter."""
    param_type = "Type" if param.IsInstance == False else "Instance"
    shared = "Shared" if param.IsShared else "Family"
    storage = str(param.StorageType)
    return (param.Definition.Name, storage, param_type, shared)

def collect_family_parameters():
    """Collect and sort FamilyParameters by group order as shown in Revit."""
    fam_mgr = doc.FamilyManager
    all_params = list(fam_mgr.GetParameters())

    grouped = {}
    for p in all_params:
        group = get_parameter_group_name(p.Definition.ParameterGroup)
        grouped.setdefault(group, []).append(p)

    # Sort groups by display order
    ordered_groups = sorted(grouped.keys(), key=lambda g: DB.LabelUtils.GetLabelFor(
        DB.BuiltInParameterGroup.get_BuiltInParameterGroupByName(g)
    ) if hasattr(DB.BuiltInParameterGroup, 'get_BuiltInParameterGroupByName') else g)

    return [(g, grouped[g]) for g in ordered_groups]

def main():
    if not doc.IsFamilyDocument:
        script.exit("Open a Family document to run this report.")

    output.print_md("### 📘 Family Parameter Report")
    output.print_md("**Family Name:** `{}`".format(doc.Title))
    output.print_md("")

    total = 0
    for group_name, params in collect_family_parameters():
        output.print_md("#### 🗂 {}".format(group_name))
        output.print_md("| Parameter | Storage Type | Type/Instance | Shared/Family |")
        output.print_md("|------------|---------------|----------------|----------------|")
        for p in sorted(params, key=lambda x: x.Definition.Name.lower()):
            name, storage, typeflag, shared = describe_parameter(p)
            output.print_md("| {} | {} | {} | {} |".format(name, storage, typeflag, shared))
            total += 1
        output.print_md("")

    output.print_md("**Total Parameters:** `{}`".format(total))


if __name__ == "__main__":
    main()
