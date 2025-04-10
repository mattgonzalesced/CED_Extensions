# -*- coding: UTF-8 -*-

import os
import math
import tempfile
from pyrevit import revit, forms, script, DB, HOST_APP


output = script.get_output()
output.close_others(True)
output.center()
output.set_title('Models Checker')
logger = script.get_logger()

doc = revit.doc


def pick_families():
    """
    Prompt user to pick from a flat list of families with category and name.
    Returns:
        The selected Revit Family object(s).
    """
    fam_collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
    logger.debug("Total families in document: {}".format(fam_collector.GetElementCount()))

    fam_options = []
    fam_lookup = {}

    for fam in fam_collector:
        fam_category = fam.FamilyCategory
        if not fam_category:
            logger.debug("Skipped family: {}, Category: {}".format(
                fam.Name, fam_category.Name if fam_category else "None"
            ))
            continue

        label = "{} | {}".format(fam.FamilyCategory.Name, fam.Name)
        fam_options.append(label)
        fam_lookup[label] = fam

    fam_options.sort()

    logger.debug("Family Options for Selection: {}".format(fam_options))

    selected_options = forms.SelectFromList.show(
        fam_options,
        title="Select a Family",
        multiselect=True
    )

    if not selected_options:
        logger.info("No family selected. Exiting script.")
        script.exit()

    selected_families = [fam_lookup[label] for label in selected_options if label in fam_lookup]

    return selected_families

FIELDS = ["Size", "Name", "Category", "Creator"]
# temporary path for saving families
temp_dir = os.path.join(tempfile.gettempdir(), "pyRevit_ListFamilySizes")
if not os.path.exists(temp_dir):
    os.mkdir(temp_dir)
save_as_options = DB.SaveAsOptions()
save_as_options.OverwriteExistingFile = True

def convert_size(size_bytes):
    if not size_bytes:
        return "N/A"
    size_unit = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    size = round(size_bytes / p, 2)
    return "{}{}".format(size, size_unit[i])


def print_totals(families):
    total_size = sum([fam_item.get("Size") or 0 for fam_item in families])
    print("%d families, found total size: %s\n\n" % (
        len(families), convert_size(total_size)))


def print_sorted(families, group_by):
    # group by provided field, use the next field in the list to sort by
    fields_rest = list(FIELDS)
    fields_rest.pop(FIELDS.index(group_by))
    sort_by = fields_rest[0]
    fields_sorted = [group_by] + fields_rest

    if group_by not in ["Creator", "Category"]: # do not group by name and size
        print("Sort by: %s" % group_by)
        print("; ".join(fields_sorted))
        families_grouped = {"": sorted(families, key=\
            lambda fam_item: fam_item.get(group_by),
                                       reverse=group_by=="Size")}
    else:
        print("Group by: %s" % group_by)
        print("Sort by: %s" % sort_by)
        print(";".join(fields_rest))
        # convert to grouped dict
        families_grouped = {}
        # reverse if sorted by Size
        for fam_item in sorted(families, key=\
                lambda fam_item: fam_item.get(sort_by),
                               reverse=sort_by=="Size"):
            group_by_value = fam_item[group_by]
            fam_item_reduced = dict(fam_item)
            fam_item_reduced.pop(group_by)
            if group_by_value not in families_grouped:
                families_grouped[group_by_value] = []
            families_grouped[group_by_value].append(fam_item_reduced)

    for group_value in sorted(families_grouped.keys()):
        print(50 * "-")
        print("%s: %s" % (group_by, group_value))
        for fam_item in families_grouped[group_value]:
            family_row = []
            for field in fields_sorted:
                value = fam_item.get(field)
                if value is None:
                    continue
                if field == "Size":
                    value = convert_size(value)
                family_row.append(value)
            print("; ".join(family_row))
        print_totals(families_grouped[group_value])

# main logic

picked_fams = pick_families()
picked_family_items = []
opened_families = [od.Title for od in HOST_APP.uiapp.Application.Documents
                   if od.IsFamilyDocument]

# ask use to choose sort option
sort_by = forms.CommandSwitchWindow.show(FIELDS,
     message='Sorting options:',
)
if not sort_by:
    script.exit()

with forms.ProgressBar(title="Getting Family File Sizes", cancellable=True) as pb:
    i = 0

    for fam in picked_fams:
        with revit.ErrorSwallower() as swallower:
            if fam.IsEditable:
                fam_doc = revit.doc.EditFamily(fam)
                fam_path = fam_doc.PathName
                # if the family path does not exists, save it temporary
                #  only if the wasn't opened when the script was started
                if fam_doc.Title not in opened_families and (
                        not fam_path or not os.path.exists(fam_path)):
                    # edit family
                    fam_doc = revit.doc.EditFamily(fam)
                    # save with temporary path, to know family size
                    fam_path = os.path.join(temp_dir, fam_doc.Title)
                    fam_doc.SaveAs(fam_path, save_as_options)

                fam_size = 0
                fam_category = fam.FamilyCategory.Name if fam.FamilyCategory \
                    else "N/A"
                fam_creator = \
                    DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc,
                                                                  fam.Id).Creator
                if fam_path and os.path.exists(fam_path):
                    fam_size = os.path.getsize(fam_path)
                picked_family_items.append({"Size": fam_size,
                                         "Creator": fam_creator,
                                         "Category": fam_category,
                                         "Name": fam.Name})
                # if the family wasn't opened before, close it
                if fam_doc.Title not in opened_families:
                    fam_doc.Close(False)
                    # remove temporary family
                    if fam_path.lower().startswith(temp_dir.lower()):
                        os.remove(fam_path)
        if pb.cancelled:
            break
        else:
            pb.update_progress(i, len(picked_fams))
        i += 1


# print results
print("Families overview:")
print_sorted(picked_family_items, sort_by)
print_totals(picked_family_items)



