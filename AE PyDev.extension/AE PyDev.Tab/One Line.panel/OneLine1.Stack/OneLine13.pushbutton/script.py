# -*- coding: utf-8 -*-
__title__ = "Tagging Tools"

from pyrevit import revit, DB, forms, script
from pyrevit.revit import Transaction

from Snippets.tagcatmap import TAG_CAT_MAP

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()


def get_tag_categories_for_element(category_obj):
    category_id = category_obj.Id.IntegerValue  # Get int ID of the element's category

    matching_tag_cats = [
        tag_cat for tag_cat, model_cat in TAG_CAT_MAP.items()
        if model_cat.value__ == category_id
    ]

    if not matching_tag_cats:
        logger.info("No tag categories found for category ID: {} ({})".format(
            category_id, category_obj.Name))
    else:
        logger.info("Found tag categories: {}".format([cat.ToString() for cat in matching_tag_cats]))

    return matching_tag_cats

def get_all_tag_family_symbols():
    tag_types = []
    for tag_cat in TAG_CAT_MAP.keys():
        tag_types += DB.FilteredElementCollector(doc) \
            .OfClass(DB.FamilySymbol) \
            .OfCategory(tag_cat) \
            .ToElements()
    return tag_types


def get_tag_types_for_category(element_category):
    tag_types = []
    tag_categories = get_tag_categories_for_element(element_category)

    all_tag_types = get_all_tag_family_symbols()

    for tag in all_tag_types:
        try:
            tag_cat = tag.Category
            if not tag_cat:
                continue
            if tag_cat.Id == DB.ElementId(DB.BuiltInCategory.OST_MultiCategoryTags):
                tag_types.append(tag)
            elif tag_cat.Id.IntegerValue in [c.value__ for c in tag_categories]:
                tag_types.append(tag)
        except Exception as e:
            logger.error("Error on tag {}: {}".format(tag.Name, e))

    logger.info("Found {} tag types for element category {}".format(len(tag_types), element_category.Name))
    return tag_types


def get_tag_label(tag_symbol):
    try:
        tag_cat = tag_symbol.Category.Name if tag_symbol.Category else "?"
        type_name = tag_symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        fam_name = tag_symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
        value = fam_name + " : " + type_name
        logger.info(value)
        return "[{}] {}".format(tag_cat, value)
    except Exception as e:
        logger.error("Error building label for tag type: {}".format(e))
        return "[?] <error>"


def get_element_bbox_center(element):
    bbox = element.get_BoundingBox(None)
    center = (bbox.Min + bbox.Max) * 0.5
    return bbox, center


def get_tag_column_locations(bbox, count, spacing=3.0):
    max_x = bbox.Max.X
    mid_y = (bbox.Min.Y + bbox.Max.Y) / 2.0
    center_z = (bbox.Min.Z + bbox.Max.Z) / 2.0

    start_y = mid_y + ((count - 1) * spacing / 2.0)
    tag_positions = []

    for i in range(count):
        pos = DB.XYZ(max_x + 8.0, start_y - i * spacing, center_z)
        tag_positions.append(pos)

    return tag_positions


def tag_element_with_types(element, tag_types):
    bbox, center = get_element_bbox_center(element)
    positions = get_tag_column_locations(bbox, len(tag_types))
    reference = DB.Reference(element)
    scale = revit.active_view.Scale
    elbow_offset = scale * 0.03  # approx 3% of a drawing scale inch

    with Transaction("Tag Element With Selected Tags"):
        for tag_type, loc in zip(tag_types, positions):
            if not tag_type.IsActive:
                tag_type.Activate()
                doc.Regenerate()

            # Create the tag
            tag = DB.IndependentTag.Create(
                doc,
                tag_type.Id,
                revit.active_view.Id,
                reference,
                True,
                DB.TagOrientation.Horizontal,
                loc
            )

            # Regenerate to ensure geometry is resolved
            doc.Regenerate()

            leader_param = tag.get_Parameter(DB.BuiltInParameter.LEADER_LINE)
            if leader_param:
                leader_param.Set(0)  # False

            # Get bounding box of the tag head
            tag_bbox = tag.get_BoundingBox(revit.active_view)

            if tag_bbox:
                tag_min_x = tag_bbox.Min.X
                elbow_pt = DB.XYZ(tag_min_x - elbow_offset, loc.Y, loc.Z)
            else:
                elbow_pt = loc + DB.XYZ(elbow_offset, 0, 0)

            # Set the tag head first
            tag.TagHeadPosition = loc

            # Re-enable leader line
            if leader_param:
                leader_param.Set(1)  # True

            # Now finalize leader
            tag.LeaderEndCondition = DB.LeaderEndCondition.Attached
            tag.SetLeaderElbow(reference, elbow_pt)


# --- Main ---
selection = revit.get_selection()
if not selection or len(selection) != 1:
    forms.alert("Select one element to tag.")
    script.exit()

element = selection[0]
category = element.Category
if not category:
    forms.alert("Selected element has no category.")
    script.exit()

valid_tags = get_tag_types_for_category(category)

if not valid_tags:
    script.exit()

options = [get_tag_label(tag) for tag in valid_tags]


selected = forms.SelectFromList.show(options, multiselect=True, title="Select Tag Types")

if not selected:
    script.exit()

selected_tag_types = [t for t in valid_tags if get_tag_label(t) in selected]

tag_element_with_types(element,selected_tag_types)




