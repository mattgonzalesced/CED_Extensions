# -*- coding: utf-8 -*-
# Copy V/G Settings to View Templates
# Revit Python 2.7 – pyRevit / Revit API

from pyrevit import revit, DB, forms, script

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView


def get_selected_categories():
    """Prompt user to select categories (model + annotation)."""
    all_cats = []
    for cat in doc.Settings.Categories:
        if cat.AllowsVisibilityControl:
            all_cats.append(cat)

    # prompt user
    names = [c.Name for c in all_cats]
    picked = forms.SelectFromList.show(
        names,
        title="Select Categories to Copy V/G Settings",
        multiselect=True
    )
    if not picked:
        return []

    return [c for c in all_cats if c.Name in picked]


def collect_overrides_and_hiding(view, categories):
    """Collect OverrideGraphicSettings and hidden states for categories + subcats."""
    overrides = {}
    hidden_cats = []
    hidden_subcats = []

    for cat in categories:
        try:
            print("---- Checking Category: {}".format(cat.Name))

            # category level
            ogs = view.GetCategoryOverrides(cat.Id)
            if ogs:
                overrides[cat.Id] = ogs
                print("  Collected override for: {}".format(cat.Name))
            if view.GetCategoryHidden(cat.Id):
                hidden_cats.append(cat.Id)
                print("  Category is hidden: {}".format(cat.Name))

            # subcategories
            if cat.SubCategories and cat.SubCategories.Size > 0:
                for subcat in cat.SubCategories:
                    try:
                        ogs = view.GetCategoryHidden(subcat.Id)
                        if view.GetCategoryHidden(subcat.Id):
                            hidden_subcats.append(subcat.Id)
                            print("    Subcategory: {} is hidden -> will be copied".format(subcat.Name))
                        else:
                            print("    Subcategory: {} is visible".format(subcat.Name))
                        if ogs:
                            overrides[subcat.Id] = ogs
                    except Exception as e:
                        print("    Error checking subcategory {}: {}".format(subcat.Name, e))
            else:
                print("    No subcategories for: {}".format(cat.Name))

        except Exception as e:
            msg = "Could not collect settings for {0}: {1}".format(cat.Name, e)
            print(msg)
            logger.debug(msg)
            continue

    print("Summary: {} overrides, {} hidden cats, {} hidden subcats".format(
        len(overrides), len(hidden_cats), len(hidden_subcats)
    ))

    return overrides, hidden_cats, hidden_subcats



def get_view_templates():
    """Return dict of available view templates with formatted names."""
    all_templates = DB.FilteredElementCollector(doc)\
                      .OfClass(DB.View)\
                      .ToElements()

    dict_views = {}
    for view in all_templates:
        if not view.IsTemplate:
            continue

        if view.ViewType == DB.ViewType.FloorPlan:
            dict_views["[FLOOR] {}".format(view.Name)] = view
        elif view.ViewType == DB.ViewType.CeilingPlan:
            dict_views["[CEIL] {}".format(view.Name)] = view
        elif view.ViewType == DB.ViewType.ThreeD:
            dict_views["[3D] {}".format(view.Name)] = view
        elif view.ViewType == DB.ViewType.Section:
            dict_views["[SEC] {}".format(view.Name)] = view
        elif view.ViewType == DB.ViewType.Elevation:
            dict_views["[EL] {}".format(view.Name)] = view
        elif view.ViewType == DB.ViewType.DraftingView:
            dict_views["[DRAFT] {}".format(view.Name)] = view
        else:
            dict_views["[?] {}".format(view.Name)] = view

    return dict_views


def apply_to_templates(view_templates, overrides, hidden_cats, hidden_subcats):
    """Apply collected settings to the selected view templates."""
    with revit.Transaction("Copy V/G Settings to View Templates"):
        for vt in view_templates:
            logger.debug("Applying to template: {0}".format(vt.Name))
            # apply category overrides
            for catid, ogs in overrides.items():
                try:
                    vt.SetCategoryOverrides(catid, ogs)
                except Exception as e:
                    logger.debug("Could not apply override to {0}: {1}".format(catid, e))
            # apply hidden categories
            for catid in hidden_cats:
                try:
                    vt.SetCategoryHidden(catid, True)
                except Exception as e:
                    logger.debug("Could not hide category {0}: {1}".format(catid, e))
            # apply hidden subcategories
            for subid in hidden_subcats:
                try:
                    vt.SetCategoryHidden(subid, True)
                except Exception as e:
                    logger.debug("Could not hide subcategory {0}: {1}".format(subid, e))


def main():
    categories = get_selected_categories()
    if not categories:
        forms.alert("No categories selected.")
        return

    overrides, hidden_cats, hidden_subcats = collect_overrides_and_hiding(active_view, categories)

    vt_dict = get_view_templates()
    if not vt_dict:
        forms.alert("No view templates found.")
        return

    picked = forms.SelectFromList.show(
        vt_dict.keys(),
        title="Select View Templates to Apply",
        multiselect=True
    )
    if not picked:
        return

    templates = [vt_dict[name] for name in picked]

    apply_to_templates(templates, overrides, hidden_cats, hidden_subcats)
    forms.alert("Copy V/G Settings complete.")


if __name__ == "__main__":
    main()
