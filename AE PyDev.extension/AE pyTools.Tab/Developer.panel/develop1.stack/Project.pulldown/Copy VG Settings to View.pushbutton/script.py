# -*- coding: utf-8 -*-
# Copy V/G Settings to View Templates
# Revit Python 2.7 – pyRevit / Revit API

from pyrevit import revit, DB, forms, script

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView
output = script.get_output()
output.close_others()
accepted_views = [
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.DraftingView,
    DB.ViewType.ThreeD,
    DB.ViewType.Section
]

# ------------------------------------------------------------
# CATEGORY SELECTION
# ------------------------------------------------------------
def get_selected_categories():
    """Prompt user to select categories (model + annotation + import)."""
    all_cats = []

    for cat in doc.Settings.Categories:
        if not cat.AllowsVisibilityControl:
            continue

        # Determine category type
        try:
            bic = cat.BuiltInCategory
            is_valid = int(bic) != -1
        except Exception:
            is_valid = False

        if not is_valid:
            cat_type = "Import"
        elif cat.CategoryType == DB.CategoryType.Model:
            cat_type = "Model"
        elif cat.CategoryType == DB.CategoryType.Annotation:
            cat_type = "Anno"
        else:
            cat_type = "Other"

        all_cats.append((cat, cat_type))

    # Sort alphabetically
    all_cats = sorted(all_cats, key=lambda x: x[0].Name.lower())

    # Build display names
    names = ["{} [{}]".format(c.Name, t) for c, t in all_cats]

    picked = forms.SelectFromList.show(
        names,
        title="Select Categories to Copy V/G Settings",
        multiselect=True
    )
    if not picked:
        return []

    return [c for c, t in all_cats if "{} [{}]".format(c.Name, t) in picked]


# ------------------------------------------------------------
# COLLECTOR FUNCTION
# ------------------------------------------------------------
def collect_overrides_and_hiding(view, categories):
    """Collect OverrideGraphicSettings and hidden states for categories + subcats."""
    overrides = {}
    hidden_ids = []

    for cat in categories:
        try:
            # --- Main category ---
            ogs = view.GetCategoryOverrides(cat.Id)
            if ogs:
                overrides[cat.Id] = ogs
            if view.GetCategoryHidden(cat.Id):
                hidden_ids.append(cat.Id)

            # --- Subcategories ---
            for subcat in cat.SubCategories:
                try:
                    sub_ogs = view.GetCategoryOverrides(subcat.Id)
                    if sub_ogs:
                        overrides[subcat.Id] = sub_ogs
                    if view.GetCategoryHidden(subcat.Id):
                        hidden_ids.append(subcat.Id)
                except Exception as e:
                    logger.warning("Subcategory error for {} › {}: {}".format(cat.Name, subcat.Name, e))

        except Exception as e:
            logger.warning("Failed collecting settings for {}: {}".format(cat.Name, e))
            continue

    logger.debug("Collected {} overrides, {} hidden categories/subcategories".format(
        len(overrides), len(hidden_ids)
    ))

    return overrides, hidden_ids




def get_view_templates():
    """Return dict of available view templates with formatted names."""
    all_templates = DB.FilteredElementCollector(doc)\
                      .OfClass(DB.View)\
                      .ToElements()

    dict_views = {}
    for view in all_templates:
        if not view.IsTemplate:
            continue
        logger.info("getting view: {}".format(view.Name))
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
        elif view.ViewType == DB.ViewType.Schedule:
            logger.info("skipping Schedule view type...")
            continue
        else:
            logger.info("skipping unknown view type...")
            continue

    return dict_views

# ------------------------------------------------------------
# APPLY FUNCTION (modular, with optional reporting)
# ------------------------------------------------------------
def apply_vg_to_view(view, overrides, hidden_ids):
    """Apply category visibility/graphics to one view (template). Returns results list."""
    results = []

    all_ids = set(overrides.keys() + hidden_ids)
    for catid in all_ids:
        result = {'id': catid, 'override': False, 'hidden': False, 'failed': False}

        try:
            if catid in overrides:
                view.SetCategoryOverrides(catid, overrides[catid])
                result['override'] = True
        except Exception as e:
            result['override'] = 'FAILED'
            result['failed'] = True
            logger.warning("Override failed in {}: {}, Category:{}".format(view.Name, e,))

        try:
            if catid in hidden_ids:
                view.SetCategoryHidden(catid, True)
                result['hidden'] = True
        except Exception as e:
            result['hidden'] = 'FAILED'
            result['failed'] = True
            logger.warning("Hide failed in {}: {}".format(view.Name, e))

        results.append(result)

    return results


def apply_to_templates(view_templates, overrides, hidden_ids, show_report=False):
    """Apply collected settings to all templates. Markdown summary optional."""


    with revit.Transaction("Copy V/G Settings to View Templates"):
        summary = {}

        for vt in view_templates:
            logger.debug("Applying to template: {}".format(vt.Name))
            results = apply_vg_to_view(vt, overrides, hidden_ids)
            summary[vt.Name] = results

    if show_report:
        output.print_md("# Copy V/G Settings Report\n")
        for vt_name, results in summary.items():
            output.print_md("## {}".format(vt_name))
            output.print_md("| Category ID | Hidden | Override |")
            output.print_md("|--------------|--------|-----------|")

            for r in results:
                output.print_md("| {} | {} | {} |".format(r['id'], r['hidden'], r['override']))

    return summary



# ------------------------------------------------------------
# MAIN EXECUTION
# ------------------------------------------------------------
def main():
    if active_view.ViewType not in accepted_views:
        forms.alert("Active View must be:\n"
                    "       Floor Plan/RCP\n"
                    "       3D View\n"
                    "       Section/Elevation\n"
                    "       or Drafting View.",
                    title= "Incompatible View Type!",
                    exitscript=True
                    )
    categories = get_selected_categories()
    if not categories:
        forms.alert("No categories selected.")
        return

    overrides, hidden_ids = collect_overrides_and_hiding(active_view, categories)

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

    apply_to_templates(templates, overrides, hidden_ids, show_report=False)
    forms.alert("Copy V/G Settings complete.")


if __name__ == "__main__":
    main()