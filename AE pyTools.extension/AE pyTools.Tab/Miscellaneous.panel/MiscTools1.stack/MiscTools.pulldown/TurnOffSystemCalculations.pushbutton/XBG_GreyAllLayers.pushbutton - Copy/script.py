# -*- coding: UTF-8 -*-

# -*- coding: UTF-8 -*-
from pyrevit import revit, DB, forms
doc = revit.doc

GREY = DB.Color(128, 128, 128)

def get_xbg_category(doc):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance)
    for imp in collector.ToElements():
        try:
            if imp.IsLinked and "X_BG" in imp.Category.Name:
                return imp.Category
        except:
            pass
    return None

def apply_grey_to_view(view, subcats):
    changed = 0
    for subcat in subcats:
        try:
            # Get current overrides without touching anything
            ogs = view.GetCategoryOverrides(subcat.Id)
            # Only set the projection line color to grey, leave everything else as-is
            ogs.SetProjectionLineColor(GREY)
            view.SetCategoryOverrides(subcat.Id, ogs)
            changed += 1
        except:
            pass
    return changed

cat = get_xbg_category(doc)
if not cat:
    forms.alert("X_BG.dwg linked CAD not found in this model.", exitscript=True)

subcats = list(cat.SubCategories)

# Collect all views and view templates
all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()

views_updated = 0
layers_updated = 0

with revit.Transaction("Grey All X_BG Layers"):
    for view in all_views:
        try:
            # Skip views that can't have overrides (schedules, legends, etc.)
            if view.ViewType in [
                DB.ViewType.Schedule,
                DB.ViewType.ColumnSchedule,
                DB.ViewType.PanelSchedule,
            ]:
                continue
            # Process both regular views and view templates
            n = apply_grey_to_view(view, subcats)
            if n > 0:
                views_updated += 1
                layers_updated += n
        except:
            pass

forms.alert(
    "Done!\n\nViews/Templates updated: {}\nLayer overrides applied: {}".format(
        views_updated, layers_updated
    )
)
