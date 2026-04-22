# -*- coding: UTF-8 -*-
 
# -*- coding: UTF-8 -*-
from pyrevit import revit, DB, forms
doc = revit.doc
GREY = DB.Color(128, 128, 128)
 
CAD_KEYWORDS = ["x_bg", "x_base", "X_BG_Mezz" , "X_BG - Mezz"]
 
def get_matching_categories(doc):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance)
    found = {}
    for imp in collector.ToElements():
        try:
            if imp.IsLinked:
                name_lower = imp.Category.Name.lower()
                if any(kw in name_lower for kw in CAD_KEYWORDS):
                    found[imp.Category.Id.IntegerValue] = imp.Category
        except:
            pass
    return list(found.values())
 
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
 
cats = get_matching_categories(doc)
if not cats:
    forms.alert("No linked CAD matching X_BG or X_BASE found in this model.", exitscript=True)
 
subcats = []
cat_names = []
for cat in cats:
    subcats.extend(list(cat.SubCategories))
    cat_names.append(cat.Name)
 
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
    "Done!\n\nCAD links greyed: {}\nViews/Templates updated: {}\nLayer overrides applied: {}".format(
        ", ".join(cat_names), views_updated, layers_updated
    )
)
 