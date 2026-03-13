# -*- coding: UTF-8 -*-
from pyrevit import revit, DB, forms, script

doc = revit.doc
active_view = doc.ActiveView

all_cats = doc.Settings.Categories
import_cats = []

for cat in all_cats:
    try:
        name = cat.Name
        low = name.lower()
        if ".dwg" in low or ".dxf" in low or ".dgn" in low or ".dwf" in low:
            import_cats.append(cat)
    except:
        pass

if not import_cats:
    forms.alert("No import categories found.", exitscript=True)

import_cat_dict = {}
for c in import_cats:
    import_cat_dict[c.Name] = c

selected_name = forms.SelectFromList.show(
    sorted(import_cat_dict.keys()),
    title="Select Import Category",
    button_name="Select",
    multiselect=False
)

if not selected_name:
    script.exit()

source_cat = import_cat_dict[selected_name]
sub_cats = source_cat.SubCategories
hidden_ids = []
visible_ids = []
hidden_names = []

for sc in sub_cats:
    try:
        if active_view.GetCategoryHidden(sc.Id):
            hidden_ids.append(sc.Id)
            hidden_names.append(sc.Name)
        else:
            visible_ids.append(sc.Id)
    except:
        pass

summary = """Source View: {0}
Import Category: {1}
Hidden Layers: {2}
Visible Layers: {3}
""".format(
    active_view.Name, selected_name, len(hidden_ids), len(visible_ids))
if hidden_names:
    summary += """
Hidden layers:
"""
    for n in sorted(hidden_names):
        summary += "  - {0}\n".format(n)

proceed = forms.alert(summary, title="Layer Visibility Summary", ok=False, yes=True, no=True)
if not proceed:
    script.exit()

ok_types = [
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.ThreeD,
    DB.ViewType.Section,
    DB.ViewType.AreaPlan,
    DB.ViewType.EngineeringPlan,
]

all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType().ToElements()

eligible_views = []
for v in all_views:
    try:
        if v.IsTemplate:
            continue
        if v.ViewType not in ok_types:
            continue
        if v.Id == active_view.Id:
            continue
        eligible_views.append(v)
    except:
        pass

if not eligible_views:
    forms.alert("No eligible views found.", exitscript=True)

type_groups = {}
for v in eligible_views:
    vt_name = str(v.ViewType)
    if vt_name == "ThreeD":
        vt_name = "3D View"
    if vt_name not in type_groups:
        type_groups[vt_name] = []
    type_groups[vt_name].append(v)

selected_types = forms.SelectFromList.show(
    sorted(type_groups.keys()),
    title="Filter by View Type",
    button_name="Next",
    multiselect=True
)
if not selected_types:
    script.exit()

filtered_views = []
for vt in selected_types:
    filtered_views.extend(type_groups[vt])
filtered_views.sort(key=lambda v: v.Name)


class ViewOption(forms.TemplateListItem):
    @property
    def name(self):
        template_info = ""
        vt_id = self.item.ViewTemplateId
        if vt_id and vt_id != DB.ElementId.InvalidElementId:
            vt_elem = doc.GetElement(vt_id)
            if vt_elem:
                template_info = "  [Template: {0}]".format(vt_elem.Name)
        return "{0}{1}".format(self.item.Name, template_info)


view_options = [ViewOption(v) for v in filtered_views]

selected_view_options = forms.SelectFromList.show(
    view_options,
    title="Select Target Views",
    button_name="Apply Settings",
    multiselect=True
)
if not selected_view_options:
    script.exit()

modified_templates = set()
views_modified = 0
templates_modified = 0
skipped_views = []

with revit.Transaction("Copy Import Layer Visibility"):
    for target_view in selected_view_options:
        try:
            vt_id = target_view.ViewTemplateId
            has_template = (vt_id is not None and vt_id != DB.ElementId.InvalidElementId)

            template = None
            if has_template:
                template = doc.GetElement(vt_id)
                if template is None:
                    has_template = False

            # Always modify the view itself
            for layer_id in hidden_ids:
                try:
                    target_view.SetCategoryHidden(layer_id, True)
                except:
                    pass

            for layer_id in visible_ids:
                try:
                    target_view.SetCategoryHidden(layer_id, False)
                except:
                    pass

            views_modified += 1

            # Also modify the template if it exists and hasn't been modified yet
            if has_template and template.Id.IntegerValue not in modified_templates:
                for layer_id in hidden_ids:
                    try:
                        template.SetCategoryHidden(layer_id, True)
                    except:
                        pass

                for layer_id in visible_ids:
                    try:
                        template.SetCategoryHidden(layer_id, False)
                    except:
                        pass

                modified_templates.add(template.Id.IntegerValue)
                templates_modified += 1

        except Exception as e:
            skipped_views.append("{0}: {1}".format(target_view.Name, str(e)))

result_msg = """Import Layer Visibility Applied!

"""
result_msg += "Source: {0}\n".format(active_view.Name)
result_msg += """Import: {0}

""".format(selected_name)
result_msg += "Views directly modified: {0}\n".format(views_modified)
result_msg += "View templates modified: {0}\n".format(templates_modified)

if skipped_views:
    result_msg += """
Skipped views:
"""
    for s in skipped_views:
        result_msg += "  - {0}\n".format(s)

forms.alert(result_msg, title="Complete")
