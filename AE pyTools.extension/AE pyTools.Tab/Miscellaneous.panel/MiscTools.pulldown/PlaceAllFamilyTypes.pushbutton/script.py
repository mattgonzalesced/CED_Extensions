# -*- coding: utf-8 -*-
__title__ = "Place All Family Types"

from pyrevit import forms, revit, DB
from pyrevit.revit import Transaction

# -----------------------------
def is_tag_family(family):
    return family.FamilyCategory and family.FamilyCategory.Id.IntegerValue == int(DB.BuiltInCategory.OST_Tags)

def is_annotation_family(family):
    return family.FamilyCategory and family.FamilyCategory.CategoryType == DB.CategoryType.Annotation

def get_validated_view():
    view = revit.uidoc.ActiveView
    if view.ViewType in [DB.ViewType.ThreeD, DB.ViewType.Section, DB.ViewType.Elevation]:
        forms.alert("Script does not support 3D, Section, or Elevation views.", exitscript=True)
    return view

def get_view_level(view):
    return revit.doc.GetElement(view.GenLevel.Id) if hasattr(view, "GenLevel") and view.GenLevel else None

def get_families_for_view(view):
    all_families = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
    families = [f for f in all_families if not is_tag_family(f)]

    if view.ViewType == DB.ViewType.DraftingView:
        families = [f for f in families if is_annotation_family(f)]
    else:
        families = [f for f in families if not is_annotation_family(f)]

    family_options = sorted(
        ["{}: {}".format(f.FamilyCategory.Name, f.Name) for f in families if f.FamilyCategory],
        key=lambda x: (x.split(": ")[0], x.split(": ")[1])
    )

    selected = forms.SelectFromList.show(
        family_options, title="Select Families to Place", button_name="Place Families", multiselect=True
    )

    return [f.split(": ")[1] for f in selected] if selected else []

def get_starting_point():
    try:
        point = revit.uidoc.Selection.PickPoint("Select Starting Point")
        if not point:
            forms.alert("No point selected. Exiting script.", exitscript=True)
        return point
    except Exception:
        forms.alert("Error picking point. Exiting script.", exitscript=True)
        return None

def get_level_and_offset(view, z_start):
    if view.ViewType in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan]:
        level = get_view_level(view)
        if not level:
            forms.alert("No level found for this view. Exiting.", exitscript=True)
        return level, z_start - level.Elevation
    return None, z_start  # e.g., Drafting view

def place_families(family_names, view, start_point):
    x_start, y_start, z_start = start_point.X, start_point.Y, start_point.Z
    level, z_offset = get_level_and_offset(view, z_start)
    is_drafting = view.ViewType == DB.ViewType.DraftingView

    y_offset = 0
    with Transaction("Place All Family Types"):
        for family_name in family_names:
            family = next((f for f in DB.FilteredElementCollector(revit.doc)
                          .OfClass(DB.Family).ToElements() if f.Name == family_name), None)
            if not family:
                continue
            place_family_types_at_point(family, x_start, y_start + y_offset, z_offset, level, view, is_drafting)
            y_offset -= 10  # move down 10 ft for next family

def place_family_types_at_point(family, x_start, y_start, z_offset, level, view, is_drafting):
    family_types = sorted(
        [revit.doc.GetElement(type_id) for type_id in family.GetFamilySymbolIds()],
        key=lambda ft: ft.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    )

    for idx, family_type in enumerate(family_types):
        if not family_type.IsActive:
            family_type.Activate()

        x_offset = idx * 10
        point = DB.XYZ(x_start + x_offset, y_start, z_offset)

        if is_drafting:
            revit.doc.Create.NewFamilyInstance(point, family_type, view)
        else:
            revit.doc.Create.NewFamilyInstance(point, family_type, level, DB.Structure.StructuralType.NonStructural)

# -----------------------------
# Main controller
def run():
    view = get_validated_view()
    family_names = get_families_for_view(view)
    if not family_names:
        return
    start_point = get_starting_point()
    if not start_point:
        return
    place_families(family_names, view, start_point)

# Execute script
run()
