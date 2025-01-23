# -*- coding: utf-8 -*-


from pyrevit import forms, revit, DB
from pyrevit.revit import Transaction


# Prompt user to select families from a checkbox list
def get_families():
    all_families = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
    family_options = sorted(
        ["{}: {}".format(family.FamilyCategory.Name, family.Name) for family in all_families if family.FamilyCategory],
        key=lambda x: (x.split(": ")[0], x.split(": ")[1])
    )

    selected_families = forms.SelectFromList.show(
        family_options,
        title="Select Families to Place",
        button_name="Place Families",
        multiselect=True
    )
    return [f.split(": ")[1] for f in selected_families] if selected_families else []


# Function to get the starting point from the user
def get_starting_point():
    try:
        point = revit.uidoc.Selection.PickPoint("Select Starting Point")
        if not point:
            forms.alert("No point selected. Exiting script.", exitscript=True)
        return point
    except Exception as e:
        forms.alert("Error picking point. Exiting script.", exitscript=True)
        return None


# Function to get the level of the current view
def get_current_view_level():
    current_view = revit.uidoc.ActiveView
    level_id = current_view.GenLevel.Id
    level = revit.doc.GetElement(level_id)
    return level


# Place all types of selected families
def place_family_types():
    selected_family_names = get_families()
    if not selected_family_names:
        forms.alert("No families selected.", exitscript=True)

    # Get starting point from the user
    starting_point = get_starting_point()
    if not starting_point:
        return
    x_start, y_start, z_start = starting_point.X, starting_point.Y, starting_point.Z

    # Get the level of the current view
    level = get_current_view_level()
    if not level:
        forms.alert("Could not retrieve level from the current view. Exiting script.", exitscript=True)

    y_offset = 0  # Offset for placing families along the Y-axis
    with Transaction("Place All Family Types"):
        for family_name in selected_family_names:
            # Retrieve family by name
            family = next((f for f in DB.FilteredElementCollector(revit.doc)
                          .OfClass(DB.Family).ToElements()
                           if f.Name == family_name), None)
            if not family:
                continue

            # Get and sort family types alphabetically by their "ALL_MODEL_TYPE_NAME" parameter
            family_types = sorted(
                [revit.doc.GetElement(type_id) for type_id in family.GetFamilySymbolIds()],
                key=lambda ft: ft.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            )

            # Place each type in the family on the same horizontal plane
            for idx, family_type in enumerate(family_types):
                if not family_type.IsActive:
                    family_type.Activate()  # Ensure type is active

                x_offset = idx * 10  # 10 ft for each type
                point = DB.XYZ(x_start + x_offset, y_start + y_offset, z_start)

                # Place the family type at the given level
                revit.doc.Create.NewFamilyInstance(
                    point,
                    family_type,
                    level,
                    DB.Structure.StructuralType.NonStructural
                )

            # Offset the Y-axis for the next family by 10 feet
            y_offset -= 10


# Run the function
place_family_types()
