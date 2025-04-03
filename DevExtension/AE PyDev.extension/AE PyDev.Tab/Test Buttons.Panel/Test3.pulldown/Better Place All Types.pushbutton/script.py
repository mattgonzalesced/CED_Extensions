# -*- coding: utf-8 -*-

from pyrevit import forms, revit, DB, script
from pyrevit.revit import Transaction

output = script.get_output()
logger = script.get_logger()
output.close_others()

doc = revit.doc
uidoc = revit.uidoc


class FamilySelector(object):
    def __init__(self, view_information):
        self.view_information = view_information

    def get_selected_family_names(self):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
        allowed_families = []

        for family in collector:
            category = family.FamilyCategory
            if not category:
                continue

            category_id = category.Id

            if self.view_information.is_drafting_view():
                if category_id == DB.ElementId(DB.BuiltInCategory.OST_DetailComponents) or \
                   category_id == DB.ElementId(DB.BuiltInCategory.OST_GenericAnnotation):
                    allowed_families.append(family)
            else:
                allowed_families.append(family)

        logger.debug("Found {} allowed families.".format(len(allowed_families)))

        family_options = sorted(
            ["{}: {}".format(family.FamilyCategory.Name, family.Name)
             for family in allowed_families],
            key=lambda name: (name.split(": ")[0], name.split(": ")[1])
        )

        selected_items = forms.SelectFromList.show(
            family_options,
            title="Select Families to Place",
            button_name="Place Families",
            multiselect=True
        )
        return [item.split(": ")[1] for item in selected_items] if selected_items else []



class PointPicker(object):
    def get_starting_point(self):
        try:
            point = revit.uidoc.Selection.PickPoint("Select Starting Point")
            if not point:
                forms.alert("No point selected. Exiting script.", exitscript=True)
            return point
        except Exception:
            forms.alert("Error picking point. Exiting script.", exitscript=True)
            return None


class ViewContext(object):
    def __init__(self):
        self.view = revit.uidoc.ActiveView
        self.level = self._get_view_level()

    def _get_view_level(self):
        if hasattr(self.view, "GenLevel") and self.view.GenLevel:
            level = revit.doc.GetElement(self.view.GenLevel.Id)
            logger.debug("View Level found: {}".format(level.Name))
            return level
        logger.debug("No level associated with the current view.")
        return None

    def is_drafting_view(self):
        is_drafting = isinstance(self.view, DB.ViewDrafting)
        logger.debug("Is Drafting View: {}".format(is_drafting))
        return is_drafting

    def get_level(self):
        return self.level

    def get_view(self):
        return self.view


class FamilyTypePlacer(object):
    def __init__(self, family_names, base_point, view_information):
        self.family_names = family_names
        self.base_point = base_point
        self.view_information = view_information
        self.document = doc

    def place_family_types(self):
        y_offset_feet = 0.0

        with revit.Transaction("Place All Family Types"):
            for family_name in self.family_names:
                family = self._get_family_by_name(family_name)
                if not family:
                    logger.warning("Family '{}' not found.".format(family_name))
                    continue

                type_ids = family.GetFamilySymbolIds()
                if not type_ids:
                    logger.warning("Family '{}' has no types.".format(family_name))
                    continue

                family_types = sorted(
                    [self.document.GetElement(type_id) for type_id in type_ids],
                    key=lambda symbol: symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                )

                for index, family_symbol in enumerate(family_types):
                    if not family_symbol.IsActive:
                        family_symbol.Activate()
                        self.document.Regenerate()

                    x_offset_feet = index * 10.0
                    placement_point = DB.XYZ(
                        self.base_point.X + x_offset_feet,
                        self.base_point.Y + y_offset_feet,
                        self.base_point.Z
                    )

                    try:
                        if self.view_information.is_drafting_view():
                            self._place_in_drafting_view(family_symbol, placement_point)
                        else:
                            self._place_in_model_view(family_symbol, placement_point)
                    except Exception as place_exception:
                        logger.debug("Failed to place {}: {}".format(family_symbol, str(place_exception)))

                y_offset_feet -= 10.0

    def _get_family_by_name(self, name):
        return next((family for family in DB.FilteredElementCollector(self.document).OfClass(DB.Family)
                     if family.Name == name), None)

    def _place_in_model_view(self, symbol, point):
        self.document.Create.NewFamilyInstance(
            point,
            symbol,
            self.view_information.get_level(),
            DB.Structure.StructuralType.NonStructural
        )
        logger.debug("Placed model family: {} at {}".format(symbol.Name, point))

    def _place_in_drafting_view(self, symbol, point):
        self.document.Create.NewFamilyInstance(
            point,
            symbol,
            self.view_information.get_view()
        )
        logger.debug("Placed detail family: {} at {}".format(symbol.Name, point))



def main():
    view_ctx = ViewContext()
    selector = FamilySelector(view_ctx)
    selected_fams = selector.get_selected_family_names()
    if not selected_fams:
        forms.alert("No families selected.", exitscript=True)

    picker = PointPicker()
    start_point = picker.get_starting_point()
    if not start_point:
        return

    if not view_ctx.is_drafting_view() and not view_ctx.get_level():
        forms.alert("Could not retrieve level from current view.", exitscript=True)

    placer = FamilyTypePlacer(selected_fams, start_point, view_ctx)
    placer.place_family_types()


if __name__ == "__main__":
    main()
