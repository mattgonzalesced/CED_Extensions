# -*- coding: utf-8 -*-
__title__ = "Select Similar Family Instances in Model"
__author__ = "Erik Frits"
__doc__ = """Version = 1.0
Date    = 22.08.2022
_____________________________________________________________________
Description:
Select all instances in the project of the same Family.
_____________________________________________________________________
How-to:
- Select a single element
- Get All instances of the same family in Model
_____________________________________________________________________
Last update:
- [22.08.2022] - 1.0 RELEASE
_____________________________________________________________________
"""
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import clr
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from pyrevit import script


logger = script.get_logger()

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
# ==================================================
# doc   = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app     = __revit__.Application
rvt_year = int(app.VersionNumber)

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝ FUNCTION
# ==================================================

def select_similar_by_family(uidoc, mode):
    doc = uidoc.Document
    selected_elements = uidoc.Selection.GetElementIds()
    list_of_filters = List[ElementFilter]()

    for el_id in selected_elements:
        selected_element = doc.GetElement(el_id)

        try:
            elem_type_id = selected_element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
            if not elem_type_id or not elem_type_id.HasValue:
                continue  # Skip non-typed/system elements

            elem_type = doc.GetElement(elem_type_id.AsElementId())
            if not elem_type:
                continue  # Safety check

            elem_family_name = elem_type.FamilyName
            if not elem_family_name:
                continue  # Skip if no family name

            f_parameter = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME))

            if rvt_year < 2023:
                f_rule = FilterStringRule(f_parameter, FilterStringEquals(), elem_family_name, True)
            else:
                f_rule = FilterStringRule(f_parameter, FilterStringEquals(), elem_family_name)

            filter_family_name = ElementParameterFilter(f_rule)
            list_of_filters.Add(filter_family_name)

        except Exception as e:
            # Log and skip problematic elements
            category_name = selected_element.Category.Name if selected_element.Category else "No Category"
            logger.debug("Skipped element {} ({}): {}".format(
                el_id.IntegerValue,
                category_name,
                str(e)
            ))
            continue

    if list_of_filters:
        #>>>>>>>>>> COMBINE FILTERS
        multiple_filters = LogicalOrFilter(list_of_filters)

        # GET ELEMENTS
        elements_by_f_name = []
        if mode   == 'model':
            elements_by_f_name = FilteredElementCollector(doc)\
                    .WherePasses(multiple_filters).WhereElementIsNotElementType().ToElementIds()
        elif mode == 'view':
            elements_by_f_name = FilteredElementCollector(doc, doc.ActiveView.Id)\
                    .WherePasses(multiple_filters).WhereElementIsNotElementType().ToElementIds()

        # SET SELECTION
        if elements_by_f_name:
            uidoc.Selection.SetElementIds(List[ElementId](elements_by_f_name))
