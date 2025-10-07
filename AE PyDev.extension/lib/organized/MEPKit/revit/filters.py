# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

def of_category(doc, bic, types=False):
    col = FilteredElementCollector(doc).OfCategory(bic)
    return (col.WhereElementIsElementType() if types else col.WhereElementIsNotElementType())

def of_class(doc, cls, types=False):
    col = FilteredElementCollector(doc).OfClass(cls)
    return (col.WhereElementIsElementType() if types else col.WhereElementIsNotElementType())

def all_levels(doc):
    from Autodesk.Revit.DB import Level
    return list(of_class(doc, Level, types=False))

def by_name(elems, name):
    nm = name.strip().lower()
    return [e for e in elems if getattr(e, 'Name', '').strip().lower() == nm]