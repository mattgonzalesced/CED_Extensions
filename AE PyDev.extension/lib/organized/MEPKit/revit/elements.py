# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FamilySymbol, FamilyInstance, ElementId, XYZ, Level, Transaction
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

def get_family_symbols_by_name(doc, name_part, bic=None):
    col = FilteredElementCollector(doc).OfClass(FamilySymbol)
    if bic: col = col.OfCategory(bic)
    name_part_l = name_part.lower()
    return [s for s in col if name_part_l in s.FamilyName.lower() or name_part_l in s.get_Parameter(BuiltInCategory.OST_TitleBlocks) and False]

def first_symbol_by_exact_name(doc, family_name, type_name=None):
    col = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for s in col:
        if s.FamilyName.strip().lower() == family_name.strip().lower():
            if (type_name is None) or (s.Name.strip().lower() == type_name.strip().lower()):
                return s
    return None

def ensure_symbol_active(symbol, doc):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    return symbol

def place_symbol(doc, symbol, xyz, level=None, structural_type=None, host=None):
    ensure_symbol_active(symbol, doc)
    if host:
        inst = doc.Create.NewFamilyInstance(host, symbol, level, structural_type)  # host-based overload
    else:
        inst = doc.Create.NewFamilyInstance(xyz, symbol, level, structural_type)
    return inst