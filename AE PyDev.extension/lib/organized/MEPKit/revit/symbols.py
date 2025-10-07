# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/symbols.py
from __future__ import absolute_import
from Autodesk.Revit.DB import (FilteredElementCollector, FamilySymbol, Level)
from Autodesk.Revit.DB.Structure import StructuralType

def resolve_symbol(doc, family_name, type_name=None):
    fam_l = (family_name or '').strip().lower()
    type_l = (type_name or '').strip().lower() if type_name else None
    for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
        if s.FamilyName.strip().lower() != fam_l: continue
        if (type_l is None) or (s.Name.strip().lower() == type_l):
            return s
    return None

def ensure_active(doc, symbol):
    if not symbol.IsActive:
        symbol.Activate(); doc.Regenerate()

def any_level(doc):
    lvl = None
    for L in FilteredElementCollector(doc).OfClass(Level):
        lvl = L; break
    return lvl

def place_hosted(doc, wall, symbol, point_xyz):
    ensure_active(doc, symbol)
    return doc.Create.NewFamilyInstance(point_xyz, symbol, wall, StructuralType.NonStructural)

def place_free(doc, symbol, point_xyz, level=None):
    ensure_active(doc, symbol)
    level = level or any_level(doc)
    return doc.Create.NewFamilyInstance(point_xyz, symbol, level, StructuralType.NonStructural)