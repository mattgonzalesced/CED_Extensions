# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/symbols.py
from __future__ import absolute_import
import os, clr
from Autodesk.Revit.DB import (FilteredElementCollector, FamilySymbol, Level, Family, IFamilyLoadOptions)
from Autodesk.Revit.DB.Structure import StructuralType





# --- loading support (so rules' "load_from" actually works) ---
class _AlwaysLoad(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues = True
        return True

def load_family(doc, family_path, logger=None):
    if not family_path or not os.path.exists(family_path):
        if logger: logger.warning("Family path not found: {}".format(family_path))
        return None
    fam_ref = clr.Reference[Family]()
    ok = doc.LoadFamily(family_path, _AlwaysLoad(), fam_ref)
    if not ok:
        if logger: logger.warning("LoadFamily failed: {}".format(family_path))
        return None
    if logger: logger.info("Loaded family from: {}".format(family_path))
    return fam_ref.Value

# --- symbol resolution & placement ---
def resolve_symbol(doc, family_name, type_name=None):
    fam_l = (family_name or '').strip().lower()
    type_l = (type_name or '').strip().lower() if type_name else None
    for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
        if s.FamilyName.strip().lower() != fam_l:
            continue
        if (type_l is None) or (s.Name.strip().lower() == type_l):
            return s
    return None

def resolve_or_load_symbol(doc, family_name, type_name=None, load_path=None, logger=None):
    sym = resolve_symbol(doc, family_name, type_name)
    if sym:
        return sym
    if load_path:
        fam = load_family(doc, load_path, logger=logger)
        if fam:
            # Try again after load
            return resolve_symbol(doc, family_name, type_name)
    # As a last resort, try a loose contains match on family name
    fam_l = (family_name or '').strip().lower()
    type_l = (type_name or '').strip().lower() if type_name else None
    for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
        if fam_l in s.FamilyName.strip().lower():
            if (type_l is None) or (s.Name.strip().lower() == type_l):
                if logger: logger.warning("Using loose match: {} :: {}".format(s.FamilyName, s.Name))
                return s
    return None

def ensure_active(doc, symbol):
    if not symbol.IsActive:
        symbol.Activate(); doc.Regenerate()

def any_level(doc):
    for L in FilteredElementCollector(doc).OfClass(Level):
        return L
    return None

def place_hosted(doc, host, symbol, point_xyz):
    ensure_active(doc, symbol)
    return doc.Create.NewFamilyInstance(point_xyz, symbol, host, StructuralType.NonStructural)

def place_free(doc, symbol, point_xyz, level=None):
    ensure_active(doc, symbol)
    level = level or any_level(doc)
    return doc.Create.NewFamilyInstance(point_xyz, symbol, level, StructuralType.NonStructural)