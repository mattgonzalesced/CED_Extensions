# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/symbols.py
from __future__ import absolute_import
import os, clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (FilteredElementCollector, FamilySymbol, Level, Family, IFamilyLoadOptions)
from Autodesk.Revit.DB.Structure import StructuralType





# --- loading support (so rules' "load_from" actually works) ---
class AlwaysLoad(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True; return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues = True; return True

def _resolve_symbol_exact(doc, family_name, type_name=None):
    fam_l = (family_name or '').strip().lower()
    type_l = (type_name or '').strip().lower() if type_name else None
    for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
        if s.FamilyName.strip().lower() != fam_l:
            continue
        if type_l is None or s.Name.strip().lower() == type_l:
            return s
    return None

def _log_family_symbols(doc, fam, logger=None, note=None):
    try:
        ids = list(fam.GetFamilySymbolIds())
        names = []
        for sid in ids:
            sym = doc.GetElement(sid)
            if sym:
                names.append(sym.Name)
        if logger:
            logger.info(u"Family '{}' types {}{}: {}".format(
                fam.Name,
                "(from {}) ".format(note) if note else "",
                "(count={})".format(len(names)),
                u", ".join(u"'{}'".format(n) for n in names) if names else "<none>"
            ))
    except:
        pass

def resolve_or_load_symbol(doc, family_name, type_name=None, load_path=None, logger=None):
    # 1) Already in project?
    sym = _resolve_symbol_exact(doc, family_name, type_name)
    if sym:
        return sym

    # 2) Type-catalog load (if type provided)
    opts = AlwaysLoad()
    if load_path and type_name:
        try:
            if doc.LoadFamilySymbol(load_path, type_name, opts):
                if logger: logger.info(u"Loaded catalog type: {} :: {}".format(family_name, type_name))
                sym = _resolve_symbol_exact(doc, family_name, type_name)
                if sym: return sym
        except:
            pass

    # 3) Load full family (non-catalog) and pick symbol from the loaded Family object
    fam_ref = clr.Reference[Family]()
    if load_path:
        try:
            if doc.LoadFamily(load_path, opts, fam_ref):
                fam = fam_ref.Value
                if logger: logger.info(u"Loaded family from: {}".format(load_path))
                _log_family_symbols(doc, fam, logger, note="loaded")
                # Prefer the requested type_name if provided
                if type_name:
                    type_l = type_name.strip().lower()
                    for sid in fam.GetFamilySymbolIds():
                        s = doc.GetElement(sid)
                        if s and s.Name.strip().lower() == type_l:
                            return s
                # Else: pick first available symbol
                for sid in fam.GetFamilySymbolIds():
                    s = doc.GetElement(sid)
                    if s: return s
        except:
            pass

    # 4) Last-ditch: loose match on family name anywhere
    fam_l = (family_name or '').strip().lower()
    for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
        if fam_l in s.FamilyName.strip().lower():
            if logger: logger.warning(u"Using loose match: {} :: {}".format(s.FamilyName, s.Name))
            return s

    if logger:
        logger.warning(u"Failed to resolve symbol: family='{}' type='{}' (path='{}')"
                       .format(family_name, type_name or '*', load_path or ''))
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