# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/symbols.py
# IronPython 2.7 / Revit 20xx-safe
from __future__ import absolute_import
import os, re, clr
clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, FamilySymbol, Family, BuiltInCategory,
    IFamilyLoadOptions, Transaction, SubTransaction,
    Level
)
from Autodesk.Revit.DB.Structure import StructuralType

# ------------------------------
# small tx helper (works inside/outside an outer tx)
def _in_tx(doc, name, fn):
    if doc.IsModifiable:
        st = SubTransaction(doc); st.Start()
        try:
            res = fn(); st.Commit(); return res
        except:
            try: st.RollBack()
            except: pass
            raise
    else:
        t = Transaction(doc, name); t.Start()
        try:
            res = fn(); t.Commit(); return res
        except:
            try: t.RollBack()
            except: pass
            raise

# ------------------------------
# std helpers you already use

def ensure_active(doc, symbol):
    if symbol and not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

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

# ------------------------------
# symbol collection & matching

def _collect_symbols(doc, electrical_only=True):
    col = FilteredElementCollector(doc).OfClass(FamilySymbol)
    if electrical_only:
        try:
            col = col.OfCategory(BuiltInCategory.OST_ElectricalFixtures)
        except:
            pass
    out = []
    for s in col:
        try:
            fam = s.Family
            out.append((fam.Name if fam else "", s.Name, s))
        except:
            pass
    return out

def _norm(s):
    return re.sub(r'\s+', ' ', (s or u'').strip().lower())

def resolve_symbol(doc, family_name, type_name=None):
    """Exact match first (case-insensitive), among Electrical Fixtures; fallback to all."""
    fam_req = _norm(family_name)
    typ_req = _norm(type_name) if type_name else None

    for elec_only in (True, False):
        for fam, typ, fs in _collect_symbols(doc, electrical_only=elec_only):
            if _norm(fam) == fam_req and (typ_req is None or _norm(typ) == typ_req):
                return fs
    return None

# ------------------------------
# type catalog support

def _parse_type_catalog(txt_path):
    if not os.path.exists(txt_path):
        return []
    with open(txt_path, 'rb') as f:
        data = f.read()
    text = None
    for enc in ('utf-8-sig', 'utf-16', 'latin-1', 'utf-8'):
        try:
            text = data.decode(enc); break
        except:
            pass
    if text is None:
        return []
    names = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith('//'):
            continue
        # first token up to comma/tab/semicolon
        cut = None
        for sep in (',', '\t', ';'):
            i = line.find(sep)
            if i != -1:
                cut = i; break
        first = line if cut is None else line[:cut]
        first = first.strip().strip('"').strip("'")
        if first:
            names.append(first)
    return names

def _choose_catalog_type(type_names, want_exact, want_regex, logger=None):
    if not type_names:
        return None
    # exact
    if want_exact and want_exact in type_names:
        if logger: logger.info(u"[CATALOG] exact: {}".format(want_exact))
        return want_exact
    # case/space loose exact
    if want_exact:
        low = {n.lower(): n for n in type_names}
        key = u" ".join(want_exact.split()).lower()
        if key in low:
            if logger: logger.info(u"[CATALOG] case/space: {}".format(low[key]))
            return low[key]
    # regex
    if want_regex:
        try:
            rx = re.compile(want_regex, re.IGNORECASE)
            for n in type_names:
                if rx.search(n):
                    if logger: logger.info(u"[CATALOG] regex: {}".format(n))
                    return n
        except:
            pass
    # fallback → first
    if logger: logger.info(u"[CATALOG] fallback first: {}".format(type_names[0]))
    return type_names[0]

# ------------------------------
# loading (family & catalog symbol)

class _AlwaysLoad(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        try: overwriteParameterValues = True
        except: pass
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        try: overwriteParameterValues = True
        except: pass
        return True

def _load_family_shell(doc, rfa_path, logger=None):
    if not (rfa_path and os.path.exists(rfa_path)):
        if logger: logger.warning(u"[LOAD] Family path not found: {}".format(rfa_path))
        return False
    def _do():
        fam_ref = clr.Reference[Family]()
        ok = doc.LoadFamily(rfa_path, _AlwaysLoad(), fam_ref)
        if logger: logger.info(u"[LOAD] Loaded family: {} (ok={})".format(rfa_path, ok))
        return ok
    try:
        return _in_tx(doc, "Load Family (Shell)", _do)
    except Exception as ex:
        if logger: logger.warning(u"[LOAD] Family shell failed: {}".format(ex))
        return False

def _load_symbol_from_catalog(doc, rfa_path, type_name, logger=None):
    if not (rfa_path and os.path.exists(rfa_path) and type_name):
        return None

    # Prefer StrongBox[FamilySymbol] overload; fall back if not available
    try:
        StrongBox = clr.StrongBox
        has_strongbox = True
    except:
        has_strongbox = False

    def _do():
        sym = None
        if has_strongbox:
            box = clr.StrongBox[FamilySymbol]()  # out param
            ok = doc.LoadFamilySymbol(rfa_path, type_name, box)
            if ok and box.Value:
                sym = box.Value
        else:
            ok = doc.LoadFamilySymbol(rfa_path, type_name, _AlwaysLoad())
            if ok:
                # resolve by type name now present
                for fam, typ, fs in _collect_symbols(doc, electrical_only=False):
                    if _norm(typ) == _norm(type_name):
                        sym = fs; break
        return sym

    try:
        sym = _in_tx(doc, "Load Family Symbol (Catalog)", _do)
        if sym:
            ensure_active(doc, sym)
            if logger: logger.info(u"[LOAD] Loaded catalog type: {}".format(type_name))
            return sym
        else:
            if logger: logger.warning(u"[LOAD] LoadFamilySymbol returned None for: {}".format(type_name))
            return None
    except Exception as ex:
        if logger: logger.warning(u"[LOAD] Exception loading catalog symbol: {}".format(ex))
        return None

# ------------------------------
# resolve (exact → catalog → loaded family’s symbols → fuzzy)

def _log_loaded_types(doc, fam_obj, logger=None):
    if not fam_obj or not logger:
        return
    names = []
    try:
        for sid in fam_obj.GetFamilySymbolIds():
            sym = doc.GetElement(sid)
            if sym: names.append(sym.Name)
    except:
        pass
    logger.info(u"[LOAD] Family '{}' types ({}): {}".format(
        fam_obj.Name, len(names),
        u", ".join(u"'{}'".format(n) for n in names) if names else u"<none>"
    ))

def resolve_or_load_symbol(doc, family_name=None, type_name=None, rfa_path=None, logger=None):
    # 0) already loaded?
    for fs in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            fam = fs.Family
            if fam and _eq(fam.Name, family_name) and (not type_name or _eq(fs.Name, type_name)):
                if not fs.IsActive:
                    try: fs.Activate()
                    except: pass
                return fs
        except: pass

    # 1) Try to load the catalog row directly by the requested type_name
    if rfa_path and os.path.exists(rfa_path) and type_name:
        sym = _try_load_symbol_from_catalog(doc, rfa_path, type_name, logger)
        if sym:
            return sym

        # 2) If that failed, read the .txt and retry with the canonical row name
        cat_path = os.path.splitext(rfa_path)[0] + ".txt"
        names = _parse_type_catalog(cat_path)  # already in your repo from receptacles
        if names:
            # tolerant match (case/space-insensitive), fallback to first if needed
            canonical = _choose_catalog_type(names, want_exact=type_name, want_regex=None, strict=False)
            if canonical and canonical != type_name:
                if logger: logger.info(u"[CATALOG] retry with canonical name: {}".format(canonical))
                sym = _try_load_symbol_from_catalog(doc, rfa_path, canonical, logger)
                if sym:
                    return sym

    # 3) As a last resort, load the family shell and then retry catalog load once more
    if rfa_path and os.path.exists(rfa_path):
        _try_load_family_shell(doc, rfa_path)
        if type_name:
            sym = _try_load_symbol_from_catalog(doc, rfa_path, type_name, logger)
            if sym:
                return sym
        # parse .txt and try canonical again
        cat_path = os.path.splitext(rfa_path)[0] + ".txt"
        names = _parse_type_catalog(cat_path)
        if names and type_name:
            canonical = _choose_catalog_type(names, want_exact=type_name, want_regex=None, strict=False)
            if canonical:
                sym = _try_load_symbol_from_catalog(doc, rfa_path, canonical, logger)
                if sym:
                    return sym

    # 4) Give up
    if logger:
        logger.warning(u"[LOAD] Could not resolve symbol → family='{}' type='{}' (path='{}')"
                       .format(family_name or "", type_name or "", rfa_path or ""))
    return None