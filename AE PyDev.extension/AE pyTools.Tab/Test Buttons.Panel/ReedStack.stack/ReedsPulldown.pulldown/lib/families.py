# -*- coding: utf-8 -*-
# lib/families.py
import os, re, clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    BuiltInCategory, FilteredElementCollector, FamilySymbol, IFamilyLoadOptions, Transaction
)

class _AlwaysLoad(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        try: overwriteParameterValues[0] = True
        except: pass
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        try: overwriteParameterValues[0] = True
        except: pass
        return True

def _collect_symbols_any(doc):
    pairs = []
    def add(col):
        for fs in col:
            try:
                fam = fs.Family
                pairs.append((fam.Name if fam else "", fs.Name, fs))
            except:
                continue
    try: add(FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_LightingFixtures))
    except: pass
    try: add(FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_LightingDevices))
    except: pass
    if not pairs:
        try: add(FilteredElementCollector(doc).OfClass(FamilySymbol))
        except: pass
    return pairs

def _try_load_family_shell(doc, rfa_path):
    if not (rfa_path and os.path.exists(rfa_path)): return False
    t = Transaction(doc, "Load Family Shell")
    t.Start()
    try:
        ok = doc.LoadFamily(rfa_path, _AlwaysLoad())
        t.Commit(); return ok
    except Exception:
        try: t.RollBack()
        except: pass
        return False

def _parse_type_catalog(txt_path):
    if not os.path.exists(txt_path): return []
    with open(txt_path, 'rb') as f: data = f.read()
    text = None
    for enc in ('utf-8-sig', 'utf-16', 'latin-1'):
        try: text = data.decode(enc); break
        except: pass
    if text is None: return []
    names = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith('//'): continue
        sep_idx = None
        for sep in (',', '\t', ';'):
            i = line.find(sep)
            if i != -1: sep_idx = i; break
        first = line if sep_idx is None else line[:sep_idx]
        first = first.strip().strip('"').strip("'")
        if first: names.append(first)
    return names

def _collapse_ws(s): return " ".join((s or "").split())

def _choose_catalog_type(type_names, want_exact, want_regex):
    if not type_names: return None
    if want_exact and want_exact in type_names: return want_exact
    if want_exact:
        lowmap = {n.lower(): n for n in type_names}
        if want_exact.lower() in lowmap: return lowmap[want_exact.lower()]
        cmap = { _collapse_ws(n).lower(): n for n in type_names }
        key = _collapse_ws(want_exact).lower()
        if key in cmap: return cmap[key]
    if want_regex:
        try:
            rx = re.compile(want_regex, re.IGNORECASE)
            for n in type_names:
                if rx.search(n): return n
        except: pass
    return type_names[0]

def _try_load_symbol_from_catalog(doc, rfa_path, type_name):
    if not (rfa_path and os.path.exists(rfa_path) and type_name): return None
    try: out = clr.StrongBox[FamilySymbol]()
    except: out = clr.StrongBox[FamilySymbol](None)
    t = Transaction(doc, "Load Family Symbol (Catalog)")
    t.Start()
    try:
        ok = doc.LoadFamilySymbol(rfa_path, type_name, out)
        if ok and out.Value:
            sym = out.Value
            try:
                if not sym.IsActive: sym.Activate()
            except: pass
            t.Commit()
            print("[LOAD] Loaded type from catalog:", type_name)
            return sym
        t.RollBack(); return None
    except Exception:
        try: t.RollBack()
        except: pass
        return None

def pick_fixture_symbol(doc, rule):
    cands = (rule.get('fixture_candidates') or [])
    if not cands:
        print("[MATCH] No fixture_candidates in rule."); return None

    for cand in cands:
        fam_req   = (cand.get('family') or "").strip()
        typ_exact = (cand.get('type_catalog_name') or cand.get('type') or "").strip()
        typ_regex = (cand.get('type_regex') or "").strip()
        rfa_path  = (cand.get('load_from') or "").strip()

        # already-loaded?
        pairs = _collect_symbols_any(doc)
        for f, t, fs in pairs:
            if f == fam_req and (not typ_exact or t == typ_exact):
                try:
                    if not fs.IsActive: fs.Activate()
                except: pass
                print("[MATCH] Exact:", f, "::", t); return fs
            if f.lower() == fam_req.lower() and (not typ_exact or t.lower() == typ_exact.lower()):
                try:
                    if not fs.IsActive: fs.Activate()
                except: pass
                print("[MATCH] Case-insensitive:", f, "::", t); return fs

        # ensure family is loaded
        if rfa_path and os.path.exists(rfa_path):
            _try_load_family_shell(doc, rfa_path)
            # type catalog?
            cat_path = os.path.splitext(rfa_path)[0] + ".txt"
            if os.path.exists(cat_path):
                names = _parse_type_catalog(cat_path)
                if names:
                    print("[CATALOG] Found {} type names.".format(len(names)))
                    chosen = _choose_catalog_type(names, typ_exact, typ_regex)
                    if chosen:
                        sym = _try_load_symbol_from_catalog(doc, rfa_path, chosen)
                        if sym: return sym
            # re-scan for built-in types
            pairs = _collect_symbols_any(doc)
            for f, t, fs in pairs:
                if f == fam_req and (not typ_exact or t == typ_exact):
                    try:
                        if not fs.IsActive: fs.Activate()
                    except: pass
                    print("[MATCH] Exact after load:", f, "::", t); return fs

        print("[MATCH] Candidate did not resolve a symbol; trying next…")

    pairs = _collect_symbols_any(doc)
    if not pairs:
        print("[MATCH] No FamilySymbol instances are loaded in this model.")
    else:
        print("[MATCH] Available symbols (Family :: Type):")
        for f, t, _ in sorted(pairs)[:50]:
            print("  - {} :: {}".format(f, t))
    return None