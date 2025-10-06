# -*- coding: utf-8 -*-
# RP — Mech Rough-In by Space Rules (ductwork + rooftop fans)
# Revit 2024/2025 | pyRevit (IronPython 2.7)
from __future__ import print_function
import os, sys, json, math, clr

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter, XYZ, Transaction,
    Level, Family, FamilySymbol, IFamilyLoadOptions,
    ElementCategoryFilter, ElementId
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.DB.Mechanical import (
    Duct, DuctType as MechDuctType, DuctSystemType, MechanicalSystemType, MechanicalUtils
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB import SpatialElement
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, DB, UI

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# ---------------- logging ----------------
def _info(m): print("[INFO] " + m)
def _warn(m):
    print("[WARN] " + m)
    try: TaskDialog.Show("MEP", m)
    except: pass
def _err(m):
    print("[ERROR] " + m)
    try: TaskDialog.Show("MEP — Error", m)
    except: pass

def feet(v):   return float(v)
def inches(v): return float(v) / 12.0

def unit_vec(v):
    L = math.sqrt(v.X*v.X + v.Y*v.Y + v.Z*v.Z)
    return XYZ(0,0,0) if L == 0 else XYZ(v.X/L, v.Y/L, v.Z/L)

# ---------------- rules I/O ----------------
def _this_folder():
    try:
        return os.path.dirname(__file__)
    except:
        return os.getcwd()

def _read_json(path):
    with open(path, "rb") as f:
        return json.loads(f.read().decode("utf-8-sig"))

def _possible_rules_roots():
    """Search upward from the pushbutton for a 'rules\\mechanical' folder."""
    pb = _this_folder()                                  # ...\Place DuctWork.pushbutton
    panel = os.path.dirname(pb)                          # ...\RP_Panel.panel
    tab   = os.path.dirname(panel)                       # ...\RP_Tab.tab
    ext   = os.path.dirname(tab)                         # ...\RP_pyTools.extension
    # Optional override via env var (handy for dev)
    env_root = os.environ.get("RP_RULES_ROOT", "").strip()
    roots = []
    if env_root:
        roots.append(env_root)
    # 1) pushbutton\rules
    roots.append(os.path.join(pb, "rules"))
    # 2) panel\rules        <-- your current layout
    roots.append(os.path.join(panel, "rules"))
    # 3) tab\rules
    roots.append(os.path.join(tab, "rules"))
    # 4) extension\rules
    roots.append(os.path.join(ext, "rules"))
    # de-dup while preserving order
    seen, unique = set(), []
    for r in roots:
        ru = os.path.abspath(r)
        if ru not in seen:
            seen.add(ru)
            unique.append(ru)
    return unique

def _find_rules_base():
    """Return the first base that contains 'mechanical\\ductwork.json' and 'mechanical\\Rooftop_fans.json'."""
    tried = []
    for base in _possible_rules_roots():
        mech_dir = os.path.join(base, "mechanical")
        duct = os.path.join(mech_dir, "ductwork.json")
        fan  = os.path.join(mech_dir, "Rooftop_fans.json")
        tried.append((duct, fan))
        if os.path.isfile(duct) and os.path.isfile(fan):
            return mech_dir, tried
    return None, tried

def load_dual_rules():
    mech_dir, tried = _find_rules_base()
    if not mech_dir:
        # Build a helpful error showing everywhere we looked
        lines = ["Missing rules: searched for both files in these locations:"]
        for (duct, fan) in tried:
            lines.append("  - {}".format(duct))
            lines.append("    {}".format(fan))
        raise IOError("\n".join(lines))
    duct_rules = _read_json(os.path.join(mech_dir, "ductwork.json"))
    fan_rules  = _read_json(os.path.join(mech_dir, "Rooftop_fans.json"))
    print("[INFO] Loaded rules from: {}".format(mech_dir))
    return duct_rules, fan_rules

def _merge(a, b):
    """Shallow merge dict b onto a (does not deep-merge nested dicts)."""
    out = dict(a or {})
    for k, v in (b or {}).items(): out[k] = v
    return out

# ---------------- space picking / classification ----------------
class RoomOrSpaceFilter(ISelectionFilter):
    def AllowElement(self, e):
        cat = e.Category
        if not cat: return False
        bid = cat.Id.IntegerValue
        return bid in (
            int(BuiltInCategory.OST_Rooms),
            int(BuiltInCategory.OST_MEPSpaces)
        )
    def AllowReference(self, r, p): return False

def pick_room_or_space():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, RoomOrSpaceFilter(), "Pick a Room or Space")
        return doc.GetElement(ref.ElementId)
    except OperationCanceledException:
        return None

def pick_point():
    try:
        _info("Pick a point (fallback if no Room/Space picked).")
        return uidoc.Selection.PickPoint("Pick a point for riser location")
    except OperationCanceledException:
        return None

def get_space_name(el):
    # Works for Room and MEP Space
    try:
        name = el.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        if name: return name
    except: pass
    try:
        return el.Name
    except: return None

def classify_space_key(el_or_name, rules_dict):
    """Return the best key from rules_dict (e.g., 'Restrooms', 'T shape space', else 'General')."""
    keys = set((rules_dict or {}).keys())
    if not keys: return "General"
    name = None
    if isinstance(el_or_name, basestring):
        name = el_or_name
    else:
        name = get_space_name(el_or_name) or ""
    name_low = name.strip().lower()
    # Exact key match prefered
    for k in keys:
        if name_low == k.strip().lower():
            return k
    # Contains match (e.g., name has 'Restroom 102')
    for k in keys:
        if k.strip().lower() in name_low and k.strip().lower() != "general":
            return k
    return "General"

# ---------------- revit queries ----------------
def get_levels(document):
    return list(FilteredElementCollector(document).OfClass(Level))

def get_highest_level(document):
    lvls = get_levels(document)
    if not lvls: return None
    lvls.sort(key=lambda L: L.Elevation)
    return lvls[-1]

def find_level_by_name_contains(document, substr):
    sub = (substr or "").strip().lower()
    if not sub: return None
    for L in get_levels(document):
        if sub in (L.Name or "").lower():
            return L
    return None

def find_duct_type_by_hints(document, hints):
    """Try to find a DuctType whose name contains any hint (case-insensitive)."""
    if not hints:
        return None
    types = list(FilteredElementCollector(document).OfClass(DuctType))
    for h in hints:
        key = (h or "").strip().lower()
        if not key:
            continue
        for dt in types:
            nm = (dt.Name or "").lower()
            if key in nm:
                return dt
    return None

def is_rectangular_duct_type(dt):
    """
    Heuristic: rectangular duct types typically have editable Width & Height params,
    and round types typically have Diameter. This works well across versions.
    """
    try:
        w = dt.LookupParameter("Width")
        h = dt.LookupParameter("Height")
        d = dt.LookupParameter("Diameter")
        return (w is not None and h is not None) and (d is None)
    except:
        return False

def find_rectangular_duct_type(document):
    types = list(FilteredElementCollector(document).OfClass(DuctType))
    # Prefer param-based detection
    for dt in types:
        if is_rectangular_duct_type(dt):
            return dt
    # Fallback: name contains “rect”
    for dt in types:
        if "rect" in (dt.Name or "").lower():
            return dt
    return None

def get_any_duct_type(document):
    return FilteredElementCollector(document).OfClass(DuctType).FirstElement()

def ensure_duct_type(document, hints=None):
    """Pick the best duct type we can: hints → rectangular → any (with warning)."""
    dt = find_duct_type_by_hints(document, hints or [])
    if dt:
        _info("Duct type via hints: {}".format(dt.Name))
        return dt

    dt = find_rectangular_duct_type(document)
    if dt:
        _info("Duct type (rectangular): {}".format(dt.Name))
        return dt

    dt = get_any_duct_type(document)
    if dt:
        _warn("No rectangular duct type found; using '{}'".format(dt.Name))
        return dt

    return None

def get_mech_system_type(document, desired_enum):
    coll = FilteredElementCollector(document).OfClass(MechanicalSystemType)
    for st in coll:
        try:
            if st.SystemType == desired_enum:
                return st
        except: pass
    return None

def parse_system_enum(name):
    n = (name or "").strip().lower()
    if   n == "exhaustair": return DuctSystemType.ExhaustAir
    if   n == "returnair":  return DuctSystemType.ReturnAir
    return DuctSystemType.SupplyAir

# ---------------- family loading / fan picking ----------------
class AlwaysLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True
    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

def try_load_family(path):
    if not path or not os.path.exists(path):
        return None
    fam_ref = clr.Reference[Family]()
    opts = AlwaysLoadOptions()
    if doc.LoadFamily(path, opts, fam_ref):
        _info("Loaded family: {}".format(os.path.basename(path)))
        return fam_ref.Value
    return None

def find_symbol_by_family_type(family_name, type_name):
    fam_low = (family_name or "").strip().lower()
    typ_low = (type_name or "").strip().lower()
    coll = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for s in coll:
        if not s.Family: continue
        if (s.Family.Name or "").strip().lower() == fam_low and (s.Name or "").strip().lower() == typ_low:
            return s
    return None

def ensure_symbol(family_name, type_name, load_from_path=None):
    sym = find_symbol_by_family_type(family_name, type_name)
    if sym: return sym
    # try to load and search again
    fam = try_load_family(load_from_path)
    if fam:
        # after loading, symbol should exist
        sym = find_symbol_by_family_type(family_name, type_name)
        if sym: return sym
        # If type not found, pick any symbol under that family
        for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
            if s.Family and (s.Family.Name or "").strip().lower() == (family_name or "").strip().lower():
                return s
    return None

def find_fan_by_hints(hints):
    hints_up = [h.upper() for h in (hints or [])]
    me_filter = ElementCategoryFilter(BuiltInCategory.OST_MechanicalEquipment)
    symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).WherePasses(me_filter))
    cands = []
    for s in symbols:
        label = (s.Family.Name + " : " + s.Name).upper()
        if any(h in label for h in hints_up):
            cands.append(s)
    if not cands: return None
    def score(sym):
        label = (sym.Family.Name + " : " + sym.Name).upper()
        sc = 0
        for h in hints_up:
            if h in label: sc += 1
        return -sc
    cands.sort(key=score)
    return cands[0]

def ensure_active(sym):
    if not sym.IsActive:
        sym.Activate()
        doc.Regenerate()

def place_family(symbol, level, xyz):
    ensure_active(symbol)
    return doc.Create.NewFamilyInstance(xyz, symbol, level, StructuralType.NonStructural)

# ---------------- duct helpers ----------------
def set_rect_size(duct, w_ft, h_ft):
    try:
        w = duct.LookupParameter("Width")
        h = duct.LookupParameter("Height")
        if w and h and not w.IsReadOnly and not h.IsReadOnly:
            w.Set(feet(w_ft)); h.Set(feet(h_ft))
    except: pass

# ---------------- main ----------------
def main():
    # load rules
    try:
        duct_rules, fan_rules = load_dual_rules()
    except Exception as e:
        _err("Could not load rules: {}".format(e)); return

    duct_defaults = duct_rules.get("defaults", {}) or {}
    duct_map = duct_rules.get("ductwork_rules_by_space", {}) or {}
    fan_defaults = fan_rules.get("defaults", {}) or {}
    fan_map = fan_rules.get("fan_rules_by_space", {}) or {}

    # pick a Room/Space for categorization (preferred)
    pick = pick_room_or_space()
    picked_point = None
    space_key = "General"

    if pick:
        nm = get_space_name(pick) or ""
        space_key = classify_space_key(nm, duct_map)  # use duct map keys to classify
        _info("Picked: '{}' → category key: '{}'".format(nm, space_key))
    else:
        # fallback: let user pick a point; we won't have a space name, so 'General'
        picked_point = pick_point()
        if not picked_point: _warn("Canceled."); return
        _info("No Room/Space selected; using 'General' rules.")

    # resolve duct profile for this space
    duct_space = _merge(duct_defaults, duct_map.get(space_key, {}))
    system_name = duct_space.get("system", "SupplyAir")
    sys_enum = parse_system_enum(system_name)
    sys_type = get_mech_system_type(doc, sys_enum)
    if not sys_type:
        _err("No MechanicalSystemType for '{}'".format(system_name)); return

    duct_type_hints = (duct_space.get("duct_type_hints")
                       or duct_rules.get("defaults", {}).get("duct_type_hints")
                       or [])
    duct_type = ensure_duct_type(doc, duct_type_hints)
    if not duct_type:
        _err("No Duct Types in this model. Add a duct type (preferably Rectangular) to your template.")
        return
    _info("Using duct type: {}".format(duct_type.Name))

    # roof level
    roof_level = None
    if duct_space.get("place_on_highest_level", True):
        roof_level = get_highest_level(doc)
    if not roof_level:
        roof_level = find_level_by_name_contains(doc, duct_space.get("roof_level_name_contains","ROOF"))
    if not roof_level:
        _err("Could not resolve a roof level."); return
    _info("Roof level: {} (elev={:.2f} ft)".format(roof_level.Name, roof_level.Elevation))

    # sizes/heights/routing
    sizes   = duct_space.get("sizes", {})
    riserW_in, riserH_in = sizes.get("riser_in", [12,12])
    mainW_in,  mainH_in  = sizes.get("main_in",  [12,9])
    riserW_ft, riserH_ft = inches(riserW_in), inches(riserH_in)
    mainW_ft,  mainH_ft  = inches(mainW_in),  inches(mainH_in)

    heights = duct_space.get("heights", {})
    riser_base_z_ft = float(heights.get("riser_base_z_ft", 9.0))
    roof_extra_ft   = float(heights.get("roof_extra_ft",   3.0))

    routing = duct_space.get("routing", {})
    orientation      = (routing.get("orientation","WorldX") or "WorldX").upper()
    horiz_run_len_ft = float(routing.get("horiz_run_len_ft", 3.0))
    fan_offset_ft    = float(routing.get("fan_offset_ft",    3.0))

    # resolve fan profile for this space (and merge defaults)
    fan_space = _merge(fan_defaults, fan_map.get(space_key, {}))
    device_candidates = fan_space.get("device_candidates", []) or []
    name_hints        = fan_space.get("fan_name_hints", ["FAN", "ROOF"])
    if "fan_offset_ft" in fan_space:
        try: fan_offset_ft = float(fan_space.get("fan_offset_ft"))
        except: pass

    # choose fan symbol: try candidates in order, then hints
    fan_sym = None
    for c in device_candidates:
        fam = c.get("family"); typ = c.get("type"); pth = c.get("load_from")
        if not fam: continue
        fan_sym = ensure_symbol(fam, typ, pth)
        if fan_sym:
            _info("Fan via candidate: {} : {}".format(fam, fan_sym.Name)); break
    if not fan_sym:
        fan_sym = find_fan_by_hints(name_hints)
        if fan_sym: _info("Fan via hints: {} : {}".format(fan_sym.Family.Name, fan_sym.Name))
    if not fan_sym:
        _err("No suitable fan symbol found. Update Rooftop_fans.json or load a family."); return

    # determine base point (riser location)
    if picked_point is None:
        # use the room/space location (rough heuristic: its location point or bbox center)
        try:
            loc = pick.Location
            if loc and hasattr(loc, "Point"): picked_point = loc.Point
        except: picked_point = None
        if picked_point is None:
            try:
                bb = pick.get_BoundingBox(None)
                picked_point = XYZ((bb.Min.X+bb.Max.X)/2.0, (bb.Min.Y+bb.Max.Y)/2.0, (bb.Min.Z+bb.Max.Z)/2.0)
            except:
                picked_point = pick_point()
                if not picked_point: _warn("Canceled."); return

    riser_base = XYZ(picked_point.X, picked_point.Y, picked_point.Z + riser_base_z_ft)
    riser_top  = XYZ(picked_point.X, picked_point.Y, roof_level.Elevation + roof_extra_ft)

    horiz_dir = unit_vec(XYZ(1,0,0) if orientation=="WORLDX" else XYZ(0,1,0))
    fan_point = riser_top + horiz_dir.Multiply(fan_offset_ft)

    # build geometry
    t = Transaction(doc, "RP: Mech Rough-In (space rules)")
    t.Start()
    try:
        # place fan
        ensure_active(fan_sym)
        fan_inst = doc.Create.NewFamilyInstance(fan_point, fan_sym, roof_level, StructuralType.NonStructural)
        _info("Placed fan '{}' on '{}'".format(fan_sym.Family.Name + " : " + fan_sym.Name, roof_level.Name))

        # riser
        vertical_duct = Duct.Create(doc, sys_type.Id, duct_type.Id, roof_level.Id, riser_base, riser_top)
        set_rect_size(vertical_duct, riserW_ft, riserH_ft)
        _info("Riser {}x{} in".format(int(riserW_in), int(riserH_in)))

        # short horizontal
        horiz_end = riser_top + horiz_dir.Multiply(horiz_run_len_ft)
        horiz_duct = Duct.Create(doc, sys_type.Id, duct_type.Id, roof_level.Id, riser_top, horiz_end)
        set_rect_size(horiz_duct, mainW_ft, mainH_ft)
        _info("Horizontal ~{:.2f} ft, {}x{} in".format(horiz_run_len_ft, int(mainW_in), int(mainH_in)))

        doc.Regenerate()
        t.Commit()
        _info("Done.")
    except Exception as e:
        try: t.RollBack()
        except: pass
        _err("Failed: {}".format(e))

if __name__ == "__main__":
    main()