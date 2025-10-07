# -*- coding: utf-8 -*-
# Circuit Spaces — nearest-panel routing per Space (lights→nearest LP, recepts→nearest RP)
# Single undo/redo via TransactionGroup.

import os, sys, clr, math

# ---------------- Path resolver (find rules/ and lib/) ----------------
SCRIPT_DIR = os.path.dirname(__file__)

def _find_dir(base_name):
    cand, anc = [], SCRIPT_DIR
    for _ in range(5):
        cand.append(os.path.join(anc, base_name))
        anc = os.path.dirname(anc)
    for root in [SCRIPT_DIR, os.path.dirname(SCRIPT_DIR), os.path.dirname(os.path.dirname(SCRIPT_DIR))]:
        try:
            for n in os.listdir(root):
                if n.endswith(".pushbutton"):
                    cand.append(os.path.join(root, n, base_name))
        except:
            pass
    for p in cand:
        if os.path.isdir(p):
            return p
    return None

RULES_DIR = _find_dir("rules")
if not RULES_DIR:
    raise EnvironmentError("[FATAL] Couldn't find a 'rules' folder near: {0}".format(SCRIPT_DIR))

IDENTIFY_SPACES = os.path.join(RULES_DIR, "identify_spaces.json")
LIGHTING        = os.path.join(RULES_DIR, "lighting_rules.json")   # used for category map/defaults only
ELEC_DIR        = os.path.join(RULES_DIR, "electrical")

LIB_DIR = _find_dir("lib")
if not LIB_DIR:
    raise EnvironmentError("[FATAL] Couldn't find a 'lib' folder near: {0}".format(SCRIPT_DIR))
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

print("[PATH] rules -> {0}".format(RULES_DIR))
print("[PATH] lib   -> {0}".format(LIB_DIR))

# ---------------- Revit bind ----------------
try:
    from pyrevit import revit
    uiapp = revit.uiapp; uidoc = revit.uidoc; doc = revit.doc
    app = uiapp.Application if uiapp else None
except Exception:
    uiapp = uidoc = doc = app = None

if app is None:
    try:
        uiapp = __revit__; app = uiapp.Application
        uidoc = uiapp.ActiveUIDocument; doc = uidoc.Document if uidoc else None
    except Exception:
        pass
if app is None or doc is None:
    raise EnvironmentError("Open a project in Revit and run from a pyRevit button.")

active_view = doc.ActiveView

# ---------------- Imports from lib ----------------
from rules_loader import build_rule_for_room
from electrical_loader import load_electrical_rules
from place_by_space import get_target_spaces, spatial_display_name

# Single-undo wrapper (from lib); fallback if not available
try:
    from single_undo import run_as_single_undo
except Exception:
    clr.AddReference('RevitAPI')
    from Autodesk.Revit.DB import TransactionGroup
    def run_as_single_undo(document, title, work_fn, *args, **kwargs):
        tg = TransactionGroup(document, title)
        tg.Start()
        try:
            result = work_fn(*args, **kwargs)
            tg.Assimilate()
            return result
        except:
            try: tg.RollBack()
            except: pass
            raise

# ---------------- Revit API ----------------
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilyInstance, Options, Transaction,
    BuiltInParameter, XYZ, ElementId, ElementSet, SpatialElementBoundaryOptions
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType
from Autodesk.Revit.DB.Structure import StructuralType
from System.Collections.Generic import List

# ---------------- Settings ----------------
VERBOSE  = True
DRY_RUN  = False           # True = report only; no circuits created
DEFAULT_VOLT = 120.0       # for VA packing when we can't read a better value

# Panel routing behavior
USE_NEAREST_FOR_LIGHTS   = True
USE_NEAREST_FOR_RECEPTS  = True     # set False to keep recepts on a fixed panel if you prefer
FALLBACK_FIXED_LIGHTS    = "x"   # used if no matching panels exist
FALLBACK_FIXED_RECEPTS   = "x"

# ---------------- Logging ----------------
def log(msg):
    if VERBOSE:
        print(msg)

# ---------------- Panel discovery & selection ----------------
def _all_panels(document):
    return list(FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .WhereElementIsNotElementType())

def _panel_name(panel_inst):
    try:
        p = panel_inst.LookupParameter("Panel Name")
        return (p.AsString() or p.AsValueString() or "").strip()
    except:
        return (getattr(panel_inst, "Name", None) or "").strip()

def _panel_role_from_name(name):
    s = (name or "").strip().lower()
    # very simple heuristics: tweak to your naming
    if s.startswith("lp") or " lighting" in s or " light" in s:
        return "lighting"
    if s.startswith("rp") or " recept" in s or " receptacle" in s or s.startswith("pp"):
        return "receptacles"
    return "other"

def _center_point(elem):
    try:
        loc = elem.Location
        if hasattr(loc, "Point") and loc.Point:
            return loc.Point
    except:
        pass
    try:
        gb = elem.get_Geometry(Options())
        if gb:
            for g in gb:
                bb = getattr(g, "GetBoundingBox", None)
                if bb:
                    b = g.GetBoundingBox()
                    if b:
                        return (b.Min + b.Max) * 0.5
    except:
        pass
    return None

def _space_centroid(space):
    # 1) try explicit location
    try:
        loc = space.Location
        if hasattr(loc, "Point") and loc.Point:
            return loc.Point
    except:
        pass
    # 2) compute from boundary loops (planar average)
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
        pts = []
        for loop in loops:
            for seg in loop:
                c = seg.GetCurve()
                pts.append(c.GetEndPoint(0))
        if pts:
            sx = sum(p.X for p in pts); sy = sum(p.Y for p in pts); sz = sum(p.Z for p in pts)
            n = float(len(pts))
            return XYZ(sx/n, sy/n, sz/n)
    except:
        pass
    # 3) fallback: origin
    return XYZ(0, 0, 0)

def _xy_dist(a, b):
    dx = (a.X - b.X); dy = (a.Y - b.Y)
    return math.sqrt(dx*dx + dy*dy)

def _index_panels_by_role(document):
    roles = {"lighting": [], "receptacles": [], "other": []}
    for p in _all_panels(document):
        nm = _panel_name(p)
        roles[_panel_role_from_name(nm)].append(p)
    return roles

def _nearest_panel(document, space, role, roles_index=None):
    if roles_index is None:
        roles_index = _index_panels_by_role(document)
    candidates = list(roles_index.get(role, []))
    if not candidates:
        # fallback to any panel if no role-matched panel exists
        candidates = [p for plist in roles_index.values() for p in plist]
    if not candidates:
        return None
    sc = _space_centroid(space)
    best = None
    bestd = 1e30
    for p in candidates:
        pc = _center_point(p) or XYZ(0,0,0)
        d = _xy_dist(sc, pc)
        if d < bestd:
            bestd = d; best = p
    return best

# ---------------- Circuit helpers ----------------
def _add_to_circuit_compat(es, dev):
    try:
        es.AddToCircuit(dev); return True
    except:
        pass
    try:
        es.AddToCircuit(dev.Id); return True
    except:
        pass
    try:
        eset = ElementSet(); eset.Insert(dev); es.AddToCircuit(eset); return True
    except Exception as ex:
        print(u"[WARN] Skipped (can't add): {0} -> {1}".format(_ft_name_lower(dev), ex))
        return False

def _param_as_double(elem, pname):
    try:
        p = elem.LookupParameter(pname)
        if p:
            v = p.AsDouble()
            if v and v > 0:
                return float(v)
    except:
        pass
    return None

def _looks_like_receptacle(inst):
    n = _ft_name_lower(inst)
    return ("recept" in n) or ("duplex" in n) or ("outlet" in n)

def _ft_name_lower(inst):
    try:
        fam = inst.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
        typ = inst.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        fam_s = (fam.AsString() or fam.AsValueString() or "") if fam else ""
        typ_s = (typ.AsString() or typ.AsValueString() or "") if typ else ""
        return (u"%s %s" % (fam_s, typ_s)).strip().lower()
    except:
        sym = getattr(inst, "Symbol", None)
        famname = ""; typname = ""
        if sym:
            try: typname = getattr(sym, "Name", "") or ""
            except: pass
            try:
                fam = getattr(sym, "Family", None)
                famname = getattr(fam, "Name", "") or "" if fam else ""
            except: pass
        return (u"%s %s" % (famname, typname)).strip().lower()

def _assign_panel(document, es, panel_inst):
    assigned = False; last_err = None
    try:
        es.SelectPanel(panel_inst); assigned = True
    except Exception as ex:
        last_err = ex
        if hasattr(es, "SetPanelId"):
            try:
                es.SetPanelId(panel_inst.Id); assigned = True
            except Exception as ex2:
                last_err = ex2
    try:
        document.Regenerate()
    except:
        pass
    if not assigned:
        try:
            pname = (panel_inst.LookupParameter("Panel Name").AsString()
                     or panel_inst.LookupParameter("Panel Name").AsValueString() or "")
        except:
            pname = "<panel>"
        print("[WARN] Couldn't assign panel '{}': {}".format(pname, last_err))
    return assigned

def label_circuit(es, space_name, kind):
    txt = u"{} - {}".format(space_name, "Lights" if kind == "ltg" else "Recepts")
    for pname in ("Circuit Description", "Load Name"):
        try:
            p = es.LookupParameter(pname)
            if p and p.StorageType.ToString() == "String":
                p.Set(txt)
        except:
            pass

def get_light_va(inst, fallback_va):
    for target in (inst, getattr(inst, "Symbol", None)):
        if not target: continue
        for pname in ("Apparent Load", "VA", "Load", "Wattage", "Watts", "Connected Load"):
            v = _param_as_double(target, pname)
            if v and v > 0: return float(v)
    return float(fallback_va)

def group_by_va(pairs, max_circuit_va):
    groups, current, tally = [], [], 0.0
    for d, dva in pairs:
        if current and (tally + dva > max_circuit_va + 1e-6):
            groups.append(list(current)); current = []; tally = 0.0
        current.append(d); tally += float(dva or 0.0)
    if current: groups.append(current)
    return groups

def devices_in_space(document, space, bic):
    col = FilteredElementCollector(document).OfCategory(bic).WhereElementIsNotElementType()
    return [inst for inst in col if _is_in_space(inst, space)]

def _is_in_space(inst, space):
    pt = _center_point(inst)
    if not pt: return False
    zt = _space_test_z(space)
    try:
        return space.IsPointInSpace(XYZ(pt.X, pt.Y, zt))
    except:
        return False

def _space_test_z(space):
    try:
        lvl = doc.GetElement(space.LevelId); base = (lvl.Elevation if lvl else 0.0)
    except:
        base = 0.0
    return base + 3.0

def set_bool_param_if_exists(inst, name, value_bool):
    try:
        p = inst.LookupParameter(name)
        if p and p.StorageType.ToString() == "Integer":
            p.Set(1 if value_bool else 0); return True
    except:
        pass
    return False

# ---------------- Rules helpers ----------------
def get_recept_rules(elec_rules, category):
    bc = elec_rules.get("branch_circuits", {}) if elec_rules else {}
    bycat = bc.get("receptacle_rules_by_category", {})
    r = bycat.get(category, {})
    gen = {
        "unit_va": bc.get("general", {}).get("receptacle_unit_load_va", 180),
        "default_circuit_a": bc.get("general", {}).get("default_circuit_ampacity_a", 20)
    }
    return r, gen

def get_device_protection(elec_rules, category):
    dp = elec_rules.get("device_protection", {}).get("by_category", {}) if elec_rules else {}
    return dp.get(category, {})

# ---------------- Create circuit ----------------
def create_circuit_for_group(document, devices, panel_inst=None):
    devs = [d for d in devices if getattr(d, "MEPModel", None) is not None]
    if not devs:
        print("[WARN] Group had no connectable devices.")
        return None
    seed_ids = List[ElementId]()
    seed_ids.Add(devs[0].Id)
    es = ElectricalSystem.Create(document, seed_ids, ElectricalSystemType.PowerCircuit)
    if es is None:
        print("[ERROR] Could not seed circuit.")
        return None
    added = 0
    for d in devs[1:]:
        if _add_to_circuit_compat(es, d): added += 1
    if panel_inst is not None:
        _assign_panel(document, es, panel_inst)
    try:
        es.BalanceLoads = True
    except:
        pass
    try:
        print("[OK] Circuit created | devices={} (seed+{})".format(es.Elements.Size, added))
    except:
        print("[OK] Circuit created")
    return es

# ---------------- Work function (wrapped as single undo) ----------------
def do_all():
    elec = load_electrical_rules(ELEC_DIR)

    spaces = get_target_spaces(doc, uidoc, active_view, only_current_level=False, prefer_selection=True)
    if not spaces:
        print("[ERROR] No MEP Space elements found."); return

    # Pre-index panels by role once
    roles_index = _index_panels_by_role(doc)

    total_made = 0
    if len(spaces) == 1:
        nm1, num1 = spatial_display_name(spaces[0]); print("[INFO] One space found: {} {}".format(nm1, num1))
    else:
        print("[INFO] {} spaces found. Processing…".format(len(spaces)))

    for sp in spaces:
        nm, num = spatial_display_name(sp)
        disp_name = (nm + (" " + num if num else "")).strip()

        # classification (reuses your existing mapping)
        category, _lighting_rule = build_rule_for_room(nm or "", IDENTIFY_SPACES, LIGHTING)

        # nearest panels for this space
        panel_ltg = None
        panel_rec = None
        if USE_NEAREST_FOR_LIGHTS:
            panel_ltg = _nearest_panel(doc, sp, role="lighting", roles_index=roles_index)
        if USE_NEAREST_FOR_RECEPTS:
            panel_rec = _nearest_panel(doc, sp, role="receptacles", roles_index=roles_index)

        # fallbacks if role-matched not found
        if panel_ltg is None and FALLBACK_FIXED_LIGHTS:
            from_name = FALLBACK_FIXED_LIGHTS.strip()
            panel_ltg = next((p for p in _all_panels(doc) if _panel_name(p).lower() == from_name.lower()), None)
        if panel_rec is None and FALLBACK_FIXED_RECEPTS:
            from_name = FALLBACK_FIXED_RECEPTS.strip()
            panel_rec = next((p for p in _all_panels(doc) if _panel_name(p).lower() == from_name.lower()), None)

        # ---------- RECEPTACLES ----------
        rr, gen = get_recept_rules(elec, category)
        prot = get_device_protection(elec, category)

        rec_all = devices_in_space(doc, sp, BuiltInCategory.OST_ElectricalFixtures)
        rec_unc = [r for r in rec_all if _looks_like_receptacle(r) and not _device_has_circuit(r)]

        unit_va_rec = float(gen.get("unit_va", 180.0))
        max_va_rec  = float(gen.get("default_circuit_a", 20.0)) * DEFAULT_VOLT
        rec_pairs   = [(d, unit_va_rec) for d in rec_unc]
        rec_groups  = group_by_va(rec_pairs, max_va_rec)
        log("[CIRCUIT] {} -> Recepts: {} uncirc -> {} group(s)".format(disp_name, len(rec_unc), len(rec_groups)))

        # ---------- LIGHTS ----------
        ltg_all = devices_in_space(doc, sp, BuiltInCategory.OST_LightingFixtures)
        ltg_unc = [l for l in ltg_all if not _device_has_circuit(l)]

        bc_general     = elec.get("branch_circuits", {}).get("general", {}) if elec else {}
        default_ltg_va = float(bc_general.get("lighting_default_va_per_fixture", 100.0))
        default_ltg_a  = float(bc_general.get("lighting_default_circuit_ampacity_a",
                                  bc_general.get("default_circuit_ampacity_a", 20.0)))
        max_va_ltg     = default_ltg_a * DEFAULT_VOLT

        ltg_pairs  = [(l, get_light_va(l, default_ltg_va)) for l in ltg_unc]
        ltg_pairs.sort(key=lambda x: -x[1])
        ltg_groups = group_by_va(ltg_pairs, max_va_ltg)
        log("[CIRCUIT] {} -> Lights: {} uncirc -> {} group(s)".format(disp_name, len(ltg_unc), len(ltg_groups)))

        if DRY_RUN:
            continue

        made_here = 0
        t = Transaction(doc, "Create Circuits (Space {0})".format(num or nm)); t.Start()
        try:
            # Receptacle circuits
            for grp in rec_groups:
                es = create_circuit_for_group(doc, grp, panel_rec)
                if es:
                    made_here += 1
                    label_circuit(es, disp_name, "rec")
                if prot:
                    for d in grp:
                        if "gfci" in prot: set_bool_param_if_exists(d, "GFCI", bool(prot["gfci"]))
                        if "afci" in prot: set_bool_param_if_exists(d, "AFCI", bool(prot["afci"]))
                        if "tamper_resistant" in prot:
                            set_bool_param_if_exists(d, "Tamper Resistant", bool(prot["tamper_resistant"]))

            # Lighting circuits
            for grp in ltg_groups:
                es = create_circuit_for_group(doc, grp, panel_ltg)
                if es:
                    made_here += 1
                    label_circuit(es, disp_name, "ltg")

            t.Commit()
        except Exception as ex:
            print("[ERROR] Circuit creation failed:", ex)
            try: t.RollBack()
            except: pass

        total_made += made_here
        pn_l = _panel_name(panel_ltg) if panel_ltg else "(none)"
        pn_r = _panel_name(panel_rec) if panel_rec else "(none)"
        print("[RESULT] {} -> created {} circuit(s) to [{} lights] [{} recepts]".format(
            disp_name, made_here, pn_l, pn_r
        ))

    if DRY_RUN:
        print("[RESULT] DRY-RUN complete.")
    else:
        print("[RESULT] Total new circuits created:", total_made)

# small helper used above
def _device_has_circuit(inst):
    try:
        mep = inst.MEPModel
        if mep is not None:
            for es in mep.ElectricalSystems:
                try:
                    if es.SystemType == ElectricalSystemType.PowerCircuit:
                        return True
                except:
                    continue
    except:
        pass
    return False

# ---------------- Single Undo wrapper call ----------------
if __name__ == "__main__":
    run_as_single_undo(doc, "Circuit Spaces (Nearest Panels)", do_all)


