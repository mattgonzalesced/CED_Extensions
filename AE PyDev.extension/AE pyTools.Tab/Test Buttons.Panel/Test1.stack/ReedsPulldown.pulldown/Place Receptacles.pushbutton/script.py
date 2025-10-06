# -*- coding: utf-8 -*-
# Receptacles AutoLayout — places receptacles along room walls using rules/electrical

import os, sys, clr, math
# ---------- Bind to Revit ----------
uiapp = uidoc = doc = app = active_view = None
try:
    from pyrevit import revit
    uiapp = revit.uiapp; uidoc = revit.uidoc; doc = revit.doc
    app = uiapp.Application if uiapp else None
except Exception:
    pass
if app is None:
    try:
        uiapp = __revit__; app = uiapp.Application
        uidoc = uiapp.ActiveUIDocument; doc = uidoc.Document if uidoc else None
    except Exception:
        pass
if app is None or doc is None:
    raise EnvironmentError("Open a project in Revit and run from a pyRevit button.")
active_view = doc.ActiveView

# ---------- Revit API ----------
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilySymbol, FamilyPlacementType,
    Transaction, SpatialElementBoundaryOptions, XYZ, Wall, HostObjectUtils, ShellLayerType,
    FamilyInstance, Curve, BuiltInParameter, Transform, RevitLinkInstance, Opening
)
from Autodesk.Revit.DB.Structure import StructuralType

# ---------- Paths / imports ----------
SCRIPT_DIR = os.path.dirname(__file__)

def _find_dir(base_name):
    # 1) walk up to 4 ancestors and check each for base_name
    anc = SCRIPT_DIR
    candidates = []
    for _ in range(5):  # script -> parent -> grandparent -> great-grandparent -> great-great
        candidates.append(os.path.join(anc, base_name))
        anc = os.path.dirname(anc)

    # 2) for the script dir and its first two parents, also look inside sibling *.pushbutton folders
    for root in [SCRIPT_DIR,
                 os.path.dirname(SCRIPT_DIR),
                 os.path.dirname(os.path.dirname(SCRIPT_DIR))]:
        try:
            for name in os.listdir(root):
                if name.endswith(".pushbutton"):
                    candidates.append(os.path.join(root, name, base_name))
        except Exception:
            pass

    for p in candidates:
        if os.path.isdir(p):
            return p
    return None

# Resolve rules/
RULES_DIR = _find_dir("rules")
if not RULES_DIR:
    raise EnvironmentError("[FATAL] Couldn't find a 'rules' folder near: {0}".format(SCRIPT_DIR))
IDENTIFY = os.path.join(RULES_DIR, "identify_spaces.json")
LIGHTING = os.path.join(RULES_DIR, "lighting_rules.json")
ELEC_DIR = os.path.join(RULES_DIR, "electrical")

# Resolve lib/
LIB_DIR = _find_dir("lib")
if not LIB_DIR:
    raise EnvironmentError("[FATAL] Couldn't find a 'lib' folder near: {0}".format(SCRIPT_DIR))
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

print("[PATH] rules -> {0}".format(RULES_DIR))
print("[PATH] lib   -> {0}".format(LIB_DIR))

from rules_loader import build_rule_for_room
from electrical_loader import load_electrical_rules
from single_undo import run_as_single_undo

# ---------- Settings ----------
VERBOSE = True
DRY_RUN = False  # set True to test without placing

def log(msg):
    if VERBOSE: print(msg)
#----------- Avoid Corners & Doors 2 ------------
def _curve_length(curve):
    try:
        return float(curve.Length)
    except:
        return 0.0

def _project_len_on_curve(curve, point):
    """
    Return (dist_along_ft, perp_ft). Robust mapping using the curve's real parameter domain.
    """
    try:
        res = curve.Project(point)
        if res is None:
            return (None, None)
        p  = float(res.Parameter)
        ps = float(curve.GetEndParameter(0))
        pe = float(curve.GetEndParameter(1))
        u  = 0.0 if abs(pe-ps) < 1e-12 else (p-ps)/(pe-ps)
        if u < 0.0: u = 0.0
        if u > 1.0: u = 1.0
        L  = _curve_length(curve)
        pt_on = getattr(res, "XYZPoint", None) or getattr(res, "XYZ", None)
        perp  = (point - pt_on).GetLength() if pt_on else None
        return (u * L, perp)
    except:
        # safe line-only fallback
        try:
            p0 = curve.GetEndPoint(0); p1 = curve.GetEndPoint(1)
            v  = p1 - p0; L = v.GetLength()
            if L <= 1e-8:
                return (0.0, (point - p0).GetLength())
            diru = (p1 - p0) / L
            dot  = max(0.0, min(L, (point - p0).X*diru.X + (point - p0).Y*diru.Y + (point - p0).Z*diru.Z))
            pt_on = p0 + diru.Multiply(dot)
            return (dot, (point - pt_on).GetLength())
        except:
            return (None, None)

def _xy_dist(a, b):
    dx = a.X - b.X; dy = a.Y - b.Y
    return (dx*dx + dy*dy) ** 0.5

def _collect_door_and_opening_centers_for_wall(doc, wall, wall_curve, include_linked=True, snap_tol_ft=6.0):
    """
    Return a list of XYZ centers (host coords) for:
      - Doors (host model)
      - Doors (linked models) — transformed into host coords
      - Rectangular 'Opening' elements hosted in/near this wall (host model)
    Only include items whose center is within 'snap_tol_ft' of the wall curve.
    """
    centers = []

    # 1) Host model doors
    try:
        for d in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType():
            loc = getattr(d.Location, "Point", None)
            if loc is None:
                continue
            _, perp = _project_len_on_curve(wall_curve, loc)
            if perp is not None and perp <= snap_tol_ft:
                centers.append(loc)
    except:
        pass

    # 2) Host model rectangular openings (passages)
    try:
        for op in FilteredElementCollector(doc).OfClass(Opening):
            # If the opening has a host wall we can bias toward the same wall; otherwise use proximity
            try:
                host = getattr(op, "Host", None)
            except:
                host = None
            bb = op.get_BoundingBox(None)
            if not bb:
                continue
            center = XYZ((bb.Min.X+bb.Max.X)/2.0, (bb.Min.Y+bb.Max.Y)/2.0, (bb.Min.Z+bb.Max.Z)/2.0)
            # If opening has a host and it's not this wall, still allow if near the wall curve (proximity)
            _, perp = _project_len_on_curve(wall_curve, center)
            if perp is not None and perp <= snap_tol_ft:
                centers.append(center)
    except:
        pass

    # 3) Linked model doors
    if include_linked:
        try:
            links = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
            for link in links:
                ldoc = link.GetLinkDocument()
                if ldoc is None:
                    continue
                T = link.GetTransform() if hasattr(link, "GetTransform") else link.GetTotalTransform()
                for d in FilteredElementCollector(ldoc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType():
                    loc = getattr(d.Location, "Point", None)
                    if loc is None:
                        continue
                    host_pt = T.OfPoint(loc)
                    _, perp = _project_len_on_curve(wall_curve, host_pt)
                    if perp is not None and perp <= snap_tol_ft:
                        centers.append(host_pt)
        except:
            pass

    return centers

def filter_points_avoid_doors_radius(doc, wall, wall_curve, pts,
                                     radius_ft=6.0, snap_tol_ft=6.0,
                                     include_linked=True):
    """
    Remove any candidate point whose XY distance to any door/opening center (near this wall)
    is less than 'radius_ft'.
    """
    if not pts:
        return []

    centers = _collect_door_and_opening_centers_for_wall(
        doc, wall, wall_curve, include_linked=include_linked, snap_tol_ft=snap_tol_ft
    )
    if not centers:
        return pts

    r = float(radius_ft)
    r2 = r * r

    out = []
    for p in pts:
        blocked = False
        for c in centers:
            # XY check (ignore Z so multi-level small offsets don’t matter)
            if _xy_dist(p, c) ** 2 < r2:
                blocked = True
                break
        if not blocked:
            out.append(p)
    return out



#----------- Avoid Corners & Doors 1 ------------
def _curve_len(curve):
    try: return float(curve.Length)
    except: return 0.0

def _unit(v):
    mag = (v.X*v.X + v.Y*v.Y + v.Z*v.Z) ** 0.5
    return XYZ(0,0,0) if mag < 1e-9 else XYZ(v.X/mag, v.Y/mag, v.Z/mag)

def _project_on_curve(curve, point):
    """Return (dist_ft, pt_on_curve, perp_ft) using the curve's real param domain."""
    try:
        res = curve.Project(point)
        if res is None:
            return (None, None, None)
        p  = float(res.Parameter)
        ps = float(curve.GetEndParameter(0))
        pe = float(curve.GetEndParameter(1))
        u  = 0.0 if abs(pe-ps) < 1e-12 else (p-ps)/(pe-ps)
        if u < 0.0: u = 0.0
        if u > 1.0: u = 1.0
        L  = _curve_len(curve)
        pt_on = getattr(res, "XYZPoint", None) or getattr(res, "XYZ", None)
        perp  = (point - pt_on).GetLength() if pt_on else None
        return (u*L, pt_on, perp)
    except:
        # safe line-only fallback
        try:
            p0 = curve.GetEndPoint(0); p1 = curve.GetEndPoint(1)
            v  = p1 - p0; L = v.GetLength()
            if L <= 1e-8: return (0.0, p0, (point - p0).GetLength())
            diru = XYZ(v.X/L, v.Y/L, v.Z/L)
            dot  = max(0.0, min(L, (point - p0).X*diru.X + (point - p0).Y*diru.Y + (point - p0).Z*diru.Z))
            pt_on = XYZ(p0.X + diru.X*dot, p0.Y + diru.Y*dot, p0.Z + diru.Z*dot)
            return (dot, pt_on, (point - pt_on).GetLength())
        except:
            return (None, None, None)

def _door_width_param_ft(door):
    # Built-in first
    try:
        p = door.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM)
        if p and p.HasValue: return float(p.AsDouble())
    except: pass
    # Common name fallbacks
    for nm in ("Width","Rough Width","Rough Opening Width","Nominal Width","Door Width"):
        try:
            p = door.LookupParameter(nm)
            if p and p.HasValue: return float(p.AsDouble())
        except: pass
    return None

def _door_width_bbox_ft(door, wall_dir_u, xform=None):
    # Project bbox along wall direction as last resort
    try:
        bb = door.get_BoundingBox(None)
        if not bb: return None
        corners = [
            XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
            XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
            XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
            XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
        ]
        if xform: corners = [xform.OfPoint(c) for c in corners]
        dots = [c.X*wall_dir_u.X + c.Y*wall_dir_u.Y + c.Z*wall_dir_u.Z for c in corners]
        return max(dots) - min(dots)
    except:
        return None

def _build_door_spans_for_wall(doc, wall, curve, avoid_doors_ft,
                               edge_margin_in=1.0, proximity_tol_ft=6.0,
                               include_linked=True):
    """
    Returns merged [(a_ft,b_ft)] spans along the wall curve to avoid.
    Uses door center + HandOrientation to compute true jamb endpoints,
    then projects those endpoints to the wall curve.
    """
    spans = []
    if avoid_doors_ft <= 0.0 or curve is None:
        return spans

    # Wall direction (unit) from curve endpoints (works for arcs via tangent chord; ok for short spans)
    try:
        w0 = curve.GetEndPoint(0); w1 = curve.GetEndPoint(1)
        wall_dir_u = _unit(w1 - w0)
    except:
        wall_dir_u = XYZ(1,0,0)

    edge_margin_ft = float(edge_margin_in) / 12.0

    def _make_span(center_pt, hand_vec_host, width_ft):
        """Compute span by projecting endpoints (center ± wall_dir * half_total)."""
        total_half = 0.5*width_ft + avoid_doors_ft + edge_margin_ft
        # pick wall-aligned direction sign from hand vector
        sgn = 1.0 if (hand_vec_host.X*wall_dir_u.X + hand_vec_host.Y*wall_dir_u.Y + hand_vec_host.Z*wall_dir_u.Z) >= 0.0 else -1.0
        along = XYZ(wall_dir_u.X*sgn, wall_dir_u.Y*sgn, wall_dir_u.Z*sgn)
        left_pt  = center_pt - along.Multiply(total_half)
        right_pt = center_pt + along.Multiply(total_half)
        da, _, perp_a = _project_on_curve(curve, left_pt)
        db, _, perp_b = _project_on_curve(curve, right_pt)
        if da is None or db is None:
            return None
        a, b = (da, db) if da <= db else (db, da)
        return (a, b)

    def _ingest(doors_iter, xform=None):
        for d in doors_iter:
            try:
                if not isinstance(d, FamilyInstance):
                    continue

                # Transform center + hand vector into host coords
                loc = getattr(d.Location, "Point", None)
                if loc is None:
                    continue
                center_host = xform.OfPoint(loc) if xform else loc

                hand = getattr(d, "HandOrientation", None)
                if hand is None:
                    # fall back to FacingOrientation ⟂ (rare)
                    hand = getattr(d, "FacingOrientation", XYZ(1,0,0))
                hand_host = _unit(xform.OfVector(hand) if xform else hand)

                # Ensure this door belongs to/near this wall (perp distance threshold)
                dist_center, _, perp = _project_on_curve(curve, center_host)
                if dist_center is None or perp is None or perp > proximity_tol_ft:
                    continue

                # Width
                width_ft = _door_width_param_ft(d)
                if width_ft is None:
                    width_ft = _door_width_bbox_ft(d, wall_dir_u, xform)
                if width_ft is None:
                    width_ft = 3.0  # conservative default

                span = _make_span(center_host, hand_host, width_ft)
                if span:
                    spans.append(span)
            except:
                pass

    # Host model doors
    try:
        host_doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
        _ingest(host_doors, None)
    except:
        pass

    # Linked model doors
    if include_linked:
        try:
            for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
                ldoc = link.GetLinkDocument()
                if ldoc is None:
                    continue
                T = link.GetTransform() if hasattr(link, "GetTransform") else link.GetTotalTransform()
                link_doors = FilteredElementCollector(ldoc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
                _ingest(link_doors, T)
        except:
            pass

    # Merge overlaps
    if not spans:
        return spans
    spans.sort(key=lambda ab: ab[0])
    merged = []
    s, e = spans[0]
    for a, b in spans[1:]:
        if a <= e + 1e-4:
            e = max(e, b)
        else:
            merged.append((s, e)); s, e = a, b
    merged.append((s, e))
    return merged

def _dist_along(curve, point):
    d, _, _ = _project_on_curve(curve, point)
    return d

def filter_points_keepouts(doc, wall, curve, pts,
                           avoid_corners_ft, avoid_doors_ft,
                           edge_margin_in=1.0, proximity_tol_ft=6.0,
                           include_linked=True):
    """Prune points near wall ends and inside doorway spans."""
    if not pts: return []
    L = _curve_len(curve)
    door_spans = _build_door_spans_for_wall(doc, wall, curve,
                                            avoid_doors_ft,
                                            edge_margin_in=edge_margin_in,
                                            include_linked=include_linked,
                                            proximity_tol_ft=proximity_tol_ft)
    out = []
    for p in pts:
        d = _dist_along(curve, p)
        if d is None:
            continue
        # corners
        if avoid_corners_ft > 0.0 and (d < avoid_corners_ft or (L - d) < avoid_corners_ft):
            continue
        # door spans
        blocked = False
        for a, b in door_spans:
            if a <= d <= b:
                blocked = True; break
        if not blocked:
            out.append(p)
    return out

# ---------- Rules access ----------
def get_receptacle_rules_for_category(elec, category):
    bc = elec.get("branch_circuits", {}) if elec else {}
    bycat = bc.get("receptacle_rules_by_category", {})
    return bycat.get(category, {})

# ---------- Type-catalog aware symbol loader (Electrical Fixtures) ----------
import re
def _collect_symbols_electrical(doc):
    pairs = []
    def add(col):
        for fs in col:
            try:
                fam = fs.Family
                pairs.append((fam.Name if fam else "", fs.Name, fs))
            except: pass
    try:
        add(FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_ElectricalFixtures))
    except: pass
    # fallback: all symbols if nothing found (rare)
    if not pairs:
        try:
            add(FilteredElementCollector(doc).OfClass(FamilySymbol))
        except: pass
    return pairs

def _parse_type_catalog(txt_path):
    if not os.path.exists(txt_path): return []
    with open(txt_path, 'rb') as f: data = f.read()
    text = None
    for enc in ('utf-8-sig', 'utf-16', 'latin-1', 'utf-8'):
        try: text = data.decode(enc); break
        except: pass
    if text is None: return []
    names = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith('//'): continue
        # type name = first token before comma/tab/semicolon
        cut = None
        for sep in (',','\t',';'):
            i = line.find(sep)
            if i != -1: cut = i; break
        first = line if cut is None else line[:cut]
        first = first.strip().strip('"').strip("'")
        if first: names.append(first)
    return names

def _choose_catalog_type(type_names, want_exact, want_regex, strict=False):
    if not type_names: return None
    if want_exact and want_exact in type_names:
        log("[CATALOG] exact: {}".format(want_exact)); return want_exact
    if want_exact:
        low = {n.lower(): n for n in type_names}
        key = " ".join(want_exact.split()).lower()
        if key in low:
            log("[CATALOG] case/space: {}".format(low[key])); return low[key]
    if want_regex:
        try:
            rx = re.compile(want_regex, re.IGNORECASE)
            for n in type_names:
                if rx.search(n):
                    log("[CATALOG] regex: {}".format(n)); return n
        except: pass
    if strict:
        log("[CATALOG] strict mode: no match"); return None
    log("[CATALOG] fallback first: {}".format(type_names[0])); return type_names[0]

def _try_load_family_shell(doc, rfa_path):
    if not (rfa_path and os.path.exists(rfa_path)): return False
    from Autodesk.Revit.DB import IFamilyLoadOptions
    class _AlwaysLoad(IFamilyLoadOptions):
        def OnFamilyFound(self, fi, ov):
            try: ov[0] = True
            except: pass
            return True
        def OnSharedFamilyFound(self, sf, fi, src, ov):
            try: ov[0] = True
            except: pass
            return True
    t = Transaction(doc, "Load Family (Receptacles)")
    t.Start()
    try:
        ok = doc.LoadFamily(rfa_path, _AlwaysLoad()); t.Commit(); return ok
    except Exception as ex:
        try: t.RollBack()
        except: pass
        log("[LOAD] Family shell failed: {}".format(ex))
        return False

def _try_load_symbol_from_catalog(doc, rfa_path, type_name):
    if not (rfa_path and os.path.exists(rfa_path) and type_name): return None
    try:
        out = clr.StrongBox[FamilySymbol]()  # out param
    except:
        out = clr.StrongBox[FamilySymbol](None)
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
            log("[LOAD] Loaded catalog type: {}".format(type_name))
            return sym
        t.RollBack(); log("[LOAD] LoadFamilySymbol returned False for: {}".format(type_name)); return None
    except Exception as ex:
        try: t.RollBack()
        except: pass
        log("[LOAD] Exception loading catalog symbol: {}".format(ex))
        return None

def pick_receptacle_symbol(doc, device_candidates):
    """Try rule candidates; if missing, search for a 'receptacle/duplex' symbol already loaded."""
    # 1) explicit candidates from rules
    for cand in (device_candidates or []):
        fam_req   = (cand.get('family') or "").strip()
        typ_exact = (cand.get('type_catalog_name') or cand.get('type') or "").strip()
        typ_regex = (cand.get('type_regex') or "").strip()
        rfa_path  = (cand.get('load_from') or "").strip()

        pairs = _collect_symbols_electrical(doc)
        # exact first
        for f, t, fs in pairs:
            if f == fam_req and (not typ_exact or t == typ_exact):
                try:
                    if not fs.IsActive: fs.Activate()
                except: pass
                log("[MATCH] Exact: {} :: {}".format(f, t)); return fs
            if f.lower() == fam_req.lower() and (not typ_exact or t.lower() == typ_exact.lower()):
                try:
                    if not fs.IsActive: fs.Activate()
                except: pass
                log("[MATCH] Case-insensitive: {} :: {}".format(f, t)); return fs

        # try to load family + (catalog) type
        if rfa_path and os.path.exists(rfa_path):
            _try_load_family_shell(doc, rfa_path)
            cat_path = os.path.splitext(rfa_path)[0] + ".txt"
            if os.path.exists(cat_path):
                names = _parse_type_catalog(cat_path)
                if names:
                    log("[CATALOG] {} type names".format(len(names)))
                    chosen = _choose_catalog_type(names, typ_exact, typ_regex, strict=False)
                    if chosen:
                        sym = _try_load_symbol_from_catalog(doc, rfa_path, chosen)
                        if sym: return sym
            # re-scan for non-catalog family types
            pairs = _collect_symbols_electrical(doc)
            for f, t, fs in pairs:
                if f == fam_req and (not typ_exact or t == typ_exact):
                    try:
                        if not fs.IsActive: fs.Activate()
                    except: pass
                    log("[MATCH] Exact after load: {} :: {}".format(f, t)); return fs

    # 2) fallback heuristic: pick any Electrical Fixture type that looks like a duplex receptacle
    pairs = _collect_symbols_electrical(doc)
    for f, t, fs in pairs:
        name = (f + " " + t).lower()
        if "recept" in name or "duplex" in name or "outlet" in name:
            try:
                if not fs.IsActive: fs.Activate()
            except: pass
            log("[MATCH] Heuristic: {} :: {}".format(f, t))
            return fs

    log("[MATCH] No Electrical Fixture symbols available.")
    return None

# ---------- Geometry along walls ----------
def room_wall_segments(doc, room):
    """Return list of (wall, curve_on_room_boundary) for walls bounding the room."""
    opts = SpatialElementBoundaryOptions()
    walls = []
    loops = room.GetBoundarySegments(opts)
    if not loops: return walls
    for loop in loops:
        for seg in loop:
            try:
                elem = doc.GetElement(seg.ElementId)
                if isinstance(elem, Wall):
                    walls.append((elem, seg.GetCurve()))
            except: pass
    return walls

def points_along_curve(curve, first_ft, next_ft):
    """Return list of XYZ (plan) along curve at first/next distances."""
    pts = []
    L = curve.Length
    if L <= 0: return pts
    d = max(0.0, float(first_ft))
    if d <= L:
        # place first point at >= first_ft from start
        while d < L + 1e-6:
            try:
                p = curve.Evaluate(d / L, True)  # normalized parameter
                pts.append(p)
            except:
                break
            d += max(0.1, float(next_ft))
    return pts

def place_on_wall(doc, sym, wall, pt_plan, z_ft):
    """Try several host methods: wall-based, face-based, one-level-based."""
    lvl = wall.LookupParameter("Base Constraint")
    level = None
    try:
        level = doc.GetElement(wall.LevelId)
    except:
        pass
    p3 = XYZ(pt_plan.X, pt_plan.Y, z_ft)

    # Try wall-hosted
    try:
        if sym.Family.FamilyPlacementType == FamilyPlacementType.WallBased:
            inst = doc.Create.NewFamilyInstance(p3, sym, wall, StructuralType.NonStructural)
            return inst
    except: pass

    # Try face-based on interior face
    try:
        if sym.Family.FamilyPlacementType.ToString().endswith("FaceBased"):
            refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior)
            if refs and len(refs) > 0:
                # orientation: use wall orientation as "hand" vector; normal arg is required
                normal = wall.Orientation  # points roughly from interior to exterior
                inst = doc.Create.NewFamilyInstance(refs[0], p3, normal, sym)
                return inst
    except: pass

    # 1-level based or non-hosted fallback
    try:
        host_level = level or doc.ActiveView.GenLevel
        inst = doc.Create.NewFamilyInstance(p3, sym, host_level, StructuralType.NonStructural)
        return inst
    except: pass

    return None


def _view_level_id(view):
    try:
        gl = getattr(view, "GenLevel", None)
        return gl.Id if gl else view.LevelId
    except:
        return None

def spatial_display_name(spatial):
    nm = ""; num = ""
    try:
        p = spatial.LookupParameter("Name");   nm  = (p.AsString() or "") if p else ""
    except: pass
    try:
        p = spatial.LookupParameter("Number"); num = (p.AsString() or "") if p else ""
    except: pass
    return (nm or "Unnamed"), (num or "")

def get_target_spaces(doc, uidoc, view, only_current_level=False, prefer_selection=True):
    """Return list of MEP Spaces honoring selection and current-level filters."""
    # prefer current selection if asked
    try:
        if prefer_selection and uidoc and uidoc.Selection:
            ids = list(uidoc.Selection.GetElementIds())
            if ids:
                out = []
                for eid in ids:
                    el = doc.GetElement(eid)
                    if el and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces):
                        out.append(el)
                if out:
                    return out
    except:
        pass

    col = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
    spaces = [s for s in col if getattr(s, "Area", 0.0) > 1e-6]  # skip unplaced/zero-area

    if only_current_level:
        vid = _view_level_id(view)
        if vid:
            spaces = [s for s in spaces if getattr(s, "LevelId", None) == vid]

    return spaces


def get_space_level(doc, space):
    try:
        return doc.GetElement(space.LevelId)
    except:
        return None

# alias if you prefer the old function name in-place:
get_room_level = get_space_level  # <-- so the rest of the code can still call get_room_level(...)

def space_display_name(space):
    # or use your existing spatial_display_name if you already defined it earlier in this file
    nm = ""; num = ""
    try:
        p = space.LookupParameter("Name");   nm  = (p.AsString() or "") if p else ""
    except: pass
    try:
        p = space.LookupParameter("Number"); num = (p.AsString() or "") if p else ""
    except: pass
    return (nm or "Unnamed"), (num or "")

# ---------- Main ----------
def do_all():
    # Load rules
    if not (os.path.exists(IDENTIFY) and os.path.exists(LIGHTING)):
        print("[FATAL] Missing identify_spaces.json / lighting_rules.json")
        return
    elec = load_electrical_rules(ELEC_DIR) if os.path.isdir(ELEC_DIR) else {}

    # Collect target Spaces (honors selection + level filter if you wire those flags)
    spaces = get_target_spaces(doc, uidoc, active_view, only_current_level=False, prefer_selection=True)
    if not spaces:
        print("[ERROR] No Space elements found.")
        print("  - Place MEP Spaces or enable Space creation from Rooms.")
        return

    if len(spaces) == 1:
        nm, num = space_display_name(spaces[0])
        print("[INFO] One space found: {} {}".format(nm, num))
    else:
        print("[INFO] {} spaces found. Processing…".format(len(spaces)))

    total_placed = 0

    for sp in spaces:
        nm, num = space_display_name(sp)

        # classify this Space to a category (reuses your existing classifier/rules)
        try:
            category, _lighting_rule = build_rule_for_room(nm or "", IDENTIFY, LIGHTING)
        except Exception as ex:
            print("[ERROR] Could not classify space '{}': {}".format(nm, ex))
            continue

        # pull category-specific receptacle rules
        rr = get_receptacle_rules_for_category(elec, category)
        if not rr:
            print("[RECEPT] {} {} → no receptacle rules for category '{}'".format(nm, num, category))
            continue

        pc_gen = (elec.get("branch_circuits", {}).get("general", {}).get("placement_constraints", {}) if elec else {})
        pc_cat = rr.get("placement_constraints", {}) or {}
        avoid_corners_ft = float(pc_cat.get("avoid_corners_ft", pc_gen.get("avoid_corners_ft", 0.0)))
        avoid_doors_ft = float(pc_cat.get("avoid_doors_ft", pc_gen.get("avoid_doors_ft", 0.0)))
        edge_margin_in = float(pc_cat.get("door_edge_margin_in", pc_gen.get("door_edge_margin_in", 1.0)))
        snap_tol_ft = float(pc_cat.get("door_snap_tolerance_ft", pc_gen.get("door_snap_tolerance_ft", 6.0)))
        avoid_doors_radius_ft = float(pc_cat.get("avoid_doors_radius_ft", pc_gen.get("avoid_doors_radius_ft", 6.0)))
        door_snap_tolerance_ft = float(pc_cat.get("door_snap_tolerance_ft", pc_gen.get("door_snap_tolerance_ft", 6.0)))

        first_ft = float(rr.get("wall_spacing_ft", {}).get("first", 6))
        next_ft  = float(rr.get("wall_spacing_ft", {}).get("next", 12))
        mh_in    = float(rr.get("mount_height_in", 16.0))
        min_per  = int(rr.get("min_per_room", 0))  # optional minimum; enforce below if you like
        cand     = rr.get("device_candidates", [])

        # choose symbol
        sym = pick_receptacle_symbol(doc, cand)
        if not sym:
            print("[RECEPT] {} {} → no receptacle FamilySymbol available (load_from missing or wrong?)".format(nm, num))
            continue

        # elevation at mount height above the Space’s level
        lvl = get_space_level(doc, sp)
        lvl_elev = (lvl.Elevation if lvl else 0.0)
        z_ft = lvl_elev + (mh_in / 12.0)

        # generate points along the Space’s bounding walls
        segs = room_wall_segments(doc, sp)  # works for Space via SpatialElementBoundaryOptions
        plan_pts = []
        wall_refs = []
        for wall, crv in segs:
            pts = points_along_curve(crv, first_ft, next_ft)
            pts = filter_points_keepouts(doc, wall, crv, pts,
                                         avoid_corners_ft, avoid_doors_ft,
                                         edge_margin_in=edge_margin_in,
                                         proximity_tol_ft=snap_tol_ft,
                                         include_linked=True)
            pts = filter_points_avoid_doors_radius(
                doc, wall, crv, pts,
                radius_ft=avoid_doors_radius_ft,
                snap_tol_ft=door_snap_tolerance_ft,
                include_linked=True
            )
            for p in pts:
                plan_pts.append(p)
                wall_refs.append(wall)

        # de-duplicate near corners
        dedup_pts, dedup_walls, seen = [], [], set()
        for p, w in zip(plan_pts, wall_refs):
            key = (round(p.X, 3), round(p.Y, 3))
            if key in seen:
                continue
            seen.add(key); dedup_pts.append(p); dedup_walls.append(w)

        # (optional) enforce min_per by sprinkling evenly across walls
        if min_per and len(dedup_pts) < min_per and segs:
            while len(dedup_pts) < min_per:
                w, crv = segs[len(dedup_pts) % len(segs)]
                try:
                    mid = crv.Evaluate(0.5, True)
                    key = (round(mid.X,3), round(mid.Y,3))
                    if key not in seen:
                        dedup_pts.append(mid); dedup_walls.append(w); seen.add(key)
                except:
                    break

        print("[RECEPT] {} {}: planning {} points (first={}’, next={}’, mh={}”)".format(nm, num, len(dedup_pts), first_ft, next_ft, mh_in))

        if DRY_RUN:
            continue

        # place
        t = Transaction(doc, "Place Receptacles (AutoLayout by Space)")
        t.Start()
        placed = 0
        try:
            for p, wall in zip(dedup_pts, dedup_walls):
                inst = place_on_wall(doc, sym, wall, p, z_ft)
                if inst: placed += 1
            t.Commit()
        except Exception as ex:
            print("[ERROR] Placement failed in {} {}: {}".format(nm, num, ex))
            try: t.RollBack()
            except: pass

        total_placed += placed
        print("[RECEPT] Placed {} receptacles in {} {}".format(placed, nm, num))

    print("[RESULT] Placed {} instances across {} space(s).".format(total_placed, len(spaces)))

# ---------------- Single Undo wrapper call ----------------
if __name__ == "__main__":
    run_as_single_undo(doc, "Place Receptacles", do_all)
