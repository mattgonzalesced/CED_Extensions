# -*- coding: utf-8 -*-
# lib/placement.py
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, XYZ, Transaction,
    HostObjectUtils, FamilyPlacementType, BuiltInParameter, Plane, SketchPlane,
    ElementTransformUtils, SpatialElementBoundaryOptions
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.DB.Mechanical import Space
from rooms import get_room_level, get_room_bbox_center, room_display_name
from geometry import propose_grid_points_from_rule
from families import pick_fixture_symbol
from rules_loader import deep_merge

NEG_Z = XYZ(0, 0, -1)

def is_space(spatial):
    try:
        return isinstance(spatial, Space)
    except:
        return False

def spatial_display_name(spatial):
    # Safely get "Name" / "Number" for rooms or spaces
    nm = ""; num = ""
    for pname in ("Name", "ROOM_NAME", "SPACE_NAME"):
        try:
            p = spatial.LookupParameter(pname)
            if p: nm = p.AsString() or nm
        except: pass
    for pname in ("Number", "ROOM_NUMBER", "SPACE_NUMBER"):
        try:
            p = spatial.LookupParameter(pname)
            if p: num = p.AsString() or num
        except: pass
    return (nm or "Unnamed"), (num or "")

def _get_spatial_level(doc, spatial):
    try:
        return doc.GetElement(spatial.LevelId)
    except:
        return None

# keep compatibility with earlier functions that expected "room"
def _get_room_level(doc, spatial):  # alias
    return _get_spatial_level(doc, spatial)

def _spatial_test_z(doc, spatial):
    """A safe Z for IsPointInRoom/IsPointInSpace (~3 ft above base level)."""
    lvl = _get_spatial_level(doc, spatial)
    base = (lvl.Elevation if lvl else 0.0)
    return base + 3.0

def spatial_point_contains(spatial, pt):
    """True if pt is inside the room/space footprint."""
    try:
        return spatial.IsPointInRoom(pt)    # Rooms
    except:
        try:
            return spatial.IsPointInSpace(pt)  # Spaces
        except:
            return False  # fallback if neither method exists

def _room_test_z(doc, room):
    """A safe Z inside the room for IsPointInRoom()."""
    lvl = _get_room_level(doc, room)
    base = (lvl.Elevation if lvl else 0.0)
    return base + 3.0  # ~3 ft above floor works for point-in-room

def _ensure_sketchplane_at_z(doc, view, z_ft):
    """Create a horizontal SketchPlane at Z=z_ft and make sure placement uses it."""
    plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, float(z_ft)))
    sp = SketchPlane.Create(doc, plane)
    # Try to set as active work plane for the view (not strictly required if we pass sp explicitly)
    try:
        view.SketchPlane = sp
        try:
            view.ShowActiveWorkPlane = True
        except:
            pass
    except:
        pass
    return sp

def _set_instance_offset_from_level(inst, offset_ft):
    for bip in (BuiltInParameter.INSTANCE_ELEVATION_PARAM,
                BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM):
        try:
            p = inst.get_Parameter(bip)
            if p and p.StorageType.ToString() == "Double":
                p.Set(float(offset_ft)); return True
        except: pass
    for pname in ("Elevation", "Offset", "Offset from Level", "Height Offset From Level"):
        try:
            p = inst.LookupParameter(pname)
            if p and p.StorageType.ToString() == "Double":
                p.Set(float(offset_ft)); return True
        except: pass
    return False

def _raise_instance_to_z(doc, inst, z_target):
    """Last-resort nudge: move the element in Z so its Location hits z_target."""
    try:
        loc = inst.Location
        if hasattr(loc, "Point") and loc.Point:
            dz = float(z_target) - float(loc.Point.Z)
            if abs(dz) > 1e-3:
                ElementTransformUtils.MoveElement(doc, inst.Id, XYZ(0, 0, dz))
                return True
    except:
        pass
    return False

def _get_room_level(doc, room):
    try:
        return doc.GetElement(room.LevelId)
    except:
        return None

def _ceiling_underside_elev_ft(doc, ceiling):
    """Level elevation + Height Offset From Level (ft)."""
    elev = 0.0
    try:
        lvl = doc.GetElement(ceiling.LevelId); elev = lvl.Elevation if lvl else 0.0
        p = ceiling.LookupParameter("Height Offset From Level")
        if p: elev += (p.AsDouble() or 0.0)
    except:
        pass
    return elev

def _find_ceiling_under_point(doc, room, pt_xy):
    """Pick a ceiling in this room’s level whose bbox contains the XY."""
    lvl = _get_room_level(doc, room)
    cands = []
    for c in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType():
        try:
            if lvl and c.LevelId != lvl.Id:
                continue
        except:
            pass
        bb = c.get_BoundingBox(doc.ActiveView) or c.get_BoundingBox(None)
        if not bb:
            continue
        if (bb.Min.X - 1e-3) <= pt_xy.X <= (bb.Max.X + 1e-3) and (bb.Min.Y - 1e-3) <= pt_xy.Y <= (bb.Max.Y + 1e-3):
            cands.append(c)
    if not cands:
        return None, None
    # pick the highest (in case of multiple)
    cands.sort(key=lambda cc: _ceiling_underside_elev_ft(doc, cc))
    host = cands[-1]
    return host, _ceiling_underside_elev_ft(doc, host)

def place_light_instance(doc, sym, room, pt_plan, rule):
    """Prefer true ceiling hosting. For WorkPlane/Level families, place on a SketchPlane at ceiling Z or set offset/move up."""
    # activate symbol safely
    try:
        if sym and not sym.IsActive:
            sym.Activate()
    except:
        pass

    # SAFE: get FamilyPlacementType without assuming it's available
    fpt = None
    try:
        fam = getattr(sym, "Family", None)
        if fam:
            fpt = fam.FamilyPlacementType
    except:
        fpt = None  # leave as unknown; we'll branch by trial

    # find ceiling + target Z
    host, z_underside = _find_ceiling_under_point(doc, room, pt_plan)
    lvl = _get_room_level(doc, room)
    lvl_elev = (lvl.Elevation if lvl else 0.0)
    mount_elev_ft = float(rule.get("mount_elev_ft", 9.0))
    z_target = (z_underside if z_underside is not None else (lvl_elev + mount_elev_ft))

    # --- 1) CeilingBased: host to ceiling element
    try:
        from Autodesk.Revit.DB import FamilyPlacementType as FPT
        if fpt == FPT.CeilingBased and host is not None:
            return doc.Create.NewFamilyInstance(pt_plan, sym, host, StructuralType.NonStructural)
    except:
        pass

    # --- 2) FaceBased: host to underside face of the ceiling
    try:
        if fpt is not None and str(fpt).endswith("FaceBased") and host is not None:
            refs = HostObjectUtils.GetBottomFaces(host)
            if refs and len(refs) > 0:
                try:
                    return doc.Create.NewFamilyInstance(refs[0], pt_plan, NEG_Z, sym)
                except:
                    return doc.Create.NewFamilyInstance(refs[0], pt_plan, XYZ.BasisZ, sym)
    except:
        pass

    # --- 3) WorkPlaneBased: set SketchPlane at ceiling Z, place there, then verify Z
    try:
        from Autodesk.Revit.DB import FamilyPlacementType as FPT
        if fpt == FPT.WorkPlaneBased:
            sp = _ensure_sketchplane_at_z(doc, doc.ActiveView, z_target)
            doc.Regenerate()  # make sure view uses the new plane

            inst = None
            # Try overload that takes SketchPlane explicitly
            try:
                inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, 0.0), sym, sp, StructuralType.NonStructural)
            except:
                # Fallback to view-based overload (uses view's active work plane)
                try:
                    inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, z_target), sym, doc.ActiveView)
                except:
                    # Oldest API fallback
                    inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, z_target), sym, sp)

            if inst:
                doc.Regenerate()
                # If still low, nudge up or set offset
                if not _raise_instance_to_z(doc, inst, z_target):
                    _set_instance_offset_from_level(inst, z_target - lvl_elev)
                return inst
    except:
        pass

    # --- 4) Unknown / Level-based: place at level, then set offset or move up
    try:
        inst = doc.Create.NewFamilyInstance(XYZ(pt_plan.X, pt_plan.Y, lvl_elev), sym, lvl, StructuralType.NonStructural)
        if not _set_instance_offset_from_level(inst, z_target - lvl_elev):
            _raise_instance_to_z(doc, inst, z_target)
        return inst
    except:
        return None

def resolve_host_elevation(doc, room, default_height_ft=8.0):
    lvl = get_room_level(doc, room)
    lvl_elev = lvl.Elevation if lvl else 0.0
    return lvl_elev + default_height_ft


def place_fixtures_in_space(doc, active_view, spatial, rule, dry_run=True, verbose=True):
    name, number = spatial_display_name(spatial)

    # allow candidate to override category fields
    cands = rule.get('fixture_candidates') or []
    eff = deep_merge(rule, cands[0]) if cands else dict(rule)

    # plan points
    pts_plan = propose_grid_points_from_rule(spatial, active_view, eff)  # works for Space (SpatialElement)
    if not isinstance(pts_plan, (list, tuple)):
        pts_plan = []

    # clip to Space
    z_test = _spatial_test_z(doc, spatial)
    pts_plan = [p for p in pts_plan if spatial_point_contains(spatial, XYZ(p.X, p.Y, z_test))]

    if verbose:
        print("[PLAN] {} {}: {} pts after clipping".format(name, number, len(pts_plan)))

    if dry_run:
        return len(pts_plan), []

    # replace existing (optional)
    if eff.get("replace_existing"):
        # simple in-place filter for lights contained by the space
        victims = []
        for inst in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_LightingFixtures).WhereElementIsNotElementType():
            lp = getattr(getattr(inst, "Location", None), "Point", None)
            if lp and spatial_point_contains(spatial, XYZ(lp.X, lp.Y, z_test)):
                victims.append(inst)
        if victims:
            tdel = Transaction(doc, "Clear Space Lights")
            tdel.Start()
            for v in victims: doc.Delete(v.Id)
            tdel.Commit()
            if verbose: print("[PLACE] Removed {} existing light(s)".format(len(victims)))

    # pick family/type
    sym = pick_fixture_symbol(doc, eff)
    if not sym:
        if verbose: print("[WARN] No matching FamilySymbol; skipping Space {}".format(number))
        return len(pts_plan), []

    # place with ceiling-aware routine you already have
    placed = []
    t = Transaction(doc, "Place Lighting (Space)")
    t.Start()
    try:
        for p in pts_plan:
            inst = place_light_instance(doc, sym, spatial, p, eff)  # spatial passed into existing function
            if inst: placed.append(inst)
        t.Commit()
        if verbose:
            print("[PLACE] Placed {} instances in {} {}".format(len(placed), name, number))
    except Exception as ex:
        print("[ERROR] Placement failed in {} {}: {}".format(name, number, ex))
        try: t.RollBack()
        except: pass

    return len(pts_plan), placed



def place_fixtures_in_room(doc, active_view, room, rule, dry_run=True, verbose=True):
    # --- names ---
    name, number = room_display_name(room)

    # --- effective rule: let candidate override category fields (spacing_ft, offset_ft, etc.) ---
    cands = rule.get('fixture_candidates') or []
    eff = deep_merge(rule, cands[0]) if cands else dict(rule)

    # --- debug: what spacing target are we actually using? ---
    print("[DBG] spacing target used:", (eff.get("spacing_ft") or {}).get("target"))

    # --- plan points strictly from JSON (no rounding math) ---
    pts_plan = propose_grid_points_from_rule(room, active_view, eff)
    if not isinstance(pts_plan, (list, tuple)):
        pts_plan = []

    # --- clip to room polygon ---
    #z_test = _room_test_z(doc, room)
    #before = len(pts_plan)
    #pts_plan = [p for p in pts_plan if room.IsPointInRoom(XYZ(p.X, p.Y, z_test))]
    #if verbose:
    #    print("[CLIP] {} {}: kept {}/{} pts inside room".format(name, number, len(pts_plan), before))

    # --- clip points to the actual Space polygon ---
    z_test = _spatial_test_z(doc, spatial)  # spatial = the Space object
    before = len(pts_plan)
    pts_plan = [p for p in pts_plan if spatial_point_contains(spatial, XYZ(p.X, p.Y, z_test))]
    print("[PLACE] Placed {} instances in {} {}".format(len(placed), name, number))

    if eff.get("require_ceiling_host"):
        kept = []
        for p in pts_plan:
            host, _ = _find_ceiling_under_point(doc, room, p)  # you already have this helper
            if host is not None:
                kept.append(p)
        if verbose:
            print("[CLIP] {} {}: removed {} pts with no ceiling host".format(
                name, number, len(pts_plan) - len(kept)))
        pts_plan = kept

    # --- debug: show actual step measured from planned points ---
    if pts_plan:
        xs = sorted({ round(p.X, 3) for p in pts_plan })
        ys = sorted({ round(p.Y, 3) for p in pts_plan })
        stepx = (xs[1] - xs[0]) if len(xs) > 1 else None
        stepy = (ys[1] - ys[0]) if len(ys) > 1 else None
    else:
        stepx = stepy = None

    if verbose:
        print("[PLAN] {} {}: proposed {} pts{}".format(
            name, number, len(pts_plan),
            "" if stepx is None else " (stepX≈{:.2f}ft, stepY≈{:.2f}ft)".format(stepx, stepy or stepx)
        ))

    # --- dry run? ---
    if dry_run:
        return len(pts_plan), []

    # --- optionally clear existing lights in this room (if you want layouts to update) ---
    if eff.get("replace_existing"):
        try:
            from qc import lights_in_room
            victims = lights_in_room(doc, room)
        except:
            victims = []
        if victims:
            tdel = Transaction(doc, "Replace Lights (clear existing)")
            tdel.Start()
            try:
                for v in victims:
                    doc.Delete(v.Id)
                tdel.Commit()
                if verbose:
                    print("[PLACE] Removed {} existing light(s) before re-layout".format(len(victims)))
            except:
                tdel.RollBack()

    # --- pick symbol using the *effective* rule (so candidate fields apply) ---
    sym = pick_fixture_symbol(doc, eff)
    if sym is None:
        if verbose:
            print("[WARN] No matching FamilySymbol for rule; skipping placement.")
        return len(pts_plan), []

    # --- place: prefer ceiling host / face host; fallback to ceiling Z ---
    placed = []
    t = Transaction(doc, "Place Lighting (AutoLayout)")
    t.Start()
    try:
        for p in pts_plan:
            inst = place_light_instance(doc, sym, room, p, eff)   # <<< ceiling-aware drop
            if inst:
                placed.append(inst)
        t.Commit()
        if verbose:
            print("[PLACE] Placed {} instances in {} {}".format(len(placed), name, number))
    except Exception as ex:
        if verbose:
            print("[ERROR] Placement failed:", ex)
        try:
            t.RollBack()
        except:
            pass

    return len(pts_plan), placed