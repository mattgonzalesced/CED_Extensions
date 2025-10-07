# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/placement.py
from __future__ import absolute_import
import traceback
from Autodesk.Revit.DB import (
    FilteredElementCollector, Level, XYZ, BuiltInParameter, HostObjectUtils, ShellLayerType, Wall
)
from Autodesk.Revit.DB.Structure import StructuralType

# --- small utilities ------------------------------------------------

# ---------- tiny utils

def ensure_active(doc, symbol):
    if symbol and not symbol.IsActive:
        symbol.Activate(); doc.Regenerate()

def any_level(doc):
    for L in FilteredElementCollector(doc).OfClass(Level):
        return L
    return None

def _host_level(doc, host):
    # try direct LevelId
    try:
        if hasattr(host, "LevelId") and host.LevelId and host.LevelId.IntegerValue > 0:
            return doc.GetElement(host.LevelId)
    except: pass
    # walls: base constraint
    try:
        p = host.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        if p:
            L = doc.GetElement(p.AsElementId())
            if L: return L
    except: pass
    return any_level(doc)

def _set_elevation_like_params(elem, height_ft):
    if height_ft is None:
        return
    # common built-ins
    for bip in (
        BuiltInParameter.INSTANCE_ELEVATION_PARAM,
        BuiltInParameter.FAMILY_LEVEL_OFFSET_PARAM,
        BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
    ):
        try:
            p = elem.get_Parameter(bip)
            if p and (not p.IsReadOnly):
                p.Set(float(height_ft))
                return
        except: pass
    # friendly names used by many content libraries
    for name in ("Elevation from Level", "Elevation", "Offset", "Mounting Height", "Device Elevation", "Height"):
        try:
            p = elem.LookupParameter(name)
            if p and (not p.IsReadOnly):
                p.Set(float(height_ft))
                return
        except: pass

def _log_exc(logger, tag, ex):
    if logger:
        try:
            logger.debug(u"[{}] {}".format(tag, ex))
            logger.debug(traceback.format_exc())
        except: pass

# ---------- public API (with fallbacks)

def place_hosted(doc, host, symbol, point_xyz, mounting_height_ft=None, logger=None):
    """
    Best-effort placement on a host (e.g., Wall) at mount height.
    Tries:
      A) host + point with Z = level + height
      B) host + point with original Z
      C) face-based fallback using a wall face Reference (if host is a Wall)
    """
    ensure_active(doc, symbol)
    lvl = _host_level(doc, host)
    base_elev = getattr(lvl, "Elevation", 0.0) if lvl else 0.0
    pA = XYZ(point_xyz.X, point_xyz.Y, base_elev + (mounting_height_ft or 0.0))
    pB = XYZ(point_xyz.X, point_xyz.Y, point_xyz.Z)

    # A) normal host placement at desired Z
    try:
        inst = doc.Create.NewFamilyInstance(pA, symbol, host, StructuralType.NonStructural)
        _set_elevation_like_params(inst, mounting_height_ft)
        return inst
    except Exception as ex:
        _log_exc(logger, "place_hosted:A", ex)

    # B) try original Z (some host-based families ignore point.Z anyway)
    try:
        inst = doc.Create.NewFamilyInstance(pB, symbol, host, StructuralType.NonStructural)
        _set_elevation_like_params(inst, mounting_height_ft)
        return inst
    except Exception as ex:
        _log_exc(logger, "place_hosted:B", ex)

    # C) face-based fallback on a wall face (for face-hosted families)
    try:
        if isinstance(host, Wall):
            # prefer exterior, then interior
            refs = list(HostObjectUtils.GetSideFaces(host, ShellLayerType.Exterior)) or \
                   list(HostObjectUtils.GetSideFaces(host, ShellLayerType.Interior))
            if refs:
                ref = refs[0]
                # point: use desired Z if possible
                pC = pA
                normal = host.Orientation  # Revit uses this as "up" on the face-based overload
                inst = doc.Create.NewFamilyInstance(ref, pC, normal, symbol)
                _set_elevation_like_params(inst, mounting_height_ft)
                return inst
    except Exception as ex:
        _log_exc(logger, "place_hosted:C(face)", ex)

    # give up
    raise Exception("Host placement failed for symbol '{}'".format(getattr(symbol, "Name", "<unknown>")))

def place_free(doc, symbol, point_xyz, level=None, mounting_height_ft=None, logger=None):
    """
    Level-based placement at mount height with fallback to original Z.
    """
    ensure_active(doc, symbol)
    lvl = level or any_level(doc)
    base_elev = getattr(lvl, "Elevation", 0.0) if lvl else 0.0
    pA = XYZ(point_xyz.X, point_xyz.Y, base_elev + (mounting_height_ft or 0.0))
    pB = XYZ(point_xyz.X, point_xyz.Y, point_xyz.Z)

    try:
        inst = doc.Create.NewFamilyInstance(pA, symbol, lvl, StructuralType.NonStructural)
        _set_elevation_like_params(inst, mounting_height_ft)
        return inst
    except Exception as ex:
        _log_exc(logger, "place_free:A", ex)

    try:
        inst = doc.Create.NewFamilyInstance(pB, symbol, lvl, StructuralType.NonStructural)
        _set_elevation_like_params(inst, mounting_height_ft)
        return inst
    except Exception as ex:
        _log_exc(logger, "place_free:B", ex)

    raise Exception("Free placement failed for symbol '{}'".format(getattr(symbol, "Name", "<unknown>")))