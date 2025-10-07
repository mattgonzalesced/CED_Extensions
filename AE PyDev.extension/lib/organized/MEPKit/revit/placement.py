# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/placement.py
from __future__ import absolute_import
import traceback
from System import Enum
from Autodesk.Revit.DB import (
    FilteredElementCollector, Level, XYZ, BuiltInParameter, HostObjectUtils, ShellLayerType, Wall
)
from Autodesk.Revit.DB.Structure import StructuralType


# --- small utilities ------------------------------------------------

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

def _set_elevation_like_params(elem, height_ft, logger=None):
    if height_ft is None:
        return
    # the bip thing here is because Built in parameter.Family_level_offset_param was being shadowed.
    def _try_bip(name):
        try:
            bip = Enum.Parse(BuiltInParameter, name)  # parse by name, avoids attribute lookup
            p = elem.get_Parameter(bip)
            if p and (not p.IsReadOnly):
                p.Set(float(height_ft))
                if logger: logger.debug(u"[ELEV] set via {} = {}".format(name, height_ft))
                return True
        except Exception as ex:
            if logger: logger.debug(u"[ELEV] {} not usable: {}".format(name, ex))
        return False

    # Built-in names to try (some may not exist in your Revit/content)
    for bip_name in ("INSTANCE_ELEVATION_PARAM",
                     "FAMILY_LEVEL_OFFSET_PARAM",
                     "INSTANCE_FREE_HOST_OFFSET_PARAM"):
        if _try_bip(bip_name):
            return

    # Friendly/project parameter names
    for name in ("Elevation from Level", "Elevation", "Offset",
                 "Mounting Height", "Device Elevation", "Height"):
        try:
            p = elem.LookupParameter(name)
            if p and (not p.IsReadOnly):
                p.Set(float(height_ft))
                if logger: logger.debug(u"[ELEV] set via '{}' = {}".format(name, height_ft))
                return
        except Exception as ex:
            if logger: logger.debug(u"[ELEV] '{}' not usable: {}".format(name, ex))

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
                pC = pA
                # pick a reference direction not parallel to the face normal (host.Orientation is the face normal on walls)
                ref_dir = XYZ.BasisZ
                try:
                    n = host.Orientation
                    if abs(ref_dir.DotProduct(n)) > 0.99:
                        ref_dir = XYZ.BasisX
                except:  # if Orientation not available, keep BasisZ
                    pass
                inst = doc.Create.NewFamilyInstance(ref, pC, ref_dir, symbol)
                _set_elevation_like_params(inst, mounting_height_ft, logger=logger)
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
        _set_elevation_like_params(inst, mounting_height_ft, logger=logger)
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