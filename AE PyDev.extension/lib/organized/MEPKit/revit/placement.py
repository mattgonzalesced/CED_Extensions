# -*- coding: utf-8 -*-
# lib/organized/MEPKit/revit/placement.py
from __future__ import absolute_import
from Autodesk.Revit.DB import (
    FilteredElementCollector, Level, XYZ, BuiltInParameter
)
from Autodesk.Revit.DB.Structure import StructuralType

# --- small utilities ------------------------------------------------

def ensure_active(doc, symbol):
    if symbol and not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

def any_level(doc):
    for L in FilteredElementCollector(doc).OfClass(Level):
        return L
    return None

def _host_level(doc, host):
    """Best-effort level for a host (e.g., Wall)."""
    try:
        # direct LevelId when present
        if hasattr(host, "LevelId") and host.LevelId and host.LevelId.IntegerValue > 0:
            return doc.GetElement(host.LevelId)
    except:
        pass
    # common “base constraint” style params (e.g., on walls)
    for bip in (BuiltInParameter.WALL_BASE_CONSTRAINT,):
        try:
            p = host.get_Parameter(bip)
            if p:
                lvl = doc.GetElement(p.AsElementId())
                if lvl: return lvl
        except:
            pass
    return any_level(doc)

def _set_elevation_like_params(elem, height_ft):
    """Try common elevation/offset params after placement."""
    if height_ft is None:
        return
    # Built-ins first
    for bip in (
        BuiltInParameter.INSTANCE_ELEVATION_PARAM,        # many 1-level families
        BuiltInParameter.FAMILY_LEVEL_OFFSET_PARAM,       # level-based offset
        BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM  # free-host offset
    ):
        try:
            p = elem.get_Parameter(bip)
            if p and (not p.IsReadOnly):
                p.Set(float(height_ft))
                return
        except:
            pass
    # Friendly/project params
    for name in ("Elevation from Level", "Elevation", "Offset", "Mounting Height", "Height"):
        try:
            p = elem.LookupParameter(name)
            if p and (not p.IsReadOnly):
                p.Set(float(height_ft))
                return
        except:
            pass

# --- public placement API -------------------------------------------

def place_hosted(doc, host, symbol, point_xyz, mounting_height_ft=None):
    """
    Place a host-based family at (X,Y) on host, set Z = Level.Elevation + mounting_height_ft,
    then try to set common elevation/offset parameters.
    """
    ensure_active(doc, symbol)
    lvl = _host_level(doc, host)
    base_elev = getattr(lvl, "Elevation", 0.0) if lvl else 0.0
    desired_z = base_elev + (mounting_height_ft or 0.0)
    p = XYZ(point_xyz.X, point_xyz.Y, desired_z)

    inst = doc.Create.NewFamilyInstance(p, symbol, host, StructuralType.NonStructural)
    _set_elevation_like_params(inst, mounting_height_ft)
    return inst

def place_free(doc, symbol, point_xyz, level=None, mounting_height_ft=None):
    """
    Place a level-based family at (X,Y), set Z = Level.Elevation + mounting_height_ft,
    then try to set common elevation/offset parameters.
    """
    ensure_active(doc, symbol)
    lvl = level or any_level(doc)
    base_elev = getattr(lvl, "Elevation", 0.0) if lvl else 0.0
    desired_z = base_elev + (mounting_height_ft or 0.0)
    p = XYZ(point_xyz.X, point_xyz.Y, desired_z)

    inst = doc.Create.NewFamilyInstance(p, symbol, lvl, StructuralType.NonStructural)
    _set_elevation_like_params(inst, mounting_height_ft)
    return inst