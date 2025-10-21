# -*- coding: utf-8 -*-
# IronPython 2.7 + Revit 2024/2025-friendly
from __future__ import absolute_import
import math
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilyInstance, Element, ElementId, XYZ
)
from Autodesk.Revit.DB import LocationPoint, LocationCurve

from organized.MEPKit.revit.filters import of_category, of_class
from organized.MEPKit.revit.params import get_param_value

# ---------- basics

def active_phase(doc, name=None):
    """Return the active/newest phase, or the one matching name."""
    phases = list(doc.Phases)
    if not phases: return None
    if name:
        for ph in phases:
            if (ph.Name or '').strip().lower() == name.strip().lower():
                return ph
    return phases[-1]

def _loc_point(elem):
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point
    if isinstance(loc, LocationCurve):
        c = loc.Curve
        p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
        return XYZ(0.5*(p0.X+p1.X), 0.5*(p0.Y+p1.Y), 0.5*(p0.Z+p1.Z))
    bb = elem.get_BoundingBox(None)
    if bb:
        return XYZ(0.5*(bb.Min.X+bb.Max.X), 0.5*(bb.Min.Y+bb.Max.Y), 0.5*(bb.Min.Z+bb.Max.Z))
    return None

def element_point(elem):
    """Safe-ish XYZ for any element (family instances preferred)."""
    return _loc_point(elem)

def _distance(a, b):
    return math.sqrt((a.X-b.X)**2 + (a.Y-b.Y)**2 + (a.Z-b.Z)**2)

# ---------- by category

def collect_family_instances(doc, categories=None):
    """Collect FamilyInstances, optionally restricting to one or many BuiltInCategory enums."""
    col = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    elems = list(col)
    if not categories:
        return elems
    if not isinstance(categories, (list, tuple)):
        categories = [categories]
    cats = set(categories)
    return [e for e in elems if e.Category and e.Category.Id.IntegerValue in [c.value__ for c in cats]]

# ---------- space / room / level lookups

def element_space_or_room(elem, doc, phase=None):
    """Try FamilyInstance.Space[phase] (MEP) then Room[phase] (Arch)."""
    phase = phase or active_phase(doc)
    # Many MEP families expose Space/Room indexers by phase:
    try:
        sp = getattr(elem, "Space", None)
        if sp:
            try:
                s = sp[phase]
                if s: return s
            except: pass
    except: pass
    try:
        rm = getattr(elem, "Room", None)
        if rm:
            try:
                r = rm[phase]
                if r: return r
            except: pass
    except: pass
    return None  # fallback: not resolved

def devices_in_space(doc, space, categories=None):
    """All devices whose Space/Room resolves to this spatial element."""
    res = []
    for e in collect_family_instances(doc, categories=categories):
        s = element_space_or_room(e, doc)
        if s and s.Id == space.Id:
            res.append(e)
    return res

def devices_in_room(doc, room, categories=None):
    """Alias of devices_in_space for clarity."""
    return devices_in_space(doc, room, categories=categories)

def devices_on_level(doc, level, categories=None):
    """All devices whose LevelId matches the given Level."""
    lvl_id = level.Id
    out = []
    for e in collect_family_instances(doc, categories=categories):
        try:
            if getattr(e, "LevelId", ElementId.InvalidElementId) == lvl_id:
                out.append(e)
            else:
                # Some families store host level via parameter "Reference Level" or built-in
                ref_lvl = get_param_value(e, "Reference Level")
                if ref_lvl and (ref_lvl.strip().lower() == (level.Name or '').strip().lower()):
                    out.append(e)
        except:
            pass
    return out

# ---------- spatial search

def devices_within_radius(doc, point, radius_ft, categories=None):
    """All devices within radius (ft) of a point."""
    r2 = float(radius_ft) ** 2.0
    hits = []
    for e in collect_family_instances(doc, categories=categories):
        p = element_point(e)
        if not p: continue
        dx = p.X - point.X; dy = p.Y - point.Y; dz = p.Z - point.Z
        if (dx*dx + dy*dy + dz*dz) <= r2:
            hits.append(e)
    return hits

# ---------- grouping helpers

def group_by_space(doc, elements):
    """Return dict: SpatialElementId (or None) -> list of elements."""
    g = {}
    for e in elements:
        s = element_space_or_room(e, doc)
        key = s.Id if s else None
        g.setdefault(key, []).append(e)
    return g

def group_by_level(elements):
    """Return dict: LevelId (or None) -> list of elements."""
    g = {}
    for e in elements:
        lid = getattr(e, "LevelId", None)
        key = lid if lid else None
        g.setdefault(key, []).append(e)
    return g