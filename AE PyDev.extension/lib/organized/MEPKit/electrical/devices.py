# -*- coding: utf-8 -*-
from __future__ import absolute_import
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB import LocationPoint, LocationCurve, XYZ

from organized.MEPKit.revit.filters import of_category
from organized.MEPKit.revit.params import get_param_value

def _elem_point(elem):
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        return loc.Point
    if isinstance(loc, LocationCurve):
        c = loc.Curve
        return XYZ((c.GetEndPoint(0).X + c.GetEndPoint(1).X)/2.0,
                   (c.GetEndPoint(0).Y + c.GetEndPoint(1).Y)/2.0,
                   (c.GetEndPoint(0).Z + c.GetEndPoint(1).Z)/2.0)
    # Fallback to bbox center
    bb = elem.get_BoundingBox(None)
    if bb:
        return XYZ((bb.Min.X + bb.Max.X)/2.0, (bb.Min.Y + bb.Max.Y)/2.0, (bb.Min.Z + bb.Max.Z)/2.0)
    return None

def location_point(elem):
    """Safe XYZ for any family instance-like element."""
    return _elem_point(elem)

def receptacles(doc):
    return list(of_category(doc, BuiltInCategory.OST_ElectricalFixtures, only_instances=True))

def lighting_fixtures(doc):
    return list(of_category(doc, BuiltInCategory.OST_LightingFixtures, only_instances=True))

def lighting_devices(doc):  # switches, sensors, etc.
    return list(of_category(doc, BuiltInCategory.OST_LightingDevices, only_instances=True))

def is_circuited(elem):
    mep = getattr(elem, "MEPModel", None)
    if not mep: return False
    try:
        # Some versions expose .ElectricalSystems or .GetElectricalSystems()
        systems = getattr(mep, "ElectricalSystems", None) or []
        if systems: return len(list(systems)) > 0
        if hasattr(mep, "GetElectricalSystems"):
            ss = mep.GetElectricalSystems()
            return ss is not None and ss.Size > 0
    except:
        pass
    return False

def uncircuited(elements):
    return [e for e in elements if not is_circuited(e)]

def apparent_load_va(elem):
    """Try common param names for VA; fallback None."""
    for n in ("Apparent Load", "Apparent Load (VA)", "Load", "Wattage", "Watts", "VA"):
        v = get_param_value(elem, n)
        if v not in (None, ""):
            try: return float(v)
            except: pass
    return None