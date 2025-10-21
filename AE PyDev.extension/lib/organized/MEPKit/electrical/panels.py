# -*- coding: utf-8 -*-
from __future__ import absolute_import
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilyInstance
from Autodesk.Revit.DB import XYZ
from organized.MEPKit.revit.filters import of_category
from organized.MEPKit.revit.params import get_param_value
from organized.MEPKit.electrical.devices import location_point
from organized.MEPKit.electrical.calc import distance_ft

def panels(doc):
    return list(of_category(doc, BuiltInCategory.OST_ElectricalEquipment, only_instances=True))

def is_lighting_panel(panel):
    """Heuristic: name/mark contains 'LP' or 'Lighting'; easy to extend later."""
    name = getattr(panel, "Name", "") or ""
    mark = get_param_value(panel, "Mark") or ""
    text = (name + " " + mark).lower()
    return (" lighting" in text) or text.startswith("lp") or " lp" in text

def collect_panels(doc, only_lighting=False):
    ps = panels(doc)
    return [p for p in ps if (is_lighting_panel(p) if only_lighting else True)]

def panel_voltage(panel):
    # Try common parameters; adapt as needed for your families
    for n in ("Voltage", "Voltage (V)", "System Voltage"):
        v = get_param_value(panel, n)
        if v not in (None, ""):
            try: return float(v)
            except: pass
    return None

def panel_point(panel):
    return location_point(panel)

def nearest_panel(doc, point, only_lighting=True):
    candidates = collect_panels(doc, only_lighting=only_lighting)
    if not candidates or not point: return None
    best = None; bestd = None
    for p in candidates:
        pp = panel_point(p)
        if not pp: continue
        d = distance_ft(point, pp)
        if best is None or d < bestd:
            best, bestd = p, d
    return best