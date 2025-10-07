# -*- coding: utf-8 -*-
# Helpers to propose and apply names/numbers for circuits, panels, and device marks.
from __future__ import absolute_import
import re

from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType

from organized.MEPKit.revit.params import get_param_value, set_param_value
from organized.MEPKit.revit.transactions import RunInTransaction

_slug_rx = re.compile(r'[^A-Za-z0-9\-]+')

def slugify(text, sep='-'):
    text = (text or u'').strip()
    text = text.replace('–','-').replace('—','-').replace('_','-')
    text = _slug_rx.sub(sep, text)
    text = re.sub(r'{0}+'.format(re.escape(sep)), sep, text).strip(sep)
    return text

# --------- CIRCUITS

def _systems_on_panel(doc, panel):
    out = []
    col = FilteredElementCollector(doc).OfClass(ElectricalSystem)
    for s in col:
        try:
            # Prefer BaseEquipment; fall back to Panel if present in your version
            be = getattr(s, 'BaseEquipment', None) or getattr(s, 'Panel', None)
            if be and be.Id == panel.Id:
                out.append(s)
        except:
            pass
    return out

_num_rx = re.compile(r'\d+')

def _used_circuit_numbers(doc, panel, power_only=True):
    """Return a sorted set of integer circuit numbers already used on the panel."""
    used = set()
    for s in _systems_on_panel(doc, panel):
        try:
            if power_only and s.SystemType != ElectricalSystemType.PowerCircuit:
                continue
            cn = getattr(s, "CircuitNumber", None)
            if not cn: continue
            # extract all integers from strings like "1", "1-3", "1,3,5", etc.
            for m in _num_rx.findall(str(cn)):
                try: used.add(int(m))
                except: pass
        except:
            pass
    return sorted(used)

def next_circuit_number(doc, panel, start=1, side='either'):
    """
    Propose the next available circuit number on a panel.
    side: 'either' | 'odd' | 'even'
    """
    used = set(_used_circuit_numbers(doc, panel))
    n = max(int(start), 1)
    if side == 'odd':
        if n % 2 == 0: n += 1
        while n in used: n += 2
        return n
    if side == 'even':
        if n % 2 != 0: n += 1
        while n in used: n += 2
        return n
    # either side
    while n in used: n += 1
    return n

def propose_circuit_name(panel, number, desc=None):
    """
    Example: LP-1-07 'Office Recepts'
    If panel has a Name or Mark, use that.
    """
    pname = getattr(panel, "Name", None) or get_param_value(panel, "Mark") or "PNL"
    if isinstance(number, int):
        numtxt = "{:02d}".format(number)
    else:
        numtxt = str(number)
    if desc:
        return u"{0}-{1} {2}".format(pname, numtxt, desc)
    return u"{0}-{1}".format(pname, numtxt)

@RunInTransaction("Electrical::ApplyCircuitName")
def set_circuit_name(doc, system, name):
    """
    Set the circuit 'name' parameter on an ElectricalSystem (if present).
    Tries known parameter names and built-ins, routed via params helper.
    """
    # Common names: "Circuit Name" (UI), built-in often maps as AsString via params helper.
    ok = set_param_value(system, "Circuit Name", name)
    if not ok:
        # Try a couple alternates you sometimes see in templates
        if not set_param_value(system, "Load Name", name):
            set_param_value(system, "Comments", name)
    return True

# --------- PANELS

def next_panel_name(doc, prefix="LP"):
    """
    Find the next available panel name like 'LP-1', 'LP-2', … scanning existing equipment Name/Mark.
    """
    existing = set()
    # Scan Electrical Equipment family instances
    for p in FilteredElementCollector(doc).OfCategoryId(None).ToElements():
        # Fallback: safer to just iterate all and pull Name/Mark (RPW route is fine too)
        pass  # We'll rely on a heuristic below—see collect below

    # Robust collect via names on all family instances that look like panels (cheap heuristic)
    from Autodesk.Revit.DB import BuiltInCategory, FamilyInstance, FilteredElementCollector
    candidates = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    for e in candidates:
        cat = e.Category
        if not cat: continue
        # ElectricalEquipment is panels, switchboards, etc.
        if cat.Id.IntegerValue == BuiltInCategory.OST_ElectricalEquipment.value__:
            nm = (getattr(e, "Name", None) or "").strip()
            mk = (get_param_value(e, "Mark") or "").strip()
            for t in (nm, mk):
                if not t: continue
                t = t.upper()
                if t.startswith(prefix.upper()+"-"):
                    existing.add(t)

    # Find lowest missing integer suffix
    i = 1
    while True:
        candidate = (prefix.upper() + "-" + str(i))
        if candidate not in existing:
            return candidate
        i += 1

# --------- DEVICE MARKS

def next_mark_with_prefix(doc, prefix="REC", width=3):
    """
    Return 'REC-001', 'REC-002', … lowest available project-wide.
    """
    seen = set()
    from Autodesk.Revit.DB import FilteredElementCollector
    elems = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    for e in elems:
        mk = get_param_value(e, "Mark")
        if not mk: continue
        mk = mk.strip().upper()
        if mk.startswith(prefix.upper()+"-"):
            seen.add(mk)
    n = 1
    while True:
        cand = u"{0}-{1}".format(prefix.upper(), str(n).zfill(width))
        if cand not in seen:
            return cand
        n += 1

@RunInTransaction("Electrical::ApplyDeviceMark")
def set_device_mark(doc, elem, mark_text):
    """Set the 'Mark' parameter on any element."""
    set_param_value(elem, "Mark", mark_text)
    return True