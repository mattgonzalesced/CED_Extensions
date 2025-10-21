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

def _get_bip(sp, bip_name):
    """Lazy BuiltInParameter access; avoids import at module import time."""
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        bip = getattr(BuiltInParameter, bip_name, None)
        if not bip:
            return None
        p = sp.get_Parameter(bip)
        return p.AsString() if p else None
    except Exception:
        return None

def _first_non_empty(*vals):
    for v in vals:
        if v:
            v = u"{}".format(v).strip()
            if v:
                return v
    return u""

def space_display_name(sp):
    """Best-effort human label for a Space/Room."""
    # Try common properties (works on Room too)
    name = _first_non_empty(getattr(sp, "Name", None),
                            _get_bip(sp, "SPACE_NAME"),
                            _get_bip(sp, "ROOM_NAME"))
    number = _first_non_empty(_get_bip(sp, "SPACE_NUMBER"),
                              _get_bip(sp, "ROOM_NUMBER"))
    if number and name:
        return u"{} {}".format(number, name)
    return name or number or u""

def space_match_text(sp):
    """Concatenate useful fields for rule matching (lowercased)."""
    name  = space_display_name(sp)
    dept  = _first_non_empty(_get_bip(sp, "ROOM_DEPARTMENT"),
                             _get_bip(sp, "SPACE_DEPARTMENT"))
    occ   = _first_non_empty(_get_bip(sp, "ROOM_OCCUPANCY"),
                             _get_bip(sp, "SPACE_OCCUPANCY"))
    lvl   = _first_non_empty(_get_bip(sp, "LEVEL_NAME"))
    num   = _first_non_empty(_get_bip(sp, "SPACE_NUMBER"),
                             _get_bip(sp, "ROOM_NUMBER"))
    parts = [name, dept, occ, lvl, num]
    txt = u" ".join([p for p in parts if p]).strip().lower()
    # normalize spaces
    return u" ".join(txt.split())

def _compile_regex(rx_text):
    try:
        import re
        return re.compile(rx_text, re.IGNORECASE)
    except Exception:
        return None

def _iter_rule_candidates(identify_rules):
    """
    Be flexible about JSON shape. Accepts either:
      - {"categories":[{"name":"Sales Floor","regex":["sales\\s*floor", ...], "contains_any":[...]} , ...]}
      - {"Sales Floor": {"regex":[...], ...}, "Back Office": {...}, ...}
    Yields (category_name, rule_dict).
    """
    if not identify_rules:
        return
    cats = identify_rules.get("categories")
    if isinstance(cats, list):
        for item in cats:
            cname = (item.get("name") or u"").strip()
            if cname:
                yield cname, item
    else:
        # dict mapping
        for cname, item in identify_rules.items():
            if isinstance(item, dict):
                yield u"{}".format(cname).strip(), item

def _rule_matches_text(rule_dict, text_lc):
    """Support 'regex', 'name_regex', 'number_regex', 'department_regex', 'contains', 'contains_any'."""
    import re
    # Plain contains (any)
    for key in ("contains", "contains_any", "any"):
        vals = rule_dict.get(key)
        if isinstance(vals, (list, tuple)):
            for v in vals:
                if v and u"{}".format(v).strip().lower() in text_lc:
                    return True

    # Regex buckets
    rx_keys = ("regex", "name_regex", "number_regex", "department_regex")
    for key in rx_keys:
        vals = rule_dict.get(key)
        if isinstance(vals, (list, tuple)):
            for pat in vals:
                rx = _compile_regex(u"{}".format(pat))
                if rx and rx.search(text_lc):
                    return True

    # Single-string convenience
    for key in rx_keys:
        val = rule_dict.get(key)
        if isinstance(val, basestring):
            rx = _compile_regex(val)
            if rx and rx.search(text_lc):
                return True

    return False

def space_category_string(sp, identify_rules=None):
    """
    Return category label for a Space/Room.
    - If identify_rules given, pick first category whose rule matches space_match_text(sp).
    - Otherwise, fallback to the space/room name.
    """
    txt = space_match_text(sp)
    if identify_rules:
        for cat_name, rule_dict in _iter_rule_candidates(identify_rules):
            try:
                if _rule_matches_text(rule_dict, txt):
                    return u"{}".format(cat_name)
            except Exception:
                # ignore malformed rule
                pass

    # Fallback: just the (number +) name
    return space_display_name(sp)

# Optional: tiny helper to normalize to lowercase consistently (for pair comparisons)
def normalized_category(sp, identify_rules=None):
    return (space_category_string(sp, identify_rules) or u"").strip().lower()

@RunInTransaction("Electrical::ApplyDeviceMark")
def set_device_mark(doc, elem, mark_text):
    """Set the 'Mark' parameter on any element."""
    set_param_value(elem, "Mark", mark_text)
    return True