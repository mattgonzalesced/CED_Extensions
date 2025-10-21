# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import StorageType

def get_param(elem, name_or_bip):
    # BuiltInParameter enum?
    if hasattr(name_or_bip, "value__"):
        return elem.get_Parameter(name_or_bip)
    # Name lookup
    ps = elem.Parameters
    for p in ps:
        d = p.Definition
        if d and d.Name == name_or_bip: return p
    return None

def get_param_value(elem, name_or_bip):
    p = get_param(elem, name_or_bip)
    if not p or not p.HasValue: return None
    st = p.StorageType
    if st == StorageType.String:  return p.AsString()
    if st == StorageType.Integer: return p.AsInteger()
    if st == StorageType.Double:  return p.AsDouble()
    # Unit-typed or element id, give value string as fallback
    try: return p.AsValueString()
    except: return None

def set_param_value(elem, name_or_bip, value):
    p = get_param(elem, name_or_bip)
    if not p: return False
    st = p.StorageType
    try:
        if st == StorageType.String:  return p.Set(str(value)) is True or True
        if st == StorageType.Integer: return p.Set(int(value))  is True or True
        if st == StorageType.Double:  return p.Set(float(value)) is True or True
        # Unit-typed etc.
        return p.Set(value) is True or True
    except:
        return False