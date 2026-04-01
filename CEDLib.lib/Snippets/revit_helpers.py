# -*- coding: utf-8 -*-
"""Small Revit API helper utilities shared by circuit tools."""

import Autodesk.Revit.DB as DB
from System import Int64


def get_elementid_value(item, default=0):
    """Return an ElementId numeric value across Revit API versions."""
    if item is None:
        return int(default or 0)
    try:
        return int(getattr(item, "Value"))
    except Exception:
        pass
    try:
        return int(getattr(item, "IntegerValue"))
    except Exception:
        return int(default or 0)


def elementid_from_value(value):
    """Create a Revit ElementId from an integer-like value."""
    numeric = int(value or 0)
    try:
        return DB.ElementId(Int64(numeric))
    except Exception:
        return DB.ElementId(numeric)


def get_type_element(element, doc=None):
    """Return an element type for an instance, or None when unavailable."""
    if element is None:
        return None
    if doc is None:
        try:
            doc = element.Document
        except Exception:
            doc = None
    if doc is None:
        return None
    try:
        type_id = element.GetTypeId()
    except Exception:
        return None
    if not type_id or type_id == DB.ElementId.InvalidElementId:
        return None
    try:
        return doc.GetElement(type_id)
    except Exception:
        return None


def get_parameter(element, name, include_type=False, case_insensitive=True, doc=None):
    """Return a parameter by name from instance, optionally falling back to type."""
    if element is None or not name:
        return None
    target_name = str(name)

    def _lookup(owner):
        if owner is None:
            return None
        try:
            param = owner.LookupParameter(target_name)
            if param is not None:
                return param
        except Exception:
            pass
        if not case_insensitive:
            return None
        try:
            for candidate in owner.Parameters:
                try:
                    definition = candidate.Definition
                    if definition and str(definition.Name).strip().lower() == target_name.strip().lower():
                        return candidate
                except Exception:
                    continue
        except Exception:
            pass
        return None

    param = _lookup(element)
    if param is not None or not include_type:
        return param
    return _lookup(get_type_element(element, doc=doc))


def get_parameter_value(parameter, default=None):
    """Return a best-effort native Python value for a Revit Parameter."""
    if parameter is None:
        return default
    try:
        storage_type = parameter.StorageType
    except Exception:
        return default
    try:
        if storage_type == DB.StorageType.String:
            value = parameter.AsString()
            if value is None:
                value = parameter.AsValueString()
            return value if value is not None else default
        if storage_type == DB.StorageType.Integer:
            return parameter.AsInteger()
        if storage_type == DB.StorageType.Double:
            return parameter.AsDouble()
        if storage_type == DB.StorageType.ElementId:
            return parameter.AsElementId()
    except Exception:
        return default
    return default


def get_parameter_text(element, name, include_type=False, case_insensitive=True, doc=None, default=""):
    """Return a parameter value as text from instance or type."""
    param = get_parameter(
        element,
        name,
        include_type=include_type,
        case_insensitive=case_insensitive,
        doc=doc,
    )
    value = get_parameter_value(param, default=None)
    if value is None:
        return default
    try:
        return str(value)
    except Exception:
        return default


def get_family_symbol_name(element, doc=None, fallback=""):
    """Return a family/type display name for an element."""
    if element is None:
        return fallback
    type_element = get_type_element(element, doc=doc)
    for candidate in (type_element, element):
        if candidate is None:
            continue
        try:
            name = getattr(candidate, "Name", None)
            if name:
                return str(name)
        except Exception:
            pass
        try:
            family_name = get_parameter_text(
                candidate,
                "Type Name",
                include_type=False,
                case_insensitive=True,
                doc=doc,
                default="",
            )
            if family_name:
                return family_name
        except Exception:
            pass
    return fallback
