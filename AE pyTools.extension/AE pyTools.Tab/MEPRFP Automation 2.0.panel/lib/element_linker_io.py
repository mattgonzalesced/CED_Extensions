# -*- coding: utf-8 -*-
"""
Read and write the Element_Linker JSON payload on a Revit element.

Reads:
    payload = read_from_element(elem)
    payload is None when no Element_Linker parameter exists or it's blank.

Writes:
    write_to_element(elem, linker)   # caller manages the Revit transaction

Pure-Python ``element_linker.py`` defines the codec; this module is the
Revit-API edge that touches actual element parameters.
"""

import clr  # noqa: F401

from element_linker import ElementLinker, ElementLinkerError, PARAMETER_NAME


class ElementLinkerIOError(Exception):
    pass


def _lookup_param(elem):
    if elem is None:
        return None
    return elem.LookupParameter(PARAMETER_NAME)


def has_element_linker_param(elem):
    return _lookup_param(elem) is not None


def read_from_element(elem):
    """Return ``ElementLinker`` or ``None`` if no payload."""
    param = _lookup_param(elem)
    if param is None:
        return None
    text = param.AsString()
    if not text or not text.strip():
        return None
    try:
        return ElementLinker.from_json(text)
    except ElementLinkerError:
        # Legacy bespoke text payload — try the migration path so
        # 2.0 readers don't choke on a project that still has old data.
        try:
            return ElementLinker.from_legacy_text(text)
        except Exception:
            return None


def write_to_element(elem, linker):
    """Set the Element_Linker parameter to the JSON encoding of ``linker``.

    Raises ``ElementLinkerIOError`` if the parameter is absent or read-only.
    Caller must have an open Revit transaction.
    """
    param = _lookup_param(elem)
    if param is None:
        raise ElementLinkerIOError(
            "Element {} has no '{}' parameter — bind the shared "
            "parameter to this category first.".format(elem.Id, PARAMETER_NAME)
        )
    if param.IsReadOnly:
        raise ElementLinkerIOError(
            "Element {} has a read-only '{}' parameter.".format(
                elem.Id, PARAMETER_NAME
            )
        )
    text = linker.to_json() if linker is not None else ""
    param.Set(text)


def clear_on_element(elem):
    """Set the Element_Linker parameter to an empty string."""
    param = _lookup_param(elem)
    if param is None or param.IsReadOnly:
        return False
    param.Set("")
    return True
