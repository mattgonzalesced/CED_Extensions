# -*- coding: utf-8 -*-
"""ISelectionFilter implementation that accepts only Door-category
elements (host or linked).

This module is **deliberately separated** from ``space_door_picker``
so the picker module can be reloaded by ``_dev_reload.purge()`` during
iterative development without re-registering the CLR type defined here.
pythonnet 3 registers ``ISelectionFilter`` subclasses as generated CLR
types at class-statement execution time; re-running that statement
under a hot-reload throws "Duplicate type name within an assembly".

Keeping the class in its own module — and excluding *only* this module
from the purge list — means the picker's flow code stays purgeable
while the filter survives across reloads by design.
"""

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    ElementId,
    RevitLinkInstance,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter  # noqa: E402


def _is_door(elem):
    """True iff ``elem`` is a Door-category element."""
    if elem is None:
        return False
    try:
        cat = elem.Category
    except Exception:
        return False
    if cat is None:
        return False
    try:
        cat_id = cat.Id
        cat_int = getattr(cat_id, "Value", None) or getattr(
            cat_id, "IntegerValue", None,
        )
        return int(cat_int) == int(BuiltInCategory.OST_Doors)
    except Exception:
        return False


class SpaceOnlyFilter(ISelectionFilter):
    """Restricts ``PickObject`` to MEP Spaces.

    Lives here (next to ``DoorOnlyFilter``) for the same reason — the
    CLR-registered ISelectionFilter subclass survives across script
    reloads because this module is excluded from
    ``_dev_reload.purge()``. Defining it at script top-level would
    crash with "Duplicate type name within an assembly" on the
    second run.
    """

    __namespace__ = "MEPRFP.Automation.SpacePickFilters.SpaceOnly"

    def AllowElement(self, element):
        cat = getattr(element, "Category", None)
        if cat is None:
            return False
        try:
            cat_id = cat.Id
            cat_int = getattr(cat_id, "Value", None) or getattr(
                cat_id, "IntegerValue", None,
            )
            return int(cat_int) == int(BuiltInCategory.OST_MEPSpaces)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return True


class DoorOnlyFilter(ISelectionFilter):
    """Picks only Doors; works for both host and linked elements.

    The trick for linked picks: when ``PickObject(ObjectType.LinkedElement,
    ...)`` is in play, ``AllowElement`` is called with the
    ``RevitLinkInstance`` (whose Category is "RVT Links"), NOT with
    the door inside the link. A naive door-category check in
    ``AllowElement`` therefore rejects every link instance and the
    user can't click any linked door. So we allow ``RevitLinkInstance``
    in ``AllowElement`` and do the actual Door-category check in
    ``AllowReference``, where we have the linked element id and can
    resolve it through the link document.
    """

    # Required by pythonnet 3 so the filter registers as a proper CLR
    # type. Without this, PickObject errors out with "object does not
    # implement ISelectionFilter".
    __namespace__ = "MEPRFP.Automation.SpaceDoorPicker"

    def __init__(self, doc):
        self._doc = doc

    def AllowElement(self, element):
        if isinstance(element, RevitLinkInstance):
            return True
        return _is_door(element)

    def AllowReference(self, reference, position):
        try:
            linked_id = reference.LinkedElementId
        except Exception:
            return True
        if linked_id is None or linked_id == ElementId.InvalidElementId:
            return True
        try:
            host_id = reference.ElementId
        except Exception:
            return False
        link_inst = self._doc.GetElement(host_id)
        if not isinstance(link_inst, RevitLinkInstance):
            return False
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            return False
        elem = link_doc.GetElement(linked_id)
        return _is_door(elem)
