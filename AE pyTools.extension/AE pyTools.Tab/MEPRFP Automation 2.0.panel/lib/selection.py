# -*- coding: utf-8 -*-
"""
Selection helpers for the MEPRFP 2.0 capture flow.

Wraps ``UIDocument.Selection.PickObject`` for picking parents and
children, and resolves linked-model picks back to the linked element +
its link instance + total transform. Callers receive ``ParentRef`` and
``ChildRef`` objects with everything needed to compute offsets.
"""

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    FamilyInstance,
    Group,
    RevitLinkInstance,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ObjectType  # noqa: E402
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402

import links


class SelectionCancelled(Exception):
    pass


class ElementRef(object):
    """An element resolved to its host-document coordinate frame.

    ``element``     the Revit element (FamilyInstance, Group, ...)
    ``host_doc``    the document the element lives in
    ``link_inst``   RevitLinkInstance if linked, else None
    ``transform``   total transform from the element's doc to the host doc
                    (Identity if element is in the host doc itself)
    """

    def __init__(self, element, host_doc, link_inst, transform):
        self.element = element
        self.host_doc = host_doc
        self.link_inst = link_inst
        self.transform = transform

    @property
    def is_linked(self):
        return self.link_inst is not None

    @property
    def element_id_value(self):
        eid = self.element.Id
        return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _resolve_reference(reference, doc):
    """Convert a Revit ``Reference`` to an ``ElementRef``."""
    if reference is None:
        return None
    link_inst_id = getattr(reference, "LinkedElementId", None)
    if link_inst_id is None or link_inst_id == ElementId.InvalidElementId:
        # Host element
        elem = doc.GetElement(reference.ElementId)
        from Autodesk.Revit.DB import Transform
        return ElementRef(elem, doc, None, Transform.Identity)

    # Linked element. reference.ElementId is the link instance; LinkedElementId is the
    # ID of the element within the linked doc.
    link_inst = doc.GetElement(reference.ElementId)
    if not isinstance(link_inst, RevitLinkInstance):
        elem = doc.GetElement(reference.ElementId)
        from Autodesk.Revit.DB import Transform
        return ElementRef(elem, doc, None, Transform.Identity)
    link_doc = link_inst.GetLinkDocument()
    if link_doc is None:
        raise SelectionCancelled("Linked document is not loaded.")
    elem = link_doc.GetElement(reference.LinkedElementId)
    transform = links.get_link_transform(link_inst)
    return ElementRef(elem, doc, link_inst, transform)


def pick_parent(uidoc, prompt="Pick parent element", from_linked=False):
    """Prompt the user for one parent element.

    ``from_linked`` chooses between ``ObjectType.Element`` (host model — the
    default) and ``ObjectType.LinkedElement`` (an element inside a loaded
    linked document). The two pick modes are mutually exclusive in Revit's
    PickObject API, so callers ask the user up front which one they want.
    """
    object_type = ObjectType.LinkedElement if from_linked else ObjectType.Element
    try:
        ref = uidoc.Selection.PickObject(object_type, prompt)
    except OperationCanceledException:
        raise SelectionCancelled("User cancelled parent pick")
    return _resolve_reference(ref, uidoc.Document)


def pick_children(uidoc, prompt="Pick child elements (Finish to commit)",
                  from_linked=False):
    """Prompt for any number of children.

    Defaults to host-model picking. Pass ``from_linked=True`` to switch
    the selector into linked-element mode. The two modes are mutually
    exclusive in Revit's PickObjects API (same constraint as
    ``pick_parent``).
    """
    object_type = ObjectType.LinkedElement if from_linked else ObjectType.Element
    try:
        refs = uidoc.Selection.PickObjects(object_type, prompt)
    except OperationCanceledException:
        raise SelectionCancelled("User cancelled child pick")
    if refs is None:
        return []
    out = []
    for r in refs:
        try:
            out.append(_resolve_reference(r, uidoc.Document))
        except Exception:
            continue
    return out


def is_capturable(elem):
    """Cheap predicate: element is something we can capture as a child."""
    return isinstance(elem, FamilyInstance) or isinstance(elem, Group)
