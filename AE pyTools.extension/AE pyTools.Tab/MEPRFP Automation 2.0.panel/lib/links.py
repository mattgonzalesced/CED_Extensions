# -*- coding: utf-8 -*-
"""
Linked-document traversal and transform-chain helpers.

This module imports the Revit API and is only usable inside a pyRevit
session. ``geometry.py`` does the math; this module bridges to the
Revit-side primitives (XYZ, Transform, RevitLinkInstance) and traverses
nested links with cycle detection.

Cycle detection uses a document key (``PathName`` or ``Title``). When a
link instance points back to a document already in the chain, that
sub-tree is skipped.
"""

import clr  # noqa: F401  -- needed before importing Autodesk.Revit.DB

from Autodesk.Revit.DB import (  # noqa: E402
    FilteredElementCollector,
    RevitLinkInstance,
    Transform,
    XYZ,
)


# ---------------------------------------------------------------------
# XYZ <-> tuple bridges
# ---------------------------------------------------------------------

def xyz_to_tuple(point):
    if point is None:
        return None
    return (point.X, point.Y, point.Z)


def tuple_to_xyz(t):
    if t is None:
        return None
    return XYZ(float(t[0]), float(t[1]), float(t[2]))


# ---------------------------------------------------------------------
# Transforms
# ---------------------------------------------------------------------

def get_link_transform(link_instance):
    """Prefer ``GetTotalTransform`` (honours survey-point shifts);
    fall back to ``GetTransform`` for older Revit builds."""
    if link_instance is None:
        return None
    try:
        return link_instance.GetTotalTransform()
    except Exception:
        pass
    try:
        return link_instance.GetTransform()
    except Exception:
        return None


def compose(parent_transform, child_transform):
    """Compose two transforms via ``Multiply``. Either side may be None."""
    if parent_transform is None:
        return child_transform
    if child_transform is None:
        return parent_transform
    return parent_transform.Multiply(child_transform)


def transform_point_tuple(transform, point_tuple):
    """Apply a Transform to a 3-tuple and return a new 3-tuple."""
    if point_tuple is None:
        return None
    if transform is None:
        return tuple(point_tuple)
    p = transform.OfPoint(tuple_to_xyz(point_tuple))
    return (p.X, p.Y, p.Z)


def transform_vector_tuple(transform, vector_tuple):
    if vector_tuple is None:
        return None
    if transform is None:
        return tuple(vector_tuple)
    v = transform.OfVector(tuple_to_xyz(vector_tuple))
    return (v.X, v.Y, v.Z)


# ---------------------------------------------------------------------
# Document traversal
# ---------------------------------------------------------------------

def _doc_key(doc):
    if doc is None:
        return None
    try:
        path = getattr(doc, "PathName", None)
        if path:
            return path
    except Exception:
        pass
    try:
        title = getattr(doc, "Title", None)
        if title:
            return title
    except Exception:
        pass
    return id(doc)


def iter_link_documents(doc, include_root=False):
    """Yield ``(link_doc, total_transform)`` for every linked document
    reachable from ``doc``.

    Depth-first traversal. Cycles are skipped via ``_doc_key``. The
    ``total_transform`` composes every Revit link transform from the
    root down to the yielded document, so calling
    ``transform.OfPoint(p_in_link)`` produces a point in the root
    document's coordinate frame.

    If ``include_root`` is True, ``(doc, identity)`` is yielded first.
    """
    if doc is None:
        return

    if include_root:
        yield doc, Transform.Identity

    seen = {_doc_key(doc)}

    def _walk(parent_doc, parent_transform):
        try:
            collector = FilteredElementCollector(parent_doc).OfClass(RevitLinkInstance)
        except Exception:
            return
        for link_inst in collector:
            link_doc = None
            try:
                link_doc = link_inst.GetLinkDocument()
            except Exception:
                pass
            if link_doc is None:
                continue
            key = _doc_key(link_doc)
            if key in seen:
                continue
            seen.add(key)
            child_transform = get_link_transform(link_inst)
            total = compose(parent_transform, child_transform)
            yield link_doc, total
            for nested in _walk(link_doc, total):
                yield nested

    for entry in _walk(doc, Transform.Identity):
        yield entry
