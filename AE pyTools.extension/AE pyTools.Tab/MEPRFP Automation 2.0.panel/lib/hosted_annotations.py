# -*- coding: utf-8 -*-
"""
Annotation classification + parameter / position extraction.

CED convention (not Revit's):
    * tags        = IndependentTag instances
    * text_notes  = TextNote instances
    * keynotes    = FamilyInstance instances of family ``GA_Keynote Symbol_CED``

The capture engine uses these helpers to:
    * decide which picked elements are fixtures vs annotations,
    * sweep auto-hosted annotations off a fixture (GetDependentElements),
    * compute world position / rotation for relative-offset math,
    * build the annotation descriptor dict that lands in YAML.

``collect_hosted_annotations`` still returns ``(tags, keynotes, notes)``
for backward compatibility with any caller mid-refactor; new code should
prefer ``annotation_kind`` + a single list.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    ElementMulticlassFilter,
    FamilyInstance,
    IndependentTag,
    LocationPoint,
    TextNote,
)
from System.Collections.Generic import List as ClrList  # noqa: E402
from System import Type  # noqa: E402


KEYNOTE_FAMILY_NAME = "GA_Keynote Symbol_CED"


# ---------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------

def _family_name(family_instance):
    sym = getattr(family_instance, "Symbol", None)
    if sym is None:
        return ""
    fam = getattr(sym, "Family", None)
    return getattr(fam, "Name", "") or ""


def annotation_kind(elem):
    """``'tag' | 'keynote' | 'text_note' | None``."""
    if isinstance(elem, IndependentTag):
        return "tag"
    if isinstance(elem, TextNote):
        return "text_note"
    if isinstance(elem, FamilyInstance):
        if _family_name(elem) == KEYNOTE_FAMILY_NAME:
            return "keynote"
    return None


def is_annotation_element(elem):
    return annotation_kind(elem) is not None


# ---------------------------------------------------------------------
# Tag-target resolution (for matching picked tags to picked fixtures)
# ---------------------------------------------------------------------

def tag_target_element_ids(tag_elem):
    """Return host-doc ElementIds the tag references. Empty list if none."""
    if not isinstance(tag_elem, IndependentTag):
        return []
    out = []
    # Modern API (Revit 2022+)
    try:
        for eid in tag_elem.GetTaggedLocalElementIds():
            out.append(eid)
    except Exception:
        pass
    # Fallback to deprecated TaggedElementId (LinkElementId-typed)
    if not out:
        try:
            lei = tag_elem.TaggedElementId
            if lei is not None:
                host_id = lei.HostElementId
                if host_id is not None and host_id != ElementId.InvalidElementId:
                    out.append(host_id)
        except Exception:
            pass
    return out


# ---------------------------------------------------------------------
# Position / rotation
# ---------------------------------------------------------------------

def annotation_world_point(elem):
    """Return ``(x, y, z)`` in feet, or None.

    For tags we use ``TagHeadPosition``; for text notes ``Coord``; for
    keynote symbols (FamilyInstance) we fall through to the standard
    ``Location.Point``.
    """
    if elem is None:
        return None
    if isinstance(elem, IndependentTag):
        try:
            p = elem.TagHeadPosition
            if p is not None:
                return (p.X, p.Y, p.Z)
        except Exception:
            pass
    if isinstance(elem, TextNote):
        try:
            p = elem.Coord
            if p is not None:
                return (p.X, p.Y, p.Z)
        except Exception:
            pass
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Point"):
        try:
            p = loc.Point
            if p is not None:
                return (p.X, p.Y, p.Z)
        except Exception:
            pass
    return None


def annotation_rotation_deg(elem):
    """Rotation in degrees, or 0.0."""
    if elem is None:
        return 0.0
    # IndependentTag and TextNote both expose a RotationAngle (radians).
    rot = getattr(elem, "RotationAngle", None)
    if rot is not None:
        try:
            return math.degrees(float(rot))
        except (TypeError, ValueError):
            pass
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            return 0.0
    return 0.0


# ---------------------------------------------------------------------
# Parameter sweep (for fixture LEDs and annotations)
# ---------------------------------------------------------------------

def collect_element_parameters(elem):
    """Return ``{name: value_string}`` for every parameter on ``elem``.

    Empty / no-value parameters (e.g. ``CKT_*`` before circuiting)
    appear with an empty-string value so they show up in the editor —
    the user wants to see the full parameter set, not just the
    populated ones.
    """
    out = {}
    if elem is None:
        return out
    try:
        params_iter = elem.Parameters
    except Exception:
        return out
    for p in params_iter:
        if p is None:
            continue
        try:
            name = p.Definition.Name
        except Exception:
            continue
        if not name or name in out:
            continue
        value = None
        try:
            value = p.AsValueString()
        except Exception:
            value = None
        if value is None:
            try:
                value = p.AsString()
            except Exception:
                value = None
        out[name] = "" if value is None else str(value)
    return out


# ---------------------------------------------------------------------
# Descriptor builders
# ---------------------------------------------------------------------

def annotation_descriptor(elem, kind=None):
    """Build the dict that becomes one ``annotations[*]`` entry.

    Caller fills in ``id`` and ``offsets`` (relative to the host fixture).

    For ``text_note`` kind we also store the text content as a top-level
    ``text`` field — that's what Place Element Annotations needs to
    recreate the text note on placement.
    """
    if elem is None:
        return None
    if kind is None:
        kind = annotation_kind(elem) or "tag"

    type_id = elem.GetTypeId()
    type_elem = elem.Document.GetElement(type_id) if type_id else None
    family_name = ""
    type_name = ""
    category_name = ""
    if type_elem is not None:
        family_name = getattr(type_elem, "FamilyName", "") or ""
        type_name = getattr(type_elem, "Name", "") or ""
    if elem.Category is not None:
        category_name = elem.Category.Name or ""

    text_content = ""
    if kind == "text_note":
        try:
            text_content = (elem.Text or "").strip()
        except Exception:
            text_content = ""

    # For text notes, the *content* is the most useful identity — show
    # that as the label, not the TextNoteType string. The full text is
    # also stored in a top-level ``text`` field for placement.
    if kind == "text_note":
        label = (text_content[:60] + "...") if len(text_content) > 60 else text_content
        if not label:
            label = "(empty text note)"
    elif family_name and type_name:
        label = "{} : {}".format(family_name, type_name)
    else:
        label = category_name or kind

    descriptor = {
        "kind": kind,
        "label": label,
        "category_name": category_name,
        "family_name": family_name,
        "type_name": type_name,
        "parameters": collect_element_parameters(elem),
    }
    if kind == "text_note":
        descriptor["text"] = text_content
    return descriptor


# Legacy alias kept for any external callers that still import this name.
def tag_descriptor(elem):
    return annotation_descriptor(elem)


# ---------------------------------------------------------------------
# Auto-sweep of hosted annotations off a fixture (GetDependentElements)
# ---------------------------------------------------------------------

def _multiclass_filter():
    types = ClrList[Type]()
    types.Add(clr.GetClrType(IndependentTag))
    types.Add(clr.GetClrType(TextNote))
    types.Add(clr.GetClrType(FamilyInstance))
    return ElementMulticlassFilter(types)


def collect_hosted_dependents(host_elem):
    """Return a list of ``(elem, kind)`` for annotations that depend on
    ``host_elem``. Only true dependents (face-hosted, attached, tagged)
    are returned.
    """
    if host_elem is None:
        return []
    doc = host_elem.Document
    try:
        ids = host_elem.GetDependentElements(_multiclass_filter())
    except Exception:
        return []
    out = []
    for eid in ids:
        elem = doc.GetElement(eid)
        if elem is None:
            continue
        kind = annotation_kind(elem)
        if kind is None:
            continue
        out.append((elem, kind))
    return out


def collect_hosted_annotations(host_elem):
    """Legacy three-list shape kept for any caller mid-refactor."""
    deps = collect_hosted_dependents(host_elem)
    tags, keynotes, notes = [], [], []
    for elem, kind in deps:
        if kind == "tag":
            tags.append(elem)
        elif kind == "keynote":
            keynotes.append(elem)
        elif kind == "text_note":
            notes.append(elem)
    return tags, keynotes, notes
