# -*- coding: utf-8 -*-
"""
Hide Existing Profiles — duplicate the active view, then hide every
linked-doc element matching one of the chosen profiles' parent_filter
rules in the duplicate. The user is switched to the duplicated view so
they can work without the matched elements visible. Closing the
duplicated view restores the original visibility automatically.

Matching reuses ``placement.profile_family_names`` so the same
suffix-strip + alias rules apply as in the placement matcher.
"""

import datetime
import math  # noqa: F401  -- kept for future use

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    LinkElementId,
    RevitLinkInstance,
    View,
    ViewDuplicateOption,
)
from System.Collections.Generic import List as ClrList  # noqa: E402

import placement


class HideError(Exception):
    pass


# ---------------------------------------------------------------------
# Match
# ---------------------------------------------------------------------

def _strip_suffix_lower(value):
    return placement.normalize_name(value or "")


def collect_targets(doc, profile_data, profile_ids=None, categories=None):
    """Walk every loaded link and return the LinkElementIds + plain
    host-doc element ids whose family name matches any included
    profile.

    Returns ``(link_element_ids, host_element_ids)``.
    """
    profiles = list(profile_data.get("equipment_definitions") or [])
    if profile_ids:
        profiles = [p for p in profiles if p.get("id") in profile_ids]
    if categories:
        profiles = [
            p for p in profiles
            if (p.get("parent_filter") or {}).get("category") in categories
        ]
    if not profiles:
        return [], []

    # Build the union of family-name keys across all included profiles.
    family_keys = set()
    for p in profiles:
        family_keys.update(placement.profile_family_names(p))
    if not family_keys:
        return [], []

    link_pairs = []
    host_ids = []

    # Linked elements.
    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = None
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            continue
        if link_doc is None:
            continue
        for klass in (FamilyInstance, Group):
            try:
                col = (
                    FilteredElementCollector(link_doc)
                    .OfClass(klass)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue
            for elem in col:
                fam = _element_family(elem)
                if not fam:
                    continue
                if _strip_suffix_lower(fam) in family_keys:
                    link_pairs.append((link_inst.Id, elem.Id))

    # Host elements (in case the user wants those hidden too — fairly
    # uncommon but consistent with the legacy behaviour).
    for klass in (FamilyInstance, Group):
        col = (
            FilteredElementCollector(doc)
            .OfClass(klass)
            .WhereElementIsNotElementType()
        )
        for elem in col:
            fam = _element_family(elem)
            if not fam:
                continue
            if _strip_suffix_lower(fam) in family_keys:
                host_ids.append(elem.Id)

    return link_pairs, host_ids


def _element_family(elem):
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        if sym is not None and sym.Family is not None:
            return sym.Family.Name or ""
    if isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        return gtype.Name if gtype else ""
    return ""


# ---------------------------------------------------------------------
# Duplicate view + hide
# ---------------------------------------------------------------------

def duplicate_active_view(doc, source_view):
    """Duplicate ``source_view`` (with detailing) and rename to a
    timestamped temp view. Caller manages the transaction. Returns the
    new ``View`` element."""
    if source_view is None:
        raise HideError("No active view to duplicate")
    try:
        new_id = source_view.Duplicate(ViewDuplicateOption.WithDetailing)
    except Exception:
        try:
            new_id = source_view.Duplicate(ViewDuplicateOption.Duplicate)
        except Exception as exc:
            raise HideError("Failed to duplicate view: {}".format(exc))
    new_view = doc.GetElement(new_id)
    if new_view is None:
        raise HideError("Duplicated view came back null")
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    base_name = source_view.Name
    new_name = "MEPRFP Hidden {} ({})".format(base_name, stamp)
    try:
        new_view.Name = new_name
    except Exception:
        pass
    return new_view


def hide_in_view(doc, view, link_pairs, host_ids):
    """Hide ``host_ids`` and ``link_pairs`` in ``view``. Caller manages
    the transaction. Returns ``(host_count, link_count, warnings_list)``.
    """
    warnings = []
    host_count = 0
    link_count = 0

    if host_ids:
        host_collection = ClrList[ElementId](host_ids)
        try:
            view.HideElements(host_collection)
            host_count = host_collection.Count
        except Exception as exc:
            warnings.append("HideElements (host) failed: {}".format(exc))

    if link_pairs:
        # Try the LinkElementId overload first (Revit 2024+); fall back
        # to host-instance hiding (less granular but available everywhere).
        link_eids = ClrList[LinkElementId]()
        for link_inst_id, linked_elem_id in link_pairs:
            try:
                link_eids.Add(LinkElementId(link_inst_id, linked_elem_id))
            except Exception:
                pass
        try:
            view.HideElements(link_eids)
            link_count = link_eids.Count
        except Exception:
            # Fallback: hide the link instances themselves. Coarse — it
            # hides the whole link rather than just matched elements —
            # but better than a hard failure.
            warnings.append(
                "View.HideElements doesn't accept LinkElementId here. "
                "Hiding the parent link instance(s) instead — entire "
                "link visibility is affected."
            )
            link_inst_ids = {pair[0] for pair in link_pairs}
            inst_list = ClrList[ElementId]([eid for eid in link_inst_ids])
            try:
                view.HideElements(inst_list)
                link_count = len(link_pairs)
            except Exception as exc:
                warnings.append(
                    "Fallback HideElements failed too: {}".format(exc)
                )
    return host_count, link_count, warnings
