# -*- coding: utf-8 -*-
"""Bounding-box helpers for Circuit Element Finder."""

from pyrevit import DB
from pyrevit import script

try:
    from System.Collections.Generic import List
except Exception:
    List = None


def get_combined_bounding_box(doc, element_ids, view=None):
    """Return a combined BoundingBoxXYZ for the given elements."""
    min_x = None
    min_y = None
    min_z = None
    max_x = None
    max_y = None
    max_z = None

    for element_id in list(element_ids or []):
        try:
            element = doc.GetElement(element_id)
        except Exception:
            element = None
        if element is None:
            continue

        bbox = None
        try:
            bbox = element.get_BoundingBox(view)
        except Exception:
            bbox = None
        if bbox is None and view is not None:
            try:
                bbox = element.get_BoundingBox(None)
            except Exception:
                bbox = None
        if bbox is None or bbox.Min is None or bbox.Max is None:
            continue

        bmin = bbox.Min
        bmax = bbox.Max

        if min_x is None:
            min_x = bmin.X
            min_y = bmin.Y
            min_z = bmin.Z
            max_x = bmax.X
            max_y = bmax.Y
            max_z = bmax.Z
            continue

        min_x = min(min_x, bmin.X)
        min_y = min(min_y, bmin.Y)
        min_z = min(min_z, bmin.Z)
        max_x = max(max_x, bmax.X)
        max_y = max(max_y, bmax.Y)
        max_z = max(max_z, bmax.Z)

    if min_x is None:
        return None

    combined = DB.BoundingBoxXYZ()
    combined.Min = DB.XYZ(min_x, min_y, min_z)
    combined.Max = DB.XYZ(max_x, max_y, max_z)
    return combined


def expand_bounding_box(box, padding_feet=3.0, minimum_half_extent=0.5):
    """Return expanded BoundingBoxXYZ with a minimum non-zero extent."""
    if box is None or box.Min is None or box.Max is None:
        return None

    try:
        pad = float(padding_feet)
    except Exception:
        pad = 3.0
    if pad < 0.0:
        pad = 0.0

    try:
        minimum_half = float(minimum_half_extent)
    except Exception:
        minimum_half = 0.5
    if minimum_half < 0.01:
        minimum_half = 0.01

    min_x = float(box.Min.X)
    min_y = float(box.Min.Y)
    min_z = float(box.Min.Z)
    max_x = float(box.Max.X)
    max_y = float(box.Max.Y)
    max_z = float(box.Max.Z)

    center_x = (min_x + max_x) * 0.5
    center_y = (min_y + max_y) * 0.5
    center_z = (min_z + max_z) * 0.5
    half_x = max((max_x - min_x) * 0.5, minimum_half)
    half_y = max((max_y - min_y) * 0.5, minimum_half)
    half_z = max((max_z - min_z) * 0.5, minimum_half)

    expanded = DB.BoundingBoxXYZ()
    expanded.Min = DB.XYZ(center_x - half_x - pad, center_y - half_y - pad, center_z - half_z - pad)
    expanded.Max = DB.XYZ(center_x + half_x + pad, center_y + half_y + pad, center_z + half_z + pad)
    return expanded


def to_element_id_list(element_ids):
    """Return a .NET List[ElementId] when available, else a Python list."""
    cleaned = [x for x in list(element_ids or []) if isinstance(x, DB.ElementId)]
    if List is None:
        return cleaned
    try:
        dotnet_ids = List[DB.ElementId]()
        for element_id in cleaned:
            dotnet_ids.Add(element_id)
        return dotnet_ids
    except Exception:
        return cleaned


def show_elements(uidoc, element_ids, logger=None):
    """Best-effort zoom/focus helper using UIDocument.ShowElements."""
    ids = to_element_id_list(element_ids)
    if not ids:
        return False
    try:
        uidoc.ShowElements(ids)
        return True
    except Exception as ex:
        log = logger or script.get_logger()
        try:
            log.debug("ShowElements failed: {0}".format(ex))
        except Exception:
            pass
        return False


def zoom_view_to_bounds(uidoc, view_id, box, logger=None):
    """Zoom target UI view to bounding box extents without switching views."""
    if uidoc is None or view_id is None or box is None or box.Min is None or box.Max is None:
        return False
    try:
        open_views = list(uidoc.GetOpenUIViews() or [])
    except Exception:
        open_views = []
    for ui_view in open_views:
        try:
            if ui_view.ViewId != view_id:
                continue
            ui_view.ZoomAndCenterRectangle(box.Min, box.Max)
            return True
        except Exception as ex:
            log = logger or script.get_logger()
            try:
                log.debug("ZoomAndCenterRectangle failed: {0}".format(ex))
            except Exception:
                pass
            return False
    return False
