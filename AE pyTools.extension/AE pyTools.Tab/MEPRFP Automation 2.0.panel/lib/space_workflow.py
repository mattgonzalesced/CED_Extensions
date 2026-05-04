# -*- coding: utf-8 -*-
"""
Stage 6 — Spaces orchestration.

Bridges the Revit-API edge (collecting Spaces, reading their name and
number) and the pure-Python pieces (``space_classifier``,
``space_bucket_model``). The Classify Spaces UI consumes ``collect`` and
``auto_classify``; the placement workflow (Stage 6 Batch 5+) consumes
``load_classifications_indexed`` to map a Space ElementId to its
assigned bucket(s).

Per-project classification *state* lives in the ES entity managed by
``space_storage`` (see ``active_yaml.load_classifications`` /
``save_classifications``). Templates (``space_buckets`` /
``space_profiles``) come through ``active_yaml.load_active_data``.
"""

import clr  # noqa: F401  -- needed before importing Autodesk.Revit.DB

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
)

import active_yaml as _active_yaml  # noqa: E402
import space_bucket_model as _bucket_model  # noqa: E402
import space_classifier as _classifier  # noqa: E402


# ---------------------------------------------------------------------
# Plain-data record for a Space row
# ---------------------------------------------------------------------

class SpaceInfo(object):
    """One Revit Space, captured in pure Python.

    The Classify Spaces UI binds row attributes by name, so the field
    set here is what the XAML can read. ``element`` is held only so the
    Revit-API edge can re-resolve the Space at Save time; UI code reads
    only the string fields.
    """

    __slots__ = (
        "element",
        "element_id",
        "unique_id",
        "name",
        "number",
        "level_name",
    )

    def __init__(self, element=None, element_id=None, unique_id="",
                 name="", number="", level_name=""):
        self.element = element
        self.element_id = element_id
        self.unique_id = unique_id or ""
        self.name = name or ""
        self.number = number or ""
        self.level_name = level_name or ""

    def __repr__(self):
        return "<SpaceInfo id={} name={!r} num={!r}>".format(
            self.element_id, self.name, self.number
        )


# ---------------------------------------------------------------------
# Revit-side collection
# ---------------------------------------------------------------------

def _element_id_int(elem_id):
    if elem_id is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return int(value)
        except Exception:
            continue
    return None


def _param_text(element, built_in_param):
    if element is None:
        return ""
    try:
        param = element.get_Parameter(built_in_param)
    except Exception:
        return ""
    if param is None:
        return ""
    for getter in ("AsString", "AsValueString"):
        try:
            value = getattr(param, getter)()
        except Exception:
            value = None
        if value:
            text = str(value).strip()
            if text:
                return text
    return ""


def _level_name(doc, space):
    try:
        lvl_id = space.LevelId
    except Exception:
        return ""
    if lvl_id is None:
        return ""
    try:
        lvl = doc.GetElement(lvl_id)
    except Exception:
        return ""
    if lvl is None:
        return ""
    name = getattr(lvl, "Name", "") or ""
    return str(name).strip()


def collect_spaces(doc):
    """Walk every placed Space in ``doc`` and return ``[SpaceInfo, ...]``.

    Sorted by (level, number, name) for stable display order.
    """
    rows = []
    if doc is None:
        return rows

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_MEPSpaces)
        .WhereElementIsNotElementType()
    )

    for space in collector:
        # Unplaced Spaces have Area == 0 and no Location; skip those —
        # they're noise in the table and have no point to anchor to.
        try:
            area = float(space.Area or 0.0)
        except Exception:
            area = 0.0
        if area <= 0.0:
            continue

        eid = getattr(space, "Id", None)
        try:
            uid = str(getattr(space, "UniqueId", "") or "")
        except Exception:
            uid = ""

        name = _param_text(space, BuiltInParameter.ROOM_NAME)
        if not name:
            try:
                name = str(getattr(space, "Name", "") or "").strip()
            except Exception:
                name = ""

        number = _param_text(space, BuiltInParameter.ROOM_NUMBER)
        level_name = _level_name(doc, space)

        rows.append(SpaceInfo(
            element=space,
            element_id=_element_id_int(eid),
            unique_id=uid,
            name=name,
            number=number,
            level_name=level_name,
        ))

    rows.sort(key=lambda r: (
        (r.level_name or "").lower(),
        (r.number or "").lower(),
        (r.name or "").lower(),
    ))
    return rows


# ---------------------------------------------------------------------
# Auto-classification (pure logic on top of collect_spaces)
# ---------------------------------------------------------------------

def auto_classify(spaces, buckets, client_key=None):
    """Return ``[(SpaceInfo, [SpaceBucket, ...]), ...]``.

    ``spaces``  — output of ``collect_spaces``.
    ``buckets`` — list of bucket dicts (or ``SpaceBucket`` wrappers) from
                  ``active_yaml.load_space_buckets``.
    ``client_key`` filters out client-restricted buckets that don't apply.
    """
    wrapped = _bucket_model.wrap_buckets(buckets) if buckets else []
    out = []
    for s in spaces or ():
        matches = _classifier.classify_space(
            s.name, wrapped, client_key=client_key,
        )
        out.append((s, matches))
    return out


# ---------------------------------------------------------------------
# Saved classifications
# ---------------------------------------------------------------------

def load_classifications_indexed(doc):
    """Return ``{space_element_id: [bucket_id, ...]}``.

    Multiple stacked entries for the same space collapse into one list
    while preserving save order. Bucket IDs not present in the loaded
    template are still returned — the placement layer is responsible
    for skipping unknown IDs.
    """
    raw = _active_yaml.load_classifications(doc) or []
    by_id = {}
    for entry in raw:
        if not isinstance(entry, dict):
            continue
        sid = entry.get("space_element_id")
        bid = (entry.get("bucket_id") or "").strip()
        if sid is None or not bid:
            continue
        try:
            sid = int(sid)
        except Exception:
            continue
        bucket_list = by_id.setdefault(sid, [])
        if bid not in bucket_list:
            bucket_list.append(bid)
    return by_id


def merge_with_saved(rows, saved_index):
    """Apply saved overrides to the auto-classified rows.

    Returns ``[(SpaceInfo, assigned_bucket_ids, auto_bucket_ids), ...]``
    where ``assigned`` defaults to the auto-detected list when no saved
    entry exists, but is fully replaced by the saved list otherwise.
    """
    out = []
    for space, auto_buckets in rows or ():
        auto_ids = [b.id for b in auto_buckets if b.id]
        saved_ids = saved_index.get(space.element_id) if saved_index else None
        if saved_ids is None:
            assigned_ids = list(auto_ids)
        else:
            assigned_ids = list(saved_ids)
        out.append((space, assigned_ids, auto_ids))
    return out


def payload_from_assignments(assignments):
    """Flatten ``[(SpaceInfo, [bucket_id, ...]), ...]`` to a save list.

    One classification record per (space, bucket_id) pair. Spaces with
    an empty assigned list contribute zero records — this represents an
    intentionally-uncategorised space and persists no row for it.
    """
    out = []
    for space, bucket_ids in assignments or ():
        for bid in bucket_ids or ():
            out.append({
                "space_element_id": space.element_id,
                "bucket_id": bid,
                "space_name": space.name,
            })
    return out
