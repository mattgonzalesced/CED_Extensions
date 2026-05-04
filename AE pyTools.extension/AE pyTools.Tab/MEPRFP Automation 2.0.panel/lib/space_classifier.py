# -*- coding: utf-8 -*-
"""
Pure-Python classifier: Revit Space name -> matching ``SpaceBucket``s.

The legacy MEP Automation panel matched a Space to a "category" by
scanning the space's name for any of the category's keywords as a
case-insensitive substring. MEPRFP 2.0 keeps that semantic
(``classification_keywords`` are substring tests, case-insensitive),
extended to *return all matching buckets* rather than the first hit so
that stacked templates work — a bakery can match both ``BAKERY`` and
``OVEN ROOM`` if both buckets define overlapping keywords, and the
placement layer unions every matching profile.

No Revit-API imports; the caller passes the space's display name in.
"""

from space_bucket_model import (
    SpaceBucket,
    wrap_buckets,
    filter_buckets_for_client,
)


# ---------------------------------------------------------------------
# Single-keyword test
# ---------------------------------------------------------------------

def keyword_matches(space_name, keyword):
    """Case-insensitive substring test, with empty inputs returning False."""
    if not space_name or not keyword:
        return False
    return str(keyword).strip().lower() in str(space_name).lower()


# ---------------------------------------------------------------------
# Bucket-level test
# ---------------------------------------------------------------------

def bucket_matches_space(bucket, space_name):
    """True if any of the bucket's keywords appear in ``space_name``.

    A bucket with no keywords never matches (it would otherwise apply
    to every space, which is rarely what authors want).
    """
    b = bucket if isinstance(bucket, SpaceBucket) else SpaceBucket(bucket)
    keywords = b.classification_keywords
    if not keywords:
        return False
    name = space_name or ""
    for kw in keywords:
        if keyword_matches(name, kw):
            return True
    return False


# ---------------------------------------------------------------------
# Top-level entry point
# ---------------------------------------------------------------------

def classify_space(space_name, buckets, client_key=None):
    """Return the list of buckets that match ``space_name``.

    ``buckets`` may be wrappers or raw dicts. ``client_key`` filters out
    buckets whose ``client_keys`` don't include this client (universal
    buckets — empty ``client_keys`` — always pass through). The order
    of the returned list mirrors the input order so YAML authoring
    controls determinism.
    """
    candidates = filter_buckets_for_client(buckets, client_key)
    return [b for b in candidates if bucket_matches_space(b, space_name)]


def classify_many(spaces, buckets, client_key=None):
    """Bulk variant: ``spaces`` is an iterable of ``(key, name)`` pairs.

    Returns ``[(key, [SpaceBucket, ...]), ...]`` preserving the input
    order. Useful for the Classify Spaces table where ``key`` is the
    space's ElementId.
    """
    wrapped = wrap_buckets(buckets) if buckets else []
    out = []
    for key, name in spaces or ():
        out.append((key, classify_space(name, wrapped, client_key=client_key)))
    return out
