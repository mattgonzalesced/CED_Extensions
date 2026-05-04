# -*- coding: utf-8 -*-
"""
Typed wrapper for a ``space_buckets[*]`` YAML entry.

A *bucket* is the keyword-indexed category that a Revit Space gets
classified into (e.g. ``BAKERY``, ``RESTROOM``, ``ELECTRICAL ROOM``).
Buckets carry the lookup keywords; the work of mapping a bucket to one
or more ``space_profiles`` (and from there to LEDs) lives in the
profile/placement layers.

Schema (``space_buckets[*]``)::

    id: str                       # stable ID (e.g. SBKT-001)
    name: str                     # display name (e.g. "BAKERY")
    client_keys: [str, ...]       # restrict to clients (empty = universal)
    classification_keywords:      # case-insensitive substring tests
      - bakery
      - bake

Following the ``profile_model._DictBacked`` convention, the wrapper owns
the underlying dict so mutations round-trip through YAML.
"""


# ---------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------

def _str_or_none(value):
    if value is None:
        return None
    return str(value)


def _ensure_str_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v) for v in value if v is not None]
    return [str(value)]


# ---------------------------------------------------------------------
# Wrapper
# ---------------------------------------------------------------------

class SpaceBucket(object):
    """Wrapper around a ``space_buckets[*]`` dict."""

    def __init__(self, data=None):
        self._data = data if data is not None else {}

    def to_dict(self):
        return self._data

    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @id.setter
    def id(self, value):
        self._data["id"] = _str_or_none(value)

    @property
    def name(self):
        return _str_or_none(self._data.get("name")) or ""

    @name.setter
    def name(self, value):
        self._data["name"] = _str_or_none(value) or ""

    @property
    def client_keys(self):
        return _ensure_str_list(self._data.get("client_keys"))

    @client_keys.setter
    def client_keys(self, value):
        self._data["client_keys"] = _ensure_str_list(value)

    @property
    def classification_keywords(self):
        return _ensure_str_list(self._data.get("classification_keywords"))

    @classification_keywords.setter
    def classification_keywords(self, value):
        self._data["classification_keywords"] = _ensure_str_list(value)

    @property
    def is_universal(self):
        """True if no client_keys constraint is set (matches all clients)."""
        return not self.client_keys

    def applies_to_client(self, client_key):
        """``client_key=None`` or empty matches universal buckets only."""
        if self.is_universal:
            return True
        if not client_key:
            return False
        ck = str(client_key).strip().lower()
        return any(k.strip().lower() == ck for k in self.client_keys)

    def __repr__(self):
        return "<SpaceBucket id={!r} name={!r} kw={}>".format(
            self.id, self.name, self.classification_keywords
        )


# ---------------------------------------------------------------------
# Collection helpers
# ---------------------------------------------------------------------

def wrap_buckets(raw):
    """Return ``[SpaceBucket, ...]`` from a list of dicts."""
    return [SpaceBucket(d) for d in (raw or []) if isinstance(d, dict)]


def find_bucket_by_id(buckets, bucket_id):
    """Lookup by ID. ``buckets`` may be wrappers or raw dicts."""
    if not bucket_id:
        return None
    target = str(bucket_id).strip()
    for b in buckets or ():
        bid = b.id if isinstance(b, SpaceBucket) else (b or {}).get("id")
        if str(bid or "").strip() == target:
            return b if isinstance(b, SpaceBucket) else SpaceBucket(b)
    return None


def filter_buckets_for_client(buckets, client_key):
    """Return only buckets that apply to ``client_key`` (or universal)."""
    out = []
    for b in buckets or ():
        bucket = b if isinstance(b, SpaceBucket) else SpaceBucket(b)
        if bucket.applies_to_client(client_key):
            out.append(bucket)
    return out
