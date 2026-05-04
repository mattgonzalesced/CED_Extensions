# -*- coding: utf-8 -*-
"""Tests for space_bucket_model + space_classifier."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from space_bucket_model import (
    SpaceBucket,
    wrap_buckets,
    find_bucket_by_id,
    filter_buckets_for_client,
)
from space_classifier import (
    keyword_matches,
    bucket_matches_space,
    classify_space,
    classify_many,
)


_FAILS = []


def _check(label, cond, detail=""):
    if cond:
        print("  PASS  {}".format(label))
    else:
        print("  FAIL  {}  {}".format(label, detail))
        _FAILS.append(label)


# ---------------------------------------------------------------------
# Bucket model
# ---------------------------------------------------------------------

def test_bucket_basic_fields():
    print("\n[buckets] basic fields")
    b = SpaceBucket({"id": "SBKT-1", "name": "BAKERY"})
    _check("id", b.id == "SBKT-1")
    _check("name", b.name == "BAKERY")
    _check("client_keys default empty", b.client_keys == [])
    _check("keywords default empty", b.classification_keywords == [])
    _check("universal when empty client_keys", b.is_universal is True)


def test_bucket_setters_round_trip():
    print("\n[buckets] setters round-trip into the dict")
    d = {}
    b = SpaceBucket(d)
    b.id = "SBKT-9"
    b.name = "RESTROOM"
    b.client_keys = ["heb"]
    b.classification_keywords = ["restroom", "wc"]
    _check("dict id", d.get("id") == "SBKT-9")
    _check("dict name", d.get("name") == "RESTROOM")
    _check("dict client_keys", d.get("client_keys") == ["heb"])
    _check(
        "dict keywords",
        d.get("classification_keywords") == ["restroom", "wc"],
    )
    _check("not universal once client_keys set", b.is_universal is False)


def test_bucket_applies_to_client():
    print("\n[buckets] applies_to_client semantics")
    universal = SpaceBucket({"id": "U", "client_keys": []})
    heb = SpaceBucket({"id": "H", "client_keys": ["heb"]})
    multi = SpaceBucket({"id": "M", "client_keys": ["heb", "pf"]})

    _check("universal matches None", universal.applies_to_client(None))
    _check("universal matches any", universal.applies_to_client("heb"))
    _check("heb matches heb", heb.applies_to_client("heb"))
    _check("heb matches HEB (case)", heb.applies_to_client("HEB"))
    _check("heb does not match pf", heb.applies_to_client("pf") is False)
    _check("heb does not match None", heb.applies_to_client(None) is False)
    _check("multi matches pf", multi.applies_to_client("pf"))


def test_find_bucket_by_id():
    print("\n[buckets] find_bucket_by_id")
    raw = [
        {"id": "A", "name": "Alpha"},
        {"id": "B", "name": "Beta"},
        {"id": "C", "name": "Gamma"},
    ]
    found = find_bucket_by_id(raw, "B")
    _check("returns wrapper", isinstance(found, SpaceBucket))
    _check("right one", found.name == "Beta")
    _check("missing -> None", find_bucket_by_id(raw, "Z") is None)
    _check("empty id -> None", find_bucket_by_id(raw, "") is None)


def test_filter_buckets_for_client():
    print("\n[buckets] filter_buckets_for_client")
    raw = [
        {"id": "U", "client_keys": []},
        {"id": "H", "client_keys": ["heb"]},
        {"id": "P", "client_keys": ["pf"]},
    ]
    heb_only = filter_buckets_for_client(raw, "heb")
    _check("heb sees U+H", sorted(b.id for b in heb_only) == ["H", "U"])
    pf_only = filter_buckets_for_client(raw, "pf")
    _check("pf sees U+P", sorted(b.id for b in pf_only) == ["P", "U"])
    none_seen = filter_buckets_for_client(raw, None)
    _check("None sees only universal",
           [b.id for b in none_seen] == ["U"])


# ---------------------------------------------------------------------
# Classifier
# ---------------------------------------------------------------------

def test_keyword_matches():
    print("\n[classifier] keyword_matches (case-insensitive substring)")
    _check("hit lower", keyword_matches("Bakery 101", "bakery"))
    _check("hit upper", keyword_matches("BAKERY 101", "bakery"))
    _check("hit mixed", keyword_matches("Bake Room", "BAKE"))
    _check("substring hit", keyword_matches("WC-101", "wc"))
    _check("miss", keyword_matches("Office 101", "bakery") is False)
    _check("empty space name", keyword_matches("", "bakery") is False)
    _check("empty keyword", keyword_matches("Bakery", "") is False)
    _check("None safe", keyword_matches(None, "bake") is False)


def test_bucket_matches_space():
    print("\n[classifier] bucket_matches_space")
    bakery = SpaceBucket({
        "id": "B", "name": "BAKERY",
        "classification_keywords": ["bakery", "bake"],
    })
    _check("bakery matches 'Bakery 101'",
           bucket_matches_space(bakery, "Bakery 101"))
    _check("bakery matches 'Bake Room'",
           bucket_matches_space(bakery, "Bake Room"))
    _check("bakery does not match 'Office'",
           bucket_matches_space(bakery, "Office") is False)
    no_kw = SpaceBucket({"id": "X", "classification_keywords": []})
    _check("no keywords -> never matches",
           bucket_matches_space(no_kw, "Anything") is False)


def test_classify_space_returns_all_matches():
    print("\n[classifier] classify_space returns ALL matching buckets")
    raw = [
        {"id": "B1", "client_keys": [],
         "classification_keywords": ["restroom"]},
        {"id": "B2", "client_keys": [],
         "classification_keywords": ["women"]},
        {"id": "B3", "client_keys": [],
         "classification_keywords": ["office"]},
    ]
    matches = classify_space("Women's Restroom 1", raw)
    ids = sorted(b.id for b in matches)
    _check("two stacked matches", ids == ["B1", "B2"])
    _check("preserves input order",
           [b.id for b in matches] == ["B1", "B2"])
    _check("no match returns empty",
           classify_space("Lobby", raw) == [])


def test_classify_space_filters_by_client():
    print("\n[classifier] classify_space respects client_keys")
    raw = [
        {"id": "U", "client_keys": [],
         "classification_keywords": ["restroom"]},
        {"id": "H", "client_keys": ["heb"],
         "classification_keywords": ["restroom"]},
        {"id": "P", "client_keys": ["pf"],
         "classification_keywords": ["restroom"]},
    ]
    heb = sorted(b.id for b in classify_space("Restroom", raw, client_key="heb"))
    _check("heb sees U+H", heb == ["H", "U"])
    pf = sorted(b.id for b in classify_space("Restroom", raw, client_key="pf"))
    _check("pf sees U+P", pf == ["P", "U"])
    nokey = [b.id for b in classify_space("Restroom", raw)]
    _check("None client sees only universal", nokey == ["U"])


def test_classify_many_preserves_keys():
    print("\n[classifier] classify_many preserves keys + order")
    raw = [
        {"id": "BAK", "client_keys": [],
         "classification_keywords": ["bakery"]},
        {"id": "WC", "client_keys": [],
         "classification_keywords": ["restroom"]},
    ]
    spaces = [
        (101, "Bakery"),
        (102, "Lobby"),
        (103, "Men's Restroom"),
    ]
    out = classify_many(spaces, raw)
    keys = [k for k, _ in out]
    _check("keys preserved", keys == [101, 102, 103])
    _check("101 -> BAK",
           [b.id for b in out[0][1]] == ["BAK"])
    _check("102 -> []",
           out[1][1] == [])
    _check("103 -> WC",
           [b.id for b in out[2][1]] == ["WC"])


# ---------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------

def main():
    print("Running space_classifier tests")
    test_bucket_basic_fields()
    test_bucket_setters_round_trip()
    test_bucket_applies_to_client()
    test_find_bucket_by_id()
    test_filter_buckets_for_client()
    test_keyword_matches()
    test_bucket_matches_space()
    test_classify_space_returns_all_matches()
    test_classify_space_filters_by_client()
    test_classify_many_preserves_keys()

    print("")
    if _FAILS:
        print("FAILED: {} test(s) — {}".format(len(_FAILS), _FAILS))
        sys.exit(1)
    print("All space_classifier tests passed.")


if __name__ == "__main__":
    main()
