# -*- coding: utf-8 -*-
from __future__ import absolute_import

import os

from organized.MEPKit.revit.symbols import resolve_or_load_symbol

_SYMBOL_CACHE = {}
_SYMBOL_FAIL = object()


def _cache_key(family, type_name, load_path):
    return (
        (family or u"").strip(),
        (type_name or u"").strip(),
        os.path.abspath(load_path) if load_path else u""
    )


def resolve_candidate_symbol(doc, candidate, logger=None):
    """
    Attempt to resolve or load a family symbol described by a candidate dict.
    Candidate keys:
        - family: family name (required for reliable loading)
        - type_catalog_name: type name from catalog
        - load_from: absolute path to .rfa or type catalog
    """
    if not candidate:
        return None

    family = (candidate.get('family') or u"").strip()
    type_name = (candidate.get('type_catalog_name') or u"").strip()
    load_path = (candidate.get('load_from') or u"").strip()

    key = _cache_key(family, type_name, load_path)
    cached = _SYMBOL_CACHE.get(key)
    if cached is _SYMBOL_FAIL:
        return None
    if cached:
        return cached

    path = load_path or None
    if path and not os.path.exists(path):
        if logger:
            logger.warning(u"[SYMBOL] Family path missing -> {} (skipping load)".format(path))
        path = None

    attempts = []
    if family and type_name:
        attempts.append((family, type_name))
    if family:
        attempts.append((family, None))
    if not attempts and type_name:
        attempts.append((None, type_name))

    for fam, typ in attempts or [(None, None)]:
        if not fam and not typ:
            continue
        sym = resolve_or_load_symbol(doc, fam, typ, load_path=path, logger=logger)
        if sym:
            if logger and type_name and typ is None and getattr(sym, "Name", None) != type_name:
                logger.warning(
                    u"[SYMBOL] Requested type '{}' not found in family '{}'; using '{}' instead."
                    .format(type_name, family or u"<any>", getattr(sym, "Name", u"<unnamed>"))
                )
            _SYMBOL_CACHE[key] = sym
            return sym

    _SYMBOL_CACHE[key] = _SYMBOL_FAIL
    if logger and (family or type_name):
        logger.warning(
            u"[SYMBOL] Failed to resolve family '{}' type '{}'. Check the type catalog or load path."
            .format(family or u"<unspecified>", type_name or u"*")
        )
    return None


def resolve_first_available_symbol(doc, candidates, logger=None):
    """Return the first symbol resolved from a list of candidate dicts."""
    for cand in (candidates or []):
        sym = resolve_candidate_symbol(doc, cand, logger=logger)
        if sym:
            return sym
    return None
