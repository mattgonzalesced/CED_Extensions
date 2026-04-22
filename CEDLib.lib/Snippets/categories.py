# -*- coding: utf-8 -*-
"""Shared helpers for electrical category resolution and binding scopes."""

import re

from pyrevit import DB, HOST_APP

from Snippets import revit_helpers

BINDING_SCOPE_ELECTRICAL_CIRCUITS = "electrical_circuits"
BINDING_SCOPE_ALL_ELECTRICAL = "all_electrical"
BINDING_SCOPE_EXPLICIT = "explicit_categories"

_BASE_FIXTURE_BIC_NAMES = (
    # Keep binding scope aligned with current downstream writer behavior.
    # We only write fixture results to OST_ElectricalFixtures today.
    "OST_ElectricalFixtures",
)
_OPTIONAL_FIXTURE_BIC_BY_MIN_VERSION = (
    # Reserved for future expansion when downstream writer supports more categories.
)

_CIRCUIT_BIC_NAMES = ("OST_ElectricalCircuit",)
_EQUIPMENT_BIC_NAMES = ("OST_ElectricalEquipment",)

_SCOPE_CIRCUITS_TOKENS = set(
    [
        "electricalcircuits",
        "electricalcircuit",
        "eelctricalcircuits",
        "eelctricalcircuit",
    ]
)
_SCOPE_ALL_TOKENS = set(
    [
        "allelectrical",
        "allelectricalcategories",
    ]
)
_EQUIPMENT_TOKENS = set(
    [
        "electricalequipment",
    ]
)
_FIXTURE_GROUP_TOKENS = set(
    [
        "electricalfixtures",
        "electricaldevices",
    ]
)

_INVALID_CATEGORY_ID_VALUE = revit_helpers.get_elementid_value(DB.ElementId.InvalidElementId, default=-1)


def _host_revit_version():
    try:
        return int(getattr(HOST_APP, "version", 0) or 0)
    except Exception:
        return 0


def category_id_value(category_or_id, default=_INVALID_CATEGORY_ID_VALUE):
    """Return a numeric category id from Category/ElementId-like inputs."""
    if category_or_id is None:
        return int(default)
    try:
        maybe_id = getattr(category_or_id, "Id", None)
    except Exception:
        maybe_id = None
    if isinstance(maybe_id, DB.ElementId):
        return revit_helpers.get_elementid_value(maybe_id, default=default)
    return revit_helpers.get_elementid_value(category_or_id, default=default)


def is_valid_category_id_value(value):
    """Return True when value is a real category id (allows negative built-in ids)."""
    try:
        numeric = int(value)
    except Exception:
        return False
    return int(numeric) != int(_INVALID_CATEGORY_ID_VALUE)


def _norm(text):
    return re.sub(r"[^a-z0-9]+", "", str(text or "").strip().lower())


def _bic_from_name(name):
    try:
        return getattr(DB.BuiltInCategory, str(name or ""), None)
    except Exception:
        return None


def _element_id_from_bic(bic):
    if bic is None:
        return None
    try:
        return revit_helpers.elementid_from_value(int(bic))
    except Exception:
        return None


def _category_exists_in_doc(doc, category_id):
    if doc is None or category_id is None:
        return True
    target_value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
    if not is_valid_category_id_value(target_value):
        return False
    try:
        for cat in list(doc.Settings.Categories or []):
            if cat is None:
                continue
            if category_id_value(cat, default=_INVALID_CATEGORY_ID_VALUE) == int(target_value):
                return True
    except Exception:
        pass
    return False


def _unique_category_ids(category_ids):
    unique = []
    seen = set()
    for category_id in list(category_ids or []):
        value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if (not is_valid_category_id_value(value)) or value in seen:
            continue
        seen.add(value)
        unique.append(category_id)
    return unique


def _category_ids_for_bic_names(doc, bic_names):
    ids = []
    for bic_name in list(bic_names or []):
        bic = _bic_from_name(bic_name)
        if bic is None:
            continue
        category_id = _element_id_from_bic(bic)
        if category_id is None:
            continue
        if not _category_exists_in_doc(doc, category_id):
            continue
        ids.append(category_id)
    return _unique_category_ids(ids)


def get_fixture_bic_names(version=None):
    revit_version = int(version or _host_revit_version() or 0)
    names = list(_BASE_FIXTURE_BIC_NAMES)
    for bic_name, min_version in list(_OPTIONAL_FIXTURE_BIC_BY_MIN_VERSION or []):
        if int(revit_version) >= int(min_version):
            names.append(bic_name)
    return names


def get_circuit_category_ids(doc=None):
    return _category_ids_for_bic_names(doc, _CIRCUIT_BIC_NAMES)


def get_equipment_category_ids(doc=None):
    return _category_ids_for_bic_names(doc, _EQUIPMENT_BIC_NAMES)


def get_fixture_category_ids(doc=None, version=None):
    return _category_ids_for_bic_names(doc, get_fixture_bic_names(version=version))


def get_all_electrical_category_ids(doc=None, version=None):
    all_ids = []
    all_ids.extend(get_circuit_category_ids(doc))
    all_ids.extend(get_equipment_category_ids(doc))
    all_ids.extend(get_fixture_category_ids(doc, version=version))
    return _unique_category_ids(all_ids)


def parse_binding_scope(categories_value):
    token = _norm(categories_value)
    if token in _SCOPE_CIRCUITS_TOKENS:
        return BINDING_SCOPE_ELECTRICAL_CIRCUITS
    if token in _SCOPE_ALL_TOKENS:
        return BINDING_SCOPE_ALL_ELECTRICAL
    return BINDING_SCOPE_EXPLICIT


def split_category_tokens(categories_value):
    return [x.strip() for x in str(categories_value or "").split(",") if x and str(x).strip()]


def resolve_explicit_category_ids(doc, category_tokens):
    resolved = []
    missing = []

    category_map = {}
    try:
        for cat in list(doc.Settings.Categories or []):
            if cat is None:
                continue
            category_map[_norm(cat.Name)] = cat
    except Exception:
        category_map = {}

    for token in list(category_tokens or []):
        token_norm = _norm(token)
        if not token_norm:
            continue

        matched_ids = []
        if token_norm in _SCOPE_CIRCUITS_TOKENS:
            matched_ids.extend(get_circuit_category_ids(doc))
        elif token_norm in _SCOPE_ALL_TOKENS:
            matched_ids.extend(get_all_electrical_category_ids(doc))
        elif token_norm in _EQUIPMENT_TOKENS:
            matched_ids.extend(get_equipment_category_ids(doc))
        elif token_norm in _FIXTURE_GROUP_TOKENS:
            matched_ids.extend(get_fixture_category_ids(doc))
        else:
            cat = category_map.get(token_norm)
            if cat is not None:
                matched_ids.append(getattr(cat, "Id", None))

        matched_ids = _unique_category_ids(matched_ids)
        if not matched_ids:
            missing.append(str(token))
            continue
        resolved.extend(matched_ids)

    return _unique_category_ids(resolved), missing


def resolve_binding_category_ids(doc, categories_value):
    scope = parse_binding_scope(categories_value)
    if scope == BINDING_SCOPE_ELECTRICAL_CIRCUITS:
        return get_circuit_category_ids(doc), []
    if scope == BINDING_SCOPE_ALL_ELECTRICAL:
        return get_all_electrical_category_ids(doc), []
    tokens = split_category_tokens(categories_value)
    return resolve_explicit_category_ids(doc, tokens)


def apply_writeback_filter(doc, category_ids, write_equipment_results, write_fixture_results):
    equipment_values = category_id_values(get_equipment_category_ids(doc))
    fixture_values = category_id_values(get_fixture_category_ids(doc))

    filtered = []
    for category_id in list(category_ids or []):
        value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if not is_valid_category_id_value(value):
            continue
        if (value in equipment_values) and (not bool(write_equipment_results)):
            continue
        if (value in fixture_values) and (not bool(write_fixture_results)):
            continue
        filtered.append(category_id)
    return _unique_category_ids(filtered)


def category_id_values(category_ids):
    values = set()
    for category_id in list(category_ids or []):
        value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if not is_valid_category_id_value(value):
            continue
        values.add(int(value))
    return values


def category_id_values_from_categories(categories):
    """Return numeric id values for Category collections."""
    return category_id_values([getattr(cat, "Id", None) for cat in list(categories or []) if cat is not None])


def merge_category_sets(doc, first_categories, second_categories):
    """Return CategorySet union of two category collections (Category/ElementId mixed)."""
    merged_values = set()
    merged_values.update(category_id_values(first_categories))
    merged_values.update(category_id_values(second_categories))
    merged_ids = [revit_helpers.elementid_from_value(int(v)) for v in sorted(list(merged_values or []))]
    return build_category_set(doc, merged_ids)


def build_category_set(doc, category_ids):
    category_set = DB.CategorySet()
    missing = []
    inserted = 0

    category_map = {}
    try:
        for cat in list(doc.Settings.Categories or []):
            if cat is None:
                continue
            cat_id_value = category_id_value(cat, default=_INVALID_CATEGORY_ID_VALUE)
            if is_valid_category_id_value(cat_id_value):
                category_map[cat_id_value] = cat
    except Exception:
        category_map = {}

    for category_id in list(_unique_category_ids(category_ids) or []):
        cat_id_value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if not is_valid_category_id_value(cat_id_value):
            continue
        category = category_map.get(cat_id_value)
        if category is None:
            missing.append(str(cat_id_value))
            continue
        category_set.Insert(category)
        inserted += 1

    return category_set, inserted, missing
