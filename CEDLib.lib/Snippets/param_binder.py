# -*- coding: utf-8 -*-
"""
Project-parameter binder for CED extensions.

Both ``CED MechTools`` and ``AE pyTools`` (MEPRFP Automation 2.0) need to
ensure a fixed list of shared parameters are bound to a fixed list of
categories in the active project. This module provides one helper that
does exactly that, idempotently:

    * load (and later restore) the application's shared-parameter file
    * locate the ``ExternalDefinition`` by name in the named group
    * compute the union of (existing categories) | (target categories)
    * if the union differs from the existing binding, ``ReInsert``
      (or ``Insert`` on first binding) with the union as an InstanceBinding
      or TypeBinding in the right parameter group

The "union not equality" wording is deliberate. Project authors sometimes
add the parameter to extra categories by hand; the goal is to *never
shrink* the binding, only extend it to cover what our tools need.

Compatibility:
    Importable under both IronPython 2.7 (pyRevit Ref Ops tools) and
    CPython 3 with PythonNet (pyRevit MEPRFP 2.0). Uses no f-strings,
    no annotations, no walrus.

Caller contract:
    ``ensure_bound`` mutates ``doc.ParameterBindings``. Open a Revit
    transaction *before* the call. ``needs_binding_update`` is a pure
    read and needs no transaction.
"""

import os

from Autodesk.Revit.DB import (
    BuiltInParameterGroup,
    Category,
    CategorySet,
    InstanceBinding,
    TypeBinding,
)


# Sentinel values for BindResult.status -----------------------------------

NOOP = "noop"
CREATED = "created"
EXTENDED = "extended"


class SharedParamError(Exception):
    """Raised when binding can't proceed for a recoverable reason."""


class ProjectParameterSpec(object):
    """Description of a project-bound shared parameter.

    Parameters
    ----------
    name : str
        Parameter name as it appears in the shared-param file *and* as
        Revit will surface it (the two must match; that's how Revit's
        ``Definitions.get_Item`` works).
    shared_param_file : str
        Absolute path to the .txt shared-parameter file holding ``name``.
    group_name_in_spfile : str
        ``*GROUP NAME`` value containing the parameter in the .txt file.
    builtin_categories : iterable of BuiltInCategory
        Target categories. Categories that don't allow bound parameters
        (or aren't present in the document) are silently skipped.
    parameter_group : BuiltInParameterGroup
        Which "Group under" the parameter appears in on the property
        palette (e.g. ``PG_DATA``, ``PG_IDENTITY_DATA``).
    instance : bool, default True
        ``True`` for InstanceBinding, ``False`` for TypeBinding.
    """

    __slots__ = (
        "name",
        "shared_param_file",
        "group_name_in_spfile",
        "builtin_categories",
        "parameter_group",
        "instance",
    )

    def __init__(
        self,
        name,
        shared_param_file,
        group_name_in_spfile,
        builtin_categories,
        parameter_group=BuiltInParameterGroup.PG_DATA,
        instance=True,
    ):
        self.name = name
        self.shared_param_file = shared_param_file
        self.group_name_in_spfile = group_name_in_spfile
        self.builtin_categories = tuple(builtin_categories)
        self.parameter_group = parameter_group
        self.instance = bool(instance)


class BindResult(object):
    """Outcome of an ``ensure_bound`` call.

    ``status`` is one of ``NOOP``, ``CREATED``, ``EXTENDED``.
    ``added_category_names`` lists the categories that were *added* to
    the binding by this call (always [] for ``NOOP``).
    ``skipped_category_names`` lists target categories that couldn't be
    bound because Revit reports they don't allow bound parameters in
    this document (uncommon, but possible for some category states).
    """

    __slots__ = ("status", "added_category_names", "skipped_category_names")

    def __init__(self, status, added_category_names=None, skipped_category_names=None):
        self.status = status
        self.added_category_names = list(added_category_names or [])
        self.skipped_category_names = list(skipped_category_names or [])

    def __repr__(self):
        return "BindResult(status={!r}, added={!r}, skipped={!r})".format(
            self.status, self.added_category_names, self.skipped_category_names
        )


# ---------------------------------------------------------------- internals


def _category_id_value(category):
    """Stable int key for a Category.Id, across Revit 2023 and 2024+.

    ``ElementId.IntegerValue`` is the IronPython-friendly path; on newer
    builds it's ``Value`` (deprecated ``IntegerValue`` still works in
    practice but a guarded fallback keeps us future-proof).
    """
    try:
        return int(category.Id.IntegerValue)
    except Exception:
        try:
            return int(category.Id.Value)
        except Exception:
            return int(str(category.Id))


def _resolve_bindable_categories(doc, builtins):
    """Return (Category objects, names skipped) for the target builtins.

    A category is skipped if Revit returns ``None`` for it or it reports
    ``AllowsBoundParameters`` False (some categories in some discipline
    configurations).
    """
    cats = []
    skipped = []
    for builtin in builtins:
        try:
            cat = Category.GetCategory(doc, builtin)
        except Exception:
            cat = None
        if cat is None:
            skipped.append(str(builtin))
            continue
        if not cat.AllowsBoundParameters:
            skipped.append(cat.Name)
            continue
        cats.append(cat)
    return cats, skipped


def _open_param_file(app, path):
    prior = app.SharedParametersFilename
    if not os.path.isfile(path):
        raise SharedParamError(
            "Shared parameter file is missing: {}".format(path)
        )
    app.SharedParametersFilename = path
    try:
        sp_file = app.OpenSharedParameterFile()
    except Exception as exc:
        app.SharedParametersFilename = prior or ""
        raise SharedParamError(
            "Failed to open shared parameter file {}: {}".format(path, exc)
        )
    if sp_file is None:
        app.SharedParametersFilename = prior or ""
        raise SharedParamError(
            "Revit returned no shared parameter file for {}".format(path)
        )
    return prior, sp_file


def _restore_param_file(app, prior):
    app.SharedParametersFilename = prior or ""


def _find_definition(sp_file, group_name, param_name):
    for group in sp_file.Groups:
        if group.Name != group_name:
            continue
        for definition in group.Definitions:
            if definition.Name == param_name:
                return definition
    return None


def _existing_binding(doc, param_name):
    """Return (Definition, Binding) for a name match in this doc, or (None, None)."""
    iterator = doc.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        defn = iterator.Key
        if getattr(defn, "Name", None) == param_name:
            return defn, iterator.Current
    return None, None


def _binding_category_ids(binding):
    if binding is None:
        return set()
    return set(_category_id_value(c) for c in binding.Categories)


def _build_category_set(categories):
    cat_set = CategorySet()
    for cat in categories:
        cat_set.Insert(cat)
    return cat_set


def _build_binding(category_set, instance):
    return InstanceBinding(category_set) if instance else TypeBinding(category_set)


# ---------------------------------------------------------------- public API


def diff_categories(doc, spec):
    """Return (missing_names, present_target_categories, skipped_names).

    * ``missing_names`` — target categories that aren't yet in the binding
      (or all target categories if the parameter isn't bound at all).
    * ``present_target_categories`` — Category objects already covered.
    * ``skipped_names`` — target categories Revit can't bind in this doc.

    Pure read; safe to call outside a transaction.
    """
    target_cats, skipped = _resolve_bindable_categories(doc, spec.builtin_categories)
    _, existing = _existing_binding(doc, spec.name)
    existing_ids = _binding_category_ids(existing)
    missing = []
    present = []
    for cat in target_cats:
        if _category_id_value(cat) in existing_ids:
            present.append(cat)
        else:
            missing.append(cat)
    missing_names = [c.Name for c in missing]
    return missing_names, present, skipped


def needs_binding_update(doc, spec):
    """True iff ``ensure_bound`` would change anything (no Tx required)."""
    missing_names, _, _ = diff_categories(doc, spec)
    return bool(missing_names)


def ensure_bound(doc, spec):
    """Insert or extend the parameter binding to cover ``spec`` categories.

    Idempotent. Caller MUST be inside a Revit transaction. Returns a
    ``BindResult`` describing what happened.

    The binding's category set is the *union* of the existing categories
    and the spec's target categories — we never shrink a binding that
    project authors may have extended by hand.
    """
    app = doc.Application
    target_cats, skipped = _resolve_bindable_categories(doc, spec.builtin_categories)
    if not target_cats:
        raise SharedParamError(
            "None of the target categories are bindable in this document "
            "(skipped: {}).".format(", ".join(skipped) or "<none>")
        )

    existing_defn, existing_binding = _existing_binding(doc, spec.name)
    existing_ids = _binding_category_ids(existing_binding)
    target_ids = set(_category_id_value(c) for c in target_cats)

    if existing_binding is not None and target_ids.issubset(existing_ids):
        # Already covers everything we need; also verify binding *kind*
        # (instance vs type) matches. If kind mismatches the user has
        # something incompatible — surface, don't silently flip.
        is_instance = isinstance(existing_binding, InstanceBinding)
        if is_instance != spec.instance:
            raise SharedParamError(
                "Parameter {!r} is bound as {} but spec requires {}. "
                "Re-bind by hand or change the spec.".format(
                    spec.name,
                    "Instance" if is_instance else "Type",
                    "Instance" if spec.instance else "Type",
                )
            )
        return BindResult(NOOP, [], skipped)

    # Need to (re)bind. Union the categories so we never shrink the set.
    union_id_to_cat = {}
    for cat in target_cats:
        union_id_to_cat[_category_id_value(cat)] = cat
    if existing_binding is not None:
        for cat in existing_binding.Categories:
            union_id_to_cat[_category_id_value(cat)] = cat
    union_cats = list(union_id_to_cat.values())
    cat_set = _build_category_set(union_cats)

    # Get the Definition. If already bound, reuse the bound Definition
    # (its GUID may differ from our SP file's; keep what's in the project
    # to avoid orphaning existing values).
    if existing_defn is not None:
        definition = existing_defn
    else:
        prior, sp_file = _open_param_file(app, spec.shared_param_file)
        try:
            definition = _find_definition(
                sp_file, spec.group_name_in_spfile, spec.name
            )
            if definition is None:
                raise SharedParamError(
                    "Definition {!r} not found under group {!r} in {}".format(
                        spec.name, spec.group_name_in_spfile, spec.shared_param_file
                    )
                )
        finally:
            _restore_param_file(app, prior)

    binding = _build_binding(cat_set, spec.instance)
    bindmap = doc.ParameterBindings
    if existing_binding is not None:
        ok = bindmap.ReInsert(definition, binding, spec.parameter_group)
        status = EXTENDED
    else:
        ok = bindmap.Insert(definition, binding, spec.parameter_group)
        status = CREATED
    if not ok:
        raise SharedParamError(
            "Revit refused to bind parameter {!r} ({} returned False).".format(
                spec.name, "ReInsert" if existing_binding is not None else "Insert"
            )
        )

    added_names = [
        union_id_to_cat[i].Name
        for i in target_ids
        if i not in existing_ids
    ]
    return BindResult(status, added_names, skipped)
