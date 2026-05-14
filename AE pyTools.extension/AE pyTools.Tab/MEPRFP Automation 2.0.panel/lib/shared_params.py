# -*- coding: utf-8 -*-
"""
Shared-parameter setup for the MEPRFP 2.0 ``Element_Linker`` parameter.

The parameter is bound to instance categories that can host placed
equipment children. On first use of any capture / placement tool,
``ensure_element_linker_bound(doc)`` is called; it swaps the
application's shared-parameter file to ours just long enough to bind
(or *extend* the binding) and then restores the prior path.

Binding strategy
----------------
The category list below is the "fixed broad list with extend-if-missing"
agreed with the team. Every time a placement runs, the helper checks
that the parameter is bound to *all* of the categories below; if any
are missing, it ``ReInsert``s with the union of (existing) | (target)
so we never shrink an author-extended binding. Add to
``_BINDING_CATEGORY_BUILTINS`` later and the next placement run
automatically picks it up.

The parameter GUID is fixed (matches ``_resources/MEPRFP_2_SharedParams.txt``)
so the parameter is recognised across projects.
"""

import os

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameterGroup,
)
from System import Guid  # noqa: F401  (kept for API back-compat)

from Snippets import param_binder
from Snippets.param_binder import SharedParamError  # noqa: F401 (re-export)


ELEMENT_LINKER_PARAM_NAME = "Element_Linker"
ELEMENT_LINKER_GUID_STR = "b4e5f6a7-1c2d-3e4f-5a6b-7c8d9e0f1a2b"
SHARED_PARAM_GROUP_NAME = "MEPRFP_2_Authoring"

_RESOURCES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_resources")
SHARED_PARAM_FILE_PATH = os.path.join(_RESOURCES_DIR, "MEPRFP_2_SharedParams.txt")


# Instance categories that host equipment children placed by MEPRFP 2.0.
# Add new builtins here when a new tool starts placing into a new
# category — the next placement run will auto-extend the binding.
_BINDING_CATEGORY_BUILTINS = (
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_IOSModelGroups,  # Model Groups
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
)


def _spec():
    """One-shot factory; cheap, no doc required."""
    return param_binder.ProjectParameterSpec(
        name=ELEMENT_LINKER_PARAM_NAME,
        shared_param_file=SHARED_PARAM_FILE_PATH,
        group_name_in_spfile=SHARED_PARAM_GROUP_NAME,
        builtin_categories=_BINDING_CATEGORY_BUILTINS,
        parameter_group=BuiltInParameterGroup.PG_DATA,
        instance=True,
    )


# ---------------------------------------------------------- public API


def is_element_linker_bound(doc):
    """True if ``Element_Linker`` is already bound to *every* target category.

    Returns False if the parameter is missing entirely or is bound to
    only a subset of ``_BINDING_CATEGORY_BUILTINS`` — that mirrors the
    contract callers want: ``False`` means "you need to call
    ``ensure_element_linker_bound`` inside a transaction."
    """
    return not param_binder.needs_binding_update(doc, _spec())


def ensure_element_linker_bound(doc):
    """Bind ``Element_Linker`` or extend the binding to all target categories.

    Idempotent. Caller is responsible for opening a Revit transaction
    *before* calling this; we mutate ``ParameterBindings``.

    Returns the ``BindResult`` from the helper so callers can decide
    whether to surface a "extended to X, Y, Z" message. Existing callers
    that ignore the return value continue to work unchanged.
    """
    return param_binder.ensure_bound(doc, _spec())


def prompt_and_bind(doc, forms_module, title, reason=None):
    """Top-level "ensure bound for this tool" helper for placement scripts.

    Behaviour:
        * already covers all target categories  → return True silently
        * binding missing some categories but exists → silently extend
          (user already authorised the parameter on this project)
        * binding doesn't exist at all → prompt user, then create

    On failure (user cancel or Revit refuses), alerts the user and
    returns False. Wraps its own Revit transaction.

    Pass any object with ``confirm(msg, title=...)`` and
    ``alert(msg, title=...)`` (e.g. ``pyrevit.forms`` or the
    ``forms_compat`` module used by MEPRFP 2.0).
    """
    # Defer Revit imports until call; helper is also import-safe outside Revit.
    from pyrevit import revit

    spec = _spec()
    if not param_binder.needs_binding_update(doc, spec):
        return True

    # If anything is bound under this name already, treat it as an
    # already-authorised parameter and silently extend.
    is_first_time = _is_first_time_bind(doc)
    if is_first_time:
        prompt = (
            "The MEPRFP 2.0 Element_Linker shared parameter is not bound "
            "in this project.\nBind it now?"
        )
        if reason:
            prompt += "\n\n({})".format(reason)
        if not forms_module.confirm(prompt, title=title):
            return False

    try:
        tx_label = (
            "Bind MEPRFP Element_Linker"
            if is_first_time
            else "Extend MEPRFP Element_Linker binding"
        )
        with revit.Transaction(tx_label, doc=doc):
            param_binder.ensure_bound(doc, spec)
    except SharedParamError as exc:
        forms_module.alert(
            "Failed to bind shared parameter:\n\n{}".format(exc),
            title=title,
        )
        return False
    return True


def _is_first_time_bind(doc):
    """True if the parameter has no binding at all (vs. partial binding)."""
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        if getattr(iterator.Key, "Name", None) == ELEMENT_LINKER_PARAM_NAME:
            return False
    return True
