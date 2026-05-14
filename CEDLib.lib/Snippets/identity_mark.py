# -*- coding: utf-8 -*-
"""
``Identity Mark_CEDT`` project parameter helper for Ref Ops tools.

System Tagger, Name Piping Systems, and Print Pipe Data all read /
write an "Identity Mark" value on pipes and mechanical equipment.
The team's canonical parameter for this is ``Identity Mark_CEDT``
(GUID ``14a01120-5dd1-48f1-8302-18fad2572601``) — defined in the master
``AE CoolSys Energy Design (CED) Shared Parameters.txt`` and mirrored
in the small ``RefOps_SharedParams.txt`` that ships with the Ref Ops
pulldown.

Callers pass in the absolute path to the .txt file (so the helper
doesn't have to know about the repo layout) plus a forms-like module
for the prompt/alert dialogs. ``ensure_bound`` returns True on success,
False if the user cancelled or Revit refused.

Compatibility: IronPython 2.7 and CPython 3 via PythonNet.
"""

from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameterGroup

from Snippets import param_binder


IDENTITY_MARK_PARAM_NAME = "Identity Mark_CEDT"
IDENTITY_MARK_GUID = "14a01120-5dd1-48f1-8302-18fad2572601"  # informational
IDENTITY_MARK_SP_GROUP = "CED Identity"

# Categories Ref Ops tools read/write this parameter on. Extend here when
# a new Ref Ops tool starts using it on another category — the next run
# will auto-extend the binding (helper does union+ReInsert).
DEFAULT_BUILTIN_CATEGORIES = (
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_MechanicalControlDevices,
)


def make_spec(shared_param_file, builtin_categories=None):
    """Build a ``ProjectParameterSpec`` for ``Identity Mark_CEDT``."""
    return param_binder.ProjectParameterSpec(
        name=IDENTITY_MARK_PARAM_NAME,
        shared_param_file=shared_param_file,
        group_name_in_spfile=IDENTITY_MARK_SP_GROUP,
        builtin_categories=builtin_categories or DEFAULT_BUILTIN_CATEGORIES,
        parameter_group=BuiltInParameterGroup.PG_IDENTITY_DATA,
        instance=True,
    )


def ensure_bound(doc, forms_module, title, shared_param_file, builtin_categories=None):
    """One-call wrapper: needs-check, prompt-on-first-bind, transactional ReInsert.

    Behaviour:
        * already bound to every target category → silent, returns True
        * partial binding → silent extend, returns True
        * not bound at all → prompts user; returns False on cancel
        * Revit error → alerts user, returns False
    """
    from pyrevit import revit

    spec = make_spec(shared_param_file, builtin_categories)
    if not param_binder.needs_binding_update(doc, spec):
        return True

    is_first_time = _is_first_time_bind(doc, spec.name)
    if is_first_time:
        prompt = (
            "The {!r} shared parameter is not bound in this project.\n"
            "Bind it now? (used by Ref Ops tools to read/write per-element identity)"
        ).format(spec.name)
        if not forms_module.confirm(prompt, title=title):
            return False

    try:
        tx_label = (
            "Bind Identity Mark_CEDT"
            if is_first_time
            else "Extend Identity Mark_CEDT binding"
        )
        with revit.Transaction(tx_label, doc=doc):
            param_binder.ensure_bound(doc, spec)
    except param_binder.SharedParamError as exc:
        forms_module.alert(
            "Failed to bind {!r}:\n\n{}".format(spec.name, exc),
            title=title,
        )
        return False
    return True


def _is_first_time_bind(doc, param_name):
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        if getattr(iterator.Key, "Name", None) == param_name:
            return False
    return True


def lookup_identity_param(elem):
    """Find the canonical Identity Mark parameter on an element.

    Tries ``Identity Mark_CEDT`` first (canonical), then ``Identity Mark``
    (legacy data already in the project), then falls back to the built-in
    ``Mark`` parameter so older models keep working until they're rebound.
    Returns the ``Parameter`` or ``None``.
    """
    if elem is None:
        return None
    from Autodesk.Revit.DB import BuiltInParameter
    for name in (IDENTITY_MARK_PARAM_NAME, "Identity Mark"):
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if param is not None:
            return param
    try:
        return elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    except Exception:
        return None
