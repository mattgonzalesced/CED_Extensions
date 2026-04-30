# -*- coding: utf-8 -*-
"""
Reusable "Add to Profile" workflow.

Used by the ``Add to Profile`` pushbutton (no preselected profile) and
by the Manage Profiles editor's "Add LED" button (preselected profile
from the modal). The flow handles:

    * shared-parameter binding (one-time per project)
    * profile selection (skipped when ``profile_id`` is given)
    * automatic parent discovery via existing Element_Linker payloads
    * child pick (host model only)
    * optional BYPARENT / BYSIBLING directives (with parent params
      populated from the discovered parent)
    * the actual capture inside a single transaction

The caller passes ``forms`` / ``wpf_dialogs`` modules so this module
stays decoupled from any specific UI implementation choice.
"""

from pyrevit import revit

import active_yaml
import capture
import directives_dialog
import selection
import shared_params


def _ensure_param_bound(doc, forms, title):
    if shared_params.is_element_linker_bound(doc):
        return True
    if not forms.confirm(
        "Bind the MEPRFP 2.0 Element_Linker shared parameter now?",
        title=title,
    ):
        return False
    try:
        with revit.Transaction("Bind MEPRFP Element_Linker", doc=doc):
            shared_params.ensure_element_linker_bound(doc)
    except shared_params.SharedParamError as exc:
        forms.alert(
            "Failed to bind shared parameter:\n\n{}".format(exc),
            title=title,
        )
        return False
    return True


def _collect_parent_param_names(parent_ref):
    if parent_ref is None:
        return []
    out, seen = [], set()

    def _add_from(target):
        if target is None:
            return
        try:
            iter_params = target.Parameters
        except Exception:
            return
        for p in iter_params:
            if p is None:
                continue
            try:
                name = p.Definition.Name
            except Exception:
                continue
            if not name or name in seen:
                continue
            seen.add(name)
            out.append(name)

    elem = parent_ref.element
    _add_from(elem)
    try:
        type_id = elem.GetTypeId()
        type_elem = elem.Document.GetElement(type_id) if type_id else None
        _add_from(type_elem)
    except Exception:
        pass
    return out


def _collect_child_param_values(child_refs):
    out = {}
    for idx, child_ref in enumerate(child_refs):
        params = {}
        try:
            iter_params = child_ref.element.Parameters
        except Exception:
            iter_params = []
        for p in iter_params:
            if p is None or p.IsReadOnly or not p.HasValue:
                continue
            try:
                name = p.Definition.Name
            except Exception:
                continue
            if not name or name in params:
                continue
            try:
                params[name] = p.AsValueString() or p.AsString() or ""
            except Exception:
                params[name] = ""
        out[idx] = params
    return out


def run(doc, uidoc, profile_data, forms, wpf_dialogs, output,
        title="Add to Profile", profile_id=None):
    """Run the Add to Profile UI flow.

    Returns the ``capture.CaptureResult`` (with ``.warnings`` etc.) on
    success, or ``None`` if the user cancelled or the flow aborted.

    The caller is responsible for committing ``profile_data`` back to
    Extensible Storage if a result is returned.
    """
    if doc is None:
        forms.alert("No active document.", title=title)
        return None
    if not _ensure_param_bound(doc, forms, title):
        return None

    # Resolve target profile.
    profiles = profile_data.get("equipment_definitions") or []
    if profile_id is not None:
        target = next(
            (p for p in profiles if isinstance(p, dict) and p.get("id") == profile_id),
            None,
        )
        if target is None:
            forms.alert(
                "Profile id {} no longer exists in the active store.".format(profile_id),
                title=title,
            )
            return None
    else:
        if not profiles:
            forms.alert(
                "No profiles in the active store yet. Use 'New Profile' first.",
                title=title,
            )
            return None
        target = wpf_dialogs.pick_from_list(
            profiles,
            title=title,
            prompt="Pick the target profile:",
            display_func=lambda p: "{}  ({})".format(
                p.get("name") or "(unnamed)", p.get("id") or "?"
            ),
        )
        if target is None:
            return None

    # Auto-discover the parent so offsets are computed in the original frame.
    parent_ref = capture.discover_parent_ref(doc, target)
    parent_note = (
        "Parent auto-discovered from existing placement: {}".format(
            capture.element_label(parent_ref.element)
        ) if parent_ref else
        "No existing placement found. Offsets will fall back to the centroid."
    )

    # Pick children.
    forms.alert(
        "Pick the CHILD elements to append to:\n\n    {}\n\n{}\n\n"
        "  - Hosted tags / keynote symbols / text notes are auto-attached.\n"
        "  - Click 'Finish' in the ribbon when done.".format(
            target.get("name") or target.get("id") or "?",
            parent_note,
        ),
        title=title,
    )
    try:
        child_refs = selection.pick_children(
            uidoc, "Pick children, then click Finish"
        )
    except selection.SelectionCancelled:
        return None
    if not child_refs:
        return None

    # Optional directives — populated from the discovered parent.
    directives = {}
    if forms.confirm(
        "Define BYPARENT / BYSIBLING directives for these children?",
        title=title,
    ):
        parent_param_names = _collect_parent_param_names(parent_ref)
        child_param_values = _collect_child_param_values(child_refs)
        rows = directives_dialog.build_rows(
            child_refs, child_param_values, parent_param_names, []
        )
        if rows:
            chosen = directives_dialog.show_dialog(rows)
            if chosen:
                directives = chosen

    # Capture inside a transaction.
    request = capture.CaptureRequest(
        append_to_profile_id=target.get("id"),
        parent=parent_ref,
        children=child_refs,
        directives=directives,
    )
    with revit.Transaction(title, doc=doc):
        result = capture.execute_capture(doc, profile_data, request)
        active_yaml.save_active_data(doc, profile_data, action=title)

    if output is not None:
        output.print_md(
            "**{} succeeded**\n\n"
            "- Profile: `{}` (`{}`)\n"
            "- LEDs added: {}\n"
            "- Annotations added: {}\n"
            "- Element_Linker writes: {} (skipped: {})\n".format(
                title,
                result.profile_name, result.profile_id,
                len(result.created_led_ids),
                len(result.created_annotation_ids),
                result.linker_writes, result.linker_skipped,
            )
        )
        if result.warnings:
            output.print_md(
                "\n**Warnings:**\n"
                + "\n".join("- {}".format(w) for w in result.warnings)
            )
    return result
