#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: New Profile"""

import os
import sys

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

import forms_compat as forms
import wpf_dialogs
import active_yaml
import capture
import directives_dialog
import selection
import shared_params

TITLE = "New Profile (MEPRFP 2.0)"


def _collect_param_names(target, into, seen):
    """Append every Definition.Name from ``target.Parameters`` (deduped)."""
    if target is None:
        return
    try:
        params_iter = target.Parameters
    except Exception:
        return
    for p in params_iter:
        if p is None:
            continue
        try:
            name = p.Definition.Name
        except Exception:
            continue
        if not name or name in seen:
            continue
        seen.add(name)
        into.append(name)


def _build_directive_inputs(parent_ref, child_refs):
    """Inputs for the directives dialog.

    Parent options: every parameter on the parent (instance + type),
    regardless of read-only / has-value, since directives only *read*
    the parent's value.

    Child rows: every writable, has-value instance parameter on each
    child — those are the parameters whose stored value we'd be
    overriding with a directive.
    """
    parent_param_names = []
    if parent_ref is not None:
        seen = set()
        elem = parent_ref.element
        _collect_param_names(elem, parent_param_names, seen)
        try:
            type_id = elem.GetTypeId()
            type_elem = elem.Document.GetElement(type_id) if type_id else None
            _collect_param_names(type_elem, parent_param_names, seen)
        except Exception:
            pass

    child_param_values = {}
    for idx, child_ref in enumerate(child_refs):
        params = {}
        for p in child_ref.element.Parameters:
            if p is None or p.IsReadOnly or not p.HasValue:
                continue
            try:
                name = p.Definition.Name
            except Exception:
                continue
            try:
                params[name] = p.AsValueString() or p.AsString() or ""
            except Exception:
                params[name] = ""
        child_param_values[idx] = params

    # Sibling options aren't useful until LED IDs are known. Stage 1
    # exposes children by index as "child[i] :: <param_name>"; the
    # capture engine will resolve to LED IDs after creation. For now we
    # keep sibling-mode conservative (empty options) so users default to
    # parent/static directives. Sibling support is a layered iteration.
    sibling_options = []
    return parent_param_names, child_param_values, sibling_options


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    # Shared parameter binding (one-time per project).
    if not shared_params.is_element_linker_bound(doc):
        if not forms.confirm(
            "The MEPRFP 2.0 Element_Linker shared parameter is not bound in this project.\n"
            "Bind it now? (required for capture to write back to elements)",
            title=TITLE,
        ):
            return
        try:
            with revit.Transaction("Bind MEPRFP Element_Linker", doc=doc):
                shared_params.ensure_element_linker_bound(doc)
        except shared_params.SharedParamError as exc:
            forms.alert("Failed to bind shared parameter:\n\n{}".format(exc), title=TITLE)
            return

    # Workflow overview before any picks.
    if not forms.confirm(
        "New Profile workflow:\n\n"
        "  1. Pick the PARENT element. The profile name is auto-derived\n"
        "     from the parent's Family : Type.\n\n"
        "  2. Pick the CHILD elements (and any hosted tags / keynote\n"
        "     symbols / text notes) that belong to the profile.\n"
        "     - Use Ctrl+Click to add or remove from the selection.\n"
        "     - Click 'Finish' in the ribbon when done.\n\n"
        "  3. (Optional) Define BYPARENT / BYSIBLING parameter directives.\n\n"
        "Click Yes to begin, No to cancel.",
        title=TITLE,
    ):
        return

    # Step 1: pick parent.
    parent_in_link = forms.confirm(
        "Is the parent element in a LINKED model?\n\n"
        "Yes: pick a linked-model element.\n"
        "No:  pick an element in the active model.",
        title=TITLE,
    )
    forms.alert(
        "Step 1 of 3: Pick the PARENT element {} model.".format(
            "in a LINKED" if parent_in_link else "in the ACTIVE"
        ),
        title=TITLE,
    )
    try:
        parent_ref = selection.pick_parent(
            uidoc,
            prompt="Step 1 of 3 — Pick parent element",
            from_linked=parent_in_link,
        )
    except selection.SelectionCancelled:
        return

    # Auto-derive the profile name from the parent.
    name = capture.element_label(parent_ref.element)
    if not name:
        forms.alert(
            "Could not derive a Family : Type name from the picked parent. "
            "Capture cancelled.",
            title=TITLE,
        )
        return

    profile_data = active_yaml.load_active_data(doc)

    if capture.find_profile_by_name(profile_data, name) is not None:
        forms.alert(
            "A profile named:\n\n    {}\n\n"
            "already exists. To add more children to it, use 'Add to Profile' instead.\n"
            "Capture cancelled.".format(name),
            title=TITLE,
        )
        return

    # Step 2: pick children.
    forms.alert(
        "Step 2 of 3: Pick CHILD elements (host model only).\n\n"
        "  - Click each child to add it to the running selection.\n"
        "  - Hosted tags and 'GA_Keynote Symbol_CED' family instances\n"
        "    that depend on a child are captured automatically — you do\n"
        "    not need to click them. Pick them only if you want them as\n"
        "    standalone LEDs.\n"
        "  - Click 'Finish' in the ribbon when done.",
        title=TITLE,
    )
    try:
        child_refs = selection.pick_children(
            uidoc, "Step 2 of 3 — Pick children, then click Finish"
        )
    except selection.SelectionCancelled:
        return
    if not child_refs:
        forms.alert("No child elements were picked. Capture cancelled.", title=TITLE)
        return

    # Step 3: optional directives.
    directives = {}
    if forms.confirm(
        "Step 3 of 3: define BYPARENT / BYSIBLING parameter directives for this capture?\n\n"
        "Yes: map parameters between parent and/or siblings.\n"
        "No:  keep all captured parameters static.",
        title=TITLE,
    ):
        parent_params, child_params, sibling_opts = _build_directive_inputs(
            parent_ref, child_refs
        )
        rows = directives_dialog.build_rows(
            child_refs, child_params, parent_params, sibling_opts
        )
        if rows:
            chosen = directives_dialog.show_dialog(rows)
            if chosen:
                directives = chosen

    request = capture.CaptureRequest(
        profile_name=name,
        parent=parent_ref,
        children=child_refs,
        directives=directives,
    )

    with revit.Transaction("New Profile (MEPRFP 2.0)", doc=doc):
        result = capture.execute_capture(doc, profile_data, request)
        active_yaml.save_active_data(doc, profile_data, action="New Profile")

    output.print_md(
        "**New Profile created**\n\n"
        "- Profile: `{}` (`{}`)\n"
        "- Set: `{}`\n"
        "- LEDs created: {}\n"
        "- Element_Linker writes: {} (skipped: {})\n".format(
            result.profile_name, result.profile_id, result.set_id,
            len(result.created_led_ids), result.linker_writes, result.linker_skipped,
        )
    )
    if result.warnings:
        output.print_md("\n**Warnings:**\n" + "\n".join("- {}".format(w) for w in result.warnings))


if __name__ == "__main__":
    main()
