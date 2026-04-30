#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Create Independent Profile"""

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
import selection
import shared_params

TITLE = "Create Independent Profile (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    if not shared_params.is_element_linker_bound(doc):
        if not forms.confirm(
            "Bind the MEPRFP 2.0 Element_Linker shared parameter now?",
            title=TITLE,
        ):
            return
        try:
            with revit.Transaction("Bind MEPRFP Element_Linker", doc=doc):
                shared_params.ensure_element_linker_bound(doc)
        except shared_params.SharedParamError as exc:
            forms.alert("Failed to bind shared parameter:\n\n{}".format(exc), title=TITLE)
            return

    profile_data = active_yaml.load_active_data(doc)

    name = wpf_dialogs.prompt_for_string("Name for the independent profile:", title=TITLE)
    if not name:
        return
    while capture.find_profile_by_name(profile_data, name) is not None:
        new_name = wpf_dialogs.prompt_for_string(
            "A profile named {!r} already exists.\nEnter a different name:".format(name),
            title=TITLE, default=name + " (2)",
        )
        if not new_name:
            return
        name = new_name

    try:
        child_refs = selection.pick_children(uidoc, "Pick child elements (no parent), then Finish")
    except selection.SelectionCancelled:
        return
    if not child_refs:
        forms.alert("No elements were picked.", title=TITLE)
        return

    request = capture.CaptureRequest(
        profile_name=name,
        parent=None,
        children=child_refs,
    )
    with revit.Transaction("Create Independent Profile (MEPRFP 2.0)", doc=doc):
        result = capture.execute_capture(doc, profile_data, request)
        active_yaml.save_active_data(doc, profile_data, action="Create Independent Profile")

    output.print_md(
        "**Independent profile created**\n\n"
        "- Profile: `{}` (`{}`)\n"
        "- Set: `{}`\n"
        "- LEDs created: {}\n"
        "- Element_Linker writes: {} (skipped: {})\n".format(
            result.profile_name, result.profile_id, result.set_id,
            len(result.created_led_ids), result.linker_writes, result.linker_skipped,
        )
    )


if __name__ == "__main__":
    main()
