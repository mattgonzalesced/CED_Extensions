#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Follow Parent"""

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
import active_yaml
import follow_parent_window

TITLE = "Follow Parent (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return
    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert("No profiles in the active store.", title=TITLE)
        return

    # Capture the user's current Revit selection BEFORE the modal
    # opens. WPF modals can't read live selection changes, so this
    # is the only useful read. The dialog's "Only selected" checkbox
    # consumes this set.
    selection_ids = []
    try:
        uidoc = revit.uidoc
        if uidoc is not None:
            for eid in uidoc.Selection.GetElementIds():
                value = (
                    getattr(eid, "Value", None)
                    or getattr(eid, "IntegerValue", None)
                )
                if value is None:
                    continue
                try:
                    selection_ids.append(int(value))
                except (TypeError, ValueError):
                    continue
    except Exception:
        selection_ids = []

    controller = follow_parent_window.show_modal(
        doc, profile_data, selection_element_ids=selection_ids,
    )
    result = getattr(controller, "_last_result", None)
    if controller.committed and result is not None:
        output.print_md(
            "**Follow Parent complete**\n\n"
            "- Moved: {}\n"
            "- Skipped aligned: {}\n"
            "- Skipped (no parent): {}\n"
            "- Warnings: {}\n".format(
                result.moved_count,
                result.skipped_aligned,
                result.skipped_no_parent,
                len(result.warnings),
            )
        )
        if result.warnings:
            output.print_md(
                "\n**Warnings:**\n"
                + "\n".join("- {}".format(w) for w in result.warnings[:50])
            )


if __name__ == "__main__":
    main()
