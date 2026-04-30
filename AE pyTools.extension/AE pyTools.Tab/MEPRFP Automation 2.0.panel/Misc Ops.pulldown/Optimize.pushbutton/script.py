#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Optimize"""

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
import optimize_window

TITLE = "Optimize (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return
    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert("No profiles in the active store.", title=TITLE)
        return
    selected_ids = []
    if uidoc is not None:
        try:
            selected_ids = list(uidoc.Selection.GetElementIds() or [])
        except Exception:
            selected_ids = []
    controller = optimize_window.show_modal(
        doc, profile_data,
        selected_element_ids=selected_ids or None,
    )
    result = getattr(controller, "_last_result", None)
    if controller.committed and result is not None:
        output.print_md(
            "**Optimize complete**\n\n"
            "- Moved:        {}\n"
            "- Re-parented:  {}\n"
            "- Skipped (no host in radius): {}\n"
            "- Warnings:     {}\n".format(
                result.moved_count,
                result.reparented_count,
                result.skipped_no_host,
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
