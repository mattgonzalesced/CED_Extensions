#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place Element Annotations"""

import os
import sys

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

import forms_compat as forms
import active_yaml
import annotation_placement
import annotation_placement_window

TITLE = "Place Element Annotations (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    active_view = doc.ActiveView
    ok, reason = annotation_placement.is_view_eligible(active_view)
    if not ok:
        forms.alert(
            "Active view can't host annotations:\n\n  {}\n\n"
            "Switch to a 2D view (plan / section / elevation / drafting / "
            "callout) and try again.".format(reason),
            title=TITLE,
        )
        return

    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert(
            "No profiles in the active store.",
            title=TITLE,
        )
        return

    controller = annotation_placement_window.show_modal(
        doc, active_view, profile_data
    )

    result = getattr(controller, "_last_result", None)
    if controller.committed and result is not None:
        output.print_md(
            "**Place Element Annotations complete**\n\n"
            "- Tags placed:        {}\n"
            "- Keynotes placed:    {}\n"
            "- Text notes placed:  {}\n"
            "- Already-placed (skipped): {}\n"
            "- Warnings: {}\n".format(
                result.placed_count_by_kind.get("tag", 0),
                result.placed_count_by_kind.get("keynote", 0),
                result.placed_count_by_kind.get("text_note", 0),
                result.skipped_duplicates,
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
