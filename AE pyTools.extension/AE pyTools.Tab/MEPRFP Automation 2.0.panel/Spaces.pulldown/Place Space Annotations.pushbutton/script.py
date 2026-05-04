#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place Space Annotations"""

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
import space_annotation_workflow
import place_space_annotations_window

TITLE = "Place Space Annotations (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    active_view = doc.ActiveView
    ok, reason = space_annotation_workflow.is_view_eligible(active_view)
    if not ok:
        forms.alert(
            "Active view can't host annotations:\n\n  {}\n\n"
            "Switch to a 2D view (plan / section / elevation / drafting / "
            "callout) and try again.".format(reason),
            title=TITLE,
        )
        return

    profile_data = active_yaml.load_active_data(doc) or {}
    if not profile_data.get("space_profiles"):
        forms.alert(
            "No space_profiles in the active YAML. Use Manage Space Profiles to "
            "define some first.",
            title=TITLE,
        )
        return

    controller = place_space_annotations_window.show_modal(
        doc, active_view, profile_data,
    )

    result = controller.last_result
    if controller.committed and result is not None:
        n_tag = result.placed_count_by_kind.get("tag", 0)
        n_kn = result.placed_count_by_kind.get("keynote", 0)
        n_tn = result.placed_count_by_kind.get("text_note", 0)
        output.print_md(
            "**Place Space Annotations complete**\n\n"
            "- Tags placed:        {}\n"
            "- Keynotes placed:    {}\n"
            "- Text notes placed:  {}\n"
            "- Warnings: {}\n".format(n_tag, n_kn, n_tn, len(result.warnings))
        )
        if result.warnings:
            output.print_md(
                "\n**Warnings:**\n"
                + "\n".join("- {}".format(w) for w in result.warnings[:50])
            )


if __name__ == "__main__":
    main()
