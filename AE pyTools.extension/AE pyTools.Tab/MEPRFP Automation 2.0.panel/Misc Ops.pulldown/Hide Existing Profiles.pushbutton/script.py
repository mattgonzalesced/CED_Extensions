#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Hide Existing Profiles"""

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
import hide_profiles_window

TITLE = "Hide Existing Profiles (MEPRFP 2.0)"


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
    controller = hide_profiles_window.show_modal(doc, profile_data)
    result = getattr(controller, "_last_result", None)
    if controller.committed and result:
        output.print_md(
            "**Hide Existing Profiles complete**\n\n"
            "- Duplicated view: `{}`\n"
            "- Host elements hidden: {}\n"
            "- Linked elements hidden: {}\n"
            "- Warnings: {}\n".format(
                result.get("view_name"),
                result.get("host_count"),
                result.get("link_count"),
                len(result.get("warnings") or []),
            )
        )
        if result.get("warnings"):
            output.print_md(
                "\n**Warnings:**\n"
                + "\n".join("- {}".format(w) for w in result["warnings"])
            )


if __name__ == "__main__":
    main()
