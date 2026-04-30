#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place from CAD or Linked Model"""

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
import placement_window
import shared_params

TITLE = "Place from CAD or Linked Model (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    if not shared_params.is_element_linker_bound(doc):
        if not forms.confirm(
            "The MEPRFP 2.0 Element_Linker shared parameter is not bound in this project.\n"
            "Bind it now? (required for placement to write back to placed elements)",
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
    if not profile_data.get("equipment_definitions"):
        forms.alert(
            "No profiles in the active store. "
            "Use 'New Profile' or 'Import YAML File' first.",
            title=TITLE,
        )
        return

    controller = placement_window.show_modal(doc, profile_data)

    result = getattr(controller, "_last_result", None)
    if controller.committed and result is not None:
        output.print_md(
            "**Placement run complete**\n\n"
            "- Fixtures placed: {}\n"
            "- Element_Linker writes: {}\n"
            "- Static parameter writes: {}\n"
            "- Already-placed (skipped): {}\n"
            "- Normalized-name matches: {}\n"
            "- Type substitutions: {}\n"
            "- Warnings: {}\n".format(
                result.placed_fixture_count,
                result.element_linker_writes,
                getattr(result, "static_param_writes", 0),
                result.skipped_already_placed,
                getattr(result, "normalized_match_count", 0),
                getattr(result, "substituted_type_count", 0),
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
