#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Manage Profiles"""

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
import append_workflow
import manage_profiles_window

TITLE = "Manage Profiles (MEPRFP 2.0)"


def _save_dirty_edits(doc, profile_data, output, action):
    with revit.Transaction(action, doc=doc):
        active_yaml.save_active_data(doc, profile_data, action=action)
    output.print_md("**Profile edits saved.**")


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
        forms.alert(
            "No profiles in the active store. "
            "Use 'New Profile' or 'Import YAML File' first.",
            title=TITLE,
        )
        return

    controller = manage_profiles_window.ManageProfilesController(profile_data)
    controller.show()

    # Persist any in-memory edits the user made before clicking Add LED
    # or before plain-closing the window.
    if controller.dirty:
        _save_dirty_edits(doc, profile_data, output, action="Manage Profiles edit")

    # If the user clicked Add LED..., run the append workflow on that
    # profile. The workflow opens its own transactions and saves.
    if controller.requested_add_to_profile_id:
        append_workflow.run(
            doc=doc,
            uidoc=uidoc,
            profile_data=profile_data,
            forms=forms,
            wpf_dialogs=wpf_dialogs,
            output=output,
            title="Add LED via Manage Profiles",
            profile_id=controller.requested_add_to_profile_id,
        )


if __name__ == "__main__":
    main()
