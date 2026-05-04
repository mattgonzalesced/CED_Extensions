#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Manage Space Profiles"""

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
import manage_space_profiles_window

TITLE = "Manage Space Profiles (MEPRFP 2.0)"


def _save_dirty_edits(doc, profile_data, output, action):
    with revit.Transaction(action, doc=doc):
        active_yaml.save_active_data(doc, profile_data, action=action)
    output.print_md("**Space-profile edits saved.**")


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc) or {}

    if not isinstance(profile_data.get("space_profiles"), list):
        profile_data["space_profiles"] = []
    if not isinstance(profile_data.get("space_buckets"), list):
        profile_data["space_buckets"] = []

    if not profile_data["space_buckets"]:
        forms.alert(
            "No space_buckets are defined in the active YAML.\n\n"
            "You can still create profiles, but they need a bucket "
            "reference to apply at placement time. Add at least one "
            "bucket by hand-editing the YAML and re-importing, or open "
            "Manage Space Profiles after importing a starter YAML.",
            title=TITLE,
        )

    controller = manage_space_profiles_window.ManageSpaceProfilesController(
        profile_data=profile_data, doc=doc,
    )
    controller.show()

    if controller.dirty:
        _save_dirty_edits(doc, profile_data, output, action="Manage Space Profiles edit")


if __name__ == "__main__":
    main()
