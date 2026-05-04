#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place Space Elements"""

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
import place_space_elements_window

TITLE = "Place Space Elements (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc) or {}
    if not profile_data.get("space_profiles"):
        forms.alert(
            "No space_profiles in the active YAML. "
            "Use Manage Space Profiles to define some first.",
            title=TITLE,
        )
        return

    place_space_elements_window.show_modal(doc=doc, profile_data=profile_data)


if __name__ == "__main__":
    main()
