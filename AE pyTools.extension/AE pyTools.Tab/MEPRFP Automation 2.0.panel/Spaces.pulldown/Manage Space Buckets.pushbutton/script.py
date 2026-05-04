#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Manage Space Buckets"""

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
import manage_space_buckets_window

TITLE = "Manage Space Buckets (MEPRFP 2.0)"


def _save_dirty_edits(doc, profile_data, output, action):
    with revit.Transaction(action, doc=doc):
        active_yaml.save_active_data(doc, profile_data, action=action)
    output.print_md("**Space-bucket edits saved.**")


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc) or {}

    if not isinstance(profile_data.get("space_buckets"), list):
        profile_data["space_buckets"] = []

    controller = manage_space_buckets_window.ManageSpaceBucketsController(
        profile_data=profile_data, doc=doc,
    )
    controller.show()

    if controller.dirty:
        _save_dirty_edits(doc, profile_data, output, action="Manage Space Buckets edit")


if __name__ == "__main__":
    main()
