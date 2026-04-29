#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Check Synced Relationships"""

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
import sync_audit_window

TITLE = "Check Synced Relationships (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert(
            "No profiles in the active store. Nothing to audit.",
            title=TITLE,
        )
        return

    def transaction_factory():
        return revit.Transaction("Apply Synced Relationship resolutions", doc=doc)

    sync_audit_window.show_modal(doc, profile_data, transaction_factory)


if __name__ == "__main__":
    main()
