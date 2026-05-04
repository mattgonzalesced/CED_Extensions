#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: SuperCircuit V5"""

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
import circuit_clients
import circuit_window
import circuit_workflow

TITLE = "SuperCircuit V5 (MEPRFP 2.0)"


def _pick_client():
    options = circuit_clients.all_clients()
    if not options:
        return None
    chosen = wpf_dialogs.pick_from_list(
        options,
        title=TITLE,
        prompt="Pick the client whose circuiting rules apply:",
        display_func=lambda c: c.display_name or c.key or "?",
    )
    return chosen


def _pick_scope(uidoc):
    selected_ids = []
    if uidoc is not None:
        try:
            selected_ids = list(uidoc.Selection.GetElementIds() or [])
        except Exception:
            selected_ids = []
    if not selected_ids:
        return circuit_workflow.SCOPE_ALL, []
    if forms.confirm(
        "You have {} element(s) selected. Limit SuperCircuit to the "
        "selection only? Choose No to walk every eligible element in "
        "the document.".format(len(selected_ids)),
        title=TITLE,
    ):
        return circuit_workflow.SCOPE_SELECTION, selected_ids
    return circuit_workflow.SCOPE_ALL, []


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    client = _pick_client()
    if client is None:
        return

    scope, selected_ids = _pick_scope(uidoc)

    profile_data = active_yaml.load_active_data(doc)

    circuit_window.show_modeless(
        doc=doc,
        uidoc=uidoc,
        client=client,
        scope=scope,
        selected_element_ids=selected_ids,
        profile_data=profile_data,
    )


if __name__ == "__main__":
    main()
