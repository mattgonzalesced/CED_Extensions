# -*- coding: utf-8 -*-
"""
Parameter Flag Settings
-----------------------
Toggle and test parent-parameter conflict checks that run after sync.
"""

import imp
import os

from pyrevit import forms, revit

TITLE = "Parameter Flag Settings"


def _load_module():
    module_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "parent_param_conflicts.py"))
    if not os.path.exists(module_path):
        forms.alert("Conflict checker not found at:\n{}".format(module_path), title=TITLE)
        return None
    try:
        return imp.load_source("ced_parent_param_conflicts_settings", module_path)
    except Exception as exc:
        forms.alert("Failed to load conflict checker:\n{}\n\n{}".format(module_path, exc), title=TITLE)
        return None


def main():
    module = _load_module()
    if module is None:
        return
    enabled = module.get_setting(default=True)
    status = "enabled" if enabled else "disabled"
    options = [
        "Enable after-sync check",
        "Disable after-sync check",
        "Run check now",
    ]
    result = forms.CommandSwitchWindow.show(
        options,
        message="Parent parameter conflict checks are currently {}.".format(status),
        title=TITLE,
    )
    if not result:
        return
    if result == "Enable after-sync check":
        module.set_setting(True)
        forms.alert("After-sync parent parameter checks are now enabled.", title=TITLE)
    elif result == "Disable after-sync check":
        module.set_setting(False)
        forms.alert("After-sync parent parameter checks are now disabled.", title=TITLE)
    elif result == "Run check now":
        doc = getattr(revit, "doc", None)
        if doc is None:
            forms.alert("No active document detected.", title=TITLE)
            return
        module.run_sync_check(doc)


if __name__ == "__main__":
    main()
