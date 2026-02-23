# -*- coding: utf-8 -*-
"""
Proximity Check: Lights-Coils
Configure after-sync flag and run the check on demand.
"""

__title__ = "Proximity Check\nLights-Coils"
__doc__ = "Check lighting fixtures within 18 inches of CED-R-KRACK coils."

import os
import imp

from pyrevit import forms, revit

TITLE = "Proximity Check: Lights-Coils"


def _load_module():
    module_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "proximity_lights_coils.py"))
    if not os.path.exists(module_path):
        forms.alert("Checker not found at:\n{}".format(module_path), title=TITLE)
        return None
    try:
        return imp.load_source("ced_proximity_lights_coils", module_path)
    except Exception as exc:
        forms.alert("Failed to load checker:\n{}\n\n{}".format(module_path, exc), title=TITLE)
        return None


def main():
    module = _load_module()
    if module is None:
        return
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected; project-based settings require an open document.", title=TITLE)
        return
    enabled = module.get_setting(default=True, doc=doc)
    status = "enabled" if enabled else "disabled"
    options = [
        "Enable after-sync check",
        "Disable after-sync check",
        "Run check now",
    ]
    result = forms.CommandSwitchWindow.show(
        options,
        message="Proximity checks are currently {}.".format(status),
        title=TITLE,
    )
    if not result:
        return
    if result == "Enable after-sync check":
        module.set_setting(True, doc=doc)
        forms.alert("After-sync proximity checks are now enabled.", title=TITLE)
    elif result == "Disable after-sync check":
        module.set_setting(False, doc=doc)
        forms.alert("After-sync proximity checks are now disabled.", title=TITLE)
    elif result == "Run check now":
        module.run_check(doc, show_ui=True, show_empty=True)


if __name__ == "__main__":
    main()
