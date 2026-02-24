# -*- coding: utf-8 -*-
"""
Create an independent profile from selected elements.
"""

import imp
import os

from pyrevit import forms, revit

TITLE = "Create Independent Profile"


def _manage_profiles_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "..",
            "Modify Profiles.stack",
            "Modify Profiles.pulldown",
            "Manage Profiles.pushbutton",
            "script.py",
        )
    )


def _load_manage_profiles():
    path = _manage_profiles_path()
    if not os.path.exists(path):
        forms.alert("Manage Profiles script not found.", title=TITLE)
        return None
    try:
        return imp.load_source("ced_manage_profiles", path)
    except Exception as exc:
        forms.alert("Failed to load Manage Profiles:\n\n{}".format(exc), title=TITLE)
        return None


def _build_state(module, data_path, raw_data):
    label = None
    try:
        label = module.get_yaml_display_name(data_path)
    except Exception:
        label = data_path or "active YAML"
    return {
        "raw_data": raw_data or {},
        "yaml_label": label,
        "yaml_path": data_path,
        "normalized_yaml_path": module._normalize_yaml_path(data_path),
    }


def _refresh_state(module, state):
    try:
        data_path, raw_data = module.load_active_yaml_data()
    except Exception:
        return
    if raw_data is not None:
        state["raw_data"] = raw_data
    if data_path:
        state["yaml_path"] = data_path
        try:
            state["normalized_yaml_path"] = module._normalize_yaml_path(data_path)
        except Exception:
            pass
        try:
            state["yaml_label"] = module.get_yaml_display_name(data_path)
        except Exception:
            pass


def main():
    module = _load_manage_profiles()
    if module is None:
        return

    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    try:
        data_path, raw_data = module.load_active_yaml_data()
    except Exception as exc:
        forms.alert("Failed to load active YAML:\n\n{}".format(exc), title=TITLE)
        return

    if raw_data is None:
        forms.alert("No active YAML data found.", title=TITLE)
        return

    cad_name = forms.ask_for_string(
        prompt="Enter a name for the new independent profile:",
        default="",
    )
    if not cad_name:
        return
    cad_name = cad_name.strip()
    if not cad_name:
        return

    state = _build_state(module, data_path, raw_data)

    def _refresh():
        _refresh_state(module, state)

    try:
        module._capture_orphan_profile(doc, cad_name, state, _refresh, state.get("yaml_label") or "")
    except Exception as exc:
        forms.alert("Failed to create independent profile:\n\n{}".format(exc), title=TITLE)


if __name__ == "__main__":
    main()
