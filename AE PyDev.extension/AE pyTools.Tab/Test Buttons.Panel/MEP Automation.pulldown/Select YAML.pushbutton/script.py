# -*- coding: utf-8 -*-
"""
Select default equipment-definition YAML for Let There Be YAML tools.
"""

import os
import sys
import io

from pyrevit import forms, revit

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402
from ExtensibleStorage.yaml_store import seed_active_yaml  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")


def main():
    cached = get_cached_yaml_path()
    init_dir = os.path.dirname(cached) if cached else os.path.dirname(DEFAULT_DATA_PATH)
    picked = forms.pick_file(
        file_ext="yaml",
        title="Select default equipment definition YAML",
        init_dir=init_dir,
    )
    if not picked:
        return
    set_cached_yaml_path(picked)
    with io.open(picked, "r", encoding="utf-8") as handle:
        raw_text = handle.read()
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected; cannot store YAML in Extensible Storage.", title="Select YAML")
        return
    seed_active_yaml(doc, picked, raw_text)
    forms.alert(
        "Loaded '{}' into the project. All YAML operations now run from Extensible Storage.\n"
        "The original file will remain untouched until you export it again.".format(picked),
        title="Select YAML",
    )


if __name__ == "__main__":
    main()
