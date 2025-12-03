# -*- coding: utf-8 -*-
"""
Select default profileData.yaml for Let There Be YAML tools.
"""

import os
import sys

from pyrevit import forms

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")


def main():
    cached = get_cached_yaml_path()
    init_dir = os.path.dirname(cached) if cached else os.path.dirname(DEFAULT_DATA_PATH)
    picked = forms.pick_file(
        file_ext="yaml",
        title="Select default profileData YAML file",
        init_dir=init_dir,
    )
    if not picked:
        return
    set_cached_yaml_path(picked)
    forms.alert(
        "Set Let There Be YAML default file to:\n\n{}".format(picked),
        title="Select profileData",
    )


if __name__ == "__main__":
    main()
