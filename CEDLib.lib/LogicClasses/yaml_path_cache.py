# -*- coding: utf-8 -*-
"""
Shared helpers for persisting the active Let There Be YAML profileData path.
"""

import io
import json
import os


CONFIG_FILE = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "LetThereBeYAML.settings.json")
)


def _read_config():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with io.open(CONFIG_FILE, "r", encoding="utf-8") as handle:  # type: ignore # noqa
            return json.load(handle)
    except Exception:
        try:
            with open(CONFIG_FILE, "r") as handle:
                return json.load(handle)
        except Exception:
            return {}


def _write_config(data):
    directory = os.path.dirname(CONFIG_FILE)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)
    with open(CONFIG_FILE, "w") as handle:
        json.dump(data, handle, indent=2)


def get_cached_yaml_path():
    data = _read_config()
    path = data.get("yaml_path")
    if not path:
        return None
    return os.path.abspath(path)


def set_cached_yaml_path(path):
    if not path:
        return
    data = _read_config()
    data["yaml_path"] = os.path.abspath(path)
    _write_config(data)


__all__ = ["get_cached_yaml_path", "set_cached_yaml_path"]
