# -*- coding: utf-8 -*-
"""
Helpers for working with the active YAML stored inside Extensible Storage.
"""

from pyrevit import revit, script

from profile_schema import load_data_from_text, dump_data_to_string  # noqa: E402
from ExtensibleStorage import ExtensibleStorage  # noqa: E402


def _get_doc(doc=None):
    if doc:
        return doc
    return getattr(revit, "doc", None)


def load_active_yaml_text(doc=None):
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    path, normalized, text = ExtensibleStorage.get_active_yaml(doc)
    if not path or text is None:
        raise RuntimeError("Select YAML first so the profile data is loaded into the project.")
    return path, text


def load_active_yaml_data(doc=None):
    path, text = load_active_yaml_text(doc)
    data = load_data_from_text(text, path)
    logger = script.get_logger()
    logger.info("[YAML Storage] loaded equipment definitions: %s", [eq.get("name") or eq.get("id") for eq in data.get("equipment_definitions") or [] if isinstance(eq, dict)])
    return path, data


def save_active_yaml_data(doc, data, action, description):
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    path, text = load_active_yaml_text(doc)
    new_text = dump_data_to_string(data)
    if new_text == text:
        return
    ExtensibleStorage.update_active_yaml(doc, path, text, new_text, action, description)


def seed_active_yaml(doc, yaml_path, raw_text):
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    ExtensibleStorage.seed_active_yaml(doc, yaml_path, raw_text or "")
