# -*- coding: utf-8 -*-
"""
Shared helpers for resolving the active Let There Be YAML profileData path.
"""

import os

try:
    from pyrevit import revit
except Exception:
    revit = None


def _get_doc(doc=None):
    if doc is not None:
        return doc
    if revit is not None:
        try:
            return getattr(revit, "doc", None)
        except Exception:
            return None
    try:
        return __revit__.ActiveUIDocument.Document
    except Exception:
        return None


def _get_storage():
    try:
        from ExtensibleStorage import ExtensibleStorage
    except Exception:
        return None
    return ExtensibleStorage


def get_cached_yaml_path(doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return None
    storage = _get_storage()
    if storage is None:
        return None
    try:
        path, _, _ = storage.get_active_yaml(doc)
    except Exception:
        return None
    if not path:
        return None
    return os.path.abspath(path)


def set_cached_yaml_path(path, doc=None):
    # Deprecated: settings are stored in-model via Extensible Storage.
    return False


def get_yaml_display_name(active_path=None, doc=None):
    """
    Returns a friendly name for the current YAML file (basename of the cached path).
    """
    path = active_path or get_cached_yaml_path(doc)
    if not path:
        return "selected YAML file"
    try:
        return os.path.basename(path)
    except Exception:
        return path


__all__ = ["get_cached_yaml_path", "set_cached_yaml_path", "get_yaml_display_name"]
