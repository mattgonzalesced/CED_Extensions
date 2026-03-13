# -*- coding: utf-8 -*-
"""Shared path resolution helpers for CED extension tools."""

import os
import sys


def find_workspace_root(start_dir, marker_dir="CEDLib.lib"):
    """Walk ancestors until a directory containing marker_dir is found."""
    current = os.path.abspath(start_dir)
    marker_name = str(marker_dir or "CEDLib.lib")
    while True:
        if os.path.isdir(os.path.join(current, marker_name)):
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


def find_named_ancestor(start_dir, folder_name):
    """Return the first ancestor whose basename matches folder_name."""
    current = os.path.abspath(start_dir)
    target = str(folder_name or "").strip().lower()
    while True:
        if os.path.basename(current).lower() == target:
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


def resolve_lib_root(start_dir, marker_dir="CEDLib.lib"):
    """Resolve absolute path to CEDLib.lib from a script directory."""
    workspace_root = find_workspace_root(start_dir, marker_dir=marker_dir)
    if workspace_root:
        return os.path.abspath(os.path.join(workspace_root, marker_dir))
    return None


def ensure_lib_root_on_syspath(start_dir, fallback_rel_parts=None):
    """Add CEDLib.lib to sys.path if needed and return resolved path."""
    lib_root = resolve_lib_root(start_dir)
    if lib_root is None:
        rel_parts = fallback_rel_parts
        if rel_parts is None:
            rel_parts = ("..", "..", "..", "..", "..", "CEDLib.lib")
        lib_root = os.path.abspath(os.path.join(os.path.abspath(start_dir), *tuple(rel_parts)))

    if lib_root and os.path.isdir(lib_root) and lib_root not in sys.path:
        sys.path.append(lib_root)
    return lib_root


def resolve_ui_resources_root(lib_root):
    """Resolve UIClasses/Resources path from CEDLib.lib path."""
    if not lib_root:
        return None
    resources_root = os.path.abspath(os.path.join(lib_root, "UIClasses", "Resources"))
    if os.path.isdir(resources_root):
        return resources_root
    return None
