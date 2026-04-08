# -*- coding: utf-8 -*-
"""Shared path resolution helpers for CED extension tools."""

import os
import sys


def _get_runtime_envvar(name):
    key = str(name or "").strip()
    if not key:
        return None
    try:
        from pyrevit import script as pyrevit_script  # type: ignore

        value = pyrevit_script.get_envvar(key)
    except Exception:
        value = None
    if value in (None, ""):
        value = os.environ.get(key)
    return value


def find_workspace_root(start_dir, marker_dir="CEDLib.lib"):
    """Walk ancestors until a directory containing marker_dir is found."""
    seeded_workspace = _get_runtime_envvar("CED_WORKSPACE_ROOT")
    if seeded_workspace and os.path.isdir(str(seeded_workspace)):
        marker_path = os.path.join(str(seeded_workspace), str(marker_dir or "CEDLib.lib"))
        if os.path.isdir(marker_path):
            return os.path.abspath(str(seeded_workspace))

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
    seeded_lib = _get_runtime_envvar("CED_LIB_ROOT")
    if seeded_lib and os.path.isdir(str(seeded_lib)):
        return os.path.abspath(str(seeded_lib))

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


def resolve_start_dir(module_name=None, fallback_dir=None):
    """Resolve absolute start directory for a module name or fallback path."""
    module = None
    if module_name:
        try:
            module = sys.modules.get(module_name)
        except Exception:
            module = None
    module_file = getattr(module, "__file__", None) if module is not None else None
    if module_file:
        try:
            return os.path.abspath(os.path.dirname(module_file))
        except Exception:
            pass
    if fallback_dir:
        try:
            return os.path.abspath(fallback_dir)
        except Exception:
            pass
    return os.path.abspath(os.getcwd())


def resolve_ui_context(start_dir, marker_dir="CEDLib.lib", fallback_rel_parts=None, ensure_syspath=True):
    """Resolve workspace/lib/resources context for UI tools.

    Returns:
        dict with keys:
            start_dir, workspace_root, lib_root, resources_root
    """
    start_abs = os.path.abspath(start_dir)
    workspace_root = find_workspace_root(start_abs, marker_dir=marker_dir)
    lib_root = resolve_lib_root(start_abs, marker_dir=marker_dir)
    if not lib_root:
        rel_parts = fallback_rel_parts
        if rel_parts is None:
            rel_parts = ("..", "..", "..", "..", "..", "CEDLib.lib")
        lib_root = os.path.abspath(os.path.join(start_abs, *tuple(rel_parts)))
    if ensure_syspath and lib_root and os.path.isdir(lib_root) and lib_root not in sys.path:
        sys.path.append(lib_root)
    resources_root = resolve_ui_resources_root(lib_root)
    return {
        "start_dir": start_abs,
        "workspace_root": workspace_root,
        "lib_root": lib_root,
        "resources_root": resources_root,
    }
