# -*- coding: utf-8 -*-
"""Startup hooks for CED ElecTools extension."""

import imp
import os
import sys

from pyrevit import forms, script

_CIRCUIT_MANAGER_REGISTERED = False


def _extension_root():
    return os.path.abspath(os.path.dirname(__file__))


def _workspace_root():
    return os.path.abspath(os.path.join(_extension_root(), ".."))


def _cedlib_root():
    return os.path.abspath(os.path.join(_workspace_root(), "CEDLib.lib"))


def _seed_runtime_paths():
    logger = script.get_logger()
    ext_root = _extension_root()
    lib_root = _cedlib_root()
    if os.path.isdir(lib_root) and lib_root not in sys.path:
        sys.path.append(lib_root)
        logger.info("CEDLib added to sys.path from CED ElecTools startup.")
    try:
        script.set_envvar("CED_EXTENSION_ROOT", ext_root)
    except Exception:
        pass
    try:
        script.set_envvar("CED_WORKSPACE_ROOT", _workspace_root())
    except Exception:
        pass
    try:
        script.set_envvar("CED_LIB_ROOT", lib_root if os.path.isdir(lib_root) else "")
    except Exception:
        pass


def _find_circuit_manager_panel_path():
    root = os.path.abspath(os.path.dirname(__file__))
    for current_root, _, files in os.walk(root):
        if "CircuitBrowserPanel.py" not in files:
            continue
        path = os.path.abspath(os.path.join(current_root, "CircuitBrowserPanel.py"))
        normalized = path.replace("/", "\\").lower()
        if "\\electrical.panel\\circuit manager.pushbutton\\" in normalized:
            return path
    return None


def _register_circuit_manager_panel():
    global _CIRCUIT_MANAGER_REGISTERED
    logger = script.get_logger()
    if _CIRCUIT_MANAGER_REGISTERED:
        logger.info("Circuit Manager panel registration skipped (already registered in this startup run).")
        return

    panel_path = _find_circuit_manager_panel_path()
    if not panel_path or not os.path.exists(panel_path):
        logger.warning("Circuit Manager panel file not found under CED ElecTools extension.")
        return

    try:
        panel_module = imp.load_source("ced_electools_circuit_manager_panel", panel_path)
    except Exception as exc:
        logger.warning("Failed to load Circuit Manager panel: %s", exc)
        return

    panel_cls = getattr(panel_module, "CircuitBrowserPanel", None)
    if panel_cls is None:
        logger.warning("Circuit Manager panel class not found in: %s", panel_path)
        return

    try:
        if not forms.is_registered_dockable_panel(panel_cls):
            forms.register_dockable_panel(panel_cls, default_visible=False)
            logger.info("Circuit Manager panel registered successfully.")
        else:
            logger.info("Circuit Manager panel already registered.")
        _CIRCUIT_MANAGER_REGISTERED = True
    except Exception as exc:
        logger.warning("Failed to register Circuit Manager panel: %s", exc)


_seed_runtime_paths()
_register_circuit_manager_panel()
