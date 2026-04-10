# -*- coding: utf-8 -*-

import imp
import os

from pyrevit import forms, script

PANEL_MODULE_NAME = "ced_circuit_browser_panel"
PANEL_CLASS_NAME = "CircuitBrowserPanel"
PANEL_ID = "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2"


def _panel_module_path():
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "CircuitBrowserPanel.py"))


def _load_panel_class():
    module = imp.load_source(PANEL_MODULE_NAME, _panel_module_path())
    return getattr(module, PANEL_CLASS_NAME, None)


def _ensure_registered():
    logger = script.get_logger()
    panel_cls = _load_panel_class()
    if panel_cls:
        try:
            if not forms.is_registered_dockable_panel(panel_cls):
                forms.register_dockable_panel(panel_cls, default_visible=False)
                logger.info("Circuit Manager panel registered from button command.")
        except Exception as reg_exc:
            logger.warning("Circuit Manager register from button failed: %s", reg_exc)
    return panel_cls


panel_class = _ensure_registered()
logger = script.get_logger()
try:
    forms.open_dockable_panel(PANEL_ID)
    if panel_class and hasattr(panel_class, "get_instance"):
        panel = panel_class.get_instance()
        if panel and hasattr(panel, "refresh_on_open"):
            panel.refresh_on_open()
except Exception as open_exc:
    logger.warning("Circuit Manager open by id failed: %s", open_exc)
    panel_class = _ensure_registered()
    if not panel_class:
        forms.alert(
            "Circuit Manager panel class could not be loaded.\n\n{}".format(open_exc),
            title="Circuit Manager",
        )
    else:
        try:
            forms.open_dockable_panel(PANEL_ID)
            if hasattr(panel_class, "get_instance"):
                panel = panel_class.get_instance()
                if panel and hasattr(panel, "refresh_on_open"):
                    panel.refresh_on_open()
        except Exception as open_exc2:
            logger.warning("Circuit Manager open by id after register failed: %s", open_exc2)
            forms.alert(
                "Circuit Manager pane is still not available.\n\n{}\n\n{}"
                .format(open_exc, open_exc2),
                title="Circuit Manager",
            )
