# -*- coding: utf-8 -*-

import sys

from Autodesk.Revit.UI import DockablePane, DockablePaneId
from pyrevit import forms, script, coreutils, HOST_APP

TITLE = "Circuit Manager"
PANEL_ID = "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2"


def _panel_is_shown():
    pane_id = DockablePaneId(coreutils.Guid.Parse(PANEL_ID))
    if not DockablePane.PaneIsRegistered(pane_id):
        return None
    pane = HOST_APP.uiapp.GetDockablePane(pane_id)
    try:
        return bool(pane.IsShown())
    except Exception:
        try:
            return bool(getattr(pane, "IsShown", False))
        except Exception:
            return False


def _refresh_panel_instance():
    module = sys.modules.get("ced_electools_circuit_manager_panel")
    if module is None:
        return
    panel_cls = getattr(module, "CircuitBrowserPanel", None)
    if panel_cls is None or not hasattr(panel_cls, "get_instance"):
        return
    panel = panel_cls.get_instance()
    if panel is None or not hasattr(panel, "refresh_on_open"):
        return
    panel.refresh_on_open()


logger = script.get_logger()
try:
    shown = _panel_is_shown()
    if shown is None:
        raise Exception("Pane is not registered.")
    if shown:
        forms.close_dockable_panel(PANEL_ID)
    else:
        forms.open_dockable_panel(PANEL_ID)
        _refresh_panel_instance()
except Exception as open_exc:
    logger.warning("Circuit Manager open by id failed: %s", open_exc)
    forms.alert(
        "Circuit Manager pane is not available.\n\n"
        "Dockable panes are expected to be registered at startup.\n"
        "Try reloading pyRevit or restarting Revit.\n\n{}"
        .format(open_exc),
        title=TITLE,
    )
