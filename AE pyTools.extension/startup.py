# -*- coding: utf-8 -*-
"""
Startup hook for after-sync parent parameter conflict checks.
"""

import imp
import os

from pyrevit import script

try:
    from Autodesk.Revit.UI.Events import DocumentSynchronizedWithCentralEventArgs as UiSyncArgs
except Exception:
    UiSyncArgs = None

try:
    from Autodesk.Revit.DB.Events import DocumentSynchronizedWithCentralEventArgs as DbSyncArgs
except Exception:
    DbSyncArgs = None

try:
    from System import EventHandler
except Exception:
    EventHandler = None

_SYNC_HANDLER_UI = None
_SYNC_HANDLER_APP = None
_MODULE = None


def _module_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "AE pyTools.Tab",
            "MEP Automation.panel",
            "Parameter Flag Settings.pushbutton",
            "parent_param_conflicts.py",
        )
    )


def _load_checker():
    global _MODULE
    if _MODULE is not None:
        return _MODULE
    path = _module_path()
    if not os.path.exists(path):
        return None
    try:
        _MODULE = imp.load_source("ced_parent_param_conflicts", path)
        return _MODULE
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to load parent param conflict checker: %s", exc)
        return None


def _on_doc_sync(sender, args):
    doc = None
    try:
        doc = getattr(args, "Document", None)
    except Exception:
        doc = None
    if doc is None:
        try:
            doc = __revit__.ActiveUIDocument.Document
        except Exception:
            doc = None
    if doc is None:
        return
    checker = _load_checker()
    if checker is None:
        return
    try:
        checker.run_sync_check(doc)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Parent param conflict check failed: %s", exc)


def _register_sync_handler():
    global _SYNC_HANDLER_UI, _SYNC_HANDLER_APP
    logger = script.get_logger()
    if EventHandler is None:
        logger.warning("Parent parameter sync handler not registered: EventHandler missing.")
        return
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    if uiapp is not None and UiSyncArgs is not None and _SYNC_HANDLER_UI is None:
        try:
            handler = EventHandler[UiSyncArgs](_on_doc_sync)
            uiapp.DocumentSynchronizedWithCentral += handler
            _SYNC_HANDLER_UI = handler
            logger.info("Parent parameter conflict UI sync handler registered.")
        except Exception as exc:
            logger.warning("UI sync handler not registered: %s", exc)
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    if app is not None and DbSyncArgs is not None and _SYNC_HANDLER_APP is None:
        try:
            handler = EventHandler[DbSyncArgs](_on_doc_sync)
            app.DocumentSynchronizedWithCentral += handler
            _SYNC_HANDLER_APP = handler
            logger.info("Parent parameter conflict app sync handler registered.")
        except Exception as exc:
            logger.warning("App sync handler not registered: %s", exc)


_register_sync_handler()
