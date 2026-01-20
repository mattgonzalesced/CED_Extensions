# -*- coding: utf-8 -*-
"""
Startup hook for after-sync parent parameter conflict checks.
"""

import imp
import json
import os
import time

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
_IS_RUNNING = False

ENV_HANDLER_KEY = "ced_parent_param_sync_handler_registered"
ENV_LAST_RUN_KEY = "ced_parent_param_sync_last_run"
ENV_RUNNING_KEY = "ced_parent_param_sync_running"


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
    global _IS_RUNNING
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
    if _IS_RUNNING:
        return
    if not _should_run_sync(doc):
        return
    checker = _load_checker()
    if checker is None:
        return
    try:
        _IS_RUNNING = True
        _set_env(ENV_RUNNING_KEY, "1")
        checker.run_sync_check(doc)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Parent param conflict check failed: %s", exc)
    finally:
        _set_env(ENV_RUNNING_KEY, "0")
        _IS_RUNNING = False


def _sync_guard_host():
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    return app or uiapp


def _get_doc_key(doc):
    if doc is None:
        return None
    try:
        return doc.PathName or doc.Title
    except Exception:
        return None


def _get_env(name, default=None):
    try:
        value = script.get_envvar(name)
    except Exception:
        return default
    if value in (None, ""):
        return default
    return value


def _set_env(name, value):
    try:
        script.set_envvar(name, value)
    except Exception:
        return False
    return True


def _load_env_payload(raw_value):
    if isinstance(raw_value, dict):
        return raw_value
    try:
        return json.loads(raw_value)
    except Exception:
        return {}


def _should_run_sync(doc):
    running = _get_env(ENV_RUNNING_KEY)
    if str(running).strip() == "1":
        return False
    doc_key = _get_doc_key(doc)
    if not doc_key:
        return True
    now = time.time()
    payload = _load_env_payload(_get_env(ENV_LAST_RUN_KEY, "{}"))
    last_key = payload.get("doc_key")
    last_ts = payload.get("timestamp") or 0.0
    if last_key == doc_key and now - last_ts < 20.0:
        return False
    payload = {"doc_key": doc_key, "timestamp": now}
    _set_env(ENV_LAST_RUN_KEY, json.dumps(payload))
    return True


def _handler_registry(uiapp):
    if uiapp is None:
        return None
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    host = app or uiapp
    registry = getattr(host, "_ced_parent_param_sync_handlers", None)
    if registry is None:
        registry = {}
        try:
            setattr(host, "_ced_parent_param_sync_handlers", registry)
        except Exception:
            return None
    return registry


def _register_sync_handler():
    global _SYNC_HANDLER_UI, _SYNC_HANDLER_APP
    logger = script.get_logger()
    if EventHandler is None:
        logger.warning("Parent parameter sync handler not registered: EventHandler missing.")
        return
    if _get_env(ENV_HANDLER_KEY):
        return
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    registry = _handler_registry(uiapp)
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    if registry is not None and registry.get("registered"):
        return
    if app is not None and DbSyncArgs is not None and _SYNC_HANDLER_APP is None:
        try:
            handler = EventHandler[DbSyncArgs](_on_doc_sync)
            app.DocumentSynchronizedWithCentral += handler
            _SYNC_HANDLER_APP = handler
            if registry is not None:
                registry["registered"] = "app"
            _set_env(ENV_HANDLER_KEY, "app")
            logger.info("Parent parameter conflict app sync handler registered.")
            return
        except Exception as exc:
            logger.warning("App sync handler not registered: %s", exc)
    if uiapp is not None and UiSyncArgs is not None and _SYNC_HANDLER_UI is None:
        try:
            handler = EventHandler[UiSyncArgs](_on_doc_sync)
            uiapp.DocumentSynchronizedWithCentral += handler
            _SYNC_HANDLER_UI = handler
            if registry is not None:
                registry["registered"] = "ui"
            _set_env(ENV_HANDLER_KEY, "ui")
            logger.info("Parent parameter conflict UI sync handler registered.")
        except Exception as exc:
            logger.warning("UI sync handler not registered: %s", exc)


_register_sync_handler()
