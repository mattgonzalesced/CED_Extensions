# -*- coding: utf-8 -*-
"""
Startup hook for after-sync parent parameter conflict checks.
"""

import getpass
import imp
import json
import os
import shutil
import time

from pyrevit import forms, script

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

_PROX_SYNC_HANDLER_UI = None
_PROX_SYNC_HANDLER_APP = None
_PROX_MODULE = None
_PROX_IS_RUNNING = False

_REF_SYNC_HANDLER_UI = None
_REF_SYNC_HANDLER_APP = None
_REF_MODULE = None
_REF_IS_RUNNING = False

_DOCKABLE_REGISTERED = False

ENV_HANDLER_KEY = "ced_parent_param_sync_handler_registered"
ENV_LAST_RUN_KEY = "ced_parent_param_sync_last_run"
ENV_RUNNING_KEY = "ced_parent_param_sync_running"

PROX_ENV_HANDLER_KEY = "ced_proximity_lights_coils_sync_handler_registered"
PROX_ENV_LAST_RUN_KEY = "ced_proximity_lights_coils_sync_last_run"
PROX_ENV_RUNNING_KEY = "ced_proximity_lights_coils_sync_running"

REF_ENV_HANDLER_KEY = "ced_ref_sched_change_sync_handler_registered"
REF_ENV_LAST_RUN_KEY = "ced_ref_sched_change_sync_last_run"
REF_ENV_RUNNING_KEY = "ced_ref_sched_change_sync_running"


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
    global _IS_RUNNING, _MODULE
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
        try:
            checker.run_sync_check(doc, modeless=True)
        except TypeError as exc:
            logger = script.get_logger()
            logger.warning(
                "Parent conflict checker rejected modeless arg (%s); reloading module and retrying modeless.",
                exc,
            )
            _MODULE = None
            checker = _load_checker()
            if checker is not None:
                checker.run_sync_check(doc, modeless=True)
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


def _proximity_module_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "AE pyTools.Tab",
            "QualityChecks.panel",
            "Proximity Check.pushbutton",
            "proximity_lights_coils.py",
        )
    )


def _load_proximity_checker():
    global _PROX_MODULE
    if _PROX_MODULE is not None:
        return _PROX_MODULE
    path = _proximity_module_path()
    if not os.path.exists(path):
        return None
    try:
        _PROX_MODULE = imp.load_source("ced_proximity_lights_coils", path)
        return _PROX_MODULE
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to load proximity checker: %s", exc)
        return None


def _on_doc_sync_proximity(sender, args):
    global _PROX_IS_RUNNING
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
    if _PROX_IS_RUNNING:
        return
    if not _should_run_proximity_sync(doc):
        return
    checker = _load_proximity_checker()
    if checker is None:
        return
    try:
        _PROX_IS_RUNNING = True
        _set_env(PROX_ENV_RUNNING_KEY, "1")
        checker.run_sync_check(doc)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Proximity check failed: %s", exc)
    finally:
        _set_env(PROX_ENV_RUNNING_KEY, "0")
        _PROX_IS_RUNNING = False


def _should_run_proximity_sync(doc):
    running = _get_env(PROX_ENV_RUNNING_KEY)
    if str(running).strip() == "1":
        return False
    doc_key = _get_doc_key(doc)
    if not doc_key:
        return True
    now = time.time()
    payload = _load_env_payload(_get_env(PROX_ENV_LAST_RUN_KEY, "{}"))
    last_key = payload.get("doc_key")
    last_ts = payload.get("timestamp") or 0.0
    if last_key == doc_key and now - last_ts < 20.0:
        return False
    payload = {"doc_key": doc_key, "timestamp": now}
    _set_env(PROX_ENV_LAST_RUN_KEY, json.dumps(payload))
    return True


def _proximity_handler_registry(uiapp):
    if uiapp is None:
        return None
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    host = app or uiapp
    registry = getattr(host, "_ced_proximity_lights_coils_handlers", None)
    if registry is None:
        registry = {}
        try:
            setattr(host, "_ced_proximity_lights_coils_handlers", registry)
        except Exception:
            return None
    return registry


def _register_proximity_sync_handler():
    global _PROX_SYNC_HANDLER_UI, _PROX_SYNC_HANDLER_APP
    logger = script.get_logger()
    if EventHandler is None:
        logger.warning("Proximity sync handler not registered: EventHandler missing.")
        return
    if _get_env(PROX_ENV_HANDLER_KEY):
        return
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    registry = _proximity_handler_registry(uiapp)
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    if registry is not None and registry.get("registered"):
        return
    if app is not None and DbSyncArgs is not None and _PROX_SYNC_HANDLER_APP is None:
        try:
            handler = EventHandler[DbSyncArgs](_on_doc_sync_proximity)
            app.DocumentSynchronizedWithCentral += handler
            _PROX_SYNC_HANDLER_APP = handler
            if registry is not None:
                registry["registered"] = "app"
            _set_env(PROX_ENV_HANDLER_KEY, "app")
            logger.info("Proximity check app sync handler registered.")
            return
        except Exception as exc:
            logger.warning("App sync handler not registered: %s", exc)
    if uiapp is not None and UiSyncArgs is not None and _PROX_SYNC_HANDLER_UI is None:
        try:
            handler = EventHandler[UiSyncArgs](_on_doc_sync_proximity)
            uiapp.DocumentSynchronizedWithCentral += handler
            _PROX_SYNC_HANDLER_UI = handler
            if registry is not None:
                registry["registered"] = "ui"
            _set_env(PROX_ENV_HANDLER_KEY, "ui")
            logger.info("Proximity check UI sync handler registered.")
        except Exception as exc:
            logger.warning("UI sync handler not registered: %s", exc)


def _ref_sched_module_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "AE pyTools.Tab",
            "QualityChecks.panel",
            "Ref Sched Change.pushbutton",
            "ref_sched_change.py",
        )
    )


def _load_ref_sched_checker():
    global _REF_MODULE
    if _REF_MODULE is not None:
        return _REF_MODULE
    path = _ref_sched_module_path()
    if not os.path.exists(path):
        return None
    try:
        _REF_MODULE = imp.load_source("ced_ref_sched_change", path)
        return _REF_MODULE
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to load Ref Sched Change checker: %s", exc)
        return None


def _on_doc_sync_ref_sched(sender, args):
    global _REF_IS_RUNNING
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
    if _REF_IS_RUNNING:
        return
    if not _should_run_ref_sched_sync(doc):
        return
    checker = _load_ref_sched_checker()
    if checker is None:
        return
    try:
        _REF_IS_RUNNING = True
        _set_env(REF_ENV_RUNNING_KEY, "1")
        checker.run_sync_check(doc, args=args)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Ref Sched Change check failed: %s", exc)
    finally:
        _set_env(REF_ENV_RUNNING_KEY, "0")
        _REF_IS_RUNNING = False


def _should_run_ref_sched_sync(doc):
    running = _get_env(REF_ENV_RUNNING_KEY)
    if str(running).strip() == "1":
        return False
    doc_key = _get_doc_key(doc)
    if not doc_key:
        return True
    now = time.time()
    payload = _load_env_payload(_get_env(REF_ENV_LAST_RUN_KEY, "{}"))
    last_key = payload.get("doc_key")
    last_ts = payload.get("timestamp") or 0.0
    if last_key == doc_key and now - last_ts < 20.0:
        return False
    payload = {"doc_key": doc_key, "timestamp": now}
    _set_env(REF_ENV_LAST_RUN_KEY, json.dumps(payload))
    return True


def _ref_sched_handler_registry(uiapp):
    if uiapp is None:
        return None
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    host = app or uiapp
    registry = getattr(host, "_ced_ref_sched_change_handlers", None)
    if registry is None:
        registry = {}
        try:
            setattr(host, "_ced_ref_sched_change_handlers", registry)
        except Exception:
            return None
    return registry


def _register_ref_sched_sync_handler():
    global _REF_SYNC_HANDLER_UI, _REF_SYNC_HANDLER_APP
    logger = script.get_logger()
    if EventHandler is None:
        logger.warning("Ref Sched Change handler not registered: EventHandler missing.")
        return
    if _get_env(REF_ENV_HANDLER_KEY):
        return
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    registry = _ref_sched_handler_registry(uiapp)
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    if registry is not None and registry.get("registered"):
        return
    if app is not None and DbSyncArgs is not None and _REF_SYNC_HANDLER_APP is None:
        try:
            handler = EventHandler[DbSyncArgs](_on_doc_sync_ref_sched)
            app.DocumentSynchronizedWithCentral += handler
            _REF_SYNC_HANDLER_APP = handler
            if registry is not None:
                registry["registered"] = "app"
            _set_env(REF_ENV_HANDLER_KEY, "app")
            logger.info("Ref Sched Change app sync handler registered.")
            return
        except Exception as exc:
            logger.warning("App sync handler not registered: %s", exc)
    if uiapp is not None and UiSyncArgs is not None and _REF_SYNC_HANDLER_UI is None:
        try:
            handler = EventHandler[UiSyncArgs](_on_doc_sync_ref_sched)
            uiapp.DocumentSynchronizedWithCentral += handler
            _REF_SYNC_HANDLER_UI = handler
            if registry is not None:
                registry["registered"] = "ui"
            _set_env(REF_ENV_HANDLER_KEY, "ui")
            logger.info("Ref Sched Change UI sync handler registered.")
        except Exception as exc:
            logger.warning("UI sync handler not registered: %s", exc)


def _dockable_panel_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "AE pyTools.Tab",
            "MEP Automation.panel",
            "Place Single Profile.pushbutton",
            "PlaceSingleProfilePanel.py",
        )
    )


def _register_place_single_profile_panel():
    global _DOCKABLE_REGISTERED
    if _DOCKABLE_REGISTERED:
        return
    panel_path = _dockable_panel_path()
    if not os.path.exists(panel_path):
        return
    try:
        panel_module = imp.load_source("ced_place_single_profile_panel", panel_path)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to load Place Single Profile panel: %s", exc)
        return
    panel_cls = getattr(panel_module, "PlaceSingleProfilePanel", None)
    if panel_cls is None:
        return
    try:
        if not forms.is_registered_dockable_panel(panel_cls):
            forms.register_dockable_panel(panel_cls, default_visible=False)
        _DOCKABLE_REGISTERED = True
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to register Place Single Profile panel: %s", exc)

def _on_app_closing(sender, args):

    log_data = {
        "username": None,
        "files_found": 0,
        "files_moved": 0,
        "files_failed": 0,
        "status": "unknown",
        "error": None
    }

    try:
        # Username
        try:
            username = getpass.getuser()
        except:
            username = os.environ.get("USERNAME", "UnknownUser")

        log_data["username"] = username

        # Destination
        base_path = r"C:\ACC\ACCDocs\CoolSys\CED Content Collection\Project Files\03 Automations\Usage"
        user_folder = os.path.join(base_path, username)

        try:
            if not os.path.exists(user_folder):
                os.makedirs(user_folder)
        except Exception as e:
            log_data["status"] = "failed_create_user_folder"
            log_data["error"] = str(e)
            # from Snippets import hooks_logger
            # hooks_logger.log_hook(__file__, log_data)
            return

        # Source
        user_home = os.path.expanduser("~")
        source_folder = os.path.join(user_home, "CED_pyTelemetry")

        if not os.path.exists(source_folder):
            log_data["status"] = "no_source_folder"
            # from Snippets import hooks_logger
            # hooks_logger.log_hook(__file__, log_data)
            return

        files = os.listdir(source_folder)
        log_data["files_found"] = len(files)

        for fname in files:
            try:
                src = os.path.join(source_folder, fname)

                if not os.path.isfile(src):
                    continue

                dst = os.path.join(user_folder, fname)

                shutil.move(src, dst)
                log_data["files_moved"] += 1

            except:
                log_data["files_failed"] += 1

        if log_data["files_failed"] > 0:
            log_data["status"] = "partial_success"
        else:
            log_data["status"] = "success"

    except Exception as e:
        log_data["status"] = "fatal_error"
        log_data["error"] = str(e)

    # Always log
    # try:
    #     from Snippets import hooks_logger
    #     hooks_logger.log_hook(__file__, log_data)
    # except:
    #     pass

def _register_shutdown_hook():
    logger = script.get_logger()

    try:
        app = __revit__
        app.ApplicationClosing += _on_app_closing
        logger.info("ApplicationClosing hook registered.")

    except Exception as exc:
        logger.warning("Failed to register ApplicationClosing hook: %s", exc)

_register_shutdown_hook()
_register_sync_handler()
_register_proximity_sync_handler()
_register_ref_sched_sync_handler()
_register_place_single_profile_panel()
_register_circuit_browser_panel()
# Temporarily disabled to prevent startup-time dockable panel activity.
# _register_place_single_profile_panel()
