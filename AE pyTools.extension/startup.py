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

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit import forms, script, telemetry

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

_DOCKABLE_REGISTERED = False

def _telemetry_source_folder():
    appdata = os.environ.get("APPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Roaming"))
    return os.path.join(appdata, "pyRevit", "Extensions", "CED_pyTelemetry")


def _ensure_telemetry_source_folder():
    source_folder = _telemetry_source_folder()
    if os.path.exists(source_folder):
        return source_folder, True, None
    try:
        os.makedirs(source_folder)
        return source_folder, True, None
    except Exception as exc:
        return source_folder, False, exc

def _normalize_path(value):
    if value in (None, ""):
        return ""
    return os.path.normcase(os.path.normpath(value))


def _event_flags_to_int(value):
    if value in (None, ""):
        return 0
    try:
        value_text = str(value).strip()
        if value_text.lower().startswith("0x"):
            return int(value_text, 16)
        return int(value_text)
    except Exception:
        return 0


def _configure_pyrevit_telemetry():
    logger = script.get_logger()
    source_folder, folder_ok, folder_error = _ensure_telemetry_source_folder()
    if not folder_ok:
        logger.warning("Telemetry folder not available: %s", folder_error)
        return

    try:
        telemetry_cfg = script.get_config("telemetry")

        expected_settings = {
            "utc_timestamps": True,
            "active": True,
            "telemetry_file_dir": source_folder,
            "telemetry_server_url": "",
            "include_hooks": True,
            "active_app": False,
            "apptelemetry_server_url": "",
            "apptelemetry_event_flags": "0x0",
        }

        current_settings = {
            setting_name: telemetry_cfg.get_option(setting_name, default_value="")
            for setting_name in expected_settings
        }

        setting_setters = {
            "utc_timestamps": telemetry.set_telemetry_utc_timestamp,
            "active": telemetry.set_telemetry_state,
            "telemetry_file_dir": telemetry.set_telemetry_file_dir,
            "telemetry_server_url": telemetry.set_telemetry_server_url,
            "include_hooks": telemetry.set_telemetry_include_hooks,
            "active_app": telemetry.set_apptelemetry_state,
            "apptelemetry_server_url": telemetry.set_apptelemetry_server_url,
            "apptelemetry_event_flags": lambda _: telemetry.set_apptelemetry_event_flags(0),
        }

        value_normalizers = {
            "telemetry_file_dir": _normalize_path,
            "apptelemetry_event_flags": _event_flags_to_int,
        }

        changed_settings = []
        for setting_name, expected_value in expected_settings.items():
            current_value = current_settings.get(setting_name)
            normalizer = value_normalizers.get(setting_name)
            if normalizer:
                current_value = normalizer(current_value)
                expected_value = normalizer(expected_value)
            if current_value != expected_value:
                setting_setters[setting_name](expected_settings[setting_name])
                changed_settings.append(setting_name)

        if changed_settings:
            # setup_telemetry() applies derived runtime state (session file path,
            # handlers, env vars) and persists the updated config once.
            telemetry.setup_telemetry()
            logger.info(
                "pyRevit telemetry updated via telemetry API. changed=%s file_dir=%s",
                ", ".join(changed_settings),
                source_folder,
            )
        else:
            logger.info(
                "pyRevit telemetry already matched required settings. "
                "No config write needed."
            )
    except Exception as exc:
        logger.warning("Failed to configure pyRevit telemetry: %s", exc)

def _find_acc_root():
    candidates = [
        r"C:\ACC\ACCDocs\CoolSys\CED Content Collection",
        os.path.join(os.path.expanduser("~"), "DC", "ACCDocs", "CoolSys", "CED Content Collection"),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

ENV_HANDLER_KEY = "ced_parent_param_sync_handler_registered"
ENV_LAST_RUN_KEY = "ced_parent_param_sync_last_run"
ENV_RUNNING_KEY = "ced_parent_param_sync_running"
ENV_APP_CLOSING_HANDLER_KEY = "ced_app_closing_handler_registered"


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

        # Destination — only proceed if ACC is actually synced
        acc_root = _find_acc_root()
        if acc_root is None:
            log_data["status"] = "acc_not_synced"
            return
        base_path = os.path.join(acc_root, "Project Files", "03 Automations", "Usage")
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
        source_folder = _telemetry_source_folder()

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
    if _get_env(ENV_APP_CLOSING_HANDLER_KEY):
        logger.info("ApplicationClosing hook already registered; skipping.")
        return

    try:
        app = __revit__
        if app is None:
            logger.warning("ApplicationClosing hook not registered: UIApplication unavailable.")
            return
        app.ApplicationClosing += _on_app_closing
        _set_env(ENV_APP_CLOSING_HANDLER_KEY, "1")
        logger.info("ApplicationClosing hook registered.")

    except Exception as exc:
        logger.warning("Failed to register ApplicationClosing hook: %s", exc)

def _check_acc_sync():
    if _find_acc_root() is not None:
        return
    from System.Windows import Window, SizeToContent, WindowStartupLocation, Thickness, TextWrapping, HorizontalAlignment
    from System.Windows.Controls import StackPanel, Image, TextBlock, Button, ScrollViewer
    from System.Windows.Media.Imaging import BitmapImage
    from System import Uri, UriKind

    img_dir = os.path.normpath(os.path.join(
        os.path.dirname(os.path.abspath(__file__)), os.pardir,
        "WM Tools.extension", "AE pyTools.Tab", "WM Tools.panel",
        "WM Tools.pulldown", "Load Electrical Content.pushbutton",
    ))
    sync_img = os.path.join(img_dir, "sync_instruction.png")
    explorer_img = os.path.join(img_dir, "file_explorer_instruction.png")

    win = Window()
    win.Title = "ACC Sync Required"
    win.SizeToContent = SizeToContent.Width
    win.Height = 700
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    scroll = ScrollViewer()
    panel = StackPanel()
    panel.Margin = Thickness(15)

    req = TextBlock()
    req.Text = "REQUIRED FOR COOLSYS EMPLOYEES:"
    req.FontSize = 14
    req.FontWeight = __import__("System.Windows", fromlist=["FontWeights"]).FontWeights.Bold
    req.Margin = Thickness(0, 0, 0, 5)
    panel.Children.Add(req)

    header = TextBlock()
    header.Text = "CED Content Collection is not synced"
    header.FontSize = 16
    header.FontWeight = __import__("System.Windows", fromlist=["FontWeights"]).FontWeights.Bold
    header.Margin = Thickness(0, 0, 0, 10)
    panel.Children.Add(header)

    msg = TextBlock()
    msg.TextWrapping = TextWrapping.Wrap
    msg.MaxWidth = 620
    msg.Text = (
        "This extension requires the CED Content Collection ACC project "
        "to be synced via Autodesk Desktop Connector.\n\n"
        "1. Click the Desktop Connector tray icon on your taskbar.\n"
        "2. Click 'Select Projects' and check 'CED Content Collection' "
        "from the CoolSys directory.\n"
        "3. Once synced, restart Revit."
    )
    msg.Margin = Thickness(0, 0, 0, 15)
    panel.Children.Add(msg)

    for img_path, caption, max_w in [(sync_img, "Select Projects in Desktop Connector", 620),
                                      (explorer_img, "ACC folder in File Explorer", 310)]:
        if os.path.exists(img_path):
            lbl = TextBlock()
            lbl.Text = caption
            lbl.FontWeight = __import__("System.Windows", fromlist=["FontWeights"]).FontWeights.SemiBold
            lbl.Margin = Thickness(0, 0, 0, 5)
            panel.Children.Add(lbl)
            img = Image()
            img.Source = BitmapImage(Uri(img_path, UriKind.Absolute))
            img.MaxWidth = max_w
            img.HorizontalAlignment = HorizontalAlignment.Left
            img.Margin = Thickness(0, 0, 0, 15)
            panel.Children.Add(img)

    btn = Button()
    btn.Content = "OK"
    btn.Width = 80
    btn.Height = 28
    btn.HorizontalAlignment = HorizontalAlignment.Center
    btn.Click += lambda s, e: win.Close()
    panel.Children.Add(btn)

    scroll.Content = panel
    win.Content = scroll
    win.ShowDialog()

_configure_pyrevit_telemetry()
_check_acc_sync()
_register_shutdown_hook()
#_register_sync_handler()
# Temporarily disabled to prevent startup-time dockable panel activity.
# _register_place_single_profile_panel()
