# -*- coding: utf-8 -*-
"""Reusable bridge between Revit UI theme and WPF tool surfaces."""

DOCK_PANE_FRAME_DARK = "#2E3440"
DOCK_PANE_FRAME_LIGHT = "#EEEEEE"

try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
except Exception:
    UIThemeManager = None
    UITheme = None

try:
    from Autodesk.Revit.UI.Events import ThemeChangedEventArgs
except Exception:
    ThemeChangedEventArgs = None

try:
    from System import EventHandler
except Exception:
    EventHandler = None

def is_revit_dark_theme():
    """True when Revit reports Dark canvas theme (2024+)."""
    if UIThemeManager is None or UITheme is None:
        return False
    try:
        return UIThemeManager.CurrentTheme == UITheme.Dark
    except Exception:
        return False


def dock_pane_frame_hex(dark_hex=DOCK_PANE_FRAME_DARK, light_hex=DOCK_PANE_FRAME_LIGHT):
    """Resolve pane frame color from Revit canvas theme, fallback light."""
    return dark_hex if is_revit_dark_theme() else light_hex


def default_ced_theme_mode(light_mode="light", dark_mode="dark"):
    """Map current Revit canvas theme to a CED theme mode name."""
    return str(dark_mode if is_revit_dark_theme() else light_mode)


class RevitThemeBridge(object):
    """Subscribes to UIApplication.ThemeChanged (when available) and invokes a callback."""

    def __init__(self, uiapp, on_theme_changed, logger=None):
        self._uiapp = uiapp
        self._on_theme_changed = on_theme_changed
        self._logger = logger
        self._attached = False
        self._attach_failed = False
        self._theme_changed_handler = None

    def _log_warning(self, message):
        if self._logger is None:
            return
        try:
            self._logger.warning(message)
        except Exception:
            pass

    def _emit_current_theme(self):
        callback = self._on_theme_changed
        if callback is None:
            return
        try:
            callback(is_revit_dark_theme())
        except Exception as ex:
            self._log_warning("RevitThemeBridge callback failed: {}".format(ex))

    def _can_subscribe(self):
        if self._uiapp is None:
            return False
        try:
            self._uiapp.ThemeChanged
            return True
        except Exception:
            return False

    def attach(self):
        self._emit_current_theme()
        if self._attached:
            return True
        if self._attach_failed:
            return False
        if not self._can_subscribe():
            return False

        generic_ex = None
        if EventHandler is not None and ThemeChangedEventArgs is not None:
            try:
                if self._theme_changed_handler is None:
                    self._theme_changed_handler = EventHandler[ThemeChangedEventArgs](self._handle_theme_changed)
                self._uiapp.ThemeChanged += self._theme_changed_handler
                self._attached = True
                return True
            except Exception as ex:
                generic_ex = ex
                self._theme_changed_handler = None

        try:
            if self._theme_changed_handler is None:
                self._theme_changed_handler = self._handle_theme_changed
            self._uiapp.ThemeChanged += self._theme_changed_handler
            self._attached = True
            return True
        except Exception as ex:
            self._attached = False
            self._attach_failed = True
            if generic_ex is not None:
                self._log_warning(
                    "RevitThemeBridge attach failed: {} | fallback failed: {}".format(generic_ex, ex)
                )
            else:
                self._log_warning("RevitThemeBridge attach failed: {}".format(ex))
            return False

    def detach(self):
        if not self._attached:
            return
        if not self._can_subscribe():
            self._attached = False
            return
        try:
            if self._theme_changed_handler is not None:
                self._uiapp.ThemeChanged -= self._theme_changed_handler
        except Exception:
            pass
        self._attached = False

    def _handle_theme_changed(self, sender, args):
        self._emit_current_theme()
