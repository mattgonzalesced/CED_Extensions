# -*- coding: utf-8 -*-
"""Base classes for standardized CED WPF windows and dockable panels."""

import os

from pyrevit import forms

from UIClasses import resource_loader


def _resolve_resources_root(value):
    if value and os.path.isdir(value):
        return os.path.abspath(value)
    return None


class CEDWindowBase(forms.WPFWindow):
    """Base window that applies CED resources/theme/accent on construction."""

    def __init__(self, xaml_path, resources_root=None, theme_mode="light", accent_mode="blue"):
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        self._ced_resources_root = _resolve_resources_root(resources_root)
        forms.WPFWindow.__init__(self, xaml_path)
        self.apply_ced_theme()

    def apply_ced_theme(self):
        return resource_loader.apply_theme(
            self,
            resources_root=self._ced_resources_root,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )


class CEDPanelBase(forms.WPFPanel):
    """Base dockable panel that applies CED resources/theme/accent on construction."""

    panel_source = None

    def __init__(self, resources_root=None, theme_mode="light", accent_mode="blue"):
        self._theme_mode = theme_mode or "light"
        self._accent_mode = accent_mode or "blue"
        self._ced_resources_root = _resolve_resources_root(resources_root)
        forms.WPFPanel.__init__(self)
        self.apply_ced_theme()

    def apply_ced_theme(self):
        return resource_loader.apply_theme(
            self,
            resources_root=self._ced_resources_root,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )
