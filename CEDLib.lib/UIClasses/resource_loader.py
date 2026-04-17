# -*- coding: utf-8 -*-
"""Reusable UI resource/theme loading helpers for pyRevit WPF UIs."""

import os

import clr
from System import Uri

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from System.Windows import ResourceDictionary

DEFAULT_BASE_RESOURCE_RELATIVE_PATHS = (
    os.path.join("Themes", "CED.Sizes.xaml"),
    os.path.join("Themes", "CED.Colors.xaml"),
    os.path.join("Themes", "CED.Brushes.xaml"),
    os.path.join("Styles", "ButtonStyles.xaml"),
    os.path.join("Styles", "TextStyles.xaml"),
    os.path.join("Styles", "InputStyles.xaml"),
    os.path.join("Styles", "ListStyles.xaml"),
    os.path.join("Styles", "BadgeStyles.xaml"),
    os.path.join("Icons", "Icons.xaml"),
    os.path.join("Templates", "ListItems.xaml"),
    os.path.join("Templates", "Cards.xaml"),
    os.path.join("Templates", "Badges.xaml"),
    os.path.join("Templates", "DataGrids.xaml"),
    os.path.join("Templates", "WindowChrome.xaml"),
    os.path.join("Templates", "ControlPrimitives.xaml"),
    os.path.join("Controls", "SearchBox.xaml"),
)

THEME_RELATIVE_PATHS = {
    "light": (
        os.path.join("Themes", "CEDTheme.Light.xaml"),
    ),
    "dark": (
        os.path.join("Themes", "CEDTheme.Dark.xaml"),
        os.path.join("Themes", "CEDTheme.Light.xaml"),
    ),
    "dark_alt": (
        os.path.join("Themes", "CEDTheme.DarkAlt.xaml"),
        os.path.join("Themes", "CEDTheme.Dark.xaml"),
        os.path.join("Themes", "CEDTheme.Light.xaml"),
    ),
}

ACCENT_BRUSH_KEY_MAP = {
    "blue": "CED.Brush.AccentBlue",
    "neutral": "CED.Brush.AccentNeutral",
}

VALID_THEME_MODES = tuple(sorted(THEME_RELATIVE_PATHS.keys()))
VALID_ACCENT_MODES = tuple(sorted(ACCENT_BRUSH_KEY_MAP.keys()))


def normalize_theme_mode(value, fallback="light"):
    mode = str(value or fallback).strip().lower()
    if mode in THEME_RELATIVE_PATHS:
        return mode
    fb = str(fallback or "light").strip().lower()
    return fb if fb in THEME_RELATIVE_PATHS else "light"


def normalize_accent_mode(value, fallback="blue"):
    mode = str(value or fallback).strip().lower()
    if mode in ACCENT_BRUSH_KEY_MAP:
        return mode
    fb = str(fallback or "blue").strip().lower()
    return fb if fb in ACCENT_BRUSH_KEY_MAP else "blue"


def _normalize_path(path):
    try:
        return os.path.abspath(path).replace("\\", "/").lower()
    except Exception:
        try:
            return str(path).replace("\\", "/").lower()
        except Exception:
            return ""


def _resource_source_path(dictionary):
    source = getattr(dictionary, "Source", None)
    if source is None:
        return ""
    try:
        local_path = source.LocalPath
        if local_path:
            return _normalize_path(local_path)
    except Exception:
        pass
    try:
        return _normalize_path(str(source))
    except Exception:
        return ""


def _is_theme_resource_dictionary(dictionary):
    source_text = _resource_source_path(dictionary)
    return (
        source_text.endswith("/cedtheme.light.xaml")
        or source_text.endswith("/cedtheme.dark.xaml")
        or source_text.endswith("/cedtheme.darkalt.xaml")
    )


def _load_dictionary(path):
    if not path or not os.path.exists(path):
        return None
    try:
        dictionary = ResourceDictionary()
        dictionary.Source = Uri(path)
        return dictionary
    except Exception:
        return None


def _resolve_resources_root(resources_root):
    if resources_root and os.path.isdir(resources_root):
        return os.path.abspath(resources_root)
    return None


def resolve_resource_paths(resources_root, relative_paths):
    root = _resolve_resources_root(resources_root)
    if not root:
        return []
    resolved = []
    for rel in list(relative_paths or []):
        if not rel:
            continue
        resolved.append(os.path.abspath(os.path.join(root, rel)))
    return resolved


def try_find_resource(owner, key):
    if owner is None or not key:
        return None
    try:
        value = owner.TryFindResource(key)
        if value is not None:
            return value
    except Exception:
        pass
    try:
        return owner.FindResource(key)
    except Exception:
        return None


def ensure_base_resources(owner, resources_root, relative_paths=None):
    try:
        resources = getattr(owner, "Resources", None)
        if resources is None:
            return False
        merged = resources.MergedDictionaries
        existing = set()
        for dictionary in list(merged):
            source_path = _resource_source_path(dictionary)
            if source_path:
                existing.add(source_path)
        added = False
        relative = relative_paths or DEFAULT_BASE_RESOURCE_RELATIVE_PATHS
        for path in resolve_resource_paths(resources_root, relative):
            normalized = _normalize_path(path)
            if not normalized or normalized in existing:
                continue
            dictionary = _load_dictionary(path)
            if dictionary is None:
                continue
            merged.Add(dictionary)
            existing.add(normalized)
            added = True
        return added
    except Exception:
        return False


def _load_theme_dictionary(resources_root, theme_mode):
    mode = normalize_theme_mode(theme_mode, "light")
    candidates = THEME_RELATIVE_PATHS.get(mode) or THEME_RELATIVE_PATHS.get("light") or ()
    for path in resolve_resource_paths(resources_root, candidates):
        dictionary = _load_dictionary(path)
        if dictionary is not None:
            return dictionary
    return None


def apply_accent(owner, accent_mode):
    try:
        resources = getattr(owner, "Resources", None)
        if resources is None:
            return False
        mode = normalize_accent_mode(accent_mode, "blue")
        key = ACCENT_BRUSH_KEY_MAP.get(mode) or ACCENT_BRUSH_KEY_MAP.get("blue")
        brush = try_find_resource(owner, key)
        if brush is None:
            brush = try_find_resource(owner, ACCENT_BRUSH_KEY_MAP.get("blue"))
        if brush is None:
            return False
        resources["CED.Brush.Accent"] = brush
        return True
    except Exception:
        return False


def apply_theme(owner, resources_root, theme_mode="light", accent_mode="blue", base_relative_paths=None):
    try:
        ensure_base_resources(owner, resources_root, base_relative_paths)
        dictionary = _load_theme_dictionary(resources_root, normalize_theme_mode(theme_mode, "light"))
        if dictionary is None:
            return False
        resources = getattr(owner, "Resources", None)
        if resources is None:
            return False
        merged = resources.MergedDictionaries
        previous = getattr(owner, "_ced_theme_dictionary", None)
        try:
            if previous is not None and previous in merged:
                merged.Remove(previous)
        except Exception:
            pass
        for existing in list(merged):
            if _is_theme_resource_dictionary(existing):
                try:
                    merged.Remove(existing)
                except Exception:
                    pass
        merged.Add(dictionary)
        owner._ced_theme_dictionary = dictionary
        apply_accent(owner, normalize_accent_mode(accent_mode, "blue"))
        return True
    except Exception:
        return False
