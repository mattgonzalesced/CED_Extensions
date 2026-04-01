# -*- coding: utf-8 -*-
"""Base classes for standardized CED WPF windows and dockable panels."""

import os

from System.Windows.Controls import ScrollViewer, TextBox
from System.Windows.Input import Keyboard, ModifierKeys
from System.Windows.Media import VisualTreeHelper
from pyrevit import forms, script

from UIClasses import pathing
from UIClasses import resource_loader

THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_THEME_KEY = "theme_mode"
THEME_CONFIG_ACCENT_KEY = "accent_mode"

try:
    import ConfigParser as _configparser  # IronPython / Py2
except Exception:
    import configparser as _configparser  # Py3 fallback


def _strip_wrapping_quotes(value):
    text = str(value or "").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ("'", '"'):
        return text[1:-1].strip()
    return text


def _read_theme_state_from_config_file(
    section_name,
    theme_key_name,
    accent_key_name,
    default_theme,
    default_accent,
):
    theme_mode = resource_loader.normalize_theme_mode(default_theme, "light")
    accent_mode = resource_loader.normalize_accent_mode(default_accent, "blue")

    cfg_path = None
    try:
        from pyrevit import userconfig as _userconfig

        cfg_path = getattr(_userconfig, "CONFIG_FILE", None)
    except Exception:
        cfg_path = None

    if not cfg_path:
        appdata = os.getenv("APPDATA", "")
        if appdata:
            cfg_path = os.path.join(appdata, "pyRevit", "pyRevit_config.ini")
    if not cfg_path or not os.path.exists(cfg_path):
        return theme_mode, accent_mode

    parser = _configparser.RawConfigParser()
    try:
        parser.read(cfg_path)
        section = str(section_name or THEME_CONFIG_SECTION)
        if not parser.has_section(section):
            return theme_mode, accent_mode

        theme_opt = str(theme_key_name or THEME_CONFIG_THEME_KEY)
        accent_opt = str(accent_key_name or THEME_CONFIG_ACCENT_KEY)
        if parser.has_option(section, theme_opt):
            raw_theme = _strip_wrapping_quotes(parser.get(section, theme_opt))
            theme_mode = resource_loader.normalize_theme_mode(raw_theme, theme_mode)
        if parser.has_option(section, accent_opt):
            raw_accent = _strip_wrapping_quotes(parser.get(section, accent_opt))
            accent_mode = resource_loader.normalize_accent_mode(raw_accent, accent_mode)
    except Exception:
        pass
    return theme_mode, accent_mode


def load_theme_state_from_config(
    section_name=None,
    theme_key_name=None,
    accent_key_name=None,
    default_theme="light",
    default_accent="blue",
):
    """Return persisted CED theme/accent from pyRevit config."""
    return _load_theme_state_from_config(
        section_name=section_name or THEME_CONFIG_SECTION,
        theme_key_name=theme_key_name or THEME_CONFIG_THEME_KEY,
        accent_key_name=accent_key_name or THEME_CONFIG_ACCENT_KEY,
        default_theme=default_theme,
        default_accent=default_accent,
    )


def _resolve_resources_root(value):
    if value and os.path.isdir(value):
        return os.path.abspath(value)
    return None


def _resolve_xaml_source(start_dir, xaml_source):
    if not xaml_source:
        return xaml_source
    source = str(xaml_source)
    if os.path.isabs(source):
        return source
    if os.path.exists(source):
        return os.path.abspath(source)
    if start_dir:
        local = os.path.abspath(os.path.join(start_dir, source))
        if os.path.exists(local):
            return local
    return source


def _infer_default_xaml_source(instance, start_dir, fallback_name=None):
    candidate_names = []
    try:
        class_name = str(instance.__class__.__name__ or "").strip()
    except Exception:
        class_name = ""
    if class_name:
        candidate_names.append(class_name + ".xaml")
    try:
        module_leaf = str((instance.__class__.__module__ or "").split(".")[-1]).strip()
    except Exception:
        module_leaf = ""
    if module_leaf:
        candidate_names.append(module_leaf + ".xaml")
    if fallback_name:
        candidate_names.append(str(fallback_name))

    seen = set()
    for name in candidate_names:
        if not name or name in seen:
            continue
        seen.add(name)
        if start_dir:
            resolved = os.path.abspath(os.path.join(start_dir, name))
            if os.path.exists(resolved):
                return resolved
    for name in candidate_names:
        if name:
            return name
    return None


def _resolve_start_dir_for_instance(instance):
    module_name = None
    try:
        module_name = instance.__class__.__module__
    except Exception:
        module_name = None
    return pathing.resolve_start_dir(module_name=module_name, fallback_dir=os.getcwd())


def _resolve_panel_source(start_dir, panel_source):
    source = panel_source
    if not source:
        return None
    source_text = str(source)
    if os.path.isabs(source_text):
        return source_text
    if os.path.exists(source_text):
        return os.path.abspath(source_text)
    if start_dir:
        local = os.path.abspath(os.path.join(start_dir, source_text))
        if os.path.exists(local):
            return local
    return source_text


def _find_visual_descendants(node, target_type):
    if node is None:
        return []
    results = []
    queue = [node]
    while queue:
        current = queue.pop(0)
        try:
            child_count = int(VisualTreeHelper.GetChildrenCount(current) or 0)
        except Exception:
            child_count = 0
        for idx in range(child_count):
            try:
                child = VisualTreeHelper.GetChild(current, idx)
            except Exception:
                continue
            if isinstance(child, target_type):
                results.append(child)
            queue.append(child)
    return results


def _find_scrollviewer_ancestor(node):
    current = node
    while current is not None:
        try:
            if isinstance(current, ScrollViewer):
                return current
        except Exception:
            pass
        try:
            current = VisualTreeHelper.GetParent(current)
        except Exception:
            return None
    return None


def _shift_wheel_to_horizontal(event_args):
    try:
        if Keyboard.Modifiers != ModifierKeys.Shift:
            return False
    except Exception:
        return False
    scroll_viewer = _find_scrollviewer_ancestor(getattr(event_args, "OriginalSource", None))
    if scroll_viewer is None:
        return False
    try:
        if float(scroll_viewer.ScrollableWidth) <= 0.0:
            return False
    except Exception:
        return False
    try:
        delta = int(getattr(event_args, "Delta", 0))
    except Exception:
        delta = 0
    if delta == 0:
        return False
    step = 36.0
    direction = -1.0 if delta > 0 else 1.0
    try:
        current_offset = float(scroll_viewer.HorizontalOffset)
        max_offset = float(scroll_viewer.ScrollableWidth)
        new_offset = max(0.0, min(max_offset, current_offset + (direction * step)))
        scroll_viewer.ScrollToHorizontalOffset(new_offset)
        event_args.Handled = True
        return True
    except Exception:
        return False


def _load_theme_state_from_config(
    section_name,
    theme_key_name,
    accent_key_name,
    default_theme="light",
    default_accent="blue",
):
    theme_mode, accent_mode = _read_theme_state_from_config_file(
        section_name=section_name,
        theme_key_name=theme_key_name,
        accent_key_name=accent_key_name,
        default_theme=default_theme,
        default_accent=default_accent,
    )
    try:
        cfg = script.get_config(str(section_name or THEME_CONFIG_SECTION))
        if cfg is None:
            return theme_mode, accent_mode
        theme_mode = resource_loader.normalize_theme_mode(
            _strip_wrapping_quotes(cfg.get_option(str(theme_key_name or THEME_CONFIG_THEME_KEY), theme_mode)),
            theme_mode,
        )
        accent_mode = resource_loader.normalize_accent_mode(
            _strip_wrapping_quotes(cfg.get_option(str(accent_key_name or THEME_CONFIG_ACCENT_KEY), accent_mode)),
            accent_mode,
        )
    except Exception:
        pass
    return theme_mode, accent_mode


def _init_ced_surface(
    instance,
    resources_root=None,
    theme_mode=None,
    accent_mode=None,
    theme_aware=None,
    use_config_theme=None,
    enable_shift_wheel_to_horizontal=None,
    auto_wire_textboxes=None,
    text_select_all_on_click=None,
    text_select_all_on_focus=None,
):
    instance._ced_start_dir = _resolve_start_dir_for_instance(instance)
    instance._ced_context = pathing.resolve_ui_context(
        instance._ced_start_dir,
        ensure_syspath=True,
    )
    instance._ced_resources_root = (
        _resolve_resources_root(resources_root)
        or _resolve_resources_root(instance._ced_context.get("resources_root"))
    )

    instance._theme_aware = bool(
        getattr(instance, "theme_aware", False) if theme_aware is None else theme_aware
    )
    instance._use_config_theme = bool(
        getattr(instance, "use_config_theme", True) if use_config_theme is None else use_config_theme
    )
    instance._shift_wheel_to_horizontal_enabled = bool(
        getattr(instance, "enable_shift_wheel_to_horizontal", True)
        if enable_shift_wheel_to_horizontal is None
        else enable_shift_wheel_to_horizontal
    )
    instance._auto_wire_textboxes = bool(
        getattr(instance, "auto_wire_textboxes", False) if auto_wire_textboxes is None else auto_wire_textboxes
    )
    instance._text_select_all_on_click = bool(
        getattr(instance, "text_select_all_on_click", False)
        if text_select_all_on_click is None
        else text_select_all_on_click
    )
    instance._text_select_all_on_focus = bool(
        getattr(instance, "text_select_all_on_focus", False)
        if text_select_all_on_focus is None
        else text_select_all_on_focus
    )

    instance._theme_mode, instance._accent_mode = _resolve_theme_state_for_instance(
        instance,
        explicit_theme_mode=theme_mode,
        explicit_accent_mode=accent_mode,
    )


def _resolve_theme_state_for_instance(instance, explicit_theme_mode=None, explicit_accent_mode=None):
    if not getattr(instance, "_theme_aware", False):
        return "light", "blue"

    default_theme_mode = getattr(instance, "default_theme_mode", "light")
    default_accent_mode = getattr(instance, "default_accent_mode", "blue")
    theme_mode = resource_loader.normalize_theme_mode(default_theme_mode, "light")
    accent_mode = resource_loader.normalize_accent_mode(default_accent_mode, "blue")

    if getattr(instance, "_use_config_theme", True):
        cfg_theme, cfg_accent = _load_theme_state_from_config(
            section_name=getattr(instance, "theme_config_section", THEME_CONFIG_SECTION),
            theme_key_name=getattr(instance, "theme_config_theme_key", THEME_CONFIG_THEME_KEY),
            accent_key_name=getattr(instance, "theme_config_accent_key", THEME_CONFIG_ACCENT_KEY),
            default_theme=theme_mode,
            default_accent=accent_mode,
        )
        theme_mode = cfg_theme
        accent_mode = cfg_accent

    if explicit_theme_mode is not None:
        theme_mode = resource_loader.normalize_theme_mode(explicit_theme_mode, theme_mode)
    if explicit_accent_mode is not None:
        accent_mode = resource_loader.normalize_accent_mode(explicit_accent_mode, accent_mode)

    return theme_mode, accent_mode


def _apply_ced_theme_for_instance(instance, theme_mode=None, accent_mode=None):
    if theme_mode is not None:
        instance._theme_mode = resource_loader.normalize_theme_mode(theme_mode, instance._theme_mode)
    if accent_mode is not None:
        instance._accent_mode = resource_loader.normalize_accent_mode(accent_mode, instance._accent_mode)
    return resource_loader.apply_theme(
        instance,
        resources_root=getattr(instance, "_ced_resources_root", None),
        theme_mode=getattr(instance, "_theme_mode", "light"),
        accent_mode=getattr(instance, "_accent_mode", "blue"),
    )


def _refresh_ced_theme_from_config_for_instance(instance):
    if not getattr(instance, "_theme_aware", False):
        instance._theme_mode = "light"
        instance._accent_mode = "blue"
        return _apply_ced_theme_for_instance(instance)
    theme_mode, accent_mode = _load_theme_state_from_config(
        section_name=getattr(instance, "theme_config_section", THEME_CONFIG_SECTION),
        theme_key_name=getattr(instance, "theme_config_theme_key", THEME_CONFIG_THEME_KEY),
        accent_key_name=getattr(instance, "theme_config_accent_key", THEME_CONFIG_ACCENT_KEY),
        default_theme=getattr(instance, "default_theme_mode", "light"),
        default_accent=getattr(instance, "default_accent_mode", "blue"),
    )
    instance._theme_mode = theme_mode
    instance._accent_mode = accent_mode
    return _apply_ced_theme_for_instance(instance)


def _resolve_textbox_targets_for_instance(instance, textbox_names=None):
    if textbox_names:
        targets = []
        for name in list(textbox_names or []):
            try:
                box = instance.FindName(str(name))
            except Exception:
                box = None
            if isinstance(box, TextBox):
                targets.append(box)
        return targets
    return list(_find_visual_descendants(instance, TextBox) or [])


def _bind_textbox_behavior_for_instance(instance, textbox, select_all_on_click=None, select_all_on_focus=None):
    if not isinstance(textbox, TextBox):
        return False
    click_enabled = (
        getattr(instance, "_text_select_all_on_click", False)
        if select_all_on_click is None
        else bool(select_all_on_click)
    )
    focus_enabled = (
        getattr(instance, "_text_select_all_on_focus", False)
        if select_all_on_focus is None
        else bool(select_all_on_focus)
    )
    try:
        setattr(textbox, "_ced_select_all_on_click", bool(click_enabled))
        setattr(textbox, "_ced_select_all_on_focus", bool(focus_enabled))
    except Exception:
        pass
    if bool(getattr(textbox, "_ced_behavior_bound", False)):
        return True
    try:
        textbox.PreviewMouseLeftButtonDown += instance._textbox_preview_mouse_down
        textbox.GotKeyboardFocus += instance._textbox_got_keyboard_focus
        textbox._ced_behavior_bound = True
        return True
    except Exception:
        return False


def _apply_textbox_behaviors_for_instance(
    instance,
    textbox_names=None,
    select_all_on_click=None,
    select_all_on_focus=None,
):
    count = 0
    for textbox in list(_resolve_textbox_targets_for_instance(instance, textbox_names=textbox_names) or []):
        if _bind_textbox_behavior_for_instance(
            instance,
            textbox,
            select_all_on_click=select_all_on_click,
            select_all_on_focus=select_all_on_focus,
        ):
            count += 1
    return count


class CEDWindowBase(forms.WPFWindow):
    """Base window with auto path resolution, theme support and input behaviors."""

    theme_aware = False
    use_config_theme = True
    default_theme_mode = "light"
    default_accent_mode = "blue"
    theme_config_section = THEME_CONFIG_SECTION
    theme_config_theme_key = THEME_CONFIG_THEME_KEY
    theme_config_accent_key = THEME_CONFIG_ACCENT_KEY
    enable_shift_wheel_to_horizontal = True
    auto_wire_textboxes = False
    text_select_all_on_click = False
    text_select_all_on_focus = False

    def __init__(
        self,
        xaml_source=None,
        resources_root=None,
        theme_mode=None,
        accent_mode=None,
        theme_aware=None,
        use_config_theme=None,
        enable_shift_wheel_to_horizontal=None,
        auto_wire_textboxes=None,
        text_select_all_on_click=None,
        text_select_all_on_focus=None,
        literal_string=False,
        handle_esc=True,
        set_owner=True,
    ):
        _init_ced_surface(
            self,
            resources_root=resources_root,
            theme_mode=theme_mode,
            accent_mode=accent_mode,
            theme_aware=theme_aware,
            use_config_theme=use_config_theme,
            enable_shift_wheel_to_horizontal=enable_shift_wheel_to_horizontal,
            auto_wire_textboxes=auto_wire_textboxes,
            text_select_all_on_click=text_select_all_on_click,
            text_select_all_on_focus=text_select_all_on_focus,
        )

        resolved_source = xaml_source or getattr(self, "xaml_source", None)
        if not literal_string:
            if not resolved_source:
                resolved_source = _infer_default_xaml_source(self, self._ced_start_dir)
            resolved_source = _resolve_xaml_source(self._ced_start_dir, resolved_source)
            if not resolved_source:
                resolved_source = _infer_default_xaml_source(self, self._ced_start_dir)

        forms.WPFWindow.__init__(
            self,
            resolved_source,
            literal_string=literal_string,
            handle_esc=handle_esc,
            set_owner=set_owner,
        )

        self.apply_ced_theme()

        if self._shift_wheel_to_horizontal_enabled:
            try:
                self.PreviewMouseWheel += self._on_preview_mouse_wheel
            except Exception:
                pass

        if self._auto_wire_textboxes and (self._text_select_all_on_click or self._text_select_all_on_focus):
            try:
                self.Loaded += self._on_loaded_wire_textboxes
            except Exception:
                pass

    def _resolve_theme_state(self, explicit_theme_mode=None, explicit_accent_mode=None):
        return _resolve_theme_state_for_instance(
            self,
            explicit_theme_mode=explicit_theme_mode,
            explicit_accent_mode=explicit_accent_mode,
        )

    def apply_ced_theme(self, theme_mode=None, accent_mode=None):
        return _apply_ced_theme_for_instance(self, theme_mode=theme_mode, accent_mode=accent_mode)

    def refresh_ced_theme_from_config(self):
        return _refresh_ced_theme_from_config_for_instance(self)

    def get_ced_context(self):
        return dict(self._ced_context or {})

    def _resolve_textbox_targets(self, textbox_names=None):
        return _resolve_textbox_targets_for_instance(self, textbox_names=textbox_names)

    def bind_textbox_behavior(self, textbox, select_all_on_click=None, select_all_on_focus=None):
        return _bind_textbox_behavior_for_instance(
            self,
            textbox,
            select_all_on_click=select_all_on_click,
            select_all_on_focus=select_all_on_focus,
        )

    def apply_textbox_behaviors(self, textbox_names=None, select_all_on_click=None, select_all_on_focus=None):
        return _apply_textbox_behaviors_for_instance(
            self,
            textbox_names=textbox_names,
            select_all_on_click=select_all_on_click,
            select_all_on_focus=select_all_on_focus,
        )

    def _on_loaded_wire_textboxes(self, sender, args):
        self.apply_textbox_behaviors()

    def _textbox_preview_mouse_down(self, sender, args):
        if not bool(getattr(sender, "_ced_select_all_on_click", False)):
            return
        try:
            if not sender.IsKeyboardFocusWithin:
                sender.Focus()
                args.Handled = True
                return
            sender.SelectAll()
            args.Handled = True
        except Exception:
            pass

    def _textbox_got_keyboard_focus(self, sender, args):
        if not bool(getattr(sender, "_ced_select_all_on_focus", False)):
            return
        try:
            sender.SelectAll()
        except Exception:
            pass

    def _on_preview_mouse_wheel(self, sender, event_args):
        _shift_wheel_to_horizontal(event_args)


class CEDPanelBase(forms.WPFPanel):
    """Base dockable panel with auto path resolution/theme behaviors."""

    panel_source = None
    theme_aware = False
    use_config_theme = True
    default_theme_mode = "light"
    default_accent_mode = "blue"
    theme_config_section = THEME_CONFIG_SECTION
    theme_config_theme_key = THEME_CONFIG_THEME_KEY
    theme_config_accent_key = THEME_CONFIG_ACCENT_KEY
    enable_shift_wheel_to_horizontal = True
    auto_wire_textboxes = False
    text_select_all_on_click = False
    text_select_all_on_focus = False

    def __init__(
        self,
        resources_root=None,
        theme_mode=None,
        accent_mode=None,
        theme_aware=None,
        use_config_theme=None,
        enable_shift_wheel_to_horizontal=None,
        auto_wire_textboxes=None,
        text_select_all_on_click=None,
        text_select_all_on_focus=None,
    ):
        _init_ced_surface(
            self,
            resources_root=resources_root,
            theme_mode=theme_mode,
            accent_mode=accent_mode,
            theme_aware=theme_aware,
            use_config_theme=use_config_theme,
            enable_shift_wheel_to_horizontal=enable_shift_wheel_to_horizontal,
            auto_wire_textboxes=auto_wire_textboxes,
            text_select_all_on_click=text_select_all_on_click,
            text_select_all_on_focus=text_select_all_on_focus,
        )
        resolved_panel_source = _resolve_panel_source(
            self._ced_start_dir,
            getattr(self, "panel_source", None),
        )
        if not resolved_panel_source:
            resolved_panel_source = _infer_default_xaml_source(self, self._ced_start_dir)
        if resolved_panel_source:
            self.panel_source = resolved_panel_source

        forms.WPFPanel.__init__(self)
        self.apply_ced_theme()

        if self._shift_wheel_to_horizontal_enabled:
            try:
                self.PreviewMouseWheel += self._on_preview_mouse_wheel
            except Exception:
                pass

        if self._auto_wire_textboxes and (self._text_select_all_on_click or self._text_select_all_on_focus):
            try:
                self.Loaded += self._on_loaded_wire_textboxes
            except Exception:
                pass

    def _resolve_theme_state(self, explicit_theme_mode=None, explicit_accent_mode=None):
        return _resolve_theme_state_for_instance(
            self,
            explicit_theme_mode=explicit_theme_mode,
            explicit_accent_mode=explicit_accent_mode,
        )

    def apply_ced_theme(self, theme_mode=None, accent_mode=None):
        return _apply_ced_theme_for_instance(self, theme_mode=theme_mode, accent_mode=accent_mode)

    def refresh_ced_theme_from_config(self):
        return _refresh_ced_theme_from_config_for_instance(self)

    def get_ced_context(self):
        return dict(self._ced_context or {})

    def _resolve_textbox_targets(self, textbox_names=None):
        return _resolve_textbox_targets_for_instance(self, textbox_names=textbox_names)

    def bind_textbox_behavior(self, textbox, select_all_on_click=None, select_all_on_focus=None):
        return _bind_textbox_behavior_for_instance(
            self,
            textbox,
            select_all_on_click=select_all_on_click,
            select_all_on_focus=select_all_on_focus,
        )

    def apply_textbox_behaviors(self, textbox_names=None, select_all_on_click=None, select_all_on_focus=None):
        return _apply_textbox_behaviors_for_instance(
            self,
            textbox_names=textbox_names,
            select_all_on_click=select_all_on_click,
            select_all_on_focus=select_all_on_focus,
        )

    def _on_loaded_wire_textboxes(self, sender, args):
        self.apply_textbox_behaviors()

    def _textbox_preview_mouse_down(self, sender, args):
        if not bool(getattr(sender, "_ced_select_all_on_click", False)):
            return
        try:
            if not sender.IsKeyboardFocusWithin:
                sender.Focus()
                args.Handled = True
                return
            sender.SelectAll()
            args.Handled = True
        except Exception:
            pass

    def _textbox_got_keyboard_focus(self, sender, args):
        if not bool(getattr(sender, "_ced_select_all_on_focus", False)):
            return
        try:
            sender.SelectAll()
        except Exception:
            pass

    def _on_preview_mouse_wheel(self, sender, event_args):
        _shift_wheel_to_horizontal(event_args)
