# -*- coding: utf-8 -*-

import os

import clr

clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from System.Windows import Visibility
from System.Windows import FontStyles
from System.Windows.Media import Brushes

from pyrevit import forms, revit, script

from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.circuit_settings import (
    CircuitSettings,
    FeederVDMethod,
    MultiPoleBranchNeutralBehavior,
    NeutralBehavior,
    IsolatedGroundBehavior,
    WireMaterialDisplay,
    WireStringSeparator,
)
from UIClasses import pathing as ui_pathing
from UIClasses import resource_loader

THIS_DIR = os.path.abspath(os.path.dirname(__file__))


def _resolve_bundle_or_local_path(file_name):
    path = None
    try:
        path = script.get_bundle_file(file_name)
    except Exception:
        path = None
    if path and os.path.exists(path):
        return path
    return os.path.join(THIS_DIR, file_name)


XAML_PATH = _resolve_bundle_or_local_path("settings.xaml")
logger = script.get_logger()
LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
UI_RESOURCES_ROOT = ui_pathing.resolve_ui_resources_root(LIB_ROOT)
THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_THEME_KEY = "theme_mode"
THEME_CONFIG_ACCENT_KEY = "accent_mode"
def _load_theme_state(default_theme="light", default_accent="blue"):
    from UIClasses import load_theme_state_from_config

    return load_theme_state_from_config(
        default_theme=default_theme,
        default_accent=default_accent,
    )


def _save_theme_state(theme_mode, accent_mode):
    cfg = script.get_config(THEME_CONFIG_SECTION)
    if cfg is None:
        return
    cfg.set_option(THEME_CONFIG_THEME_KEY, resource_loader.normalize_theme_mode(theme_mode, "light"))
    cfg.set_option(THEME_CONFIG_ACCENT_KEY, resource_loader.normalize_accent_mode(accent_mode, "blue"))
    script.save_config()


def _verify_project_parameters(doc, app, settings=None):
    active_settings = settings or settings_manager.load_circuit_settings(doc)
    return settings_manager.sync_electrical_parameter_bindings(
        doc,
        logger=logger,
        settings=active_settings,
        check_ownership=True,
        transaction_name="Verify Electrical Parameters",
    )


class CircuitSettingsWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_PATH)
        self._theme_mode, self._accent_mode = _load_theme_state("light", "blue")
        self._apply_theme()
        self._refresh_theme_brushes()
        self.doc = revit.doc
        if self.doc is None:
            try:
                uidoc = __revit__.ActiveUIDocument
                self.doc = uidoc.Document if uidoc is not None else None
            except Exception:
                self.doc = None
        if self.doc is None:
            raise Exception("No active document is available.")
        self.defaults = CircuitSettings()
        self.settings = settings_manager.load_circuit_settings(self.doc)
        self._previous_equipment_write = bool(self.settings.write_equipment_results)
        self._previous_fixture_write = bool(self.settings.write_fixture_results)
        self._last_clear_equipment_disabled = bool(getattr(self.settings, 'last_clear_equipment_disabled', False))
        self._last_clear_fixtures_disabled = bool(getattr(self.settings, 'last_clear_fixtures_disabled', False))
        self._last_clear_success = bool(getattr(self.settings, 'last_clear_success', False))
        self._help_key = None
        self._is_normalizing = False
        self._is_loading_ui = True

        self._bind_events()
        self._load_defaults_panel()
        self._load_values()
        self._refresh_styles()
        self._set_help_context('min_conduit_size')
        self._is_loading_ui = False

    def _apply_theme(self):
        resource_loader.apply_theme(
            self,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )

    def _refresh_theme_brushes(self):
        try:
            self._primary_text_brush = self.TryFindResource("CED.Brush.PrimaryText")
        except Exception:
            self._primary_text_brush = None
        try:
            self._secondary_text_brush = self.TryFindResource("CED.Brush.SecondaryText")
        except Exception:
            self._secondary_text_brush = None

    # ------------- UI wiring -----------------
    def _bind_events(self):
        self.save_btn.Click += self._on_save
        self.cancel_btn.Click += self._on_cancel
        self.reset_btn.Click += self._on_reset
        self.help_btn.Click += self._on_help
        self.verify_parameters_btn.Click += self._on_verify_parameters
        self.clear_writeback_btn.Click += self._on_clear_persistent
        self.min_conduit_size_cb.SelectionChanged += self._on_value_changed
        self.min_conduit_size_cb.GotFocus += lambda s, e: self._set_help_context('min_conduit_size')
        self.max_conduit_fill_tb.GotFocus += lambda s, e: self._set_help_context('max_conduit_fill')
        self.max_conduit_fill_tb.PreviewMouseLeftButtonDown += self._percent_box_preview_mouse_down
        self.max_conduit_fill_tb.GotKeyboardFocus += self._percent_box_got_focus
        self.max_conduit_fill_tb.LostFocus += lambda s, e: self._normalize_percent_on_blur(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4)
        self.max_branch_vd_tb.GotFocus += lambda s, e: self._set_help_context('max_branch_voltage_drop')
        self.max_branch_vd_tb.PreviewMouseLeftButtonDown += self._percent_box_preview_mouse_down
        self.max_branch_vd_tb.GotKeyboardFocus += self._percent_box_got_focus
        self.max_branch_vd_tb.LostFocus += lambda s, e: self._normalize_percent_on_blur(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05)
        self.max_feeder_vd_tb.GotFocus += lambda s, e: self._set_help_context('max_feeder_voltage_drop')
        self.max_feeder_vd_tb.PreviewMouseLeftButtonDown += self._percent_box_preview_mouse_down
        self.max_feeder_vd_tb.GotKeyboardFocus += self._percent_box_got_focus
        self.max_feeder_vd_tb.LostFocus += lambda s, e: self._normalize_percent_on_blur(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05)
        self.multi_pole_branch_neutral_behavior_cb.SelectionChanged += self._on_value_changed
        self.multi_pole_branch_neutral_behavior_cb.GotFocus += lambda s, e: self._set_help_context('multi_pole_branch_neutral_behavior')
        self.neutral_behavior_cb.SelectionChanged += self._on_value_changed
        self.neutral_behavior_cb.GotFocus += lambda s, e: self._set_help_context('neutral_behavior')
        self.isolated_ground_behavior_cb.SelectionChanged += self._on_value_changed
        self.isolated_ground_behavior_cb.GotFocus += lambda s, e: self._set_help_context('isolated_ground_behavior')
        self.wire_material_display_cb.SelectionChanged += self._on_value_changed
        self.wire_material_display_cb.GotFocus += lambda s, e: self._set_help_context('wire_material_display')
        self.wire_string_separator_cb.SelectionChanged += self._on_value_changed
        self.wire_string_separator_cb.GotFocus += lambda s, e: self._set_help_context('wire_string_separator')
        self.feeder_vd_method_cb.SelectionChanged += self._on_value_changed
        self.feeder_vd_method_cb.GotFocus += lambda s, e: self._set_help_context('feeder_vd_method')
        self.write_equipment_cb.Checked += self._on_value_changed
        self.write_equipment_cb.Unchecked += self._on_value_changed
        self.write_equipment_cb.GotFocus += lambda s, e: self._set_help_context('write_results')
        self.write_fixtures_cb.Checked += self._on_value_changed
        self.write_fixtures_cb.Unchecked += self._on_value_changed
        self.write_fixtures_cb.GotFocus += lambda s, e: self._set_help_context('write_results')
        self.clear_writeback_btn.GotFocus += lambda s, e: self._set_help_context('clear_writebacks')
        self.theme_mode_cb.SelectionChanged += self._on_theme_value_changed
        self.theme_mode_cb.GotFocus += lambda s, e: self._set_help_context('theme_mode')
        self.accent_mode_cb.SelectionChanged += self._on_theme_value_changed
        self.accent_mode_cb.GotFocus += lambda s, e: self._set_help_context('accent_mode')
        self.verify_parameters_btn.GotFocus += lambda s, e: self._set_help_context('verify_parameters')

        self.max_conduit_fill_tb.TextChanged += lambda s, e: self._on_percent_value_changed(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4)
        self.max_branch_vd_tb.TextChanged += lambda s, e: self._on_percent_value_changed(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05)
        self.max_feeder_vd_tb.TextChanged += lambda s, e: self._on_percent_value_changed(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05)

    # ------------- Data helpers --------------
    def _load_defaults_panel(self):
        self.min_conduit_default.Text = u"(Default: {})".format(self.defaults.min_conduit_size)
        self.max_conduit_fill_default.Text = u"(Default: {}%)".format(self._percent_value(self.defaults.max_conduit_fill))
        self.multi_pole_branch_neutral_behavior_default.Text = u"(Default: {})".format(
            self._describe_multipole_branch_neutral(self.defaults.multi_pole_branch_neutral_behavior)
        )
        self.neutral_behavior_default.Text = u"(Default: {})".format(self._describe_neutral(self.defaults.neutral_behavior))
        self.isolated_ground_behavior_default.Text = u"(Default: {})".format(
            self._describe_isolated_ground(self.defaults.isolated_ground_behavior)
        )
        self.wire_material_display_default.Text = u"(Default: {})".format(
            self._describe_material_display(self.defaults.wire_material_display)
        )
        self.wire_string_separator_default.Text = u"(Default: {})".format(
            self._describe_wire_separator(self.defaults.wire_string_separator)
        )
        self.max_branch_vd_default.Text = u"(Default: {}%)".format(self._percent_value(self.defaults.max_branch_voltage_drop))
        self.max_feeder_vd_default.Text = u"(Default: {}%)".format(self._percent_value(self.defaults.max_feeder_voltage_drop))
        self.feeder_vd_method_default.Text = u"(Default: {})".format(self._describe_feeder_method(self.defaults.feeder_vd_method))
        self.write_results_default.Text = u"(Defaults: Equipment ✓, Fixtures ✕)"
        self.theme_mode_default.Text = u"(Current: {})".format(self._describe_theme_mode(self._theme_mode))
        self.accent_mode_default.Text = u"(Current: {})".format(self._describe_accent_mode(self._accent_mode))

    def _load_values(self):
        self._select_combo_by_tag(self.min_conduit_size_cb, self.settings.min_conduit_size)
        self._set_percent_field(self.max_conduit_fill_tb, self.settings.max_conduit_fill)
        self._set_percent_field(self.max_branch_vd_tb, self.settings.max_branch_voltage_drop)
        self._set_percent_field(self.max_feeder_vd_tb, self.settings.max_feeder_voltage_drop)

        self._select_combo_by_tag(
            self.multi_pole_branch_neutral_behavior_cb,
            self.settings.multi_pole_branch_neutral_behavior,
        )
        self._select_combo_by_tag(self.neutral_behavior_cb, self.settings.neutral_behavior)
        self._select_combo_by_tag(self.isolated_ground_behavior_cb, self.settings.isolated_ground_behavior)
        self._select_combo_by_tag(self.wire_material_display_cb, self.settings.wire_material_display)
        self._select_combo_by_tag(self.wire_string_separator_cb, self.settings.wire_string_separator)
        self._select_combo_by_tag(self.feeder_vd_method_cb, self.settings.feeder_vd_method)
        self._select_combo_by_tag(self.theme_mode_cb, self._theme_mode)
        self._select_combo_by_tag(self.accent_mode_cb, self._accent_mode)

        self.write_equipment_cb.IsChecked = bool(self.settings.write_equipment_results)
        self.write_fixtures_cb.IsChecked = bool(self.settings.write_fixture_results)

        self._refresh_clear_alert()
        self._set_verify_status("")

    def _select_combo_by_tag(self, combo, tag_value):
        for item in combo.Items:
            if getattr(item, 'Tag', None) == tag_value:
                combo.SelectedItem = item
                break
        else:
            if combo.Items and combo.SelectedItem is None:
                combo.SelectedIndex = 0

    def _get_combo_tag(self, combo):
        if combo.SelectedItem:
            return getattr(combo.SelectedItem, 'Tag', None)
        return None

    def _update_settings_from_ui(self):
        updated = CircuitSettings.from_json(self.settings.to_json())
        updated.set('min_conduit_size', self._get_combo_tag(self.min_conduit_size_cb))
        updated.set('max_conduit_fill', self._parse_percent_field(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4))
        updated.set('multi_pole_branch_neutral_behavior', self._get_combo_tag(self.multi_pole_branch_neutral_behavior_cb))
        updated.set('neutral_behavior', self._get_combo_tag(self.neutral_behavior_cb))
        updated.set('isolated_ground_behavior', self._get_combo_tag(self.isolated_ground_behavior_cb))
        updated.set('wire_material_display', self._get_combo_tag(self.wire_material_display_cb))
        updated.set('wire_string_separator', self._get_combo_tag(self.wire_string_separator_cb))
        updated.set('max_branch_voltage_drop', self._parse_percent_field(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05))
        updated.set('max_feeder_voltage_drop', self._parse_percent_field(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05))
        updated.set('feeder_vd_method', self._get_combo_tag(self.feeder_vd_method_cb))
        updated.set('write_equipment_results', bool(self.write_equipment_cb.IsChecked))
        updated.set('write_fixture_results', bool(self.write_fixtures_cb.IsChecked))
        updated.set('pending_clear_failed', bool(getattr(self.settings, 'pending_clear_failed', False)))
        updated.set('last_clear_equipment_disabled', bool(getattr(self.settings, 'last_clear_equipment_disabled', False)))
        updated.set('last_clear_fixtures_disabled', bool(getattr(self.settings, 'last_clear_fixtures_disabled', False)))
        updated.set('last_clear_success', bool(getattr(self.settings, 'last_clear_success', False)))
        return updated

    # ------------- Styling helpers -----------
    def _is_default(self, key, value):
        default_value = self.defaults.get(key)
        try:
            return float(value) == float(default_value)
        except Exception:
            try:
                return str(value).strip() == str(default_value).strip()
            except Exception:
                return False

    def _apply_default_style(self, control, is_default):
        if control is None:
            return
        type_name = ""
        try:
            type_name = control.GetType().Name
        except Exception:
            pass
        is_option_control = type_name in ("ComboBox", "CheckBox")
        control.FontStyle = FontStyles.Normal if is_option_control else (FontStyles.Italic if is_default else FontStyles.Normal)
        control.Foreground = self._primary_text_brush or Brushes.Black

    def _refresh_styles(self, sender=None, args=None):
        self._apply_default_style(self.min_conduit_size_cb, self._is_default('min_conduit_size', self._get_combo_tag(self.min_conduit_size_cb)))
        self._apply_default_style(self.max_conduit_fill_tb, self._is_default('max_conduit_fill', self._parse_percent_field(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4, silent=True)))
        self._apply_default_style(self.max_branch_vd_tb, self._is_default('max_branch_voltage_drop', self._parse_percent_field(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05, silent=True)))
        self._apply_default_style(self.max_feeder_vd_tb, self._is_default('max_feeder_voltage_drop', self._parse_percent_field(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05, silent=True)))

        self._apply_default_style(
            self.multi_pole_branch_neutral_behavior_cb,
            self._is_default(
                'multi_pole_branch_neutral_behavior',
                self._get_combo_tag(self.multi_pole_branch_neutral_behavior_cb),
            ),
        )
        nb_value = self._get_combo_tag(self.neutral_behavior_cb)
        fd_value = self._get_combo_tag(self.feeder_vd_method_cb)
        self._apply_default_style(self.neutral_behavior_cb, self._is_default('neutral_behavior', nb_value))
        self._apply_default_style(
            self.isolated_ground_behavior_cb,
            self._is_default('isolated_ground_behavior', self._get_combo_tag(self.isolated_ground_behavior_cb)),
        )
        self._apply_default_style(
            self.wire_material_display_cb,
            self._is_default('wire_material_display', self._get_combo_tag(self.wire_material_display_cb)),
        )
        self._apply_default_style(
            self.wire_string_separator_cb,
            self._is_default('wire_string_separator', self._get_combo_tag(self.wire_string_separator_cb)),
        )
        self._apply_default_style(self.feeder_vd_method_cb, self._is_default('feeder_vd_method', fd_value))
        self._apply_default_style(self.write_equipment_cb, self._is_default('write_equipment_results', bool(self.write_equipment_cb.IsChecked)))
        self._apply_default_style(self.write_fixtures_cb, self._is_default('write_fixture_results', bool(self.write_fixtures_cb.IsChecked)))
        self._apply_default_style(self.theme_mode_cb, False)
        self._apply_default_style(self.accent_mode_cb, False)

        self._refresh_validation_state()
        self._update_help_preview()

    def _on_value_changed(self, sender, args):
        self._refresh_styles()
        self._refresh_clear_alert()

    def _on_theme_value_changed(self, sender, args):
        if self._is_loading_ui:
            return
        theme_mode = resource_loader.normalize_theme_mode(self._get_combo_tag(self.theme_mode_cb), self._theme_mode)
        accent_mode = resource_loader.normalize_accent_mode(self._get_combo_tag(self.accent_mode_cb), self._accent_mode)
        if theme_mode == self._theme_mode and accent_mode == self._accent_mode:
            return
        self._theme_mode = theme_mode
        self._accent_mode = accent_mode
        self._apply_theme()
        self._refresh_theme_brushes()
        self._refresh_styles()
        self.theme_mode_default.Text = u"(Current: {})".format(self._describe_theme_mode(self._theme_mode))
        self.accent_mode_default.Text = u"(Current: {})".format(self._describe_accent_mode(self._accent_mode))

    # ------------- Helpers -------------------
    def _describe_neutral(self, value):
        return {
            NeutralBehavior.MATCH_HOT: "Match hot conductors",
            NeutralBehavior.MANUAL: "Manual neutral",
        }.get(value, value)

    def _describe_multipole_branch_neutral(self, value):
        return {
            MultiPoleBranchNeutralBehavior.INCLUDE_BY_DEFAULT: "Include by default",
            MultiPoleBranchNeutralBehavior.EXCLUDE_BY_DEFAULT: "Exclude by default",
        }.get(value, value)

    def _describe_isolated_ground(self, value):
        return {
            IsolatedGroundBehavior.MATCH_GROUND: "Match ground conductors",
            IsolatedGroundBehavior.MANUAL: "Manual isolated ground",
        }.get(value, value)

    def _describe_material_display(self, value):
        return {
            WireMaterialDisplay.AL_ONLY: "Aluminum only",
            WireMaterialDisplay.ALL: "Show material for Copper and Aluminum",
        }.get(value, value)

    def _describe_wire_separator(self, value):
        return {
            WireStringSeparator.PLUS: "Use \"+\" separators",
            WireStringSeparator.COMMA: "Use \",\" separators",
        }.get(value, value)

    def _describe_feeder_method(self, value):
        return {
            FeederVDMethod.DEMAND: "Demand Load",
            FeederVDMethod.CONNECTED: "Connected Load",
            FeederVDMethod.EIGHTY_PERCENT: "80% of Breaker",
            FeederVDMethod.HUNDRED_PERCENT: "100% of Breaker",
        }.get(value, value)

    def _describe_theme_mode(self, value):
        return {
            "light": "Light",
            "dark": "Dark",
            "dark_alt": "Dark Alt",
        }.get(str(value or "").strip().lower(), value)

    def _describe_accent_mode(self, value):
        return {
            "blue": "Blue",
            "neutral": "Neutral",
        }.get(str(value or "").strip().lower(), value)

    def _help_texts(self):
        return {
            'min_conduit_size': "Smallest conduit size proposed during automatic calculations (has no effect on manual user overrides).",
            'max_conduit_fill': "Maximum allowable conduit fill as a percentage. In automatic mode, the conduit will be upsized until this fill is not exceeded. In manual override mode, the tool will alert the user if this value is exceeded.",
            'multi_pole_branch_neutral_behavior': "For first-time calculations on 2/3-pole branch circuits, choose whether neutral is included or excluded by default.",
            'neutral_behavior': "Determines how neutrals are sized when in manual override mode (in automatic mode, neutral size always matches the hot size).",
            'isolated_ground_behavior': "Determines how isolated grounds are sized when in manual override mode (in automatic mode, isolated ground size always matches the ground size).",
            'wire_material_display': "Controls when the material suffix (CU/AL) is shown in wire string outputs.",
            'wire_string_separator': "Controls the separator used between wire parts in wire string outputs.",
            'max_branch_voltage_drop': "Target maximum voltage drop for branch circuits. In automatic mode, calculated sizes will grow until this threshold is met. In manual override mode, the tool will alert the user if this threshold is exceeded.",
            'max_feeder_voltage_drop': "Target maximum voltage drop for feeder circuits. In automatic mode, calculated sizes will grow until this threshold is met. In manual override mode, the tool will alert the user if this threshold is exceeded.",
            'feeder_vd_method': "Which feeder load basis to use for voltage drop calculations and automatic sizing (only applies to feeder circuits that supply panels, switchboards, and transformers). Branch circuits are always based on connected load.",
            'write_results': "Toggle whether calculated results push to downstream elements when present.",
            'clear_writebacks': "Clear persistent data on categories that are currently disabled for write-back. This keeps the window open and honors ownership locks.",
            'theme_mode': "Select the UI theme for CED electrical tools.",
            'accent_mode': "Select the accent color used for highlights and primary actions.",
            'verify_parameters': "Verify required electrical shared parameters in the active project and update bindings when needed.",
        }

    def _option_help(self):
        return {
            'feeder_vd_method': {
                FeederVDMethod.EIGHTY_PERCENT: "[80% of Breaker] Uses 80% of the breaker rating for volt drop calculations unless the actual demand load is higher.",
                FeederVDMethod.HUNDRED_PERCENT: "[100% of Breaker] Uses 100% of the breaker rating for volt drop calculations unless the actual demand load is higher.",
                FeederVDMethod.DEMAND: "[Demand Load] Uses the estimated demand load for volt drop calculations.",
                FeederVDMethod.CONNECTED: "[Connected Load] Uses the connected VA for volt drop calculations with no demand factors applied.",
            },
            'neutral_behavior': {
                NeutralBehavior.MATCH_HOT: "[Match hot conductors] Neutral size will always match hot size.",
                NeutralBehavior.MANUAL: "[Manual Neutral] Neutral size is specified independently in manual override mode.",
            },
            'multi_pole_branch_neutral_behavior': {
                MultiPoleBranchNeutralBehavior.INCLUDE_BY_DEFAULT: "[Include by Default] First-time 2/3-pole branch calculations include neutral.",
                MultiPoleBranchNeutralBehavior.EXCLUDE_BY_DEFAULT: "[Exclude by Default] First-time 2/3-pole branch calculations exclude neutral.",
            },
            'isolated_ground_behavior': {
                IsolatedGroundBehavior.MATCH_GROUND: "[Match ground conductors] Isolated ground size will always match ground size.",
                IsolatedGroundBehavior.MANUAL: "[Manual Isolated Ground] Isolated ground size is specified independently in manual override mode.",
            },
            'wire_material_display': {
                WireMaterialDisplay.AL_ONLY: "[Aluminum only] Only AL circuits include the material suffix.",
                WireMaterialDisplay.ALL: "[Show material for Copper and Aluminum] Both CU and AL circuits include the material suffix.",
            },
            'wire_string_separator': {
                WireStringSeparator.PLUS: "[Use \"+\" separators] Wire parts are separated with plus signs.",
                WireStringSeparator.COMMA: "[Use \",\" separators] Wire parts are separated with commas.",
            },
            'min_conduit_size': {
                '1/2"': u"Selected: 1/2\"",
                '3/4"': u"Selected: 3/4\"",
            },
            'theme_mode': {
                'light': "[Light] Uses the light CED theme.",
                'dark': "[Dark] Uses the dark CED theme.",
                'dark_alt': "[Dark Alt] Uses the alternate dark CED theme.",
            },
            'accent_mode': {
                'blue': "[Blue] Uses blue as the tool accent color.",
                'neutral': "[Neutral] Uses neutral gray as the tool accent color.",
            },
        }

    def _set_help_context(self, key):
        self._help_key = key
        self._update_help_preview()

    def _update_help_preview(self):
        key = self._help_key or ''
        preview = self._help_texts().get(key, "Select a field to see what it controls.")
        option_detail = None
        if key in (
            'feeder_vd_method',
            'multi_pole_branch_neutral_behavior',
            'neutral_behavior',
            'isolated_ground_behavior',
            'wire_material_display',
            'wire_string_separator',
            'min_conduit_size',
            'theme_mode',
            'accent_mode',
        ):
            combo = getattr(self, key + '_cb', None)
            if combo:
                option_detail = self._option_help().get(key, {}).get(self._get_combo_tag(combo), None)

        combined = preview if option_detail is None else u"{}\n{}".format(preview, option_detail)
        self.help_preview.Text = combined

    def _refresh_clear_alert(self):
        has_pending = bool(getattr(self.settings, 'pending_clear_failed', False))
        last_clear_success = bool(getattr(self.settings, 'last_clear_success', False))
        last_equipment_disabled = bool(getattr(self.settings, 'last_clear_equipment_disabled', False))
        last_fixtures_disabled = bool(getattr(self.settings, 'last_clear_fixtures_disabled', False))

        current_equipment_disabled = not bool(self.write_equipment_cb.IsChecked)
        current_fixtures_disabled = not bool(self.write_fixtures_cb.IsChecked)

        stale = last_clear_success and (
            last_equipment_disabled != current_equipment_disabled
            or last_fixtures_disabled != current_fixtures_disabled
        )

        show_alert = has_pending or stale
        self.clear_writeback_alert.Visibility = Visibility.Visible if show_alert else Visibility.Collapsed
        self.clear_writeback_alert.ToolTip = (
            "Some downstream equipment or devices may still have out-of-date data. "
            "Use 'Clear persistent write-back data' to resync disabled categories."
        ) if show_alert else None

    def _set_verify_status(self, text, level=None):
        message = str(text or "").strip()
        self.verify_parameters_status.Text = message
        if not message:
            self.verify_parameters_status.Foreground = self._secondary_text_brush or self._primary_text_brush or Brushes.Black
            return
        if level == "ok":
            brush = self.TryFindResource("CED.Brush.AccentGreen")
        elif level == "warn":
            brush = self.TryFindResource("CED.Brush.BadgeWarningText")
        elif level == "error":
            brush = self.TryFindResource("CED.Brush.AccentRed")
        else:
            brush = self._primary_text_brush
        self.verify_parameters_status.Foreground = brush or self._primary_text_brush or Brushes.Black

    # ------------- Event handlers ------------
    def _on_save(self, sender, args):
        try:
            updated = self._update_settings_from_ui()
        except Exception as ex:
            forms.alert("Could not save settings. Please check your inputs.\n\n{}".format(ex))
            return

        clear_equipment = self._previous_equipment_write and not updated.write_equipment_results
        clear_fixtures = self._previous_fixture_write and not updated.write_fixture_results

        already_cleared = (
            bool(getattr(self.settings, 'last_clear_success', False))
            and not bool(getattr(self.settings, 'pending_clear_failed', False))
            and bool(getattr(self.settings, 'last_clear_equipment_disabled', False)) == (not updated.write_equipment_results)
            and bool(getattr(self.settings, 'last_clear_fixtures_disabled', False)) == (not updated.write_fixture_results)
        )

        if (clear_equipment or clear_fixtures) and not already_cleared:
            msg_parts = [
                "Turning off write-back will clear stored circuit data (numbers to 0, text to blank) on:"
            ]
            if clear_equipment:
                msg_parts.append("• Electrical Equipment")
            if clear_fixtures:
                msg_parts.append("• Fixtures and Devices")
            msg_parts.append("This will run after you save settings.")
            choice = forms.alert("\n".join(msg_parts), ok=False, yes=True, no=True, options=["Proceed and Clear", "Cancel"])
            if choice != "Proceed and Clear":
                return

        if (clear_equipment or clear_fixtures) and not already_cleared:
            try:
                cleared_equip, cleared_fix, locked = settings_manager.clear_downstream_results(
                    self.doc,
                    clear_equipment=clear_equipment,
                    clear_fixtures=clear_fixtures,
                    logger=logger,
                    check_ownership=True,
                )
                if locked:
                    updated.set('pending_clear_failed', True)
                    updated.set('last_clear_success', False)
                    locked_msg = [
                        "Some elements could not be cleared because they are owned by other users.",
                        "Equipment/fixtures cleared: {} / {}".format(cleared_equip, cleared_fix),
                    ]
                    forms.alert("\n".join(locked_msg))
                else:
                    updated.set('pending_clear_failed', False)
                    updated.set('last_clear_success', True)
                    updated.set('last_clear_equipment_disabled', not updated.write_equipment_results)
                    updated.set('last_clear_fixtures_disabled', not updated.write_fixture_results)
                    forms.alert(
                        "Cleared stored circuit data on {} equipment and {} fixtures.".format(
                            cleared_equip, cleared_fix
                        )
                    )
            except Exception as ex:
                updated.set('pending_clear_failed', True)
                updated.set('last_clear_success', False)
                logger.error("Failed to clear downstream circuit data: {}".format(ex))

        settings_manager.save_circuit_settings(self.doc, updated)
        self.settings = updated
        self._theme_mode = resource_loader.normalize_theme_mode(self._get_combo_tag(self.theme_mode_cb), self._theme_mode)
        self._accent_mode = resource_loader.normalize_accent_mode(self._get_combo_tag(self.accent_mode_cb), self._accent_mode)
        _save_theme_state(self._theme_mode, self._accent_mode)
        self._previous_equipment_write = bool(self.settings.write_equipment_results)
        self._previous_fixture_write = bool(self.settings.write_fixture_results)
        self._last_clear_equipment_disabled = bool(getattr(self.settings, 'last_clear_equipment_disabled', False))
        self._last_clear_fixtures_disabled = bool(getattr(self.settings, 'last_clear_fixtures_disabled', False))
        self._last_clear_success = bool(getattr(self.settings, 'last_clear_success', False))
        self._refresh_clear_alert()

        forms.alert("Calculate Circuits settings saved to project.")
        self.Close()

    def _on_cancel(self, sender, args):
        self.Close()

    def _on_reset(self, sender, args):
        self.settings = CircuitSettings()
        self._load_values()
        self._refresh_styles()


    def _on_help(self,sender, args):
        output = script.get_output()
        md_path = _resolve_bundle_or_local_path("CalculateCircuits_UserManual.md")
        if not os.path.exists(md_path):
            forms.alert("User manual not found.")
            return
        with open(md_path, "r") as f:
            text = f.read().decode("utf-8")
            output.print_md(text)

    def _on_verify_parameters(self, sender, args):
        self.verify_parameters_btn.IsEnabled = False
        self._set_verify_status("Verifying project parameters...", level=None)
        try:
            preview_settings = self._update_settings_from_ui()
            result = _verify_project_parameters(self.doc, self.doc.Application, settings=preview_settings)
            status = str(result.get("status") or "").lower()
            warnings = list(result.get("warnings") or [])
            errors = list(result.get("errors") or [])
            locked = list(result.get("locked") or [])
            if status == "failed":
                self._set_verify_status("Parameter verification failed", level="error")
                forms.alert(
                    "Verify Parameters failed.\n\n{}".format(
                        str(result.get("reason") or "Unknown error.")
                    )
                )
                return
            if not warnings and not errors and not locked:
                self._set_verify_status("All Project Parameters configured", level="ok")
                return

            self._set_verify_status("Parameter verification completed with warnings", level="warn")
            summary = [
                "Parameter verification completed with warnings.",
                "",
                "Updated: {}".format(result.get("updated", 0)),
                "Unchanged: {}".format(result.get("unchanged", 0)),
                "Skipped: {}".format(result.get("skipped", 0)),
                "Category unbind updates: {}".format(result.get("unbound", 0)),
            ]
            if locked:
                summary.append("Skipped (owned by others): {}".format(len(locked)))
            if warnings:
                summary.append("")
                summary.append("Warnings:")
                for message in warnings[:12]:
                    summary.append("- {}".format(message))
                if len(warnings) > 12:
                    summary.append("- ...and {} more warnings".format(len(warnings) - 12))
            if locked:
                summary.append("")
                summary.append("Owned By Other User:")
                for item in locked[:12]:
                    summary.append(
                        "- {} (owner: {})".format(
                            str(item.get("parameter") or "Unnamed Parameter"),
                            str(item.get("owner") or "Unknown"),
                        )
                    )
                if len(locked) > 12:
                    summary.append("- ...and {} more locked parameters".format(len(locked) - 12))
            if errors:
                summary.append("")
                summary.append("Errors:")
                for message in errors[:12]:
                    summary.append("- {}".format(message))
                if len(errors) > 12:
                    summary.append("- ...and {} more errors".format(len(errors) - 12))
            forms.alert("\n".join(summary))
        except Exception as ex:
            self._set_verify_status("Parameter verification failed", level="error")
            forms.alert("Verify Parameters failed.\n\n{}".format(ex))
        finally:
            self.verify_parameters_btn.IsEnabled = True

    def _on_clear_persistent(self, sender, args):
        clear_equipment = not bool(self.write_equipment_cb.IsChecked)
        clear_fixtures = not bool(self.write_fixtures_cb.IsChecked)

        if not (clear_equipment or clear_fixtures):
            forms.alert("Both categories are enabled for write-back; nothing to clear.")
            return

        try:
            cleared_equip, cleared_fix, locked = settings_manager.clear_downstream_results(
                self.doc,
                clear_equipment=clear_equipment,
                clear_fixtures=clear_fixtures,
                logger=logger,
                check_ownership=True,
            )
            if locked:
                self.settings.set('pending_clear_failed', True)
                self.settings.set('last_clear_success', False)
                forms.alert("Some elements are owned by others and could not be cleared. Please try again later.")
            else:
                self.settings.set('pending_clear_failed', False)
                self.settings.set('last_clear_success', True)
                self.settings.set('last_clear_equipment_disabled', clear_equipment)
                self.settings.set('last_clear_fixtures_disabled', clear_fixtures)
                forms.alert(
                    "Cleared stored circuit data on {} equipment and {} fixtures.".format(
                        cleared_equip, cleared_fix
                    )
                )
            settings_manager.save_circuit_settings(self.doc, self.settings)
        except Exception as ex:
            self.settings.set('pending_clear_failed', True)
            self.settings.set('last_clear_success', False)
            logger.error("Failed to clear downstream circuit data: {}".format(ex))

        self._refresh_clear_alert()

    def _percent_value(self, decimal_value):
        return round(float(decimal_value) * 100, 3)

    def _percent_string(self, decimal_value):
        return u"{} %".format(self._strip_trailing_zeros(self._percent_value(decimal_value)))

    def _strip_trailing_zeros(self, number):
        text = ("{0:.5f}".format(number)).rstrip('0').rstrip('.')
        return text

    def _set_percent_field(self, textbox, decimal_value):
        if decimal_value is None:
            textbox.Text = ""
            return
        textbox.Text = self._percent_string(decimal_value)

    def _parse_percent_field(self, textbox, min_value, max_value, warning_block=None, warn_threshold=None, silent=False):
        text = textbox.Text.strip()
        label = getattr(textbox, 'Tag', None) or textbox.Name
        if not text:
            if silent:
                return None
            raise ValueError("A value is required for {}.".format(label))

        has_percent = '%' in text
        raw = text.replace('%', '').strip()
        try:
            numeric = float(raw)
        except Exception:
            if silent:
                return None
            raise ValueError("{} must be numeric.".format(label))

        if has_percent:
            decimal_value = numeric / 100.0
        else:
            decimal_value = numeric / 100.0 if numeric > 1 else numeric

        decimal_value = round(decimal_value, 3)

        if decimal_value < min_value:
            decimal_value = min_value
        elif decimal_value > max_value:
            decimal_value = max_value

        if warning_block and warn_threshold is not None:
            warning_block.Visibility = Visibility.Visible if decimal_value > warn_threshold else Visibility.Collapsed

        if not silent:
            self._normalize_percent_text(textbox, decimal_value)

        return decimal_value

    def _normalize_percent_text(self, textbox, decimal_value):
        if self._is_normalizing:
            return
        self._is_normalizing = True
        try:
            textbox.Text = self._percent_string(decimal_value)
            textbox.CaretIndex = len(textbox.Text or "")
        finally:
            self._is_normalizing = False

    def _normalize_percent_on_blur(self, textbox, min_value, max_value, warning_block, warn_threshold):
        try:
            self._parse_percent_field(textbox, min_value, max_value, warning_block, warn_threshold, silent=False)
        except Exception:
            self._normalize_percent_text(textbox, min_value)
            if warning_block and warn_threshold is not None:
                warning_block.Visibility = Visibility.Visible if min_value > warn_threshold else Visibility.Collapsed

    def _on_percent_value_changed(self, textbox, min_value, max_value, warning_block, warn_threshold):
        self._sanitize_percent_input(textbox)
        self._update_warning(textbox, min_value, max_value, warning_block, warn_threshold)
        self._refresh_styles()

    def _sanitize_percent_input(self, textbox):
        if self._is_normalizing:
            return
        text = textbox.Text
        allowed = []
        dot_seen = False
        percent_seen = False
        for ch in text:
            if ch.isdigit():
                allowed.append(ch)
            elif ch == '.' and not dot_seen:
                allowed.append(ch)
                dot_seen = True
            elif ch == '%' and not percent_seen:
                percent_seen = True
            # other characters are dropped

        if percent_seen:
            allowed.append('%')

        cleaned = ''.join(allowed)
        if cleaned != text:
            caret = textbox.CaretIndex
            textbox.Text = cleaned
            textbox.CaretIndex = min(caret, len(cleaned))

    def _percent_box_preview_mouse_down(self, sender, args):
        if sender is None:
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

    def _percent_box_got_focus(self, sender, args):
        try:
            sender.SelectAll()
        except Exception:
            pass

    def _refresh_validation_state(self):
        self._update_warning(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4)
        self._update_warning(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05)
        self._update_warning(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05)

    def _update_warning(self, textbox, min_value, max_value, warning_block, warn_threshold):
        value = self._parse_percent_field(textbox, min_value, max_value, warning_block, warn_threshold, silent=True)
        if warning_block and warn_threshold is not None:
            if value is None:
                warning_block.Visibility = Visibility.Collapsed
            else:
                warning_block.Visibility = Visibility.Visible if value > warn_threshold else Visibility.Collapsed


if __name__ == '__main__':
    window = CircuitSettingsWindow()
    window.show_dialog()
