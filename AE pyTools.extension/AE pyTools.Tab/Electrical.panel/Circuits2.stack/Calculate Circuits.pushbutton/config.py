# -*- coding: utf-8 -*-

from pyrevit import forms, revit, script
from System.Windows import FontStyles, Visibility
from System.Windows.Media import Brushes

from CEDElectrical.Model.circuit_settings import (
    CircuitSettings,
    FeederVDMethod,
    NeutralBehavior,
)
from CEDElectrical.Domain import settings_manager

XAML_PATH = script.get_bundle_file('settings.xaml')


class CircuitSettingsWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_PATH)
        self.doc = revit.doc
        self.defaults = CircuitSettings()
        self.settings = settings_manager.load_circuit_settings(self.doc)
        self._help_key = None
        self._is_normalizing = False

        self._bind_events()
        self._load_defaults_panel()
        self._load_values()
        self._refresh_styles()
        self._set_help_context('min_conduit_size')

    # ------------- UI wiring -----------------
    def _bind_events(self):
        self.save_btn.Click += self._on_save
        self.cancel_btn.Click += self._on_cancel
        self.reset_btn.Click += self._on_reset
        self.min_conduit_size_cb.SelectionChanged += self._on_value_changed
        self.min_conduit_size_cb.GotFocus += lambda s, e: self._set_help_context('min_conduit_size')
        self.max_conduit_fill_tb.GotFocus += lambda s, e: self._set_help_context('max_conduit_fill')
        self.max_conduit_fill_tb.LostFocus += lambda s, e: self._normalize_percent_on_blur(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4)
        self.max_branch_vd_tb.GotFocus += lambda s, e: self._set_help_context('max_branch_voltage_drop')
        self.max_branch_vd_tb.LostFocus += lambda s, e: self._normalize_percent_on_blur(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05)
        self.max_feeder_vd_tb.GotFocus += lambda s, e: self._set_help_context('max_feeder_voltage_drop')
        self.max_feeder_vd_tb.LostFocus += lambda s, e: self._normalize_percent_on_blur(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05)
        self.neutral_behavior_cb.SelectionChanged += self._on_value_changed
        self.neutral_behavior_cb.GotFocus += lambda s, e: self._set_help_context('neutral_behavior')
        self.feeder_vd_method_cb.SelectionChanged += self._on_value_changed
        self.feeder_vd_method_cb.GotFocus += lambda s, e: self._set_help_context('feeder_vd_method')
        self.write_equipment_cb.Checked += self._on_value_changed
        self.write_equipment_cb.Unchecked += self._on_value_changed
        self.write_equipment_cb.GotFocus += lambda s, e: self._set_help_context('write_results')
        self.write_fixtures_cb.Checked += self._on_value_changed
        self.write_fixtures_cb.Unchecked += self._on_value_changed
        self.write_fixtures_cb.GotFocus += lambda s, e: self._set_help_context('write_results')

        self.max_conduit_fill_tb.TextChanged += lambda s, e: self._on_percent_value_changed(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4)
        self.max_branch_vd_tb.TextChanged += lambda s, e: self._on_percent_value_changed(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05)
        self.max_feeder_vd_tb.TextChanged += lambda s, e: self._on_percent_value_changed(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05)

    # ------------- Data helpers --------------
    def _load_defaults_panel(self):
        self.min_conduit_default.Text = u"(Default: {})".format(self.defaults.min_conduit_size)
        self.max_conduit_fill_default.Text = u"(Default: {}%)".format(self._percent_value(self.defaults.max_conduit_fill))
        self.neutral_behavior_default.Text = u"(Default: {})".format(self._describe_neutral(self.defaults.neutral_behavior))
        self.max_branch_vd_default.Text = u"(Default: {}%)".format(self._percent_value(self.defaults.max_branch_voltage_drop))
        self.max_feeder_vd_default.Text = u"(Default: {}%)".format(self._percent_value(self.defaults.max_feeder_voltage_drop))
        self.feeder_vd_method_default.Text = u"(Default: {})".format(self._describe_feeder_method(self.defaults.feeder_vd_method))
        self.write_results_default.Text = u"(Defaults: Equipment ✓, Fixtures ✕)"

    def _load_values(self):
        self._select_combo_by_tag(self.min_conduit_size_cb, self.settings.min_conduit_size)
        self._set_percent_field(self.max_conduit_fill_tb, self.settings.max_conduit_fill)
        self._set_percent_field(self.max_branch_vd_tb, self.settings.max_branch_voltage_drop)
        self._set_percent_field(self.max_feeder_vd_tb, self.settings.max_feeder_voltage_drop)

        self._select_combo_by_tag(self.neutral_behavior_cb, self.settings.neutral_behavior)
        self._select_combo_by_tag(self.feeder_vd_method_cb, self.settings.feeder_vd_method)

        self.write_equipment_cb.IsChecked = bool(self.settings.write_equipment_results)
        self.write_fixtures_cb.IsChecked = bool(self.settings.write_fixture_results)

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
        updated.set('neutral_behavior', self._get_combo_tag(self.neutral_behavior_cb))
        updated.set('max_branch_voltage_drop', self._parse_percent_field(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05))
        updated.set('max_feeder_voltage_drop', self._parse_percent_field(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05))
        updated.set('feeder_vd_method', self._get_combo_tag(self.feeder_vd_method_cb))
        updated.set('write_equipment_results', bool(self.write_equipment_cb.IsChecked))
        updated.set('write_fixture_results', bool(self.write_fixtures_cb.IsChecked))
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
        control.FontStyle = FontStyles.Italic if is_default else FontStyles.Normal
        control.Foreground = Brushes.Gray if is_default else Brushes.Black

    def _refresh_styles(self, sender=None, args=None):
        self._apply_default_style(self.min_conduit_size_cb, self._is_default('min_conduit_size', self._get_combo_tag(self.min_conduit_size_cb)))
        self._apply_default_style(self.max_conduit_fill_tb, self._is_default('max_conduit_fill', self._parse_percent_field(self.max_conduit_fill_tb, 0.1, 1.0, self.max_conduit_fill_warn, 0.4, silent=True)))
        self._apply_default_style(self.max_branch_vd_tb, self._is_default('max_branch_voltage_drop', self._parse_percent_field(self.max_branch_vd_tb, 0.001, 1.0, self.max_branch_vd_warn, 0.05, silent=True)))
        self._apply_default_style(self.max_feeder_vd_tb, self._is_default('max_feeder_voltage_drop', self._parse_percent_field(self.max_feeder_vd_tb, 0.001, 1.0, self.max_feeder_vd_warn, 0.05, silent=True)))

        nb_value = self._get_combo_tag(self.neutral_behavior_cb)
        fd_value = self._get_combo_tag(self.feeder_vd_method_cb)
        self._apply_default_style(self.neutral_behavior_cb, self._is_default('neutral_behavior', nb_value))
        self._apply_default_style(self.feeder_vd_method_cb, self._is_default('feeder_vd_method', fd_value))
        self._apply_default_style(self.write_equipment_cb, self._is_default('write_equipment_results', bool(self.write_equipment_cb.IsChecked)))
        self._apply_default_style(self.write_fixtures_cb, self._is_default('write_fixture_results', bool(self.write_fixtures_cb.IsChecked)))

        self._refresh_validation_state()
        self._update_help_preview()

    def _on_value_changed(self, sender, args):
        self._refresh_styles()

    # ------------- Helpers -------------------
    def _describe_neutral(self, value):
        return {
            NeutralBehavior.MATCH_HOT: "Match hot conductors",
            NeutralBehavior.MANUAL: "Manual neutral",
        }.get(value, value)

    def _describe_feeder_method(self, value):
        return {
            FeederVDMethod.DEMAND: "Demand",
            FeederVDMethod.CONNECTED: "Connected",
            FeederVDMethod.EIGHTY_PERCENT: "80% of Max",
        }.get(value, value)

    def _help_texts(self):
        return {
            'min_conduit_size': "Smallest conduit size that will ever be proposed by the calculation. This limits conduit downsizing in tight runs.",
            'max_conduit_fill': "Maximum allowable conduit fill as a percentage (e.g., 36%). The conduit will be upsized until this fill is not exceeded.",
            'neutral_behavior': "Choose whether neutral conductors match the quantity of hots automatically or are entered manually for each calculation.",
            'max_branch_voltage_drop': "Target maximum voltage drop for branch circuits (enter as a percentage). Calculated sizes will grow until this threshold is met.",
            'max_feeder_voltage_drop': "Target maximum voltage drop for feeders (enter as a percentage). Applies to feeder sizing logic and may differ from branch criteria.",
            'feeder_vd_method': "Which feeder load basis to use for voltage drop checks: demand load, connected load, or 80% of the maximum.",
            'write_results': "Toggle whether calculated results push to downstream elements when present.",
        }

    def _option_help(self):
        return {
            'feeder_vd_method': {
                FeederVDMethod.DEMAND: "Uses demand load (after demand factors) for voltage drop calculations.",
                FeederVDMethod.CONNECTED: "Uses the raw connected load without demand factors for voltage drop checks.",
                FeederVDMethod.EIGHTY_PERCENT: "Uses 80% of the maximum load to balance conservatism and practicality.",
            },
            'neutral_behavior': {
                NeutralBehavior.MATCH_HOT: "Neutrals will always mirror the quantity of hot conductors.",
                NeutralBehavior.MANUAL: "You will be prompted to specify neutral counts when calculating.",
            },
            'min_conduit_size': {
                '1/2"': "Choose for the tightest runs where smaller conduit is acceptable.",
                '3/4"': "Common baseline that balances pullability and capacity.",
            },
        }

    def _set_help_context(self, key):
        self._help_key = key
        self._update_help_preview()

    def _update_help_preview(self):
        key = self._help_key or ''
        preview = self._help_texts().get(key, "Select a field to see what it controls.")
        option_detail = None
        if key in ('feeder_vd_method', 'neutral_behavior', 'min_conduit_size'):
            combo = getattr(self, key + '_cb', None)
            if combo:
                option_detail = self._option_help().get(key, {}).get(self._get_combo_tag(combo), None)

        combined = preview if option_detail is None else u"{}\n{}".format(preview, option_detail)
        self.help_preview.Text = combined

    # ------------- Event handlers ------------
    def _on_save(self, sender, args):
        try:
            updated = self._update_settings_from_ui()
        except Exception as ex:
            forms.alert("Could not save settings. Please check your inputs.\n\n{}".format(ex))
            return

        settings_manager.save_circuit_settings(self.doc, updated)
        forms.alert("Calculate Circuits settings saved to project.")
        self.Close()

    def _on_cancel(self, sender, args):
        self.Close()

    def _on_reset(self, sender, args):
        self.settings = CircuitSettings()
        self._load_values()
        self._refresh_styles()

    def _percent_value(self, decimal_value):
        return round(float(decimal_value) * 100, 2)

    def _percent_string(self, decimal_value):
        return u"{}%".format(self._strip_trailing_zeros(self._percent_value(decimal_value)))

    def _strip_trailing_zeros(self, number):
        text = ("{0:.4f}".format(number)).rstrip('0').rstrip('.')
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

        if decimal_value < min_value or decimal_value > max_value:
            if silent:
                return decimal_value
            raise ValueError("{} must be between {}% and {}%.".format(label, self._percent_value(min_value), self._percent_value(max_value)))

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
            textbox.CaretIndex = len(textbox.Text)
        finally:
            self._is_normalizing = False

    def _normalize_percent_on_blur(self, textbox, min_value, max_value, warning_block, warn_threshold):
        try:
            self._parse_percent_field(textbox, min_value, max_value, warning_block, warn_threshold, silent=False)
        except Exception:
            # Leave value as-is; save flow will surface the validation
            pass

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
