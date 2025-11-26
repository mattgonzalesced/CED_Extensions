# -*- coding: utf-8 -*-

from pyrevit import forms, revit, script
from System.Windows import FontStyles
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
        self._help_window = None
        self._help_key = None

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
        self.help_btn.Click += self._on_help

        self.min_conduit_size_tb.TextChanged += self._on_value_changed
        self.min_conduit_size_tb.GotFocus += lambda s, e: self._set_help_context('min_conduit_size')
        self.max_conduit_fill_tb.TextChanged += self._on_value_changed
        self.max_conduit_fill_tb.GotFocus += lambda s, e: self._set_help_context('max_conduit_fill')
        self.max_branch_vd_tb.TextChanged += self._on_value_changed
        self.max_branch_vd_tb.GotFocus += lambda s, e: self._set_help_context('max_branch_voltage_drop')
        self.max_feeder_vd_tb.TextChanged += self._on_value_changed
        self.max_feeder_vd_tb.GotFocus += lambda s, e: self._set_help_context('max_feeder_voltage_drop')
        self.neutral_behavior_cb.SelectionChanged += self._on_value_changed
        self.neutral_behavior_cb.GotFocus += lambda s, e: self._set_help_context('neutral_behavior')
        self.feeder_vd_method_cb.SelectionChanged += self._on_value_changed
        self.feeder_vd_method_cb.GotFocus += lambda s, e: self._set_help_context('feeder_vd_method')

    # ------------- Data helpers --------------
    def _load_defaults_panel(self):
        self.min_conduit_default.Text = u"(Default: {})".format(self.defaults.min_conduit_size)
        self.max_conduit_fill_default.Text = u"(Default: {})".format(self.defaults.max_conduit_fill)
        self.neutral_behavior_default.Text = u"(Default: {})".format(self._describe_neutral(self.defaults.neutral_behavior))
        self.max_branch_vd_default.Text = u"(Default: {})".format(self.defaults.max_branch_voltage_drop)
        self.max_feeder_vd_default.Text = u"(Default: {})".format(self.defaults.max_feeder_voltage_drop)
        self.feeder_vd_method_default.Text = u"(Default: {})".format(self._describe_feeder_method(self.defaults.feeder_vd_method))

    def _load_values(self):
        self.min_conduit_size_tb.Text = self.settings.min_conduit_size
        self.max_conduit_fill_tb.Text = str(self.settings.max_conduit_fill)
        self.max_branch_vd_tb.Text = str(self.settings.max_branch_voltage_drop)
        self.max_feeder_vd_tb.Text = str(self.settings.max_feeder_voltage_drop)

        self._select_combo_by_tag(self.neutral_behavior_cb, self.settings.neutral_behavior)
        self._select_combo_by_tag(self.feeder_vd_method_cb, self.settings.feeder_vd_method)

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
        updated.set('min_conduit_size', self.min_conduit_size_tb.Text.strip())
        updated.set('max_conduit_fill', float(self.max_conduit_fill_tb.Text))
        updated.set('neutral_behavior', self._get_combo_tag(self.neutral_behavior_cb))
        updated.set('max_branch_voltage_drop', float(self.max_branch_vd_tb.Text))
        updated.set('max_feeder_voltage_drop', float(self.max_feeder_vd_tb.Text))
        updated.set('feeder_vd_method', self._get_combo_tag(self.feeder_vd_method_cb))
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
        self._apply_default_style(self.min_conduit_size_tb, self._is_default('min_conduit_size', self.min_conduit_size_tb.Text))
        self._apply_default_style(self.max_conduit_fill_tb, self._is_default('max_conduit_fill', self.max_conduit_fill_tb.Text))
        self._apply_default_style(self.max_branch_vd_tb, self._is_default('max_branch_voltage_drop', self.max_branch_vd_tb.Text))
        self._apply_default_style(self.max_feeder_vd_tb, self._is_default('max_feeder_voltage_drop', self.max_feeder_vd_tb.Text))

        nb_value = self._get_combo_tag(self.neutral_behavior_cb)
        fd_value = self._get_combo_tag(self.feeder_vd_method_cb)
        self._apply_default_style(self.neutral_behavior_cb, self._is_default('neutral_behavior', nb_value))
        self._apply_default_style(self.feeder_vd_method_cb, self._is_default('feeder_vd_method', fd_value))

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
            'max_conduit_fill': "Maximum allowable conduit fill as a decimal (e.g., 0.36 for 36%). The conduit will be upsized until this fill is not exceeded.",
            'neutral_behavior': "Choose whether neutral conductors match the quantity of hots automatically or are entered manually for each calculation.",
            'max_branch_voltage_drop': "Target maximum voltage drop for branch circuits. Calculated sizes will grow until this threshold is met.",
            'max_feeder_voltage_drop': "Target maximum voltage drop for feeders. Applies to feeder sizing logic and may differ from branch criteria.",
            'feeder_vd_method': "Which feeder load basis to use for voltage drop checks: demand load, connected load, or 80% of the maximum.",
        }

    def _set_help_context(self, key):
        self._help_key = key
        self._update_help_preview()
        if self._help_window:
            self._help_window.help_text.Text = self._help_texts().get(key, "")

    def _update_help_preview(self):
        key = self._help_key or ''
        preview = self._help_texts().get(key, "Select a field to see what it controls.")
        self.help_preview.Text = preview

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

    def _on_help(self, sender, args):
        if not self._help_window:
            self._help_window = self._create_help_window()

        # Always update text for current context
        self._help_window.help_text.Text = self._help_texts().get(self._help_key, "")
        self._help_window.Show()

    def _create_help_window(self):
        xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Calculate Circuits Help" Height="240" Width="420" Background="#FFFDFDFD"
        Foreground="#222" WindowStartupLocation="CenterOwner" ResizeMode="CanResize">
    <Border Margin="14" Padding="12" Background="#FFF7F7F7" BorderBrush="#DDD" BorderThickness="1">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
            <TextBlock x:Name="help_text" TextWrapping="Wrap" FontSize="13"/>
        </ScrollViewer>
    </Border>
</Window>
"""
        hw = forms.WPFWindow(xaml)
        hw.Owner = self
        return hw


if __name__ == '__main__':
    window = CircuitSettingsWindow()
    window.show_dialog()
