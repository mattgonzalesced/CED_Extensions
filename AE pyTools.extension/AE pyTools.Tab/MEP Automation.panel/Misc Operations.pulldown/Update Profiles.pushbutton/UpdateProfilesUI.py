# -*- coding: utf-8 -*-
"""
UpdateProfilesUI
----------------
WPF window to resolve Update Profiles discrepancies in a single grid.
"""

from pyrevit import forms
from System.Windows import Thickness, GridLength, GridUnitType, Visibility
from System.Windows.Controls import (
    Grid,
    RowDefinition,
    ColumnDefinition,
    TextBlock,
    ComboBox,
)
from System.Windows import FontWeights


class UpdateProfilesWindow(forms.WPFWindow):
    def __init__(self, xaml_path, discrepancies, param_names, include_tags, include_notes):
        forms.WPFWindow.__init__(self, xaml_path)
        self._discrepancies = discrepancies or []
        self._param_names = param_names or []
        self._include_tags = bool(include_tags)
        self._include_notes = bool(include_notes)
        self._combo_meta = {}
        self.decisions = {}
        self._replace_checkbox = None
        self._tag_set_combo = None
        self._keynote_set_combo = None
        self._note_set_combo = None

        header = self.FindName("HeaderText")
        if header is not None:
            header.Text = "Choose how each discrepancy should update the profile."

        self._init_global_controls()
        self._build_grid()
        apply_btn = self.FindName("ApplyButton")
        cancel_btn = self.FindName("CancelButton")
        if apply_btn is not None:
            apply_btn.Click += self._on_apply
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    def _display_value(self, value):
        if value is None:
            return "<none>"
        if isinstance(value, float):
            return "{:.4f}".format(value)
        return str(value)

    def _init_global_controls(self):
        self._replace_checkbox = self.FindName("ReplaceModeCheckBox")
        self._tag_set_combo = self.FindName("TagSetCombo")
        self._keynote_set_combo = self.FindName("KeynoteSetCombo")
        self._note_set_combo = self.FindName("TextNoteSetCombo")
        tag_row = self.FindName("TagSetRow")
        keynote_row = self.FindName("KeynoteSetRow")
        note_row = self.FindName("TextNoteSetRow")
        options_panel = self.FindName("ReplaceOptionsPanel")

        if not self._include_tags:
            if tag_row is not None:
                tag_row.Visibility = Visibility.Collapsed
            if keynote_row is not None:
                keynote_row.Visibility = Visibility.Collapsed
        if not self._include_notes:
            if note_row is not None:
                note_row.Visibility = Visibility.Collapsed
        if not self._include_tags and not self._include_notes:
            if self._replace_checkbox is not None:
                self._replace_checkbox.Visibility = Visibility.Collapsed
            if options_panel is not None:
                options_panel.Visibility = Visibility.Collapsed

        option_labels = [
            "Use items found on any instance (union)",
            "Use only items found on every instance (common)",
        ]
        for combo in (self._tag_set_combo, self._keynote_set_combo, self._note_set_combo):
            if combo is None:
                continue
            combo.Items.Clear()
            for label in option_labels:
                combo.Items.Add(label)
            combo.SelectedIndex = 0

        if self._replace_checkbox is not None:
            self._replace_checkbox.Checked += self._on_replace_toggle
            self._replace_checkbox.Unchecked += self._on_replace_toggle
        self._update_replace_controls()

    def _update_replace_controls(self):
        enabled = False
        if self._replace_checkbox is not None:
            enabled = bool(self._replace_checkbox.IsChecked)
        for combo in (self._tag_set_combo, self._keynote_set_combo, self._note_set_combo):
            if combo is None:
                continue
            combo.IsEnabled = enabled

    def _on_replace_toggle(self, sender, args):
        self._update_replace_controls()

    def _selected_set(self, combo):
        selection = combo.SelectedItem if combo is not None else None
        if isinstance(selection, list):
            selection = selection[0] if selection else None
        if selection and str(selection).startswith("Use only"):
            return "common"
        return "union"

    def _build_grid(self):
        grid = self.FindName("DiscrepancyGrid")
        if grid is None:
            raise Exception("DiscrepancyGrid not found in XAML.")

        columns = ["Profile", "Type"]
        for name in self._param_names:
            columns.append(name)

        for _ in columns:
            col_def = ColumnDefinition()
            col_def.Width = GridLength(240, GridUnitType.Pixel)
            grid.ColumnDefinitions.Add(col_def)

        header_row = RowDefinition()
        header_row.Height = GridLength(28, GridUnitType.Pixel)
        grid.RowDefinitions.Add(header_row)

        for col_idx, title in enumerate(columns):
            cell = TextBlock()
            cell.Text = title
            cell.FontWeight = FontWeights.Bold
            cell.Margin = Thickness(0, 0, 6, 4)
            Grid.SetRow(cell, 0)
            Grid.SetColumn(cell, col_idx)
            grid.Children.Add(cell)

        row_idx = 1
        for item in self._discrepancies:
            row_def = RowDefinition()
            row_def.Height = GridLength(32, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)

            profile_label = item.get("profile_name") or item.get("profile") or item.get("label") or item.get("led_id")
            type_label = item.get("type_label") or item.get("label") or item.get("led_id")

            profile_cell = TextBlock()
            profile_cell.Text = profile_label
            profile_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(profile_cell, row_idx)
            Grid.SetColumn(profile_cell, 0)
            grid.Children.Add(profile_cell)

            type_cell = TextBlock()
            type_cell.Text = type_label
            type_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(type_cell, row_idx)
            Grid.SetColumn(type_cell, 1)
            grid.Children.Add(type_cell)
            col_offset = 2
            for name in self._param_names:
                param_info = item.get("param_conflicts", {}).get(name)
                if not param_info:
                    cell = TextBlock()
                    cell.Text = "â€”"
                    cell.Margin = Thickness(0, 0, 6, 2)
                    Grid.SetRow(cell, row_idx)
                    Grid.SetColumn(cell, col_offset)
                    grid.Children.Add(cell)
                    col_offset += 1
                    continue
                combo = ComboBox()
                option_map = {}
                for entry in sorted(param_info.values(), key=lambda e: e["count"], reverse=True):
                    value = entry["value"]
                    label = "Use {} ({} instances)".format(self._display_value(value), entry["count"])
                    combo.Items.Add(label)
                    option_map[label] = ("value", value)
                existing_value = item.get("existing_params", {}).get(name)
                keep_label = "Keep existing profile value: {}".format(self._display_value(existing_value))
                combo.Items.Add(keep_label)
                option_map[keep_label] = ("keep", None)
                skip_label = "Skip updating this parameter"
                combo.Items.Add(skip_label)
                option_map[skip_label] = ("skip", None)
                combo.SelectedIndex = 0
                tooltip_lines = []
                for entry in sorted(param_info.values(), key=lambda e: e["count"], reverse=True):
                    tooltip_lines.append("{} ({} instances)".format(self._display_value(entry["value"]), entry["count"]))
                combo.ToolTip = "\n".join(tooltip_lines)
                Grid.SetRow(combo, row_idx)
                Grid.SetColumn(combo, col_offset)
                grid.Children.Add(combo)
                self._combo_meta[combo] = {
                    "led_id": item["led_id"],
                    "kind": "param",
                    "param_name": name,
                    "options": option_map,
                }
                col_offset += 1

            row_idx += 1

    def _on_apply(self, sender, args):
        global_settings = {
            "replace_mode": bool(self._replace_checkbox and self._replace_checkbox.IsChecked),
            "tag_set": self._selected_set(self._tag_set_combo),
            "keynote_set": self._selected_set(self._keynote_set_combo),
            "note_set": self._selected_set(self._note_set_combo),
        }
        self.decisions["_global"] = global_settings
        for combo, meta in self._combo_meta.items():
            selection = combo.SelectedItem
            if isinstance(selection, list):
                selection = selection[0] if selection else None
            if not selection:
                continue
            led_id = meta["led_id"]
            entry = self.decisions.setdefault(led_id, {"params": {}})
            if meta.get("kind") != "param":
                continue
            action, value = meta["options"].get(selection, ("skip", None))
            entry["params"][meta["param_name"]] = (action, value)
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

