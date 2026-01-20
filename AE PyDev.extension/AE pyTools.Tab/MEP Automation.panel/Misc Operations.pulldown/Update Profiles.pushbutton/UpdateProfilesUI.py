# -*- coding: utf-8 -*-
"""
UpdateProfilesUI
----------------
WPF window to resolve Update Profiles discrepancies in a single grid.
"""

from pyrevit import forms
from System.Windows import Thickness, HorizontalAlignment, GridLength, GridUnitType
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

        header = self.FindName("HeaderText")
        if header is not None:
            header.Text = "Choose how each discrepancy should update the profile."

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

    def _build_grid(self):
        grid = self.FindName("DiscrepancyGrid")
        if grid is None:
            raise Exception("DiscrepancyGrid not found in XAML.")

        columns = ["Profile", "Type"]
        if self._include_tags:
            columns.append("Tags/Keynotes")
        if self._include_notes:
            columns.append("Text Notes")
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
            if self._include_tags:
                missing = item.get("missing_tags") or []
                if missing:
                    combo = ComboBox()
                    options = [
                        "Add items found on any instance (union)",
                        "Only update items found on every instance",
                        "Skip tag updates for this definition",
                    ]
                    for opt in options:
                        combo.Items.Add(opt)
                    combo.SelectedIndex = 0
                    combo.ToolTip = "Missing on some instances:\n" + "\n".join(missing[:8])
                    Grid.SetRow(combo, row_idx)
                    Grid.SetColumn(combo, col_offset)
                    grid.Children.Add(combo)
                    self._combo_meta[combo] = {"led_id": item["led_id"], "kind": "tags"}
                else:
                    cell = TextBlock()
                    cell.Text = "—"
                    cell.Margin = Thickness(0, 0, 6, 2)
                    Grid.SetRow(cell, row_idx)
                    Grid.SetColumn(cell, col_offset)
                    grid.Children.Add(cell)
                col_offset += 1

            if self._include_notes:
                missing = item.get("missing_notes") or []
                if missing:
                    combo = ComboBox()
                    options = [
                        "Add items found on any instance (union)",
                        "Only update items found on every instance",
                        "Skip text note updates for this definition",
                    ]
                    for opt in options:
                        combo.Items.Add(opt)
                    combo.SelectedIndex = 0
                    combo.ToolTip = "Missing on some instances:\n" + "\n".join(missing[:8])
                    Grid.SetRow(combo, row_idx)
                    Grid.SetColumn(combo, col_offset)
                    grid.Children.Add(combo)
                    self._combo_meta[combo] = {"led_id": item["led_id"], "kind": "notes"}
                else:
                    cell = TextBlock()
                    cell.Text = "—"
                    cell.Margin = Thickness(0, 0, 6, 2)
                    Grid.SetRow(cell, row_idx)
                    Grid.SetColumn(cell, col_offset)
                    grid.Children.Add(cell)
                col_offset += 1

            for name in self._param_names:
                param_info = item.get("param_conflicts", {}).get(name)
                if not param_info:
                    cell = TextBlock()
                    cell.Text = "—"
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
        for combo, meta in self._combo_meta.items():
            selection = combo.SelectedItem
            if isinstance(selection, list):
                selection = selection[0] if selection else None
            if not selection:
                continue
            led_id = meta["led_id"]
            entry = self.decisions.setdefault(led_id, {"params": {}})
            kind = meta["kind"]
            if kind == "tags":
                if selection.startswith("Only update"):
                    entry["tags"] = "common"
                elif selection.startswith("Skip"):
                    entry["tags"] = "skip"
                else:
                    entry["tags"] = "union"
            elif kind == "notes":
                if selection.startswith("Only update"):
                    entry["notes"] = "common"
                elif selection.startswith("Skip"):
                    entry["notes"] = "skip"
                else:
                    entry["notes"] = "union"
            else:
                action, value = meta["options"].get(selection, ("skip", None))
                entry["params"][meta["param_name"]] = (action, value)
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()
