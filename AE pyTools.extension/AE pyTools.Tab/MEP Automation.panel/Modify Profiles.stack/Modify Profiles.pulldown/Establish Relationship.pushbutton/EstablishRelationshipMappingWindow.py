# -*- coding: utf-8 -*-
"""Single-grid mapping window for Establish Relationship parameter rules."""

from pyrevit import forms
from System.Windows import Thickness, GridLength, GridUnitType
from System.Windows import FontWeights
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, TextBlock, ComboBox, Button

MODE_STATIC = "static"
MODE_BYPARENT = "BYPARENT"
MODE_BYSIBLING = "BYSIBLING"
LABEL_NONE = "<none>"
LABEL_UNAVAILABLE = "<no source available>"


class RelationshipMappingWindow(forms.WPFWindow):
    def __init__(self, xaml_path, mapping_rows, title_text):
        forms.WPFWindow.__init__(self, xaml_path)
        self._child_rows = mapping_rows or []
        self._line_items = []
        self._line_index = 1
        self._plus_meta = {}
        self._param_meta = {}
        self._mode_meta = {}
        self._target_meta = {}
        self.results = []

        header = self.FindName("HeaderText")
        if header is not None:
            header.Text = (
                "Configure parameter mappings for all selected children. "
                "Use + to add more tracked parameters per child."
            )

        for child_row in self._child_rows:
            line = self._new_line(child_row, preferred_param=child_row.get("selected_param"))
            if line:
                self._line_items.append(line)

        self._render_grid()

        apply_btn = self.FindName("ApplyButton")
        cancel_btn = self.FindName("CancelButton")
        if apply_btn is not None:
            apply_btn.Click += self._on_apply
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    def _selected_text(self, combo):
        if combo is None:
            return ""
        selection = combo.SelectedItem
        if isinstance(selection, list):
            selection = selection[0] if selection else None
        return str(selection or "")

    def _new_line(self, child_row, preferred_param=None):
        options = list(child_row.get("available_params") or [])
        if not options:
            return None
        selected_param = preferred_param if preferred_param in options else None
        if not selected_param:
            selected_param = options[0]
        line = {
            "line_id": "line-{:05d}".format(self._line_index),
            "child_row": child_row,
            "param_name": selected_param,
            "mode": MODE_STATIC,
            "target": LABEL_NONE,
        }
        self._line_index += 1
        return line

    def _current_value_text(self, line):
        child_row = line.get("child_row") or {}
        current_values = child_row.get("current_values") or {}
        value = current_values.get(line.get("param_name"), "")
        if value is None:
            return ""
        text = str(value)
        text = text.replace("\n", " ").replace("\r", " ").strip()
        if len(text) > 64:
            text = text[:61] + "..."
        return text

    def _ordered_parent_options(self, line):
        child_row = line.get("child_row") or {}
        parent_param_names = list(child_row.get("parent_param_names") or [])
        selected_param = line.get("param_name") or ""
        options = []
        seen = set()
        if selected_param:
            options.append(selected_param)
            seen.add(selected_param.lower())
        for pname in parent_param_names:
            key = (pname or "").lower()
            if not key or key in seen:
                continue
            seen.add(key)
            options.append(pname)
        return options

    def _source_options(self, line):
        mode = line.get("mode") or MODE_STATIC
        child_row = line.get("child_row") or {}
        if mode == MODE_STATIC:
            return [LABEL_NONE]
        if mode == MODE_BYPARENT:
            options = self._ordered_parent_options(line)
            return options if options else [LABEL_UNAVAILABLE]
        if mode == MODE_BYSIBLING:
            options = list(child_row.get("sibling_options") or [])
            return options if options else [LABEL_UNAVAILABLE]
        return [LABEL_NONE]

    def _choose_next_param_for_child(self, child_row):
        options = list(child_row.get("available_params") or [])
        if not options:
            return None
        used = set()
        for line in self._line_items:
            if line.get("child_row") is child_row:
                used.add(line.get("param_name"))
        for option in options:
            if option not in used:
                return option
        return options[0]

    def _on_add_line(self, sender, args):
        line = self._plus_meta.get(sender)
        if not line:
            return
        child_row = line.get("child_row")
        if child_row is None:
            return
        preferred = self._choose_next_param_for_child(child_row)
        new_line = self._new_line(child_row, preferred_param=preferred)
        if not new_line:
            return
        new_line["mode"] = line.get("mode") or MODE_STATIC
        new_line["target"] = line.get("target") or LABEL_NONE
        try:
            index = self._line_items.index(line)
        except Exception:
            index = len(self._line_items) - 1
        self._line_items.insert(index + 1, new_line)
        self._render_grid()

    def _on_param_changed(self, sender, args):
        line = self._param_meta.get(sender)
        if not line:
            return
        line["param_name"] = self._selected_text(sender)
        line["target"] = LABEL_NONE
        self._render_grid()

    def _on_mode_changed(self, sender, args):
        line = self._mode_meta.get(sender)
        if not line:
            return
        line["mode"] = self._selected_text(sender) or MODE_STATIC
        line["target"] = LABEL_NONE
        self._render_grid()

    def _on_target_changed(self, sender, args):
        line = self._target_meta.get(sender)
        if not line:
            return
        line["target"] = self._selected_text(sender)

    def _render_grid(self):
        grid = self.FindName("MappingGrid")
        if grid is None:
            raise Exception("MappingGrid not found in XAML.")

        grid.Children.Clear()
        grid.RowDefinitions.Clear()
        grid.ColumnDefinitions.Clear()
        self._plus_meta = {}
        self._param_meta = {}
        self._mode_meta = {}
        self._target_meta = {}

        columns = [
            "",
            "Child",
            "LED ID",
            "Parameter",
            "Current Value",
            "Mode",
            "Source",
        ]
        widths = [36, 280, 220, 280, 260, 140, 360]

        for index, _ in enumerate(columns):
            col_def = ColumnDefinition()
            col_def.Width = GridLength(widths[index], GridUnitType.Pixel)
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
        for line in self._line_items:
            child_row = line.get("child_row") or {}
            row_def = RowDefinition()
            row_def.Height = GridLength(34, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)

            add_btn = Button()
            add_btn.Content = "+"
            add_btn.Width = 24
            add_btn.Height = 22
            add_btn.Margin = Thickness(0, 0, 6, 2)
            add_btn.Click += self._on_add_line
            self._plus_meta[add_btn] = line
            Grid.SetRow(add_btn, row_idx)
            Grid.SetColumn(add_btn, 0)
            grid.Children.Add(add_btn)

            child_cell = TextBlock()
            child_cell.Text = child_row.get("child_label") or ""
            child_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(child_cell, row_idx)
            Grid.SetColumn(child_cell, 1)
            grid.Children.Add(child_cell)

            led_cell = TextBlock()
            led_cell.Text = child_row.get("child_led_id") or ""
            led_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(led_cell, row_idx)
            Grid.SetColumn(led_cell, 2)
            grid.Children.Add(led_cell)

            param_combo = ComboBox()
            param_options = list(child_row.get("available_params") or [])
            for option in param_options:
                param_combo.Items.Add(option)
            if line.get("param_name") in param_options:
                param_combo.SelectedItem = line.get("param_name")
            elif param_options:
                param_combo.SelectedItem = param_options[0]
                line["param_name"] = param_options[0]
            param_combo.SelectionChanged += self._on_param_changed
            self._param_meta[param_combo] = line
            Grid.SetRow(param_combo, row_idx)
            Grid.SetColumn(param_combo, 3)
            grid.Children.Add(param_combo)

            value_cell = TextBlock()
            value_cell.Text = self._current_value_text(line)
            value_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(value_cell, row_idx)
            Grid.SetColumn(value_cell, 4)
            grid.Children.Add(value_cell)

            mode_combo = ComboBox()
            mode_combo.Items.Add(MODE_STATIC)
            mode_combo.Items.Add(MODE_BYPARENT)
            mode_combo.Items.Add(MODE_BYSIBLING)
            mode = line.get("mode") or MODE_STATIC
            if mode not in (MODE_STATIC, MODE_BYPARENT, MODE_BYSIBLING):
                mode = MODE_STATIC
            mode_combo.SelectedItem = mode
            line["mode"] = mode
            mode_combo.SelectionChanged += self._on_mode_changed
            self._mode_meta[mode_combo] = line
            Grid.SetRow(mode_combo, row_idx)
            Grid.SetColumn(mode_combo, 5)
            grid.Children.Add(mode_combo)

            source_combo = ComboBox()
            source_options = self._source_options(line)
            for option in source_options:
                source_combo.Items.Add(option)
            if line.get("target") in source_options:
                source_combo.SelectedItem = line.get("target")
            elif source_options:
                source_combo.SelectedItem = source_options[0]
                line["target"] = source_options[0]
            mode = line.get("mode") or MODE_STATIC
            if mode == MODE_STATIC or line.get("target") == LABEL_UNAVAILABLE:
                source_combo.IsEnabled = False
            source_combo.SelectionChanged += self._on_target_changed
            self._target_meta[source_combo] = line
            Grid.SetRow(source_combo, row_idx)
            Grid.SetColumn(source_combo, 6)
            grid.Children.Add(source_combo)

            row_idx += 1

    def _on_apply(self, sender, args):
        decisions = []
        for line in self._line_items:
            child_row = line.get("child_row") or {}
            child_led_id = child_row.get("child_led_id")
            param_name = line.get("param_name") or ""
            mode = line.get("mode") or MODE_STATIC
            target = line.get("target") or ""
            if not child_led_id or not param_name:
                continue
            if mode != MODE_STATIC and (not target or target in (LABEL_NONE, LABEL_UNAVAILABLE)):
                forms.alert(
                    "Child '{}' parameter '{}' requires a valid source for mode '{}'."
                    .format(child_row.get("child_label") or child_led_id, param_name, mode),
                    title="Establish Relationship",
                )
                return
            decisions.append({
                "child_led_id": child_led_id,
                "param_name": param_name,
                "mode": mode,
                "target": target,
            })
        self.results = decisions
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()
