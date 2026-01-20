# -*- coding: utf-8 -*-
"""
WPF window for resolving parent parameter conflicts after sync.
"""

from pyrevit import forms
from System.Windows import Thickness, GridLength, GridUnitType, FontWeights
from System.Windows.Controls import (
    Grid,
    RowDefinition,
    ColumnDefinition,
    TextBlock,
    ComboBox,
    StackPanel,
    Orientation,
)

ACTION_UPDATE_CHILD = "update_child"
ACTION_UPDATE_PARENT = "update_parent"
ACTION_SKIP = "skip"

LABEL_UPDATE_CHILD = "Update child from parent"
LABEL_UPDATE_PARENT = "Update parent from child"
LABEL_SKIP = "Skip"
LABEL_NO_OVERRIDE = "<no override>"


class ParentParamConflictsWindow(forms.WPFWindow):
    def __init__(self, xaml_path, conflicts, param_keys):
        forms.WPFWindow.__init__(self, xaml_path)
        self._conflicts = conflicts or []
        self._param_keys = param_keys or []
        self._combo_meta = {}
        self._combos_by_param = {}
        self._default_meta = {}
        self.decisions = {}

        header = self.FindName("HeaderText")
        if header is not None:
            header.Text = "Resolve parent parameter discrepancies detected after sync."

        self._build_param_defaults()
        self._build_grid()

        apply_btn = self.FindName("ApplyButton")
        cancel_btn = self.FindName("CancelButton")
        if apply_btn is not None:
            apply_btn.Click += self._on_apply
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    def _build_param_defaults(self):
        panel = self.FindName("ParamDefaultsPanel")
        if panel is None:
            return
        if not self._param_keys:
            return

        header = TextBlock()
        header.Text = "Apply to all (per parameter):"
        header.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(header)

        for param_key in self._param_keys:
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 0, 0, 4)

            label = TextBlock()
            label.Text = param_key
            label.Width = 360
            label.Margin = Thickness(0, 0, 8, 0)
            row.Children.Add(label)

            combo = ComboBox()
            combo.Width = 220
            combo.Items.Add(LABEL_NO_OVERRIDE)
            combo.Items.Add(LABEL_UPDATE_CHILD)
            combo.Items.Add(LABEL_UPDATE_PARENT)
            combo.Items.Add(LABEL_SKIP)
            combo.SelectedIndex = 0
            combo.SelectionChanged += self._on_default_changed

            self._default_meta[combo] = param_key
            row.Children.Add(combo)
            panel.Children.Add(row)

    def _build_grid(self):
        grid = self.FindName("ConflictGrid")
        if grid is None:
            raise Exception("ConflictGrid not found in XAML.")

        columns = [
            "Parent",
            "Child",
            "Parameter",
            "Parent Value",
            "Child Value",
            "Action",
        ]

        widths = [260, 260, 220, 140, 140, 180]
        for idx, _ in enumerate(columns):
            col_def = ColumnDefinition()
            width = widths[idx] if idx < len(widths) else 180
            col_def.Width = GridLength(width, GridUnitType.Pixel)
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
        for conflict in self._conflicts:
            row_def = RowDefinition()
            row_def.Height = GridLength(30, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)

            param_label = conflict.get("param_key") or ""
            cells = [
                conflict.get("parent_label"),
                conflict.get("child_label"),
                param_label,
                conflict.get("parent_display"),
                conflict.get("child_display"),
            ]

            for col_idx, value in enumerate(cells):
                cell = TextBlock()
                cell.Text = value or ""
                cell.Margin = Thickness(0, 0, 6, 2)
                Grid.SetRow(cell, row_idx)
                Grid.SetColumn(cell, col_idx)
                grid.Children.Add(cell)

            combo = ComboBox()
            options = []
            label_map = {}
            if conflict.get("allow_update_child"):
                options.append(LABEL_UPDATE_CHILD)
                label_map[LABEL_UPDATE_CHILD] = ACTION_UPDATE_CHILD
            if conflict.get("allow_update_parent"):
                options.append(LABEL_UPDATE_PARENT)
                label_map[LABEL_UPDATE_PARENT] = ACTION_UPDATE_PARENT
            options.append(LABEL_SKIP)
            label_map[LABEL_SKIP] = ACTION_SKIP

            for opt in options:
                combo.Items.Add(opt)
            combo.SelectedItem = LABEL_SKIP
            if len(options) == 1:
                combo.IsEnabled = False
            Grid.SetRow(combo, row_idx)
            Grid.SetColumn(combo, 5)
            grid.Children.Add(combo)

            param_key = conflict.get("param_key") or ""
            self._combo_meta[combo] = {
                "conflict_id": conflict.get("id"),
                "param_key": param_key,
                "label_map": label_map,
            }
            if param_key:
                self._combos_by_param.setdefault(param_key, []).append(combo)

            row_idx += 1

    def _on_default_changed(self, sender, args):
        param_key = self._default_meta.get(sender)
        if not param_key:
            return
        selection = sender.SelectedItem
        if selection in (None, LABEL_NO_OVERRIDE):
            return
        combos = self._combos_by_param.get(param_key) or []
        for combo in combos:
            if selection in list(combo.Items):
                combo.SelectedItem = selection

    def _on_apply(self, sender, args):
        for combo, meta in self._combo_meta.items():
            selection = combo.SelectedItem
            if isinstance(selection, list):
                selection = selection[0] if selection else None
            if not selection:
                continue
            action = meta["label_map"].get(selection, ACTION_SKIP)
            conflict_id = meta.get("conflict_id")
            if conflict_id:
                self.decisions[conflict_id] = action
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()
