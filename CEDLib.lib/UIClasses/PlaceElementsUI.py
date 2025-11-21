# -*- coding: utf-8 -*-
"""
PlaceElementsUI
---------------
Simple WPF window to map CAD names to YAML profile labels for placement.
"""

from pyrevit import forms
from System.Windows import Thickness, HorizontalAlignment, GridLength, GridUnitType
from System.Windows.Controls import StackPanel, TextBlock, ComboBox, Button, Grid, ColumnDefinition
from System.Windows.Controls import Orientation

try:
    basestring
except NameError:  # Python 3
    basestring = str


class PlaceElementsWindow(forms.WPFWindow):
    def __init__(self, xaml_path, cad_names, profile_repo, initial_mapping=None):
        forms.WPFWindow.__init__(self, xaml_path)

        self._cad_names = cad_names or []
        self._repo = profile_repo
        self._initial_mapping = initial_mapping or {}
        self._all_labels = self._collect_all_labels()

        # cad_name -> [ComboBox, ...]
        self._combo_controls = {}
        self.result_mapping = {}

        self._build_rows()
        self._wire_events()

    def _collect_all_labels(self):
        seen = set()
        labels = []
        for cad in self._repo.cad_names():
            for lbl in self._repo.labels_for_cad(cad):
                if lbl and lbl not in seen:
                    seen.add(lbl)
                    labels.append(lbl)
        return labels

    # ------------------------------------------------------------------ #
    #  UI build
    # ------------------------------------------------------------------ #
    def _build_rows(self):
        panel = self.FindName("CadBlocksPanel")
        if panel is None:
            raise Exception("CadBlocksPanel not found in XAML.")

        unique_names = sorted(set([n for n in self._cad_names if n]))

        for cad_name in unique_names:
            outer = StackPanel()
            outer.Orientation = Orientation.Vertical
            outer.Margin = Thickness(0, 4, 0, 4)

            header = StackPanel()
            header.Orientation = Orientation.Vertical

            name_row = StackPanel()
            name_row.Orientation = Orientation.Horizontal

            name_block = TextBlock()
            name_block.Text = cad_name
            name_block.Width = 220
            name_block.Margin = Thickness(0, 0, 8, 0)
            name_row.Children.Add(name_block)
            header.Children.Add(name_row)
            outer.Children.Add(header)

            items_grid = Grid()
            items_grid.Margin = Thickness(24, 2, 0, 0)
            left_col = ColumnDefinition()
            left_col.Width = GridLength(1, GridUnitType.Star)
            items_grid.ColumnDefinitions.Add(left_col)
            right_col = ColumnDefinition()
            right_col.Width = GridLength.Auto
            items_grid.ColumnDefinitions.Add(right_col)

            left_panel = StackPanel()
            left_panel.Orientation = Orientation.Vertical
            items_grid.Children.Add(left_panel)

            right_panel = StackPanel()
            right_panel.Orientation = Orientation.Vertical
            right_panel.HorizontalAlignment = HorizontalAlignment.Center
            Grid.SetColumn(right_panel, 1)
            items_grid.Children.Add(right_panel)

            add_btn = Button()
            add_btn.Content = "+"
            add_btn.Width = 24
            add_btn.Margin = Thickness(0, 0, 0, 4)
            add_btn.Tag = (cad_name, left_panel, right_panel)
            add_btn.Click += self._on_add_clicked
            right_panel.Children.Add(add_btn)

            labels_for_cad = self._repo.labels_for_cad(cad_name)
            initial_labels = self._initial_mapping.get(cad_name) or labels_for_cad
            if isinstance(initial_labels, basestring if "basestring" in globals() else str):
                initial_labels = [initial_labels]
            if initial_labels:
                for lbl in initial_labels:
                    self._add_label_row(cad_name, left_panel, right_panel, selected_label=lbl, allow_fallback=True)
            else:
                self._add_label_row(cad_name, left_panel, right_panel, allow_fallback=False)

            # attach grid containing combos/buttons
            outer.Children.Add(items_grid)
            panel.Children.Add(outer)

    def _add_label_row(self, cad_name, left_panel, right_panel, selected_label=None, allow_fallback=True):
        labels = self._repo.labels_for_cad(cad_name)
        if not labels and allow_fallback:
            labels = self._all_labels

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 2, 0, 2)

        combo = ComboBox()
        combo.Width = 500
        combo.Margin = Thickness(0, 0, 8, 0)

        if labels:
            for lbl in labels:
                combo.Items.Add(lbl)
            if selected_label and selected_label in labels:
                combo.SelectedItem = selected_label
            else:
                combo.SelectedIndex = 0
        else:
            combo.Items.Add("<no YAML profile defined>")
            combo.SelectedIndex = 0
            combo.IsEnabled = False

        row.Children.Add(combo)

        remove_btn = Button()
        remove_btn.Content = "X"
        remove_btn.Width = 24
        remove_btn.Margin = Thickness(0, 2, 0, 2)
        remove_btn.Tag = (cad_name, row, combo, "label")
        remove_btn.Click += self._on_remove_clicked

        left_panel.Children.Add(row)
        right_panel.Children.Add(remove_btn)
        self._combo_controls.setdefault(cad_name, []).append(combo)

    # ------------------------------------------------------------------ #
    #  Events
    # ------------------------------------------------------------------ #
    def _wire_events(self):
        ok_button = self.FindName("OkButton")
        cancel_button = self.FindName("CancelButton")
        if ok_button is not None:
            ok_button.Click += self._on_ok_clicked
        if cancel_button is not None:
            cancel_button.Click += self._on_cancel_clicked

    def _on_add_clicked(self, sender, args):
        tag = sender.Tag
        if not tag or len(tag) != 3:
            return
        cad_name, left_panel, right_panel = tag
        self._add_label_row(cad_name, left_panel, right_panel, allow_fallback=True)

    def _on_remove_clicked(self, sender, args):
        tag = sender.Tag
        if not tag or len(tag) != 4:
            return
        cad_name, row_panel, combo_obj, _kind = tag

        parent_panel = row_panel.Parent
        if parent_panel is not None:
            parent_panel.Children.Remove(row_panel)
        btn_parent = sender.Parent
        if btn_parent is not None:
            btn_parent.Children.Remove(sender)

        combos = self._combo_controls.get(cad_name, [])
        if combo_obj in combos:
            combos.remove(combo_obj)

    def _on_ok_clicked(self, sender, args):
        mapping = {}
        for cad_name, combos in self._combo_controls.items():
            labels = []
            for combo in combos:
                if combo is None or not combo.IsEnabled or combo.SelectedItem is None:
                    continue
                lbl = combo.SelectedItem.ToString()
                if lbl == "<no YAML profile defined>":
                    continue
                labels.append(lbl)
            if labels:
                mapping[cad_name] = labels
        self.result_mapping = mapping
        self.DialogResult = True
        self.Close()

    def _on_cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


def show_place_elements_window(xaml_path, cad_names, repo, initial_mapping=None):
    win = PlaceElementsWindow(xaml_path, cad_names, repo, initial_mapping)
    win.show_dialog()
    return win.DialogResult, win.result_mapping
