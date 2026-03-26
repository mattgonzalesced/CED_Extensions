# -*- coding: utf-8 -*-
"""
OptimizeUI
----------
WPF window for the Optimize Element Placement command.

Left panel  - Available family:types for the selected category.
              Single-click any item to add it to the Selected list.
Right panel - Selected family:types (can come from any category).
              "Remove Selected" button removes the highlighted item.
Bottom      - One rule row per selected type: Mode + Corner (Corner mode only).
"""

from pyrevit import forms
from System.Windows import (
    FontWeights,
    GridLength,
    GridUnitType,
    Thickness,
    VerticalAlignment,
    Visibility,
)
from System.Windows.Controls import (
    ColumnDefinition,
    ComboBox,
    Grid,
    RowDefinition,
    TextBlock,
)

OPTIMIZATION_MODES = ["Wall", "Ceiling", "Floor", "Door", "Corner"]

CORNER_OPTIONS = ["Lower Left", "Lower Right", "Upper Left", "Upper Right"]

DEFAULT_CORNER = "Lower Left"


class OptimizeWindow(forms.WPFWindow):
    """
    Args:
        xaml_path (str): Absolute path to OptimizeUI.xaml.
        category_map (dict): {category_name: [family_type_label, ...]}
            Only categories/types that have at least one element with a
            non-blank Element_Linker parameter should be included.
    """

    def __init__(self, xaml_path, category_map):
        forms.WPFWindow.__init__(self, xaml_path)
        self._category_map = category_map or {}
        self._mode_combos = {}    # ft_label -> ComboBox
        self._corner_combos = {}  # ft_label -> ComboBox
        self._adding = False      # re-entrancy guard for AvailableList events
        self.confirmed = False
        self.rules = {}           # {ft_label: {"mode": str, "corner": str}}
        self._init_controls()

    # ------------------------------------------------------------------
    # Setup
    # ------------------------------------------------------------------

    def _init_controls(self):
        cat_combo = self.FindName("CategoryCombo")
        available_list = self.FindName("AvailableList")
        selected_list = self.FindName("SelectedList")
        remove_btn = self.FindName("RemoveButton")
        run_btn = self.FindName("RunButton")
        cancel_btn = self.FindName("CancelButton")

        if cat_combo is not None:
            cat_combo.Items.Clear()
            for name in sorted(self._category_map.keys()):
                cat_combo.Items.Add(name)
            cat_combo.SelectionChanged += self._on_category_changed
            if cat_combo.Items.Count > 0:
                cat_combo.SelectedIndex = 0

        if available_list is not None:
            available_list.SelectionChanged += self._on_available_selection_changed

        if selected_list is not None:
            selected_list.SelectionChanged += self._on_selected_list_selection_changed

        if remove_btn is not None:
            remove_btn.Click += self._on_remove_clicked

        if run_btn is not None:
            run_btn.Click += self._on_run
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _on_category_changed(self, sender, args):
        cat_combo = self.FindName("CategoryCombo")
        available_list = self.FindName("AvailableList")
        if cat_combo is None or available_list is None:
            return
        selected_cat = cat_combo.SelectedItem
        if not selected_cat:
            return
        family_types = self._category_map.get(str(selected_cat), [])
        available_list.Items.Clear()
        for ft_label in sorted(family_types):
            available_list.Items.Add(ft_label)

    def _on_available_selection_changed(self, sender, args):
        """Single-click an available type to add it to the Selected list."""
        if self._adding:
            return
        available_list = self.FindName("AvailableList")
        if available_list is None:
            return
        item = available_list.SelectedItem
        if not item:
            return

        ft_label = str(item)

        # Add to Selected list if not already present
        selected_list = self.FindName("SelectedList")
        if selected_list is not None:
            existing = [str(x) for x in selected_list.Items]
            if ft_label not in existing:
                selected_list.Items.Add(ft_label)
                self._rebuild_rules_grid()
                self._update_run_button()

        # Clear the available list selection without re-triggering this handler
        self._adding = True
        try:
            available_list.SelectedItem = None
        finally:
            self._adding = False

    def _on_selected_list_selection_changed(self, sender, args):
        """Enable/disable Remove button based on whether anything is highlighted."""
        selected_list = self.FindName("SelectedList")
        remove_btn = self.FindName("RemoveButton")
        if selected_list is None or remove_btn is None:
            return
        remove_btn.IsEnabled = selected_list.SelectedItem is not None

    def _on_remove_clicked(self, sender, args):
        """Remove the highlighted item from the Selected list."""
        selected_list = self.FindName("SelectedList")
        if selected_list is None:
            return
        item = selected_list.SelectedItem
        if item is None:
            return
        selected_list.Items.Remove(item)
        self._rebuild_rules_grid()
        self._update_run_button()

    def _on_run(self, sender, args):
        self.rules = {}
        for ft_label, mode_combo in self._mode_combos.items():
            mode = str(mode_combo.SelectedItem) if mode_combo.SelectedItem else "Wall"
            corner = DEFAULT_CORNER
            if mode == "Corner":
                corner_combo = self._corner_combos.get(ft_label)
                if corner_combo and corner_combo.SelectedItem:
                    corner = str(corner_combo.SelectedItem)
            self.rules[ft_label] = {"mode": mode, "corner": corner}
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()

    # ------------------------------------------------------------------
    # Rules grid helpers
    # ------------------------------------------------------------------

    def _get_selected_types(self):
        """Return all labels currently in the Selected list."""
        selected_list = self.FindName("SelectedList")
        if selected_list is None:
            return []
        return [str(item) for item in selected_list.Items]

    def _clear_rules_grid(self):
        grid = self.FindName("RulesGrid")
        if grid is None:
            return
        grid.Children.Clear()
        grid.RowDefinitions.Clear()
        grid.ColumnDefinitions.Clear()
        self._mode_combos.clear()
        self._corner_combos.clear()

    def _rebuild_rules_grid(self):
        self._clear_rules_grid()
        selected_types = self._get_selected_types()
        if not selected_types:
            return
        grid = self.FindName("RulesGrid")
        if grid is None:
            return

        # Three columns: label | mode | corner
        for width in [280, 175, 160]:
            col = ColumnDefinition()
            col.Width = GridLength(width, GridUnitType.Pixel)
            grid.ColumnDefinitions.Add(col)

        # Header row
        hdr_row = RowDefinition()
        hdr_row.Height = GridLength(26, GridUnitType.Pixel)
        grid.RowDefinitions.Add(hdr_row)
        for col_idx, title in enumerate(["Family : Type", "Mode", "Corner"]):
            cell = TextBlock()
            cell.Text = title
            cell.FontWeight = FontWeights.Bold
            cell.Margin = Thickness(2, 2, 6, 2)
            Grid.SetRow(cell, 0)
            Grid.SetColumn(cell, col_idx)
            grid.Children.Add(cell)

        # One row per selected family:type
        for row_idx, ft_label in enumerate(selected_types, start=1):
            row_def = RowDefinition()
            row_def.Height = GridLength(30, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)

            # Label
            lbl = TextBlock()
            lbl.Text = ft_label
            lbl.Margin = Thickness(2, 4, 6, 2)
            lbl.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(lbl, row_idx)
            Grid.SetColumn(lbl, 0)
            grid.Children.Add(lbl)

            # Mode ComboBox
            mode_combo = ComboBox()
            mode_combo.Margin = Thickness(0, 2, 6, 2)
            for mode in OPTIMIZATION_MODES:
                mode_combo.Items.Add(mode)
            mode_combo.SelectedIndex = 0
            mode_combo.Tag = ft_label
            mode_combo.SelectionChanged += self._on_mode_changed
            self._mode_combos[ft_label] = mode_combo
            Grid.SetRow(mode_combo, row_idx)
            Grid.SetColumn(mode_combo, 1)
            grid.Children.Add(mode_combo)

            # Corner ComboBox (hidden unless mode == Corner)
            corner_combo = ComboBox()
            corner_combo.Margin = Thickness(0, 2, 0, 2)
            for corner in CORNER_OPTIONS:
                corner_combo.Items.Add(corner)
            corner_combo.SelectedIndex = 0  # Lower Left default
            corner_combo.Visibility = Visibility.Collapsed
            self._corner_combos[ft_label] = corner_combo
            Grid.SetRow(corner_combo, row_idx)
            Grid.SetColumn(corner_combo, 2)
            grid.Children.Add(corner_combo)

    def _on_mode_changed(self, sender, args):
        try:
            ft_label = str(sender.Tag)
        except Exception:
            return
        selected_mode = sender.SelectedItem
        corner_combo = self._corner_combos.get(ft_label)
        if corner_combo is not None:
            if str(selected_mode) == "Corner":
                corner_combo.Visibility = Visibility.Visible
            else:
                corner_combo.Visibility = Visibility.Collapsed

    def _update_run_button(self):
        run_btn = self.FindName("RunButton")
        if run_btn is None:
            return
        run_btn.IsEnabled = len(self._get_selected_types()) > 0
