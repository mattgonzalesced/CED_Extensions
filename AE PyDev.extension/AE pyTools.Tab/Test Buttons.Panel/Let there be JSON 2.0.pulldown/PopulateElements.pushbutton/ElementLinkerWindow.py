# -*- coding: utf-8 -*-
"""
ElementLinkerWindow
-------------------
WPF UI class for the Element Linker tool.

Supports multiple selections per CAD block:
- Each CAD block gets a header row with the CAD name and a [+] button
- Below that, there is a vertical list of ComboBox rows
- Each ComboBox row has an [X] button to remove that selection
- On OK, we return { cad_name: [label1, label2, ...] }

Pairs with:
- Element_Linker.CAD_BLOCK_PROFILES (CadBlockProfile / TypeConfig)
- ElementLinkerEngine.ElementPlacementEngine
"""

from pyrevit import forms
from System.Windows import Thickness, HorizontalAlignment, GridLength, GridUnitType
from System.Windows.Controls import StackPanel, TextBlock, ComboBox, Button, Grid, ColumnDefinition
from System.Windows.Controls import Orientation


class ElementLinkerWindow(forms.WPFWindow):
    def __init__(self, xaml_path, cad_names, cad_block_profiles, initial_mapping=None):
        """
        cad_names: list of CAD block names found in the CSV
        cad_block_profiles: Element_Linker.CAD_BLOCK_PROFILES
                            { cad_name: CadBlockProfile }
        initial_mapping: optional { cad_name: [labels...] } to pre-populate rows
        """
        forms.WPFWindow.__init__(self, xaml_path)

        self._cad_names = cad_names
        self._profiles = cad_block_profiles
        self._initial_mapping = initial_mapping or {}
        self._all_labels = self._collect_all_labels()

        # cad_name -> [ComboBox, ComboBox, ...] (labels only)
        self._combo_controls = {}

        # Final result: { cad_name: [label1, label2, ...] }
        self.result_mapping = {}

        # IMPORTANT: do NOT set self.DialogResult here.
        # It can only be set after ShowDialog() has been called.

        self._build_rows()
        self._wire_events()

    def _collect_all_labels(self):
        """Return a deduped master list of all type labels across profiles."""
        seen = set()
        labels = []
        for profile in (self._profiles or {}).values():
            try:
                for lbl in profile.get_type_labels():
                    if not lbl or lbl in seen:
                        continue
                    seen.add(lbl)
                    labels.append(lbl)
            except Exception:
                continue
        return labels

    # ------------------------------------------------------------------ #
    #  Build UI rows
    # ------------------------------------------------------------------ #
    def _build_rows(self):
        panel = self.FindName('CadBlocksPanel')
        if panel is None:
            raise Exception("CadBlocksPanel not found in XAML.")

        unique_names = sorted(set([n for n in self._cad_names if n]))

        for cad_name in unique_names:
            # Outer container for this CAD block
            outer = StackPanel()
            outer.Orientation = Orientation.Vertical
            outer.Margin = Thickness(0, 4, 0, 4)

            # Header: CAD name + [+] button above (stacked vertically)
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

            # Grid with two columns: left = combos, right = buttons (add + removes)
            items_grid = Grid()
            items_grid.Margin = Thickness(24, 2, 0, 0)
            left_col = ColumnDefinition()
            left_col.Width = GridLength(1, GridUnitType.Star)
            items_grid.ColumnDefinitions.Add(left_col)  # left
            right_col = ColumnDefinition()
            right_col.Width = GridLength.Auto
            items_grid.ColumnDefinitions.Add(right_col)  # right

            left_panel = StackPanel()
            left_panel.Orientation = Orientation.Vertical
            items_grid.Children.Add(left_panel)

            right_panel = StackPanel()
            right_panel.Orientation = Orientation.Vertical
            right_panel.HorizontalAlignment = HorizontalAlignment.Center
            Grid.SetColumn(right_panel, 1)
            items_grid.Children.Add(right_panel)

            # Add button sits at top of the right column (above all X buttons)
            add_btn = Button()
            add_btn.Content = "+"
            add_btn.Width = 24
            add_btn.Margin = Thickness(0, 0, 0, 4)
            add_btn.Click += self._on_add_clicked

            # Now that panels exist, tag the add button with them
            add_btn.Tag = (cad_name, left_panel, right_panel)
            right_panel.Children.Add(add_btn)

            outer.Children.Add(items_grid)

            panel.Children.Add(outer)

            # Initialize combo control list for this cad_name
            self._combo_controls[cad_name] = []

            # Add initial rows: either provided mapping labels or a single default
            initial_labels = self._initial_mapping.get(cad_name)
            if initial_labels:
                for lbl in initial_labels:
                    self._add_label_row(cad_name, left_panel, right_panel, selected_label=lbl, allow_fallback=True)
            else:
                # no profile mapping for this CAD block; start with a placeholder row (no fallback labels)
                self._add_label_row(cad_name, left_panel, right_panel, allow_fallback=False)

    def _add_label_row(self, cad_name, left_panel, right_panel, selected_label=None, allow_fallback=True):
        """Add a single label ComboBox row using profile labels.
        If allow_fallback is True and no profile labels exist, we fall back to all labels.
        """
        profile = self._profiles.get(cad_name)
        labels = []
        if profile:
            labels.extend(profile.get_type_labels())
        # If no labels from this profile, optionally fall back to all known labels
        if not labels and allow_fallback:
            labels.extend(self._all_labels)

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
            combo.Items.Add("<no Element_Linker profile defined>")
            combo.SelectedIndex = 0
            combo.IsEnabled = False

        row.Children.Add(combo)

        # remove button lives in the right panel column
        remove_btn = Button()
        remove_btn.Content = "X"
        remove_btn.Width = 24
        remove_btn.Margin = Thickness(0, 2, 0, 2)
        # Store info we need to remove: which CAD, which row, which combo
        remove_btn.Tag = (cad_name, row, combo, "label")
        remove_btn.Click += self._on_remove_clicked

        left_panel.Children.Add(row)
        right_panel.Children.Add(remove_btn)

        # Track this combo
        self._combo_controls.setdefault(cad_name, []).append(combo)

    # ------------------------------------------------------------------ #
    #  Button event handlers
    # ------------------------------------------------------------------ #
    def _wire_events(self):
        ok_button = self.FindName('OkButton')
        cancel_button = self.FindName('CancelButton')

        if ok_button is not None:
            ok_button.Click += self._on_ok_clicked
        if cancel_button is not None:
            cancel_button.Click += self._on_cancel_clicked

    def _on_add_clicked(self, sender, args):
        """
        [ + ] button: add another ComboBox row for this CAD name.
        """
        tag = sender.Tag
        if not tag or len(tag) != 3:
            return
        cad_name, left_panel, right_panel = tag
        self._add_label_row(cad_name, left_panel, right_panel, allow_fallback=True)

    def _on_remove_clicked(self, sender, args):
        """
        [ X ] button: remove this ComboBox row.
        """
        tag = sender.Tag
        if not tag or len(tag) != 4:
            return

        cad_name, row_panel, combo_obj, kind = tag

        # Remove row from visual tree
        parent_panel = row_panel.Parent
        if parent_panel is not None:
            parent_panel.Children.Remove(row_panel)

        # Remove the X button from its panel
        btn_parent = sender.Parent
        if btn_parent is not None:
            btn_parent.Children.Remove(sender)

        # Remove combo tuple from tracking list
        combos = self._combo_controls.get(cad_name, [])
        if combo_obj in combos:
            combos.remove(combo_obj)

    def _on_ok_clicked(self, sender, args):
        """
        Gather selections into result_mapping and close with DialogResult=True.
        """
        mapping = {}

        for cad_name, combos in self._combo_controls.items():
            labels = []
            for combo in combos:
                if combo is None or not combo.IsEnabled or combo.SelectedItem is None:
                    continue
                lbl = combo.SelectedItem.ToString()
                if lbl == "<no Element_Linker profile defined>":
                    continue
                labels.append(lbl)

            if labels:
                mapping[cad_name] = labels

        self.result_mapping = mapping
        # Now it's legal to set DialogResult, because window is shown as dialog
        self.DialogResult = True
        self.Close()

    def _on_cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()
