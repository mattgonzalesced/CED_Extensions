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
from System.Windows import Thickness
from System.Windows.Controls import StackPanel, TextBlock, ComboBox, Button
from System.Windows.Controls import Orientation


class ElementLinkerWindow(forms.WPFWindow):
    def __init__(self, xaml_path, cad_names, cad_block_profiles):
        """
        cad_names: list of CAD block names found in the CSV
        cad_block_profiles: Element_Linker.CAD_BLOCK_PROFILES
                            { cad_name: CadBlockProfile }
        """
        forms.WPFWindow.__init__(self, xaml_path)

        self._cad_names = cad_names
        self._profiles = cad_block_profiles

        # cad_name -> [ComboBox, ComboBox, ...]
        self._combo_controls = {}

        # Final result: { cad_name: [label1, label2, ...] }
        self.result_mapping = {}

        # IMPORTANT: do NOT set self.DialogResult here.
        # It can only be set after ShowDialog() has been called.

        self._build_rows()
        self._wire_events()

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

            # Header: CAD name + [+] button
            header = StackPanel()
            header.Orientation = Orientation.Horizontal

            name_block = TextBlock()
            name_block.Text = cad_name
            name_block.Width = 220
            name_block.Margin = Thickness(0, 0, 8, 0)
            header.Children.Add(name_block)

            add_btn = Button()
            add_btn.Content = "+"
            add_btn.Width = 24
            add_btn.Margin = Thickness(0, 0, 0, 0)
            add_btn.Tag = cad_name
            add_btn.Click += self._on_add_clicked
            header.Children.Add(add_btn)

            outer.Children.Add(header)

            # Panel where we put all ComboBox rows for this CAD block
            items_panel = StackPanel()
            items_panel.Orientation = Orientation.Vertical
            items_panel.Margin = Thickness(24, 2, 0, 0)
            outer.Children.Add(items_panel)

            panel.Children.Add(outer)

            # Initialize combo control list for this cad_name
            self._combo_controls[cad_name] = []

            # Add an initial ComboBox row (so user sees something)
            self._add_combo_row(cad_name, items_panel)

    def _add_combo_row(self, cad_name, container_panel):
        """
        Add a single ComboBox row for a given CAD name into container_panel.
        """
        profile = self._profiles.get(cad_name)

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 2, 0, 2)

        combo = ComboBox()
        combo.Width = 500
        combo.Margin = Thickness(0, 0, 8, 0)

        if profile:
            labels = profile.get_type_labels()
            for lbl in labels:
                combo.Items.Add(lbl)
            if labels:
                combo.SelectedIndex = 0
        else:
            combo.Items.Add("<no Element_Linker profile defined>")
            combo.SelectedIndex = 0
            combo.IsEnabled = False

        row.Children.Add(combo)

        remove_btn = Button()
        remove_btn.Content = "X"
        remove_btn.Width = 24
        remove_btn.Margin = Thickness(0, 0, 0, 0)
        # Store info we need to remove: which CAD, which row, which combo
        remove_btn.Tag = (cad_name, row, combo)
        remove_btn.Click += self._on_remove_clicked

        row.Children.Add(remove_btn)
        container_panel.Children.Add(row)

        # Track this combo
        self._combo_controls[cad_name].append(combo)

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
        cad_name = sender.Tag
        # Visual tree:
        #   header (this button's Parent) -> outer (Parent of header)
        outer = sender.Parent.Parent
        children = list(outer.Children)
        if len(children) < 2:
            return
        items_panel = children[1]
        self._add_combo_row(cad_name, items_panel)

    def _on_remove_clicked(self, sender, args):
        """
        [ X ] button: remove this ComboBox row.
        """
        tag = sender.Tag
        if not tag or len(tag) != 3:
            return

        cad_name, row_panel, combo = tag

        # Remove row from visual tree
        parent_panel = row_panel.Parent
        if parent_panel is not None:
            parent_panel.Children.Remove(row_panel)

        # Remove combo from tracking list
        combos = self._combo_controls.get(cad_name, [])
        if combo in combos:
            combos.remove(combo)

    def _on_ok_clicked(self, sender, args):
        """
        Gather selections into result_mapping and close with DialogResult=True.
        """
        mapping = {}

        for cad_name, combos in self._combo_controls.items():
            labels = []
            for combo in combos:
                if combo.IsEnabled and combo.SelectedItem is not None:
                    label = combo.SelectedItem.ToString()
                    # Skip placeholder if no profile
                    if label == "<no Element_Linker profile defined>":
                        continue
                    labels.append(label)

            if labels:
                mapping[cad_name] = labels

        self.result_mapping = mapping
        # Now it's legal to set DialogResult, because window is shown as dialog
        self.DialogResult = True
        self.Close()

    def _on_cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()
