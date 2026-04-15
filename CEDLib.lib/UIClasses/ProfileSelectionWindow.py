# -*- coding: utf-8 -*-
"""
ProfileSelectionWindow
----------------------
Simple WPF window that shows profile group choices with search and
select-all/none actions.
"""

from pyrevit import forms
from System.Windows import Thickness, Visibility
from System.Windows.Controls import CheckBox


class ProfileSelectionWindow(forms.WPFWindow):
    def __init__(self, xaml_path, options):
        forms.WPFWindow.__init__(self, xaml_path)
        self._options = options or []
        self._checkboxes = []
        self.selected_options = []
        self._build_options()

    def _build_options(self):
        panel = self.FindName("OptionsPanel")
        if panel is None:
            raise Exception("OptionsPanel not found in XAML.")

        for option in self._options:
            label = option.get("label") or ""
            checkbox = CheckBox()
            checkbox.Content = label
            checkbox.Margin = Thickness(0, 2, 0, 2)
            checkbox.IsChecked = False
            checkbox.Tag = option
            panel.Children.Add(checkbox)
            self._checkboxes.append({
                "checkbox": checkbox,
                "label_norm": label.strip().lower(),
            })

    def _apply_filter(self, search_text):
        normalized = (search_text or "").strip().lower()
        for entry in self._checkboxes:
            checkbox = entry["checkbox"]
            label_norm = entry["label_norm"]
            visible = True if not normalized else (normalized in label_norm)
            checkbox.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def _iter_visible(self):
        for entry in self._checkboxes:
            checkbox = entry["checkbox"]
            if checkbox.Visibility == Visibility.Visible:
                yield checkbox

    def SearchBox_TextChanged(self, sender, args):
        text = ""
        if sender is not None:
            text = getattr(sender, "Text", "") or ""
        self._apply_filter(text)

    def SelectAllButton_Click(self, sender, args):
        for checkbox in self._iter_visible():
            checkbox.IsChecked = True

    def SelectNoneButton_Click(self, sender, args):
        for checkbox in self._iter_visible():
            checkbox.IsChecked = False

    def OkButton_Click(self, sender, args):
        selected = []
        for entry in self._checkboxes:
            checkbox = entry["checkbox"]
            if checkbox.IsChecked:
                option = checkbox.Tag
                if option:
                    selected.append(option)
        self.selected_options = selected
        self.DialogResult = True
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


def show_profile_selection_window(xaml_path, options):
    win = ProfileSelectionWindow(xaml_path, options)
    win.show_dialog()
    return win.DialogResult, win.selected_options
