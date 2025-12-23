# -*- coding: utf-8 -*-
"""WPF window for Sync One-Line."""

from pyrevit import forms
from System.Windows.Media import Brushes


class SyncOneLineListItem(object):
    def __init__(self, association, display_text, status_symbol, status_brush):
        self.association = association
        self.display_text = display_text
        self.status_symbol = status_symbol
        self.status_brush = status_brush
        self.is_checked = False


class SyncOneLineWindow(forms.WPFWindow):
    def __init__(self, xaml_path, items, detail_symbols, tag_symbols):
        forms.WPFWindow.__init__(self, xaml_path)

        self._all_items = items
        self._detail_symbols = detail_symbols or []
        self._tag_symbols = tag_symbols or []
        self.requested_action = None

        self._build_detail_combo()
        self._build_tag_combo()
        self._refresh_list(self._all_items)

    def _build_detail_combo(self):
        families = []
        seen = set()
        for symbol in self._detail_symbols:
            family_name = symbol.Family.Name
            if family_name not in seen:
                families.append(symbol.Family)
                seen.add(family_name)

        self.DetailFamilyCombo.ItemsSource = families
        if families:
            self.DetailFamilyCombo.SelectedIndex = 0

    def _build_tag_combo(self):
        families = []
        seen = set()
        for symbol in self._tag_symbols:
            family_name = symbol.Family.Name
            if family_name not in seen:
                families.append(symbol.Family)
                seen.add(family_name)

        families.insert(0, None)
        self.TagFamilyCombo.ItemsSource = families
        self.TagFamilyCombo.SelectedIndex = 0

    def _refresh_list(self, items):
        self.ElementsList.ItemsSource = items

    def _filter_items(self, search_text):
        if not search_text:
            return self._all_items

        search_text = search_text.lower()
        return [item for item in self._all_items if search_text in item.display_text.lower()]

    def _update_detail_types(self):
        family = self.DetailFamilyCombo.SelectedItem
        if not family:
            self.DetailTypeCombo.ItemsSource = []
            return

        types = [sym for sym in self._detail_symbols if sym.Family.Id == family.Id]
        self.DetailTypeCombo.ItemsSource = types
        if types:
            self.DetailTypeCombo.SelectedIndex = 0

    def _update_tag_types(self):
        family = self.TagFamilyCombo.SelectedItem
        if not family:
            self.TagTypeCombo.ItemsSource = []
            return

        types = [sym for sym in self._tag_symbols if sym.Family.Id == family.Id]
        self.TagTypeCombo.ItemsSource = types
        if types:
            self.TagTypeCombo.SelectedIndex = 0

    def get_selected_associations(self):
        return [item.association for item in self._all_items if item.is_checked]

    def get_selected_detail_symbol(self):
        return self.DetailTypeCombo.SelectedItem

    def get_selected_tag_symbol(self):
        return self.TagTypeCombo.SelectedItem

    def SearchBox_TextChanged(self, sender, args):
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def DetailFamilyCombo_SelectionChanged(self, sender, args):
        self._update_detail_types()

    def TagFamilyCombo_SelectionChanged(self, sender, args):
        self._update_tag_types()

    def CreateButton_Click(self, sender, args):
        self.requested_action = "create"
        self.Close()

    def SyncButton_Click(self, sender, args):
        self.requested_action = "sync"
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.requested_action = None
        self.Close()


def status_symbol(status):
    if status == "linked":
        return "●", Brushes.Green
    if status == "outdated":
        return "●", Brushes.Orange
    return "●", Brushes.Red
