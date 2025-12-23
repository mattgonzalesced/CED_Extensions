# -*- coding: utf-8 -*-
"""WPF window for Sync One-Line."""

from pyrevit import forms
from System.Windows.Media import Brushes


class SyncOneLineListItem(object):
    def __init__(self, association, base_text, tree_text, status_symbol, status_brush):
        self.association = association
        self.base_text = base_text
        self.tree_text = tree_text
        self.display_text = base_text
        self.status_symbol = status_symbol
        self.status_brush = status_brush
        self.is_checked = False


class SyncOneLineWindow(forms.WPFWindow):
    def __init__(self, xaml_path, items, detail_symbols, tag_symbols, on_sync=None, on_create=None,
                 on_select_model=None, on_select_detail=None, on_selection_changed=None):
        forms.WPFWindow.__init__(self, xaml_path)

        self._all_items = items
        self._detail_symbols = detail_symbols or []
        self._tag_symbols = tag_symbols or []
        self._on_sync = on_sync
        self._on_create = on_create
        self._on_select_model = on_select_model
        self._on_select_detail = on_select_detail
        self._on_selection_changed = on_selection_changed
        self._sort_mode = "Flat"

        self._build_detail_combo()
        self._build_tag_combo()
        self.SortModeCombo.SelectedIndex = 0
        self._refresh_list(self._all_items)

    def _build_detail_combo(self):
        families = []
        seen = set()
        for symbol in self._detail_symbols:
            family = getattr(symbol, "Family", None)
            if not family:
                continue
            family_name = family.Name
            if family_name not in seen:
                families.append(family)
                seen.add(family_name)

        self.DetailFamilyCombo.ItemsSource = families
        if families:
            self.DetailFamilyCombo.SelectedIndex = 0

    def _build_tag_combo(self):
        families = []
        seen = set()
        for symbol in self._tag_symbols:
            family = getattr(symbol, "Family", None)
            if not family:
                continue
            family_name = family.Name
            if family_name not in seen:
                families.append(family)
                seen.add(family_name)

        families.insert(0, None)
        self.TagFamilyCombo.ItemsSource = families
        self.TagFamilyCombo.SelectedIndex = 0

    def _refresh_list(self, items):
        for item in items:
            if self._sort_mode == "Tree":
                item.display_text = item.tree_text
            else:
                item.display_text = item.base_text
        self.ElementsList.ItemsSource = None
        self.ElementsList.ItemsSource = list(items)

    def _status_allowed(self, status):
        if status == "linked":
            return bool(self.FilterLinked.IsChecked)
        if status == "outdated":
            return bool(self.FilterOutdated.IsChecked)
        if status == "missing":
            return bool(self.FilterMissing.IsChecked)
        return True

    def _filter_items(self, search_text):
        search_text = (search_text or "").lower()
        filtered = []
        for item in self._all_items:
            if not self._status_allowed(item.association.status):
                continue
            text = item.display_text.lower()
            if not search_text or search_text in text:
                filtered.append(item)
        return filtered

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

    def FilterStatus_Changed(self, sender, args):
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def SortModeCombo_SelectionChanged(self, sender, args):
        item = self.SortModeCombo.SelectedItem
        if item:
            self._sort_mode = item.Content
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def DetailFamilyCombo_SelectionChanged(self, sender, args):
        self._update_detail_types()

    def TagFamilyCombo_SelectionChanged(self, sender, args):
        self._update_tag_types()

    def CreateButton_Click(self, sender, args):
        if self._on_create:
            self._on_create()

    def SyncButton_Click(self, sender, args):
        if self._on_sync:
            self._on_sync()

    def CancelButton_Click(self, sender, args):
        self.Close()

    def ItemCheckBox_Click(self, sender, args):
        selected = list(self.ElementsList.SelectedItems)
        if not selected:
            selected = [sender.DataContext]
        new_state = sender.IsChecked
        for item in selected:
            item.is_checked = new_state
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def SelectAllButton_Click(self, sender, args):
        for item in self._all_items:
            item.is_checked = True
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def SelectNoneButton_Click(self, sender, args):
        for item in self._all_items:
            item.is_checked = False
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def refresh_items(self):
        self._refresh_list(self._filter_items(self.SearchBox.Text))

    def ElementsList_SelectionChanged(self, sender, args):
        selected_item = None
        if self.ElementsList.SelectedItems and self.ElementsList.SelectedItems.Count > 0:
            selected_item = self.ElementsList.SelectedItems[0]
        if self._on_selection_changed:
            self._on_selection_changed(selected_item)

    def SelectModelButton_Click(self, sender, args):
        if self._on_select_model:
            self._on_select_model()

    def SelectDetailButton_Click(self, sender, args):
        if self._on_select_detail:
            self._on_select_detail()

    def set_detail_panel(self, label_text, model_id_text, detail_id_text, status_lines):
        self.SelectedLabelText.Text = label_text or "(none)"
        self.ModelIdText.Text = model_id_text or "-"
        self.DetailIdText.Text = detail_id_text or "-"
        self.DetailStatusList.ItemsSource = status_lines or []


def status_symbol(status):
    if status == "linked":
        return "●", Brushes.Green
    if status == "outdated":
        return "●", Brushes.Orange
    return "●", Brushes.Red
