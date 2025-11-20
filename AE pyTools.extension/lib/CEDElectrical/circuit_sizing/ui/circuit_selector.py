import os
from System.Windows.Controls import ListBoxItem
from pyrevit import forms


class CircuitSelectorWindow(forms.WPFWindow):
    def __init__(self, provider):
        xaml_path = os.path.join(os.path.dirname(__file__), 'circuit_selector.xaml')
        super(CircuitSelectorWindow, self).__init__(xaml_path)
        self.provider = provider
        self.selected_ids = set()
        self.selected_circuits = []

        self.PreserveSelection.IsChecked = True
        self.PanelFilter.Items.Add('All Panels')
        for panel_name in provider.panels:
            self.PanelFilter.Items.Add(panel_name)
        self.PanelFilter.Items.Add('<No Panel>')
        self.PanelFilter.SelectedIndex = 0

        self.SearchBox.TextChanged += self._refresh_list
        self.PanelFilter.SelectionChanged += self._on_panel_changed
        self.PreserveSelection.Checked += self._on_preserve_toggle
        self.PreserveSelection.Unchecked += self._on_preserve_toggle
        self.CircuitList.SelectionChanged += self._on_selection_changed
        self.SelectAllButton.Click += self._select_all
        self.SelectNoneButton.Click += self._select_none
        self.CalculateButton.Click += self._calculate
        self.CancelButton.Click += self._close

        self._refresh_list(None, None)

    def _on_panel_changed(self, sender, args):
        self._refresh_list(sender, args)

    def _on_preserve_toggle(self, sender, args):
        if not self.PreserveSelection.IsChecked:
            self.selected_ids = set()
        self._refresh_list(sender, args)

    def _refresh_list(self, sender, args):
        search_text = self.SearchBox.Text
        panel_filter = self.PanelFilter.SelectedItem
        items = self.provider.filter(search_text, panel_filter)
        self.CircuitList.Items.Clear()
        for item in items:
            list_item = ListBoxItem()
            list_item.Content = item['label']
            list_item.Tag = item['circuit']
            self.CircuitList.Items.Add(list_item)
            if self.PreserveSelection.IsChecked and item['id'] in self.selected_ids:
                list_item.IsSelected = True

    def _on_selection_changed(self, sender, args):
        self._capture_selection()

    def _capture_selection(self):
        self.selected_ids = set()
        self.selected_circuits = []
        for item in self.CircuitList.SelectedItems:
            self.selected_circuits.append(item.Tag)
            try:
                self.selected_ids.add(item.Tag.Id.IntegerValue)
            except Exception:
                pass

    def _select_all(self, sender, args):
        self.CircuitList.SelectAll()
        self._capture_selection()

    def _select_none(self, sender, args):
        self.CircuitList.UnselectAll()
        self._capture_selection()

    def _calculate(self, sender, args):
        self._capture_selection()
        if not self.selected_circuits:
            forms.alert('Select at least one circuit to calculate.')
            return
        self.Close()

    def _close(self, sender, args):
        self.selected_circuits = []
        self.Close()

    def show_dialog(self):
        self.ShowDialog()
        return self.selected_circuits
