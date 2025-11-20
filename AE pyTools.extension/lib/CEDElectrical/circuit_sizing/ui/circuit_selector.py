import os
from System.Collections.Generic import List
from System.Windows.Controls import CheckBox, TreeViewItem
from pyrevit import forms
from pyrevit import revit
from pyrevit import DB


class CircuitSelectorWindow(forms.WPFWindow):
    def __init__(self, reader, external_event, uidoc):
        xaml_path = os.path.join(os.path.dirname(__file__), 'circuit_selector.xaml')
        super(CircuitSelectorWindow, self).__init__(xaml_path)
        self.reader = reader
        self.external_event = external_event
        self.uidoc = uidoc
        self.selected_circuit_nodes = []

        self.SortMode.SelectedIndex = 0
        self._build_tree(self.reader.hierarchy)

        self.SearchBox.TextChanged += self._on_search
        self.SortMode.SelectionChanged += self._on_sort_changed
        self.CircuitTree.MouseDoubleClick += self._on_double_click
        self.SelectButton.Click += self._select_active
        self.CalculateButton.Click += self._run_calculation

    def _on_search(self, sender, args):
        term = self.SearchBox.Text
        data = self.reader.search(term)
        self._build_tree(data)

    def _on_sort_changed(self, sender, args):
        choice = self.SortMode.SelectedItem.Content.ToString()
        self.reader.refresh(choice)
        term = self.SearchBox.Text
        data = self.reader.search(term)
        self._build_tree(data)

    def _on_double_click(self, sender, args):
        item = self.CircuitTree.SelectedItem
        target = getattr(item, 'Tag', None)
        if target:
            self._select_element(target)

    def _select_active(self, sender, args):
        item = self.CircuitTree.SelectedItem
        target = getattr(item, 'Tag', None)
        if target:
            self._select_element(target)

    def _select_element(self, element):
        if not element:
            return
        try:
            ids = List[DB.ElementId]([element.Id])
            self.uidoc.Selection.SetElementIds(ids)
            self.uidoc.ShowElements(element)
        except Exception:
            pass

    def _run_calculation(self, sender, args):
        self.reader.selected_circuits = []
        for node in self.CircuitTree.Items:
            self._collect_checked_circuits(node)
        if not self.reader.selected_circuits:
            forms.alert("Select at least one circuit to calculate.")
            return
        self.external_event.Raise()

    def _collect_checked_circuits(self, tree_item):
        header = getattr(tree_item, 'Header', None)
        if isinstance(header, CheckBox) and header.IsChecked:
            circuit = getattr(tree_item, 'Tag', None)
            if circuit and isinstance(circuit, DB.Electrical.ElectricalSystem):
                self.reader.selected_circuits.append(circuit)
        for child in getattr(tree_item, 'Items', []):
            self._collect_checked_circuits(child)

    def _build_tree(self, hierarchy):
        self.CircuitTree.Items.Clear()
        for panel in hierarchy:
            panel_checkbox = CheckBox()
            panel_checkbox.Content = panel['label']
            panel_checkbox.IsThreeState = True

            panel_item = TreeViewItem()
            panel_item.Header = panel_checkbox
            panel_item.Tag = panel.get('element')
            for circuit in panel['children']:
                circuit_checkbox = CheckBox()
                circuit_checkbox.Content = circuit['label']

                circuit_item = TreeViewItem()
                circuit_item.Header = circuit_checkbox
                circuit_item.Tag = circuit['circuit']
                for device in circuit['devices']:
                    device_item = TreeViewItem()
                    device_item.Header = device['label']
                    device_item.Tag = device.get('element')
                    circuit_item.Items.Add(device_item)
                panel_item.Items.Add(circuit_item)
            self.CircuitTree.Items.Add(panel_item)

    def show(self):
        self.Show()
        return self
