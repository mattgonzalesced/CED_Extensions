# -*- coding: utf-8 -*-
"""
Modal directive-mapping dialog. Returns
``{child_index: {param_name: directive_value}}`` or ``{}`` if cancelled
or all rows left static.

Implementation note: WPF data binding via pythonnet doesn't reliably
resolve Python ``@property`` accessors, and an INotifyPropertyChanged
subclass defines a CLR type that conflicts with itself on module
re-import (``_dev_reload`` purge -> "Duplicate type name within an
assembly"). So this dialog skips bindings entirely: rows are plain
Python objects, the UI is built programmatically inside a
``StackPanel``, and ``Apply`` reads selections directly from each
row's combo boxes.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import GridUnitType, Thickness  # noqa: E402
from System.Windows.Controls import (  # noqa: E402
    Border,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    TextBlock,
)
from System.Windows.Media import Brushes  # noqa: E402

import directives as _dir
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "DirectivesDialog.xaml"
)


# ---------------------------------------------------------------------
# Plain data row -- no .NET inheritance, no INotifyPropertyChanged
# ---------------------------------------------------------------------

class Row(object):
    """One (child, parameter) entry the user can configure."""

    def __init__(self, child_index, child_label, parameter_name, captured_value,
                 parent_options, sibling_options):
        self.child_index = child_index
        self.child_label = child_label
        self.parameter_name = parameter_name
        self.captured_value = "" if captured_value is None else str(captured_value)
        self.parent_options = list(parent_options)
        self.sibling_options = list(sibling_options)


# ---------------------------------------------------------------------
# UI construction
# ---------------------------------------------------------------------

_COLUMN_WEIGHTS = (3.0, 3.0, 2.0, 0.0, 3.0)
_MODE_COLUMN_PIXELS = 120


def _add_columns(grid):
    for weight in _COLUMN_WEIGHTS:
        col = ColumnDefinition()
        if weight == 0.0:
            from System.Windows import GridLength
            col.Width = GridLength(_MODE_COLUMN_PIXELS)
        else:
            from System.Windows import GridLength
            col.Width = GridLength(weight, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col)


def _set_col(elem, idx):
    Grid.SetColumn(elem, idx)


def _make_textblock(text, padding_left=0):
    tb = TextBlock()
    tb.Text = "" if text is None else str(text)
    tb.Margin = Thickness(padding_left, 4, 6, 4)
    return tb


def _make_combo(options, selected_index=0):
    combo = ComboBox()
    combo.Margin = Thickness(0, 2, 6, 2)
    for label in options:
        item = ComboBoxItem()
        item.Content = label
        combo.Items.Add(item)
    if 0 <= selected_index < combo.Items.Count:
        combo.SelectedIndex = selected_index
    return combo


def _refresh_target_options(target_combo, mode, parent_opts, sibling_opts):
    target_combo.Items.Clear()
    if mode == "parent":
        opts = parent_opts
    elif mode == "sibling":
        opts = sibling_opts
    else:
        opts = []
    for opt in opts:
        item = ComboBoxItem()
        item.Content = opt
        target_combo.Items.Add(item)
    target_combo.IsEnabled = bool(opts)
    if target_combo.Items.Count > 0:
        target_combo.SelectedIndex = 0


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class _DirectivesController(object):

    def __init__(self, rows):
        self.window = _wpf.load_xaml(_XAML_PATH)
        self.rows_panel = self.window.FindName("RowsPanel")
        self._rows = rows
        # Each entry: (row_data, mode_combo, target_combo)
        self._row_widgets = []
        self._build_rows()

        self.window.FindName("ApplyButton").Click += self._on_apply
        self.window.FindName("CancelButton").Click += self._on_cancel
        self.window.FindName("ResetButton").Click += self._on_reset

        self.result = None

    def _build_rows(self):
        for row in self._rows:
            grid = Grid()
            _add_columns(grid)

            cell_child = _make_textblock(row.child_label, padding_left=6)
            _set_col(cell_child, 0)
            grid.Children.Add(cell_child)

            cell_param = _make_textblock(row.parameter_name)
            _set_col(cell_param, 1)
            grid.Children.Add(cell_param)

            cell_value = _make_textblock(row.captured_value)
            _set_col(cell_value, 2)
            grid.Children.Add(cell_value)

            mode_combo = _make_combo(["static", "parent", "sibling"], selected_index=0)
            _set_col(mode_combo, 3)
            grid.Children.Add(mode_combo)

            target_combo = _make_combo([])
            target_combo.IsEnabled = False
            _set_col(target_combo, 4)
            grid.Children.Add(target_combo)

            # Wire mode change -> repopulate target options.
            def make_handler(r=row, mc=mode_combo, tc=target_combo):
                def _handler(sender, e):
                    if mc.SelectedItem is None:
                        return
                    mode = str(mc.SelectedItem.Content)
                    _refresh_target_options(tc, mode, r.parent_options, r.sibling_options)
                return _handler
            mode_combo.SelectionChanged += make_handler()

            border = Border()
            border.BorderBrush = Brushes.LightGray
            border.BorderThickness = Thickness(0, 0, 0, 1)
            border.Child = grid
            self.rows_panel.Children.Add(border)

            self._row_widgets.append((row, mode_combo, target_combo))

    # -- buttons ------------------------------------------------------

    def _on_reset(self, sender, e):
        for _row, mode_combo, target_combo in self._row_widgets:
            mode_combo.SelectedIndex = 0
            target_combo.Items.Clear()
            target_combo.IsEnabled = False

    def _on_cancel(self, sender, e):
        self.result = None
        self.window.Close()

    def _on_apply(self, sender, e):
        out = {}
        for row, mode_combo, target_combo in self._row_widgets:
            mode_item = mode_combo.SelectedItem
            if mode_item is None:
                continue
            mode = str(mode_item.Content)
            if mode == "static":
                continue
            target_item = target_combo.SelectedItem
            if target_item is None:
                continue
            target = str(target_item.Content)
            if not target:
                continue
            out.setdefault(row.child_index, {})
            if mode == "parent":
                out[row.child_index][row.parameter_name] = _dir.parent_directive(target)
            elif mode == "sibling":
                if " :: " not in target:
                    continue
                led_id, _, pname = target.partition(" :: ")
                out[row.child_index][row.parameter_name] = _dir.sibling_directive(
                    led_id.strip(), pname.strip()
                )
        self.result = out
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self.result


# ---------------------------------------------------------------------
# Public API (preserved)
# ---------------------------------------------------------------------

def show_dialog(rows):
    """Show the dialog. Returns ``{child_index: {param_name: directive}}`` or ``None``."""
    return _DirectivesController(rows).show()


def build_rows(child_refs, child_param_values, parent_param_names, sibling_options):
    """Convenience constructor: map raw inputs to ``Row`` objects."""
    rows = []
    for idx, child_ref in enumerate(child_refs):
        label = "[{}] {}".format(idx + 1, _format_child_label(child_ref))
        for param_name, captured in (child_param_values.get(idx) or {}).items():
            rows.append(Row(
                idx, label, param_name, captured,
                sorted(parent_param_names),
                sibling_options,
            ))
    return rows


def _format_child_label(child_ref):
    elem = getattr(child_ref, "element", None)
    if elem is None:
        return "(unknown)"
    cat = getattr(elem, "Category", None)
    cat_name = cat.Name if cat else type(elem).__name__
    eid = getattr(elem, "Id", None)
    eid_val = getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
    return "{} #{}".format(cat_name, eid_val)
