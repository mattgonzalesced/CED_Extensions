# -*- coding: utf-8 -*-
"""
Tree-view UI for the synced-relationship audit (option C).

Hierarchy: Profile -> LED -> Conflict. Each conflict row carries a
ComboBox per ``Conflict.UPDATE_CHILD / UPDATE_PARENT / SKIP`` choice.
On Apply, the controller calls ``sync_audit.apply_resolution`` for each
non-skip choice inside one Revit transaction.
"""

import os

import clr  # noqa: F401

from System.Windows.Controls import (  # noqa: E402
    ComboBox,
    ComboBoxItem,
    Grid,
    GridLength,
    ColumnDefinition,
    StackPanel,
    TextBlock,
    TreeViewItem,
)
from System.Windows import Thickness, GridUnitType  # noqa: E402

import sync_audit
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "SyncAuditWindow.xaml"
)


# Maps display label -> action constant.
_ACTION_LABELS = [
    ("Update child", sync_audit.Conflict.UPDATE_CHILD),
    ("Update parent", sync_audit.Conflict.UPDATE_PARENT),
    ("Skip", sync_audit.Conflict.SKIP),
]


class _ConflictRow(object):
    """One leaf in the tree. Holds the ComboBox so we can read its choice on Apply."""
    def __init__(self, conflict, combo):
        self.conflict = conflict
        self.combo = combo

    @property
    def chosen_action(self):
        item = self.combo.SelectedItem
        if item is None:
            return sync_audit.Conflict.SKIP
        return item.Tag


class SyncAuditController(object):

    def __init__(self, doc, profile_data, transaction_factory):
        """``transaction_factory`` is a callable returning a context
        manager (e.g. ``lambda: revit.Transaction("...", doc=doc)``).
        We let the caller own transaction semantics.
        """
        self.doc = doc
        self.profile_data = profile_data
        self.transaction_factory = transaction_factory
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = []
        self._wire_controls()
        self.refresh()

    # ---- wiring ------------------------------------------------------

    def _wire_controls(self):
        f = self.window.FindName
        self.tree = f("ConflictTree")
        self.summary = f("SummaryLabel")
        self.bulk_combo = f("BulkActionCombo")
        f("RefreshButton").Click += self._on_refresh
        f("ApplyButton").Click += self._on_apply
        f("CloseButton").Click += self._on_close
        self.bulk_combo.SelectionChanged += self._on_bulk

    # ---- public ------------------------------------------------------

    def show(self):
        self.window.ShowDialog()

    def refresh(self):
        conflicts = sync_audit.detect_conflicts(self.doc, self.profile_data)
        self._populate_tree(conflicts)
        self.summary.Text = "{} conflict(s) across {} profile(s)".format(
            len(conflicts),
            len({c.profile_id for c in conflicts}),
        )

    # ---- tree --------------------------------------------------------

    def _populate_tree(self, conflicts):
        self.tree.Items.Clear()
        self._rows = []
        # Group: profile -> led -> [conflicts]
        by_profile = {}
        for c in conflicts:
            by_profile.setdefault(
                (c.profile_id, c.profile_name), {}
            ).setdefault((c.led_id, c.led_label), []).append(c)

        for (profile_id, profile_name), led_groups in by_profile.items():
            profile_node = TreeViewItem()
            profile_node.Header = "{}  ({})  -  {} conflict(s)".format(
                profile_name, profile_id,
                sum(len(v) for v in led_groups.values()),
            )
            profile_node.IsExpanded = True
            for (led_id, led_label), led_conflicts in led_groups.items():
                led_node = TreeViewItem()
                led_node.Header = "{}  ({})  -  {} conflict(s)".format(
                    led_label, led_id, len(led_conflicts)
                )
                led_node.IsExpanded = True
                for conflict in led_conflicts:
                    led_node.Items.Add(self._build_conflict_node(conflict))
                profile_node.Items.Add(led_node)
            self.tree.Items.Add(profile_node)

    def _build_conflict_node(self, conflict):
        node = TreeViewItem()
        node.IsExpanded = False
        grid = Grid()
        for w in (3, 2, 2, 2):
            col = ColumnDefinition()
            col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        param = TextBlock()
        param.Text = "{}  ({})".format(conflict.parameter_name, conflict.kind)
        param.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(param, 0)
        grid.Children.Add(param)

        actual = TextBlock()
        actual.Text = "Actual: {}".format(_short(conflict.actual_value))
        actual.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(actual, 1)
        grid.Children.Add(actual)

        expected = TextBlock()
        expected.Text = "Expected: {}".format(_short(conflict.expected_value))
        expected.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(expected, 2)
        grid.Children.Add(expected)

        combo = ComboBox()
        combo.Margin = Thickness(0)
        for label, action in _ACTION_LABELS:
            item = ComboBoxItem()
            item.Content = label
            item.Tag = action
            combo.Items.Add(item)
        combo.SelectedIndex = 2  # default = skip
        Grid.SetColumn(combo, 3)
        grid.Children.Add(combo)

        node.Header = grid
        self._rows.append(_ConflictRow(conflict, combo))
        return node

    # ---- handlers ----------------------------------------------------

    def _on_refresh(self, sender, e):
        self.refresh()

    def _on_close(self, sender, e):
        self.window.Close()

    def _on_bulk(self, sender, e):
        idx = self.bulk_combo.SelectedIndex
        # 0 = placeholder, 1 = update child, 2 = update parent, 3 = skip
        if idx <= 0:
            return
        target = (
            sync_audit.Conflict.UPDATE_CHILD,
            sync_audit.Conflict.UPDATE_PARENT,
            sync_audit.Conflict.SKIP,
        )[idx - 1]
        for row in self._rows:
            for i in range(row.combo.Items.Count):
                if row.combo.Items[i].Tag == target:
                    row.combo.SelectedIndex = i
                    break

    def _on_apply(self, sender, e):
        applied = 0
        skipped = 0
        with self.transaction_factory():
            for row in self._rows:
                ok = sync_audit.apply_resolution(
                    self.doc, row.conflict, row.chosen_action
                )
                if ok:
                    applied += 1
                else:
                    skipped += 1
        self.refresh()
        self.summary.Text = "Applied {}, skipped {}".format(applied, skipped)


def _short(value):
    if value is None:
        return "(empty)"
    s = str(value)
    if len(s) > 40:
        return s[:37] + "..."
    return s


def show_modal(doc, profile_data, transaction_factory):
    SyncAuditController(doc, profile_data, transaction_factory).show()
