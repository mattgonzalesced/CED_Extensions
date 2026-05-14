# -*- coding: utf-8 -*-
"""
Tabbed QA/QC report window with row actions.

Each tab supports multiple stacked filter rules. Each rule has an
operator (``Contains`` / ``Does not contain``) and a text input;
rules are AND-combined and match against the row's profile,
description, parent, and child columns. Add / Clear-all buttons
manage the rule list per tab.

Tab 2 has a per-row "Place" button that runs the placement engine
directly (no dockable pane) via a callback supplied by script.py.
"""

from pyrevit import forms
from System.Windows import (
    Thickness, GridLength, GridUnitType, FontWeights, TextWrapping,
    VerticalAlignment, HorizontalAlignment,
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, TextBlock, Button,
    StackPanel, ComboBox, TextBox, Orientation,
)


# Tab metadata — drives initial build, filter rebuilds, and which
# columns appear per tab.
_TAB_DEFS = (
    {
        "storage_key": "tab1", "title": "No Matching Parents",
        "grid": "Tab1Grid", "tab": "Tab1Item",
        "rules_panel": "Tab1FilterRules", "add_btn": "Tab1AddFilter",
        "clear_btn": "Tab1ClearFilters", "count": "Tab1MatchCount",
        "include_adjust": False, "include_fix_id": False, "include_place": False,
    },
    {
        "storage_key": "tab2", "title": "Parents Found, Children Missing",
        "grid": "Tab2Grid", "tab": "Tab2Item",
        "rules_panel": "Tab2FilterRules", "add_btn": "Tab2AddFilter",
        "clear_btn": "Tab2ClearFilters", "count": "Tab2MatchCount",
        "include_adjust": False, "include_fix_id": False, "include_place": True,
    },
    {
        "storage_key": "tab3", "title": "Original Parent Missing",
        "grid": "Tab3Grid", "tab": "Tab3Item",
        "rules_panel": "Tab3FilterRules", "add_btn": "Tab3AddFilter",
        "clear_btn": "Tab3ClearFilters", "count": "Tab3MatchCount",
        "include_adjust": False, "include_fix_id": False, "include_place": False,
    },
    {
        "storage_key": "tab4", "title": "Parent Type Changed (Profile Exists)",
        "grid": "Tab4Grid", "tab": "Tab4Item",
        "rules_panel": "Tab4FilterRules", "add_btn": "Tab4AddFilter",
        "clear_btn": "Tab4ClearFilters", "count": "Tab4MatchCount",
        "include_adjust": False, "include_fix_id": False, "include_place": False,
    },
    {
        "storage_key": "tab5", "title": "Parent Type Changed (No Profile)",
        "grid": "Tab5Grid", "tab": "Tab5Item",
        "rules_panel": "Tab5FilterRules", "add_btn": "Tab5AddFilter",
        "clear_btn": "Tab5ClearFilters", "count": "Tab5MatchCount",
        "include_adjust": False, "include_fix_id": False, "include_place": False,
    },
    {
        "storage_key": "tab6", "title": "Far from Parent",
        "grid": "Tab6Grid", "tab": "Tab6Item",
        "rules_panel": "Tab6FilterRules", "add_btn": "Tab6AddFilter",
        "clear_btn": "Tab6ClearFilters", "count": "Tab6MatchCount",
        "include_adjust": True, "include_fix_id": False, "include_place": False,
    },
    {
        "storage_key": "tab7", "title": "ID Discrepancies",
        "grid": "Tab7Grid", "tab": "Tab7Item",
        "rules_panel": "Tab7FilterRules", "add_btn": "Tab7AddFilter",
        "clear_btn": "Tab7ClearFilters", "count": "Tab7MatchCount",
        "include_adjust": False, "include_fix_id": True, "include_place": False,
    },
)

_OP_CONTAINS = "Contains"
_OP_NOT_CONTAINS = "Does not contain"


def _row_haystack(row):
    """Lower-cased concatenation of the row's text columns, used for
    substring matching by every filter rule.
    """
    parts = []
    for key in ("profile", "description", "parent_text", "child_text"):
        value = row.get(key) or ""
        if value:
            parts.append(str(value).lower())
    return "\n".join(parts)


class _FilterRule(object):
    """One filter rule row — operator combo + text box + remove button.

    Built in the tab's ``FilterRules`` StackPanel. The owning window
    calls ``matches(row, haystack)`` for each row; the rule contributes
    one AND-condition.
    """

    def __init__(self, window, tab_def, container_panel):
        self._window = window
        self._tab_def = tab_def
        self._container = container_panel

        # Layout: horizontal StackPanel with [op combo] [text box] [×]
        self.row_panel = StackPanel()
        self.row_panel.Orientation = Orientation.Horizontal
        self.row_panel.Margin = Thickness(0, 0, 0, 4)

        self.op_combo = ComboBox()
        self.op_combo.Width = 150
        self.op_combo.Items.Add(_OP_CONTAINS)
        self.op_combo.Items.Add(_OP_NOT_CONTAINS)
        self.op_combo.SelectedIndex = 0
        self.op_combo.VerticalAlignment = VerticalAlignment.Center
        self.op_combo.Margin = Thickness(0, 0, 6, 0)
        self.op_combo.SelectionChanged += self._on_change
        self.row_panel.Children.Add(self.op_combo)

        self.text_box = TextBox()
        self.text_box.Width = 320
        self.text_box.Padding = Thickness(2)
        self.text_box.VerticalAlignment = VerticalAlignment.Center
        self.text_box.Margin = Thickness(0, 0, 6, 0)
        self.text_box.TextChanged += self._on_change
        self.row_panel.Children.Add(self.text_box)

        self.remove_btn = Button()
        self.remove_btn.Content = "Remove"
        self.remove_btn.Width = 70
        self.remove_btn.VerticalAlignment = VerticalAlignment.Center
        self.remove_btn.Click += self._on_remove
        self.row_panel.Children.Add(self.remove_btn)

        self._container.Children.Add(self.row_panel)

    def _on_change(self, sender, args):
        self._window._refresh_tab(self._tab_def)

    def _on_remove(self, sender, args):
        self._window._remove_rule(self._tab_def, self)

    def matches(self, row, haystack):
        text = ""
        try:
            text = (self.text_box.Text or "").strip().lower()
        except Exception:
            text = ""
        if not text:
            # An empty rule is a no-op — doesn't constrain the row set.
            return True
        try:
            op = self.op_combo.SelectedItem
        except Exception:
            op = _OP_CONTAINS
        contains = text in haystack
        if op == _OP_NOT_CONTAINS:
            return not contains
        return contains


class QAQCReportWindow(forms.WPFWindow):
    def __init__(
        self,
        xaml_path,
        tab_rows,
        summary_text="",
        select_child_callback=None,
        select_parent_callback=None,
        snap_callback=None,
        adjust_callback=None,
        fix_id_callback=None,
        place_callback=None,
    ):
        forms.WPFWindow.__init__(self, xaml_path)

        self._tab_rows = tab_rows or {}
        self._select_child_callback = select_child_callback
        self._select_parent_callback = select_parent_callback
        self._snap_callback = snap_callback
        self._adjust_callback = adjust_callback
        self._fix_id_callback = fix_id_callback
        self._place_callback = place_callback
        self._button_meta = {}
        # storage_key -> list of _FilterRule
        self._rules_by_tab = {tab_def["storage_key"]: [] for tab_def in _TAB_DEFS}

        summary = self.FindName("SummaryText")
        if summary is not None:
            summary.Text = summary_text or ""

        for tab_def in _TAB_DEFS:
            self._wire_tab(tab_def)
            self._refresh_tab(tab_def)

        total_issues = sum(
            len(self._tab_rows.get(tab_def["storage_key"]) or [])
            for tab_def in _TAB_DEFS
            if tab_def["storage_key"] != "tab1"
        )
        footer = self.FindName("FooterText")
        if footer is not None:
            footer.Text = "Total issues (excluding Tab 1): {}".format(total_issues)

        close_btn = self.FindName("CloseButton")
        if close_btn is not None:
            close_btn.Click += self._on_close

    # ------------------------------------------------------------------
    # Tab wiring
    # ------------------------------------------------------------------

    def _wire_tab(self, tab_def):
        add_btn = self.FindName(tab_def["add_btn"])
        clear_btn = self.FindName(tab_def["clear_btn"])
        if add_btn is not None:
            def _on_add(sender, args, td=tab_def):
                self._add_rule(td)
            add_btn.Click += _on_add
        if clear_btn is not None:
            def _on_clear(sender, args, td=tab_def):
                self._clear_rules(td)
            clear_btn.Click += _on_clear

    def _add_rule(self, tab_def):
        panel = self.FindName(tab_def["rules_panel"])
        if panel is None:
            return
        rule = _FilterRule(self, tab_def, panel)
        self._rules_by_tab.setdefault(tab_def["storage_key"], []).append(rule)
        # No grid rebuild needed — empty rule doesn't affect anything.

    def _remove_rule(self, tab_def, rule):
        rules = self._rules_by_tab.get(tab_def["storage_key"]) or []
        if rule not in rules:
            return
        rules.remove(rule)
        panel = self.FindName(tab_def["rules_panel"])
        if panel is not None and rule.row_panel in panel.Children:
            panel.Children.Remove(rule.row_panel)
        self._refresh_tab(tab_def)

    def _clear_rules(self, tab_def):
        rules = self._rules_by_tab.get(tab_def["storage_key"]) or []
        panel = self.FindName(tab_def["rules_panel"])
        if panel is not None:
            for rule in list(rules):
                if rule.row_panel in panel.Children:
                    panel.Children.Remove(rule.row_panel)
        self._rules_by_tab[tab_def["storage_key"]] = []
        self._refresh_tab(tab_def)

    # ------------------------------------------------------------------
    # Refresh / render
    # ------------------------------------------------------------------

    def _refresh_tab(self, tab_def):
        rows_all = self._tab_rows.get(tab_def["storage_key"]) or []
        rules = self._rules_by_tab.get(tab_def["storage_key"]) or []

        if not rules:
            rows_visible = list(rows_all)
        else:
            rows_visible = []
            for row in rows_all:
                haystack = _row_haystack(row)
                if all(rule.matches(row, haystack) for rule in rules):
                    rows_visible.append(row)

        tab = self.FindName(tab_def["tab"])
        if tab is not None:
            tab.Header = "{} ({})".format(tab_def["title"], len(rows_all))

        count_lbl = self.FindName(tab_def["count"])
        if count_lbl is not None:
            if rules:
                count_lbl.Text = "Showing {} of {}".format(
                    len(rows_visible), len(rows_all)
                )
            else:
                count_lbl.Text = "{} row(s)".format(len(rows_all))

        grid = self.FindName(tab_def["grid"])
        if grid is None:
            return
        self._build_grid(
            grid,
            rows_visible,
            include_adjust=tab_def["include_adjust"],
            include_fix_id=tab_def["include_fix_id"],
            include_place=tab_def["include_place"],
        )

    def _build_grid(
        self, grid, rows,
        include_adjust=False, include_fix_id=False, include_place=False,
    ):
        grid.Children.Clear()
        grid.RowDefinitions.Clear()
        grid.ColumnDefinitions.Clear()

        columns = [
            ("Profile", 210),
            ("Description", 540),
            ("Parent", 300),
            ("Child", 300),
            ("Select Child", 95),
            ("Select Parent", 105),
            ("Snap", 80),
        ]
        if include_adjust:
            columns.append(("Adjust", 90))
        if include_fix_id:
            columns.append(("Fix ID", 90))
        if include_place:
            columns.append(("Place", 90))
        for _title, width in columns:
            col = ColumnDefinition()
            col.Width = GridLength(width, GridUnitType.Pixel)
            grid.ColumnDefinitions.Add(col)

        header = RowDefinition()
        header.Height = GridLength(28, GridUnitType.Pixel)
        grid.RowDefinitions.Add(header)

        for idx, col_data in enumerate(columns):
            cell = TextBlock()
            cell.Text = col_data[0]
            cell.FontWeight = FontWeights.Bold
            cell.Margin = Thickness(0, 0, 6, 4)
            Grid.SetRow(cell, 0)
            Grid.SetColumn(cell, idx)
            grid.Children.Add(cell)

        if not rows:
            row_def = RowDefinition()
            row_def.Height = GridLength(26, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)
            cell = TextBlock()
            cell.Text = "No items in this category."
            cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(cell, 1)
            Grid.SetColumn(cell, 0)
            Grid.SetColumnSpan(cell, len(columns))
            grid.Children.Add(cell)
            return

        # Optional-button column offsets, counted from the right.
        # Order matches the `if include_*` appends above.
        optional_cols = []  # list of (key, col_index)
        next_col = 7  # first optional column
        if include_adjust:
            optional_cols.append(("adjust", next_col))
            next_col += 1
        if include_fix_id:
            optional_cols.append(("fix_id", next_col))
            next_col += 1
        if include_place:
            optional_cols.append(("place", next_col))
            next_col += 1

        row_index = 1
        for row in rows:
            row_def = RowDefinition()
            row_def.Height = GridLength(30, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)

            self._add_text_cell(grid, row_index, 0, row.get("profile"))
            self._add_text_cell(grid, row_index, 1, row.get("description"))
            self._add_text_cell(grid, row_index, 2, row.get("parent_text"))
            self._add_text_cell(grid, row_index, 3, row.get("child_text"))

            child_btn = self._add_button_cell(grid, row_index, 4, "Select")
            parent_btn = self._add_button_cell(grid, row_index, 5, "Select")
            snap_btn = self._add_button_cell(grid, row_index, 6, "Snap")

            self._button_meta[child_btn] = {"row": row, "kind": "child"}
            self._button_meta[parent_btn] = {"row": row, "kind": "parent"}
            self._button_meta[snap_btn] = {"row": row, "kind": "snap"}

            child_btn.IsEnabled = row.get("child_id") not in (None, "")
            parent_btn.IsEnabled = row.get("parent_id") not in (None, "")
            snap_btn.IsEnabled = (
                row.get("snap_point") is not None
                or row.get("child_id") not in (None, "")
            )

            child_btn.Click += self._on_action
            parent_btn.Click += self._on_action
            snap_btn.Click += self._on_action

            for kind, col_index in optional_cols:
                caption = (
                    "Adjust" if kind == "adjust"
                    else "Fix ID" if kind == "fix_id"
                    else "Place"
                )
                btn = self._add_button_cell(grid, row_index, col_index, caption)
                self._button_meta[btn] = {"row": row, "kind": kind}
                if kind == "adjust":
                    btn.IsEnabled = bool(row.get("adjust_enabled"))
                elif kind == "fix_id":
                    btn.IsEnabled = bool(row.get("fix_id_enabled"))
                elif kind == "place":
                    btn.IsEnabled = bool((row.get("profile") or "").strip())
                btn.Click += self._on_action

            row_index += 1

    def _add_text_cell(self, grid, row_index, col_index, text):
        cell = TextBlock()
        cell.Text = text or ""
        cell.TextWrapping = TextWrapping.Wrap
        cell.Margin = Thickness(0, 0, 6, 2)
        Grid.SetRow(cell, row_index)
        Grid.SetColumn(cell, col_index)
        grid.Children.Add(cell)

    def _add_button_cell(self, grid, row_index, col_index, caption):
        btn = Button()
        btn.Content = caption
        btn.Width = 82
        btn.Margin = Thickness(0, 0, 6, 2)
        Grid.SetRow(btn, row_index)
        Grid.SetColumn(btn, col_index)
        grid.Children.Add(btn)
        return btn

    # ------------------------------------------------------------------
    # Row actions
    # ------------------------------------------------------------------

    def _on_action(self, sender, args):
        meta = self._button_meta.get(sender) or {}
        row = meta.get("row") or {}
        kind = meta.get("kind")
        if kind == "child" and self._select_child_callback:
            self._select_child_callback(row)
            return
        if kind == "parent" and self._select_parent_callback:
            self._select_parent_callback(row)
            return
        if kind == "snap" and self._snap_callback:
            self._snap_callback(row)
            return
        if kind == "adjust" and self._adjust_callback:
            self._adjust_callback(row)
            return
        if kind == "fix_id" and self._fix_id_callback:
            self._fix_id_callback(row)
            return
        if kind == "place" and self._place_callback:
            self._place_callback(row)

    def _on_close(self, sender, args):
        try:
            self.DialogResult = True
        except Exception:
            pass
        self.Close()
