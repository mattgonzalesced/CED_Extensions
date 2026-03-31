# -*- coding: utf-8 -*-
"""
Tabbed QA/QC report window with row actions.
"""

from pyrevit import forms
from System.Windows import Thickness, GridLength, GridUnitType, FontWeights, TextWrapping
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, TextBlock, Button


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
    ):
        forms.WPFWindow.__init__(self, xaml_path)

        self._tab_rows = tab_rows or {}
        self._select_child_callback = select_child_callback
        self._select_parent_callback = select_parent_callback
        self._snap_callback = snap_callback
        self._adjust_callback = adjust_callback
        self._button_meta = {}

        summary = self.FindName("SummaryText")
        if summary is not None:
            summary.Text = summary_text or ""

        self._build_tab("Tab1Grid", "Tab1Item", self._tab_rows.get("tab1") or [], "No Matching Parents")
        self._build_tab("Tab2Grid", "Tab2Item", self._tab_rows.get("tab2") or [], "Parents Found, Children Missing")
        self._build_tab("Tab3Grid", "Tab3Item", self._tab_rows.get("tab3") or [], "Original Parent Missing")
        self._build_tab("Tab4Grid", "Tab4Item", self._tab_rows.get("tab4") or [], "Parent Type Changed (Profile Exists)")
        self._build_tab("Tab5Grid", "Tab5Item", self._tab_rows.get("tab5") or [], "Parent Type Changed (No Profile)")
        self._build_tab(
            "Tab6Grid",
            "Tab6Item",
            self._tab_rows.get("tab6") or [],
            "Far from Parent",
            include_adjust=True,
        )

        total_issues = (
            len(self._tab_rows.get("tab2") or [])
            + len(self._tab_rows.get("tab3") or [])
            + len(self._tab_rows.get("tab4") or [])
            + len(self._tab_rows.get("tab5") or [])
            + len(self._tab_rows.get("tab6") or [])
        )
        footer = self.FindName("FooterText")
        if footer is not None:
            footer.Text = "Total issues (excluding Tab 1): {}".format(total_issues)

        close_btn = self.FindName("CloseButton")
        if close_btn is not None:
            close_btn.Click += self._on_close

    def _build_tab(self, grid_name, tab_name, rows, title, include_adjust=False):
        grid = self.FindName(grid_name)
        tab = self.FindName(tab_name)
        if tab is not None:
            tab.Header = "{} ({})".format(title, len(rows))
        if grid is None:
            return
        self._build_grid(grid, rows, include_adjust=include_adjust)

    def _build_grid(self, grid, rows, include_adjust=False):
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
            snap_btn.IsEnabled = row.get("snap_point") is not None or row.get("child_id") not in (None, "")

            child_btn.Click += self._on_action
            parent_btn.Click += self._on_action
            snap_btn.Click += self._on_action

            if include_adjust:
                adjust_col = len(columns) - 1
                adjust_btn = self._add_button_cell(grid, row_index, adjust_col, "Adjust")
                self._button_meta[adjust_btn] = {"row": row, "kind": "adjust"}
                adjust_btn.IsEnabled = bool(row.get("adjust_enabled"))
                adjust_btn.Click += self._on_action

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

    def _on_close(self, sender, args):
        try:
            self.DialogResult = True
        except Exception:
            pass
        self.Close()
