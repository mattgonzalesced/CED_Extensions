# -*- coding: utf-8 -*-
"""Modal dialog for the Follow Parent workflow."""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import (  # noqa: E402
    GridLength,
    GridUnitType,
    RoutedEventHandler,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (  # noqa: E402
    CheckBox,
    ColumnDefinition,
    Grid,
    TextBlock,
)

import follow_parent_workflow as _fp
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "FollowParentWindow.xaml",
)


class _MatchRow(object):
    def __init__(self, candidate, checkbox):
        self.candidate = candidate
        self.checkbox = checkbox

    @property
    def checked(self):
        return bool(self.checkbox.IsChecked)


class FollowParentController(object):

    def __init__(self, doc, profile_data):
        self.doc = doc
        self.profile_data = profile_data
        self.profiles = list(profile_data.get("equipment_definitions") or [])
        self.candidates = []
        self._match_rows = []
        self.committed = False
        self._last_result = None
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_filters()
        self._wire_events()
        self._set_status("Tweak filters, then Match.")

    def _lookup_controls(self):
        f = self.window.FindName
        self.category_list = f("CategoryList")
        self.profile_list = f("ProfileList")
        self.skip_aligned = f("SkipAlignedCheck")
        self.match_btn = f("MatchButton")
        self.check_all_btn = f("CheckAllButton")
        self.uncheck_all_btn = f("UncheckAllButton")
        self.follow_btn = f("FollowButton")
        self.close_btn = f("CloseButton")
        self.summary_label = f("SummaryLabel")
        self.status_label = f("StatusLabel")
        self.match_rows_panel = f("MatchRowsPanel")

    def _delegate(self, label, fn):
        def wrapped(s, e):
            try:
                self._set_status("[{}] running...".format(label))
                fn(s, e)
            except Exception as exc:
                self._set_status("[{}] error: {}".format(label, exc))
                raise
        return RoutedEventHandler(wrapped)

    def _wire_events(self):
        self._h_match = self._delegate("match", lambda s, e: self._on_match(s, e))
        self._h_check_all = self._delegate("check-all", lambda s, e: self._on_check_all(s, e))
        self._h_uncheck_all = self._delegate("uncheck-all", lambda s, e: self._on_uncheck_all(s, e))
        self._h_follow = self._delegate("follow", lambda s, e: self._on_follow(s, e))
        self._h_close = self._delegate("close", lambda s, e: self.window.Close())
        self.match_btn.Click += self._h_match
        self.check_all_btn.Click += self._h_check_all
        self.uncheck_all_btn.Click += self._h_uncheck_all
        self.follow_btn.Click += self._h_follow
        self.close_btn.Click += self._h_close

    def _populate_filters(self):
        cats = sorted({
            (p.get("parent_filter") or {}).get("category") or ""
            for p in self.profiles if isinstance(p, dict)
        })
        cats = [c for c in cats if c]
        self.category_list.Items.Clear()
        for c in cats:
            self.category_list.Items.Add(c)
        self.profile_list.Items.Clear()
        for p in self.profiles:
            self.profile_list.Items.Add(
                "{}  ({})".format(p.get("name") or "(unnamed)", p.get("id") or "?")
            )

    def _selected_profile_ids(self):
        ids = set()
        for label in self.profile_list.SelectedItems:
            label = str(label)
            if "(" in label and label.endswith(")"):
                ids.add(label.rsplit("(", 1)[1].rstrip(")"))
        return ids or None

    def _selected_categories(self):
        out = {str(item) for item in self.category_list.SelectedItems}
        return out or None

    def _on_match(self, sender, e):
        filters = _fp.CollectFilters(
            profile_ids=self._selected_profile_ids(),
            categories=self._selected_categories(),
        )
        try:
            self.candidates = _fp.collect_candidates(
                self.doc, self.profile_data, filters, refuse_linked=True
            )
        except ValueError as exc:
            self._set_status(str(exc))
            self.candidates = []
            self._render([])
            return
        if self.skip_aligned.IsChecked:
            _fp.mark_aligned_skips(self.candidates)
        self._render(self.candidates)
        n_total = len(self.candidates)
        n_aligned = sum(1 for c in self.candidates if c.skip)
        self.summary_label.Text = "{} candidate(s); {} already aligned".format(
            n_total, n_aligned
        )
        self._set_status(
            "Review the list, uncheck to skip, then Follow." if n_total
            else "No candidates matched the filters."
        )

    def _render(self, candidates):
        self.match_rows_panel.Children.Clear()
        self._match_rows = []
        for c in candidates:
            grid, cb = self._row(c)
            self.match_rows_panel.Children.Add(grid)
            self._match_rows.append(_MatchRow(c, cb))
        self.follow_btn.IsEnabled = bool(candidates)

    def _row(self, c):
        grid = Grid()
        for w in (0.0, 1.5, 2.5, 3.0, 2.0):
            col = ColumnDefinition()
            if w == 0.0:
                col.Width = GridLength(28)
            else:
                col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        cb = CheckBox()
        cb.IsChecked = not c.skip
        cb.Margin = Thickness(4, 2, 0, 2)
        cb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(cb, 0)
        grid.Children.Add(cb)

        led_tb = TextBlock()
        led_tb.Text = "{}  ({})".format(c.led_label or "?", c.led_id or "?")
        led_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(led_tb, 1)
        grid.Children.Add(led_tb)

        cur_tb = TextBlock()
        cur_tb.Text = "Now: ({:.2f}, {:.2f}, {:.2f}) {:.1f}°".format(
            c.current_pt[0], c.current_pt[1], c.current_pt[2], c.current_rot
        )
        cur_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(cur_tb, 2)
        grid.Children.Add(cur_tb)

        tgt_tb = TextBlock()
        tgt_tb.Text = "Target: ({:.2f}, {:.2f}, {:.2f}) {:.1f}°".format(
            c.target_pt[0], c.target_pt[1], c.target_pt[2], c.target_rot
        )
        tgt_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(tgt_tb, 3)
        grid.Children.Add(tgt_tb)

        status_tb = TextBlock()
        status_tb.Text = c.skip_reason if c.skip and c.skip_reason else "ready"
        if c.skip:
            from System.Windows.Media import Brushes
            status_tb.Foreground = Brushes.Gray
        status_tb.Margin = Thickness(0, 4, 0, 4)
        Grid.SetColumn(status_tb, 4)
        grid.Children.Add(status_tb)

        return grid, cb

    def _on_check_all(self, sender, e):
        for r in self._match_rows:
            r.checkbox.IsChecked = True

    def _on_uncheck_all(self, sender, e):
        for r in self._match_rows:
            r.checkbox.IsChecked = False

    def _on_follow(self, sender, e):
        from pyrevit import revit
        chosen = []
        for r in self._match_rows:
            if not r.checked:
                continue
            r.candidate.skip = False  # override aligned-flag
            chosen.append(r.candidate)
        if not chosen:
            self._set_status("Nothing checked to follow")
            return
        with revit.Transaction("Follow Parent (MEPRFP 2.0)", doc=self.doc):
            result = _fp.execute_follow(self.doc, chosen)
        self.committed = True
        self._last_result = result
        self._set_status(
            "Moved {}, skipped {} aligned, {} warning(s).".format(
                result.moved_count, result.skipped_aligned, len(result.warnings)
            )
        )

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data):
    return FollowParentController(doc, profile_data).show()
