# -*- coding: utf-8 -*-
"""Modal dialog for the Optimize workflow."""

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
    ComboBox,
    Grid,
    TextBlock,
)
from System.Windows.Media import Brushes  # noqa: E402

import optimize_workflow as _opt
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "OptimizeWindow.xaml",
)


class _ModeRow(object):
    def __init__(self, label, combo):
        self.label = label
        self.combo = combo

    @property
    def selected_mode(self):
        item = self.combo.SelectedItem
        if item is None:
            return _opt.MODE_NONE
        return getattr(item, "Tag", None) or _opt.MODE_NONE


class _MatchRow(object):
    def __init__(self, candidate, checkbox):
        self.candidate = candidate
        self.checkbox = checkbox

    @property
    def checked(self):
        return bool(self.checkbox.IsChecked)


class OptimizeController(object):

    def __init__(self, doc, profile_data, selected_element_ids=None):
        self.doc = doc
        self.profile_data = profile_data
        self.selected_element_ids = list(selected_element_ids or [])
        self.labels = _opt.collect_family_type_labels(
            doc, profile_data,
            selected_element_ids=self.selected_element_ids or None,
        )
        self.candidates = []
        self._mode_rows = []
        self._match_rows = []
        self.committed = False
        self._last_result = None
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_modes()
        self._wire_events()
        if self.selected_element_ids:
            self._set_status(
                "Selection mode — {} element(s) preselected. {} family:type "
                "row(s) shown. Pick a mode, then Match.".format(
                    len(self.selected_element_ids), len(self.labels)
                )
            )
        else:
            self._set_status(
                "Pick a mode per family:type, then Match. Default radius 5 ft."
            )

    def _lookup_controls(self):
        f = self.window.FindName
        self.radius_box = f("RadiusBox")
        self.match_btn = f("MatchButton")
        self.check_all_btn = f("CheckAllButton")
        self.uncheck_all_btn = f("UncheckAllButton")
        self.optimize_btn = f("OptimizeButton")
        self.close_btn = f("CloseButton")
        self.summary_label = f("SummaryLabel")
        self.status_label = f("StatusLabel")
        self.mode_rows_panel = f("ModeRowsPanel")
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
        self._h_optimize = self._delegate("optimize", lambda s, e: self._on_optimize(s, e))
        self._h_close = self._delegate("close", lambda s, e: self.window.Close())
        self.match_btn.Click += self._h_match
        self.check_all_btn.Click += self._h_check_all
        self.uncheck_all_btn.Click += self._h_uncheck_all
        self.optimize_btn.Click += self._h_optimize
        self.close_btn.Click += self._h_close

    def _populate_modes(self):
        self.mode_rows_panel.Children.Clear()
        self._mode_rows = []
        if not self.labels:
            tb = TextBlock()
            tb.Text = "No family:type labels in the active store."
            tb.Foreground = Brushes.Gray
            self.mode_rows_panel.Children.Add(tb)
            return
        for label in self.labels:
            grid = Grid()
            for w in (3.5, 2.0):
                col = ColumnDefinition()
                col.Width = GridLength(w, GridUnitType.Star)
                grid.ColumnDefinitions.Add(col)

            label_tb = TextBlock()
            label_tb.Text = label
            label_tb.Margin = Thickness(0, 4, 8, 4)
            label_tb.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(label_tb, 0)
            grid.Children.Add(label_tb)

            combo = ComboBox()
            combo.Margin = Thickness(0, 2, 0, 2)
            default_mode = _opt.default_mode_for_label(label)
            default_index = 0
            for idx, mode_key in enumerate(_opt.ALL_MODES):
                from System.Windows.Controls import ComboBoxItem
                item = ComboBoxItem()
                item.Content = _opt.MODE_LABELS[mode_key]
                item.Tag = mode_key
                combo.Items.Add(item)
                if mode_key == default_mode:
                    default_index = idx
            combo.SelectedIndex = default_index
            Grid.SetColumn(combo, 1)
            grid.Children.Add(combo)

            self.mode_rows_panel.Children.Add(grid)
            self._mode_rows.append(_ModeRow(label, combo))

    def _read_radius(self):
        text = (self.radius_box.Text or "").strip()
        try:
            v = float(text)
        except (TypeError, ValueError):
            return _opt.DEFAULT_SEARCH_RADIUS_FT
        if v <= 0:
            return _opt.DEFAULT_SEARCH_RADIUS_FT
        return v

    def _read_modes(self):
        out = {}
        for row in self._mode_rows:
            mode = row.selected_mode
            if mode and mode != _opt.MODE_NONE:
                out[row.label] = mode
        return out

    def _on_match(self, sender, e):
        modes = self._read_modes()
        if not modes:
            self._set_status("Pick at least one mode (other than '(skip)') before matching.")
            self.candidates = []
            self._render([])
            return
        opts = _opt.OptimizeOptions(
            search_radius_ft=self._read_radius(),
            mode_by_family_type=modes,
        )
        self.candidates = _opt.collect_candidates(
            self.doc, self.profile_data, opts,
            selected_element_ids=self.selected_element_ids or None,
        )
        self._render(self.candidates)
        n_total = len(self.candidates)
        n_skipped = sum(1 for c in self.candidates if c.skip)
        self.summary_label.Text = "{} candidate(s); {} with no host in radius".format(
            n_total, n_skipped
        )
        self._set_status(
            "Review the list, uncheck to skip, then Optimize." if n_total
            else "No candidates matched the configured modes."
        )

    def _render(self, candidates):
        self.match_rows_panel.Children.Clear()
        self._match_rows = []
        for c in candidates:
            grid, cb = self._row(c)
            self.match_rows_panel.Children.Add(grid)
            self._match_rows.append(_MatchRow(c, cb))
        self.optimize_btn.IsEnabled = bool(candidates)

    def _row(self, c):
        grid = Grid()
        for w in (0.0, 1.0, 2.5, 2.5, 2.5, 2.0):
            col = ColumnDefinition()
            if w == 0.0:
                col.Width = GridLength(28)
            else:
                col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        cb = CheckBox()
        cb.IsChecked = not c.skip
        cb.IsEnabled = not c.skip
        cb.Margin = Thickness(4, 2, 0, 2)
        cb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(cb, 0)
        grid.Children.Add(cb)

        mode_tb = TextBlock()
        mode_tb.Text = "[{}]".format(_opt.MODE_LABELS.get(c.mode, c.mode))
        mode_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(mode_tb, 1)
        grid.Children.Add(mode_tb)

        target_tb = TextBlock()
        target_tb.Text = "{}  (id {})".format(c.led_label or "?", c.child_id)
        target_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(target_tb, 2)
        grid.Children.Add(target_tb)

        cur_tb = TextBlock()
        cur_tb.Text = "Now: ({:.2f}, {:.2f}, {:.2f}) {:.1f}°".format(
            c.current_pt[0], c.current_pt[1], c.current_pt[2], c.current_rot
        )
        cur_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(cur_tb, 3)
        grid.Children.Add(cur_tb)

        tgt_tb = TextBlock()
        if c.skip:
            tgt_tb.Text = "—"
            tgt_tb.Foreground = Brushes.Gray
        else:
            tgt_tb.Text = "Target: ({:.2f}, {:.2f}, {:.2f}) {:.1f}°".format(
                c.target_pt[0], c.target_pt[1], c.target_pt[2], c.target_rot
            )
        tgt_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(tgt_tb, 4)
        grid.Children.Add(tgt_tb)

        host_tb = TextBlock()
        if c.skip:
            host_tb.Text = c.skip_reason
            host_tb.Foreground = Brushes.OrangeRed
        else:
            host_tb.Text = c.host_description
            host_tb.Foreground = Brushes.Gray
        host_tb.Margin = Thickness(0, 4, 0, 4)
        Grid.SetColumn(host_tb, 5)
        grid.Children.Add(host_tb)

        return grid, cb

    def _on_check_all(self, sender, e):
        for r in self._match_rows:
            if not r.candidate.skip:
                r.checkbox.IsChecked = True

    def _on_uncheck_all(self, sender, e):
        for r in self._match_rows:
            r.checkbox.IsChecked = False

    def _on_optimize(self, sender, e):
        from pyrevit import revit
        chosen = []
        for r in self._match_rows:
            if not r.checked or r.candidate.skip:
                continue
            chosen.append(r.candidate)
        if not chosen:
            self._set_status("Nothing checked")
            return
        with revit.Transaction("Optimize (MEPRFP 2.0)", doc=self.doc):
            result = _opt.execute_optimize(self.doc, chosen)
        self.committed = True
        self._last_result = result
        self._set_status(
            "Moved {}, re-parented {}, skipped {} no-host, {} warning(s).".format(
                result.moved_count,
                result.reparented_count,
                result.skipped_no_host,
                len(result.warnings),
            )
        )

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data, selected_element_ids=None):
    return OptimizeController(
        doc, profile_data, selected_element_ids=selected_element_ids
    ).show()
