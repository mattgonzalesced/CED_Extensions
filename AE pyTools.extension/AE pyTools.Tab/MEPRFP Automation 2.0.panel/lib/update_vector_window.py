# -*- coding: utf-8 -*-
"""Modal dialog for Update Vector (selection-based)."""

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
from System.Windows.Media import Brushes  # noqa: E402

import update_vector_workflow as _uv
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "UpdateVectorWindow.xaml",
)


class _MatchRow(object):
    def __init__(self, candidate, checkbox):
        self.candidate = candidate
        self.checkbox = checkbox

    @property
    def checked(self):
        return bool(self.checkbox.IsChecked)


class UpdateVectorController(object):

    def __init__(self, doc, profile_data, selected_element_ids):
        self.doc = doc
        self.profile_data = profile_data
        self.selected_ids = list(selected_element_ids or [])
        self.candidates = []
        self._match_rows = []
        self.committed = False
        self._last_result = None
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._wire_events()
        self._set_status(
            "Click 'Match selection' to compute updated offsets for the {} "
            "selected element(s).".format(len(self.selected_ids))
        )

    def _lookup_controls(self):
        f = self.window.FindName
        self.match_btn = f("MatchButton")
        self.check_all_btn = f("CheckAllButton")
        self.uncheck_all_btn = f("UncheckAllButton")
        self.update_btn = f("UpdateButton")
        self.close_btn = f("CloseButton")
        self.summary_label = f("SummaryLabel")
        self.status_label = f("StatusLabel")
        self.match_rows_panel = f("MatchRowsPanel")

    def _wire_events(self):
        def safe(label, fn):
            def wrapped(s, e):
                try:
                    self._set_status("[{}] running...".format(label))
                    fn(s, e)
                except Exception as exc:
                    self._set_status("[{}] error: {}".format(label, exc))
                    raise
            return RoutedEventHandler(wrapped)

        self._h_match = safe("match", lambda s, e: self._on_match(s, e))
        self._h_check_all = safe("check-all", lambda s, e: self._on_check_all(s, e))
        self._h_uncheck_all = safe("uncheck-all", lambda s, e: self._on_uncheck_all(s, e))
        self._h_update = safe("update", lambda s, e: self._on_update(s, e))
        self._h_close = safe("close", lambda s, e: self.window.Close())
        self.match_btn.Click += self._h_match
        self.check_all_btn.Click += self._h_check_all
        self.uncheck_all_btn.Click += self._h_uncheck_all
        self.update_btn.Click += self._h_update
        self.close_btn.Click += self._h_close

    def _on_match(self, sender, e):
        self.candidates = _uv.collect_candidates_from_selection(
            self.doc, self.profile_data, self.selected_ids
        )
        self._render(self.candidates)
        self.summary_label.Text = "{} candidate(s); {} with divergence warnings".format(
            len(self.candidates),
            sum(1 for c in self.candidates if c.diverged_others),
        )
        self._set_status(
            "Review the list, uncheck to skip, then Update." if self.candidates
            else "No candidates from the current selection."
        )

    def _render(self, candidates):
        self.match_rows_panel.Children.Clear()
        self._match_rows = []
        for c in candidates:
            grid, cb = self._row(c)
            self.match_rows_panel.Children.Add(grid)
            self._match_rows.append(_MatchRow(c, cb))
        self.update_btn.IsEnabled = bool(candidates)

    def _row(self, c):
        grid = Grid()
        for w in (0.0, 1.0, 2.5, 3.0, 3.0, 1.5):
            col = ColumnDefinition()
            if w == 0.0:
                col.Width = GridLength(28)
            else:
                col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        cb = CheckBox()
        cb.IsChecked = True
        cb.Margin = Thickness(4, 2, 0, 2)
        cb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(cb, 0)
        grid.Children.Add(cb)

        kind_tb = TextBlock()
        kind_tb.Text = "[{}]".format(c.kind)
        kind_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(kind_tb, 1)
        grid.Children.Add(kind_tb)

        target_tb = TextBlock()
        target_tb.Text = c.led_id if c.kind == "led" else "{}  ({})".format(c.ann_id, c.led_id)
        target_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(target_tb, 2)
        grid.Children.Add(target_tb)

        old_tb = TextBlock()
        old_tb.Text = "Old: {}".format(self._format_offset(c.old_offset))
        old_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(old_tb, 3)
        grid.Children.Add(old_tb)

        new_tb = TextBlock()
        new_tb.Text = "New: {}".format(self._format_offset(c.new_offset))
        new_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(new_tb, 4)
        grid.Children.Add(new_tb)

        warn_tb = TextBlock()
        if c.diverged_others:
            warn_tb.Text = "{} diverged".format(len(c.diverged_others))
            warn_tb.Foreground = Brushes.OrangeRed
            warn_tb.ToolTip = "; ".join(
                "id {}: {}".format(eid, msg) for eid, msg in c.diverged_others
            )
        else:
            warn_tb.Text = ""
        warn_tb.Margin = Thickness(0, 4, 0, 4)
        Grid.SetColumn(warn_tb, 5)
        grid.Children.Add(warn_tb)

        return grid, cb

    def _format_offset(self, offset):
        return "x={:.2f} y={:.2f} z={:.2f} r={:.1f}°".format(
            offset.get("x_inches", 0.0),
            offset.get("y_inches", 0.0),
            offset.get("z_inches", 0.0),
            offset.get("rotation_deg", 0.0),
        )

    def _on_check_all(self, sender, e):
        for r in self._match_rows:
            r.checkbox.IsChecked = True

    def _on_uncheck_all(self, sender, e):
        for r in self._match_rows:
            r.checkbox.IsChecked = False

    def _on_update(self, sender, e):
        from pyrevit import revit
        import active_yaml
        chosen = []
        for r in self._match_rows:
            if not r.checked:
                continue
            r.candidate.skip = False
            chosen.append(r.candidate)
        if not chosen:
            self._set_status("Nothing checked")
            return
        with revit.Transaction("Update Vector (MEPRFP 2.0)", doc=self.doc):
            result = _uv.execute_update(self.doc, chosen)
            active_yaml.save_active_data(self.doc, self.profile_data, action="Update Vector")
        self.committed = True
        self._last_result = result
        self._set_status(
            "Updated {} LED(s), {} annotation(s); {} warning(s).".format(
                result.led_updates, result.ann_updates, len(result.warnings)
            )
        )

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data, selected_element_ids):
    return UpdateVectorController(doc, profile_data, selected_element_ids).show()
