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
        labels = [
            "{}  ({})".format(p.get("name") or "(unnamed)", p.get("id") or "?")
            for p in self.profiles if isinstance(p, dict)
        ]
        labels.sort(key=lambda s: s.lower())
        for label in labels:
            self.profile_list.Items.Add(label)

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
        stats = _fp.FollowParentScanStats()
        try:
            self.candidates = _fp.collect_candidates(
                self.doc, self.profile_data, filters, refuse_linked=True, stats=stats
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
        self.summary_label.Text = (
            "{} candidate(s); {} already aligned   |   {}".format(
                n_total, n_aligned, stats.summary_line()
            )
        )
        if n_total:
            self._set_status("Review the list, uncheck to skip, then Follow.")
        else:
            # No candidates — surface the most likely cause inline.
            reason = self._diagnose_zero_candidates(stats)
            self._set_status(reason)

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

    def _diagnose_zero_candidates(self, stats):
        """Pick the most actionable explanation given the stats.
        Drives the status-bar message when Match returns nothing."""
        if stats.elements_scanned == 0:
            return "No FamilyInstance / Group elements found in the active document."
        if stats.no_element_linker == stats.elements_scanned:
            return (
                "{} element(s) scanned but none carry an Element_Linker — "
                "place fixtures via the panel before Follow Parent.".format(
                    stats.elements_scanned
                )
            )
        if stats.led_not_in_yaml and stats.led_not_in_yaml > stats.candidates_built:
            sample = ", ".join(stats.sample_orphan_led_ids[:5]) or "?"
            return (
                "{} placed fixture(s) reference led_ids that aren't in the "
                "current YAML (e.g. {}). Re-import the YAML version that "
                "matches when the fixtures were placed.".format(
                    stats.led_not_in_yaml, sample
                )
            )
        if stats.filtered_by_profile and not stats.candidates_built:
            # Show the top profiles the placed fixtures actually map to.
            top = sorted(
                stats.profile_matches.items(),
                key=lambda kv: -kv[1],
            )[:5]
            top_text = ", ".join(
                "{} ({}x)".format(pid, n) for pid, n in top
            ) or "(none)"
            return (
                "All {} fixtures with Element_Linker resolve to other "
                "profiles. Top matches: {}. Pick the profile that owns "
                "your placed fixtures, or uncheck the profile filter.".format(
                    stats.filtered_by_profile, top_text
                )
            )
        if stats.filtered_by_category and not stats.candidates_built:
            return (
                "All matched fixtures were excluded by the category filter."
            )
        if stats.parent_unresolved and not stats.candidates_built:
            return (
                "{} fixture(s) had Element_Linker but their parent element "
                "couldn't be resolved (linked CAD unloaded? parent deleted? "
                "host_name mismatch?).".format(stats.parent_unresolved)
            )
        return "No candidates matched the filters."

    def _on_check_all(self, sender, e):
        if not self._match_rows:
            self._set_status("No rows yet — click Match first")
            return
        for r in self._match_rows:
            r.checkbox.IsChecked = True
        self._set_status("Checked {} row(s)".format(len(self._match_rows)))

    def _on_uncheck_all(self, sender, e):
        if not self._match_rows:
            self._set_status("No rows yet — click Match first")
            return
        for r in self._match_rows:
            r.checkbox.IsChecked = False
        self._set_status("Unchecked {} row(s)".format(len(self._match_rows)))

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
