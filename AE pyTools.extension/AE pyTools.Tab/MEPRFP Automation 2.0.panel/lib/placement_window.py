# -*- coding: utf-8 -*-
"""
The Placement dialog controller.

The dialog is a single window with three sections:

    1. Source (radio): Linked Revit model / CSV / DWG link
       + a source-specific picker (combo or file browse)

    2. Filters (two list boxes): category multi-select + profile-name
       multi-select. Both default to "select none" = include all.

    3. Match preview: per-row check + (target, profile) labels. Match
       button populates this; user can toggle individual rows; Place
       button commits.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import Thickness, VerticalAlignment, Visibility  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402
from System.Windows.Controls import (  # noqa: E402
    CheckBox,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    TextBlock,
)
from System.Windows import GridUnitType, GridLength  # noqa: E402

import placement
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "PlacementWindow.xaml"
)


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

class _SourceItem(object):
    """Wraps a source-side selection (a RevitLinkInstance, an
    ImportInstance, or a CSV path) for display in the source combo."""

    def __init__(self, label, kind, value):
        self.label = label
        self.kind = kind
        self.value = value

    def __str__(self):
        return self.label


class _MatchRow(object):
    """Per-row UI state alongside a placement.Match."""

    def __init__(self, match, ui_grid, checkbox):
        self.match = match
        self.grid = ui_grid
        self.checkbox = checkbox

    @property
    def checked(self):
        return bool(self.checkbox.IsChecked)


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class PlacementController(object):

    def __init__(self, doc, profile_data):
        self.doc = doc
        self.profile_data = profile_data
        self.profiles = list(profile_data.get("equipment_definitions") or [])
        self.matches = []
        self._match_rows = []
        self._csv_path = None
        self._all_profile_labels = []        # alphabetically sorted, full list
        self._selected_profile_labels = set()  # survives search-filtering
        self._suppress_profile_selection = False
        self.committed = False
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_filters()
        self._wire_events()
        self._switch_source("host_model")
        self._set_status("Pick a source, optionally filter, then Match.")

    # ---- bootstrapping ---------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.src_host_radio = f("SrcHostRadio")
        self.src_linked_radio = f("SrcLinkedRevitRadio")
        self.src_csv_radio = f("SrcCsvRadio")
        self.src_dwg_radio = f("SrcDwgRadio")
        self.src_label = f("SrcLabel")
        self.src_combo = f("SrcCombo")
        self.src_browse_btn = f("SrcBrowseButton")
        self.category_list = f("CategoryList")
        self.profile_list = f("ProfileList")
        self.profile_search_box = f("ProfileSearchBox")
        self.skip_placed_check = f("SkipPlacedCheck")
        self.one_profile_per_target_check = f("OneProfilePerTargetCheck")
        self.allow_type_sub_check = f("AllowTypeSubCheck")
        self.match_btn = f("MatchButton")
        self.check_all_btn = f("CheckAllButton")
        self.uncheck_all_btn = f("UncheckAllButton")
        self.place_btn = f("PlaceButton")
        self.close_btn = f("CloseButton")
        self.summary_label = f("SummaryLabel")
        self.status_label = f("StatusLabel")
        self.match_rows_panel = f("MatchRowsPanel")

    def _make_delegate(self, label, fn):
        """Wrap a Python callable into a retained ``RoutedEventHandler``.

        Each handler:
            * is wrapped in try/except so a Python error becomes a status
              message instead of being silently swallowed by pythonnet,
            * writes ``[label] ...`` to the status label up front so we
              can tell from the UI whether the handler ever fires,
            * is converted to a ``RoutedEventHandler`` delegate explicitly
              (pythonnet 3's implicit conversion has been unreliable
              specifically for RoutedEventHandler in some builds),
            * is stored as an attribute so neither pythonnet nor Python's
              GC drops the delegate / target.
        """
        def wrapped(sender, e):
            try:
                self._set_status("[{}] running...".format(label))
                fn(sender, e)
            except Exception as exc:
                self._set_status("[{}] error: {}".format(label, exc))
                raise

        delegate = RoutedEventHandler(wrapped)
        return delegate

    def _wire_events(self):
        # Build + retain delegates for every event we subscribe to.
        self._h_src_host = self._make_delegate(
            "src=host", lambda s, e: self._switch_source("host_model"))
        self._h_src_linked = self._make_delegate(
            "src=linked", lambda s, e: self._switch_source("linked_revit"))
        self._h_src_csv = self._make_delegate(
            "src=csv", lambda s, e: self._switch_source("csv"))
        self._h_src_dwg = self._make_delegate(
            "src=dwg", lambda s, e: self._switch_source("dwg"))
        self._h_browse = self._make_delegate(
            "browse", lambda s, e: self._on_browse_clicked(s, e))
        self._h_match = self._make_delegate(
            "match", lambda s, e: self._on_match_clicked(s, e))
        self._h_check_all = self._make_delegate(
            "check-all", lambda s, e: self._on_check_all(s, e))
        self._h_uncheck_all = self._make_delegate(
            "uncheck-all", lambda s, e: self._on_uncheck_all(s, e))
        self._h_place = self._make_delegate(
            "place", lambda s, e: self._on_place_clicked(s, e))
        self._h_close = self._make_delegate(
            "close", lambda s, e: self.window.Close())

        self.src_host_radio.Checked += self._h_src_host
        self.src_linked_radio.Checked += self._h_src_linked
        self.src_csv_radio.Checked += self._h_src_csv
        self.src_dwg_radio.Checked += self._h_src_dwg
        self.src_browse_btn.Click += self._h_browse
        self.match_btn.Click += self._h_match
        self.check_all_btn.Click += self._h_check_all
        self.uncheck_all_btn.Click += self._h_uncheck_all
        self.place_btn.Click += self._h_place
        self.close_btn.Click += self._h_close

        # Profile search + selection-survival wiring. pythonnet wraps
        # the bound methods automatically; the bound-method instances are
        # kept alive by the class itself.
        self.profile_search_box.TextChanged += self._on_profile_search
        self.profile_list.SelectionChanged += self._on_profile_selection

    # ---- source handling -------------------------------------------

    def _switch_source(self, kind):
        """``kind`` in ('host_model', 'linked_revit', 'csv', 'dwg')."""
        self._source_kind = kind
        self.src_combo.Items.Clear()
        if kind == "host_model":
            self.src_label.Text = "Active document"
            self.src_browse_btn.Visibility = Visibility.Collapsed
            self.src_combo.IsEnabled = False
            self.src_combo.Items.Add(
                _SourceItem("(this document)", "host_model", self.doc)
            )
            self.src_combo.SelectedIndex = 0
        elif kind == "linked_revit":
            self.src_label.Text = "Linked model:"
            self.src_browse_btn.Visibility = Visibility.Collapsed
            self.src_combo.IsEnabled = True
            for inst in placement.collect_linked_revit_link_instances(self.doc):
                link_doc = inst.GetLinkDocument()
                title = getattr(link_doc, "Title", "") or "(unnamed)"
                self.src_combo.Items.Add(_SourceItem(title, "linked_revit", inst))
            if self.src_combo.Items.Count > 0:
                self.src_combo.SelectedIndex = 0
        elif kind == "csv":
            self.src_label.Text = "CSV file:"
            self.src_browse_btn.Visibility = Visibility.Visible
            self.src_combo.IsEnabled = True
            if self._csv_path:
                self.src_combo.Items.Add(
                    _SourceItem(self._csv_path, "csv", self._csv_path)
                )
                self.src_combo.SelectedIndex = 0
        elif kind == "dwg":
            self.src_label.Text = "DWG link:"
            self.src_browse_btn.Visibility = Visibility.Collapsed
            self.src_combo.IsEnabled = True
            for inst in placement.collect_dwg_link_instances(self.doc):
                category = inst.Category.Name if inst.Category else ""
                try:
                    name = inst.LookupParameter("Name")
                    label = name.AsString() if name else "(import)"
                except Exception:
                    label = "(import)"
                self.src_combo.Items.Add(_SourceItem(
                    "{} — {}".format(label, category), "dwg", inst
                ))
            if self.src_combo.Items.Count > 0:
                self.src_combo.SelectedIndex = 0
        # Clear preview when source changes.
        self._clear_match_rows()
        self._set_status("Source switched. Match again.")

    def _on_browse_clicked(self, sender, e):
        if self._source_kind != "csv":
            return
        # Reuse forms_compat.pick_file via the wpf path.
        import forms_compat as forms
        path = forms.pick_file(file_ext="csv", title="Pick rebased-coords CSV")
        if not path:
            return
        self._csv_path = path
        self.src_combo.Items.Clear()
        self.src_combo.Items.Add(_SourceItem(path, "csv", path))
        self.src_combo.SelectedIndex = 0

    # ---- filter population -----------------------------------------

    def _populate_filters(self):
        # Categories come from each LED's ``category`` field across every
        # profile — i.e. the categories of the *fixture children*, not
        # the profile's parent. Picking a category here keeps every
        # profile that contains *any* LED of that category, and
        # placement runs all of that profile's LEDs (not just the
        # matching ones).
        cats = set()
        for p in self.profiles:
            if not isinstance(p, dict):
                continue
            for s in p.get("linked_sets") or []:
                if not isinstance(s, dict):
                    continue
                for led in s.get("linked_element_definitions") or []:
                    if not isinstance(led, dict):
                        continue
                    c = (led.get("category") or "").strip()
                    if c:
                        cats.add(c)
        self.category_list.Items.Clear()
        for c in sorted(cats):
            self.category_list.Items.Add(c)

        labels = []
        for p in self.profiles:
            if not isinstance(p, dict):
                continue
            labels.append("{}  ({})".format(
                p.get("name") or "(unnamed)", p.get("id") or "?"
            ))
        labels.sort(key=lambda s: s.lower())
        self._all_profile_labels = labels
        self._render_profile_list("")

    def _render_profile_list(self, search_text):
        needle = (search_text or "").strip().lower()
        self._suppress_profile_selection = True
        try:
            self.profile_list.Items.Clear()
            visible = []
            for label in self._all_profile_labels:
                if needle and needle not in label.lower():
                    continue
                self.profile_list.Items.Add(label)
                visible.append(label)
            for label in visible:
                if label in self._selected_profile_labels:
                    self.profile_list.SelectedItems.Add(label)
        finally:
            self._suppress_profile_selection = False

    def _on_profile_search(self, sender, e):
        try:
            self._render_profile_list(self.profile_search_box.Text or "")
        except Exception as exc:
            self._set_status("[search] error: {}".format(exc))

    def _on_profile_selection(self, sender, e):
        if self._suppress_profile_selection:
            return
        try:
            for item in e.AddedItems:
                self._selected_profile_labels.add(str(item))
            for item in e.RemovedItems:
                self._selected_profile_labels.discard(str(item))
        except Exception as exc:
            self._set_status("[profile-select] error: {}".format(exc))

    def _filtered_profiles(self):
        selected_cats = {str(item) for item in self.category_list.SelectedItems}
        selected_names = {
            label.split("  (", 1)[0]
            for label in self._selected_profile_labels
        }
        out = []
        for p in self.profiles:
            if not isinstance(p, dict):
                continue
            if selected_cats:
                # Keep the profile if ANY of its LEDs is in a selected
                # category. Once kept, all of the profile's LEDs are
                # placement candidates — we don't filter LEDs by category.
                led_cats = set()
                for s in p.get("linked_sets") or []:
                    if not isinstance(s, dict):
                        continue
                    for led in s.get("linked_element_definitions") or []:
                        if not isinstance(led, dict):
                            continue
                        c = (led.get("category") or "").strip()
                        if c:
                            led_cats.add(c)
                if not (led_cats & selected_cats):
                    continue
            if selected_names:
                if (p.get("name") or "") not in selected_names:
                    continue
            out.append(p)
        return out

    # ---- match button ----------------------------------------------

    def _selected_source_value(self):
        item = self.src_combo.SelectedItem
        return item.value if item is not None else None

    def _on_match_clicked(self, sender, e):
        source_value = self._selected_source_value()
        if source_value is None and self._source_kind != "csv":
            self._set_status("Pick a source first")
            return

        targets = []
        mode = placement.MATCH_FAMILY_NAME_STRIP_SUFFIX
        if self._source_kind == "host_model":
            targets = placement.find_targets_in_host_model(self.doc)
            mode = placement.MATCH_FAMILY_NAME_STRIP_SUFFIX
        elif self._source_kind == "linked_revit":
            targets = placement.find_targets_in_linked_revit(source_value)
            mode = placement.MATCH_FAMILY_NAME_STRIP_SUFFIX
        elif self._source_kind == "csv":
            if not self._csv_path:
                self._set_status("Browse to a CSV first")
                return
            try:
                targets = placement.find_targets_in_csv(self._csv_path)
            except placement.CsvParseError as exc:
                self._set_status(str(exc))
                return
            mode = placement.MATCH_CAD_ALIASES
        elif self._source_kind == "dwg":
            if source_value is None:
                self._set_status("Pick a DWG link")
                return
            targets = placement.find_targets_in_dwg_link(source_value)
            mode = placement.MATCH_CAD_ALIASES

        profiles = self._filtered_profiles()
        if not profiles:
            self._set_status("No profiles match the current filters")
            self.matches = []
            self._render_matches([])
            return
        if not targets:
            self._set_status("No targets found in the selected source")
            self.matches = []
            self._render_matches([])
            return

        raw_matches = placement.match_targets(targets, profiles, mode)
        deduped_count = 0
        if self.one_profile_per_target_check.IsChecked:
            self.matches = placement.dedupe_matches_per_target(raw_matches)
            deduped_count = len(raw_matches) - len(self.matches)
        else:
            self.matches = raw_matches
        self.matches.sort(key=lambda m: (
            (m.profile.get("name") or "").lower(),
            (m.target.name or "").lower(),
        ))
        self._render_matches(self.matches)
        summary = "{} target(s) -> {} match(es) across {} profile(s)".format(
            len(targets), len(self.matches), len(profiles),
        )
        if deduped_count:
            summary += "  ({} duplicate match(es) suppressed)".format(deduped_count)
        self.summary_label.Text = summary
        self._set_status(
            "Review the list, uncheck rows to skip, then Place." if self.matches
            else "No matches. Try different filters or a different source."
        )

    # ---- preview rendering -----------------------------------------

    def _clear_match_rows(self):
        self.match_rows_panel.Children.Clear()
        self._match_rows = []
        self.summary_label.Text = ""
        self.place_btn.IsEnabled = False

    def _render_matches(self, matches):
        self._clear_match_rows()
        for match in matches:
            grid, checkbox = self._build_match_row(match)
            self.match_rows_panel.Children.Add(grid)
            self._match_rows.append(_MatchRow(match, grid, checkbox))
        self.place_btn.IsEnabled = bool(matches)

    def _build_match_row(self, match):
        grid = Grid()
        for star in (0.0, 4.0, 4.0, 3.0):
            col = ColumnDefinition()
            if star == 0.0:
                col.Width = GridLength(28)
            else:
                col.Width = GridLength(star, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        checkbox = CheckBox()
        checkbox.IsChecked = True
        checkbox.Margin = Thickness(4, 2, 0, 2)
        checkbox.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(checkbox, 0)
        grid.Children.Add(checkbox)

        target_label = TextBlock()
        target_label.Text = "{}  @ ({:.2f}, {:.2f}, {:.2f})  rot {:.1f}°".format(
            match.target.name,
            match.target.world_pt[0], match.target.world_pt[1], match.target.world_pt[2],
            match.target.rotation_deg,
        )
        target_label.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(target_label, 1)
        grid.Children.Add(target_label)

        arrow = TextBlock()
        arrow.Text = "  ->  "
        arrow.Margin = Thickness(0, 4, 0, 4)
        Grid.SetColumn(arrow, 2)
        grid.Children.Add(arrow)

        profile_label = TextBlock()
        profile_label.Text = "{}  ({})".format(
            match.profile.get("name") or "(unnamed)",
            match.profile.get("id") or "?",
        )
        profile_label.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(profile_label, 3)
        grid.Children.Add(profile_label)

        return grid, checkbox

    def _on_check_all(self, sender, e):
        for row in self._match_rows:
            row.checkbox.IsChecked = True

    def _on_uncheck_all(self, sender, e):
        for row in self._match_rows:
            row.checkbox.IsChecked = False

    # ---- place ------------------------------------------------------

    def _on_place_clicked(self, sender, e):
        from pyrevit import revit
        import active_yaml

        chosen = []
        for row in self._match_rows:
            if not row.checked:
                continue
            chosen.append(row.match)
        if not chosen:
            self._set_status("Nothing checked to place")
            return

        options = placement.PlacementOptions(
            skip_already_placed=bool(self.skip_placed_check.IsChecked),
            allow_type_substitution=bool(self.allow_type_sub_check.IsChecked),
        )
        with revit.Transaction("Place from CAD or Linked Model (MEPRFP 2.0)", doc=self.doc):
            result = placement.execute_placement(self.doc, chosen, options)

        self.committed = True
        self._last_result = result
        self._set_status(
            "Placed {} fixture(s); skipped {} already-placed; {} warning(s).".format(
                result.placed_fixture_count,
                result.skipped_already_placed,
                len(result.warnings),
            )
        )

    # ---- misc -------------------------------------------------------

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data):
    return PlacementController(doc, profile_data).show()
