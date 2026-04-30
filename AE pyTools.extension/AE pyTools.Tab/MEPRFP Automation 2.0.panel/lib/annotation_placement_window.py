# -*- coding: utf-8 -*-
"""
Place Element Annotations dialog.

Same overall shape as the Place from CAD or Linked Model dialog —
filters, match button, per-row preview with check, place button —
but the workflow underneath is annotation-specific:

    Active view  ->  candidates from collect_candidates()  ->  dedup
    pass  ->  preview list  ->  place checked.
"""

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

import annotation_placement
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources",
    "AnnotationPlacementWindow.xaml",
)


# ---------------------------------------------------------------------
# Per-row record (preview)
# ---------------------------------------------------------------------

class _MatchRow(object):
    def __init__(self, candidate, ui_grid, checkbox):
        self.candidate = candidate
        self.grid = ui_grid
        self.checkbox = checkbox

    @property
    def checked(self):
        return bool(self.checkbox.IsChecked)


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class AnnotationPlacementController(object):

    def __init__(self, doc, view, profile_data):
        self.doc = doc
        self.view = view
        self.profile_data = profile_data
        self.profiles = list(profile_data.get("equipment_definitions") or [])
        self.candidates = []
        self._match_rows = []
        self.committed = False
        self._last_result = None
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_filters()
        self._populate_active_view_label()
        self._wire_events()
        self._set_status("Tweak filters / kinds, then Match.")

    # ---- bootstrapping ---------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.active_view_label = f("ActiveViewLabel")
        self.kind_tag = f("KindTagCheck")
        self.kind_keynote = f("KindKeynoteCheck")
        self.kind_text_note = f("KindTextNoteCheck")
        self.category_list = f("CategoryList")
        self.profile_list = f("ProfileList")
        self.skip_dupes_check = f("SkipDupesCheck")
        self.match_btn = f("MatchButton")
        self.check_all_btn = f("CheckAllButton")
        self.uncheck_all_btn = f("UncheckAllButton")
        self.place_btn = f("PlaceButton")
        self.close_btn = f("CloseButton")
        self.summary_label = f("SummaryLabel")
        self.status_label = f("StatusLabel")
        self.match_rows_panel = f("MatchRowsPanel")

    def _make_handler(self, label, fn):
        def wrapped(sender, e):
            try:
                self._set_status("[{}] running...".format(label))
                fn(sender, e)
            except Exception as exc:
                self._set_status("[{}] error: {}".format(label, exc))
                raise
        return RoutedEventHandler(wrapped)

    def _wire_events(self):
        self._h_match = self._make_handler(
            "match", lambda s, e: self._on_match_clicked(s, e))
        self._h_check_all = self._make_handler(
            "check-all", lambda s, e: self._on_check_all(s, e))
        self._h_uncheck_all = self._make_handler(
            "uncheck-all", lambda s, e: self._on_uncheck_all(s, e))
        self._h_place = self._make_handler(
            "place", lambda s, e: self._on_place_clicked(s, e))
        self._h_close = self._make_handler(
            "close", lambda s, e: self.window.Close())
        self.match_btn.Click += self._h_match
        self.check_all_btn.Click += self._h_check_all
        self.uncheck_all_btn.Click += self._h_uncheck_all
        self.place_btn.Click += self._h_place
        self.close_btn.Click += self._h_close

    def _populate_filters(self):
        cats = sorted({
            (p.get("parent_filter") or {}).get("category") or ""
            for p in self.profiles
            if isinstance(p, dict)
        })
        cats = [c for c in cats if c]
        self.category_list.Items.Clear()
        for c in cats:
            self.category_list.Items.Add(c)

        self.profile_list.Items.Clear()
        for p in self.profiles:
            label = "{}  ({})".format(
                p.get("name") or "(unnamed)", p.get("id") or "?"
            )
            self.profile_list.Items.Add(label)

    def _populate_active_view_label(self):
        if self.view is None:
            self.active_view_label.Text = "Active view: (none)"
            return
        self.active_view_label.Text = "Active view: {} ({})".format(
            self.view.Name or "(unnamed)", self.view.ViewType
        )

    # ---- filter readers --------------------------------------------

    def _selected_kinds(self):
        out = set()
        if self.kind_tag.IsChecked:
            out.add(annotation_placement.KIND_TAG)
        if self.kind_keynote.IsChecked:
            out.add(annotation_placement.KIND_KEYNOTE)
        if self.kind_text_note.IsChecked:
            out.add(annotation_placement.KIND_TEXT_NOTE)
        return out

    def _selected_profile_ids(self):
        ids = set()
        for label in self.profile_list.SelectedItems:
            label = str(label)
            # Format: "name  (id)"
            if "(" in label and label.endswith(")"):
                pid = label.rsplit("(", 1)[1].rstrip(")")
                ids.add(pid)
        return ids or None

    def _selected_categories(self):
        out = {str(item) for item in self.category_list.SelectedItems}
        return out or None

    # ---- match -----------------------------------------------------

    def _on_match_clicked(self, sender, e):
        kinds = self._selected_kinds()
        if not kinds:
            self._set_status("Pick at least one kind (tag / keynote / text note)")
            return
        filters = annotation_placement.CollectFilters(
            kinds=kinds,
            profile_ids=self._selected_profile_ids(),
            categories=self._selected_categories(),
            active_view_only=True,
        )
        self.candidates = annotation_placement.collect_candidates(
            self.doc, self.view, self.profile_data, filters
        )
        if self.skip_dupes_check.IsChecked:
            annotation_placement.mark_duplicates(
                self.doc, self.view, self.candidates
            )
        self._render_matches(self.candidates)
        n_total = len(self.candidates)
        n_dupes = sum(1 for c in self.candidates if c.skip)
        self.summary_label.Text = (
            "{} candidate(s); {} flagged as already-placed".format(n_total, n_dupes)
        )
        self._set_status(
            "Review the list — uncheck any rows to skip, then Place." if n_total
            else "No candidates. Try different filters or pick a different view."
        )

    # ---- preview rendering -----------------------------------------

    def _clear_match_rows(self):
        self.match_rows_panel.Children.Clear()
        self._match_rows = []
        self.summary_label.Text = ""
        self.place_btn.IsEnabled = False

    def _render_matches(self, candidates):
        self._clear_match_rows()
        for c in candidates:
            grid, checkbox = self._build_match_row(c)
            self.match_rows_panel.Children.Add(grid)
            self._match_rows.append(_MatchRow(c, grid, checkbox))
        self.place_btn.IsEnabled = bool(candidates)

    def _build_match_row(self, candidate):
        grid = Grid()
        for w in (0.0, 1.5, 2.0, 4.0, 2.0):
            col = ColumnDefinition()
            if w == 0.0:
                col.Width = GridLength(28)
            else:
                col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        checkbox = CheckBox()
        checkbox.IsChecked = not candidate.skip
        checkbox.Margin = Thickness(4, 2, 0, 2)
        checkbox.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(checkbox, 0)
        grid.Children.Add(checkbox)

        kind_tb = TextBlock()
        kind_tb.Text = "[{}]".format(candidate.annotation.get("kind") or "?")
        kind_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(kind_tb, 1)
        grid.Children.Add(kind_tb)

        led_tb = TextBlock()
        led_tb.Text = "{}  ({})".format(
            candidate.led_label or "?", candidate.led_id or "?"
        )
        led_tb.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(led_tb, 2)
        grid.Children.Add(led_tb)

        ann_tb = TextBlock()
        ann_tb.Text = _annotation_display(candidate.annotation)
        ann_tb.Margin = Thickness(0, 4, 8, 4)
        ann_tb.TextTrimming = _text_trimming_ellipsis()
        ann_tb.ToolTip = ann_tb.Text  # full text on hover
        Grid.SetColumn(ann_tb, 3)
        grid.Children.Add(ann_tb)

        status_tb = TextBlock()
        if candidate.skip and candidate.duplicate_reason:
            status_tb.Text = "{}".format(candidate.duplicate_reason)
            status_tb.Foreground = self._gray_brush()
        else:
            status_tb.Text = "ready"
        status_tb.Margin = Thickness(0, 4, 0, 4)
        Grid.SetColumn(status_tb, 4)
        grid.Children.Add(status_tb)

        return grid, checkbox

    def _gray_brush(self):
        from System.Windows.Media import Brushes
        return Brushes.Gray

    def _on_check_all(self, sender, e):
        for row in self._match_rows:
            row.checkbox.IsChecked = True

    def _on_uncheck_all(self, sender, e):
        for row in self._match_rows:
            row.checkbox.IsChecked = False

    # ---- place ------------------------------------------------------

    def _on_place_clicked(self, sender, e):
        from pyrevit import revit
        chosen = []
        for row in self._match_rows:
            if not row.checked:
                continue
            row.candidate.skip = False  # in case dupe flag was overridden
            chosen.append(row.candidate)
        if not chosen:
            self._set_status("Nothing checked to place")
            return
        with revit.Transaction("Place Element Annotations (MEPRFP 2.0)", doc=self.doc):
            result = annotation_placement.execute_placement(self.doc, self.view, chosen)
        self.committed = True
        self._last_result = result
        self._set_status(
            "Placed {} (tags {}, keynotes {}, text notes {}); "
            "skipped {} dupes; {} warning(s).".format(
                result.total_placed,
                result.placed_count_by_kind.get("tag", 0),
                result.placed_count_by_kind.get("keynote", 0),
                result.placed_count_by_kind.get("text_note", 0),
                result.skipped_duplicates,
                len(result.warnings),
            )
        )

    # ---- misc -------------------------------------------------------

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def _text_trimming_ellipsis():
    """Resolve TextTrimming.CharacterEllipsis explicitly — pythonnet 3
    won't accept a bare int."""
    from System.Windows import TextTrimming
    return TextTrimming.CharacterEllipsis


# Names we recognise as carrying the keynote value / description.
_KEYNOTE_VALUE_KEYS = ("Keynote", "Keynote Number", "Keynote Value")
_KEYNOTE_DESC_KEYS = ("Keynote Description", "Description", "Note")


def _first_param(params, names):
    if not isinstance(params, dict):
        return None
    for n in names:
        if n in params and str(params[n]).strip():
            return params[n]
    return None


def _annotation_display(ann):
    """Build the cell text for an annotation in the preview list.

    * text_note  -> the actual text content (truncated for display).
    * keynote    -> ``Family : Type    [VALUE - DESCRIPTION]``, both
                    pulled from the captured parameters.
    * tag        -> ``Family : Type``.
    """
    if not isinstance(ann, dict):
        return ""
    kind = ann.get("kind") or ""
    family_name = ann.get("family_name") or ""
    type_name = ann.get("type_name") or ""
    base = (
        "{} : {}".format(family_name, type_name)
        if (family_name and type_name)
        else (ann.get("label") or "")
    )

    if kind == "text_note":
        text = ann.get("text") or ann.get("label") or ""
        if not text:
            return base or "(empty text note)"
        if len(text) > 120:
            return text[:117] + "..."
        return text

    if kind == "keynote":
        params = ann.get("parameters") or {}
        value = _first_param(params, _KEYNOTE_VALUE_KEYS)
        desc = _first_param(params, _KEYNOTE_DESC_KEYS)
        bits = []
        if value:
            bits.append(str(value))
        if desc:
            bits.append(str(desc))
        if bits:
            return "{}    [{}]".format(base, "  -  ".join(bits))
        return base

    return base


def show_modal(doc, view, profile_data):
    return AnnotationPlacementController(doc, view, profile_data).show()
