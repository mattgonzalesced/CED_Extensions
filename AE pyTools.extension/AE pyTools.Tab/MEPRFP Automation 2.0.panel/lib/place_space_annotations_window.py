# -*- coding: utf-8 -*-
"""
Modal preview UI for Place Space Annotations.

Renders the candidate annotations from
``space_annotation_workflow.collect_space_candidates`` and commits
the lot in one Revit transaction on Place all. Mirrors the equipment
annotation tool but explicitly scoped to space-based fixtures (those
stamped with ``Element_Linker.space_id`` by Place Space Elements).
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402

from pyrevit import revit  # noqa: E402

import annotation_placement as _ap  # noqa: E402
import space_annotation_workflow as _saw  # noqa: E402
import wpf as _wpf  # noqa: E402


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "PlaceSpaceAnnotationsWindow.xaml",
)


# ---------------------------------------------------------------------
# Row binding object
# ---------------------------------------------------------------------

class _PreviewRow(object):

    def __init__(self, candidate):
        self.candidate = candidate

    @property
    def KindLabel(self):
        return self.candidate.annotation.get("kind") or "?"

    @property
    def ProfileLabel(self):
        return "{} ({})".format(
            self.candidate.profile_name or "?", self.candidate.profile_id or "?"
        )

    @property
    def LedLabel(self):
        return self.candidate.led_label or self.candidate.led_id or ""

    @property
    def AnnLabel(self):
        ann = self.candidate.annotation or {}
        return ann.get("label") or ann.get("text") or ann.get("keynote_id") or ""

    @property
    def SkipReason(self):
        return self.candidate.duplicate_reason if self.candidate.skip else ""

    @property
    def XText(self):
        return _fmt(self.candidate.target_pt[0]) if self.candidate.target_pt else ""

    @property
    def YText(self):
        return _fmt(self.candidate.target_pt[1]) if self.candidate.target_pt else ""

    @property
    def FixtureIdText(self):
        f = self.candidate.fixture
        if f is None:
            return ""
        eid = getattr(f, "Id", None)
        if eid is None:
            return ""
        for attr in ("Value", "IntegerValue"):
            v = getattr(eid, attr, None)
            if v is not None:
                return str(v)
        return ""


def _fmt(v):
    try:
        return "{:.3f}".format(float(v))
    except (ValueError, TypeError):
        return ""


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class PlaceSpaceAnnotationsController(object):

    def __init__(self, doc, view, profile_data=None):
        self.doc = doc
        self.view = view
        self.profile_data = profile_data or {}
        self.committed = False
        self.last_result = None

        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = ObservableCollection[_NetObject]()

        self._lookup_controls()
        self._wire_events()
        self._refresh()

    def _lookup_controls(self):
        f = self.window.FindName
        self.summary_label = f("SummaryLabel")
        self.preview_grid = f("PreviewGrid")
        self.refresh_btn = f("RefreshButton")
        self.place_btn = f("PlaceButton")
        self.close_btn = f("CloseButton")
        self.status_label = f("StatusLabel")
        self.kind_tag_check = f("KindTagCheck")
        self.kind_keynote_check = f("KindKeynoteCheck")
        self.kind_text_check = f("KindTextNoteCheck")
        self.skip_dupes_check = f("SkipDuplicatesCheck")
        self.preview_grid.ItemsSource = self._rows

    def _wire_events(self):
        self._h_refresh = RoutedEventHandler(
            lambda s, e: self._safe(self._refresh, "refresh")
        )
        self._h_place = RoutedEventHandler(
            lambda s, e: self._safe(self._on_place, "place")
        )
        self._h_close = RoutedEventHandler(
            lambda s, e: self.window.Close()
        )
        self.refresh_btn.Click += self._h_refresh
        self.place_btn.Click += self._h_place
        self.close_btn.Click += self._h_close

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def _selected_kinds(self):
        kinds = set()
        if self.kind_tag_check.IsChecked:
            kinds.add(_ap.KIND_TAG)
        if self.kind_keynote_check.IsChecked:
            kinds.add(_ap.KIND_KEYNOTE)
        if self.kind_text_check.IsChecked:
            kinds.add(_ap.KIND_TEXT_NOTE)
        return kinds

    # ----- pipeline ------------------------------------------------

    def _refresh(self):
        kinds = self._selected_kinds()
        if not kinds:
            self._set_status("Tick at least one annotation kind.")
            self._rows.Clear()
            self.summary_label.Text = ""
            return

        skip_dupes = bool(self.skip_dupes_check.IsChecked)
        candidates = _saw.collect_space_candidates(
            self.doc, self.view, self.profile_data,
            kinds=kinds,
            skip_duplicates=skip_dupes,
        )
        self._rows.Clear()
        for c in candidates:
            self._rows.Add(_PreviewRow(c))
        n_total = len(candidates)
        n_skipped = sum(1 for c in candidates if c.skip)
        self.summary_label.Text = "{} candidate(s); {} flagged as already-placed".format(
            n_total, n_skipped,
        )
        self._set_status(
            "Ready. Click 'Place all' to commit." if n_total
            else "No space-based fixtures with annotations in this view."
        )

    def _on_place(self):
        if not self._rows.Count:
            self._set_status("Nothing to place.")
            return
        candidates = [r.candidate for r in self._rows]
        self.place_btn.IsEnabled = False
        self.refresh_btn.IsEnabled = False
        try:
            with revit.Transaction("Place Space Annotations (MEPRFP 2.0)", doc=self.doc):
                self.last_result = _saw.execute_placement(
                    self.doc, self.view, candidates,
                )
            self.committed = True
        finally:
            self.place_btn.IsEnabled = True
            self.refresh_btn.IsEnabled = True

        result = self.last_result
        n_placed = sum(result.placed_count_by_kind.values()) if result else 0
        n_warns = len(result.warnings) if result else 0
        self._set_status(
            "Placed {}; {} warning(s).".format(n_placed, n_warns)
        )


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

def show_modal(doc, view, profile_data=None):
    controller = PlaceSpaceAnnotationsController(
        doc=doc, view=view, profile_data=profile_data,
    )
    controller.window.ShowDialog()
    return controller
