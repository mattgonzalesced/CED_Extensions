# -*- coding: utf-8 -*-
"""
Modal dialog for QAQC.

The window stays open across "Refresh" — the audit re-runs in place
without closing/reopening. The Misc Ops > QAQC pushbutton only opens
the modal; once closed, the only way to reopen it is to click the
pushbutton again. Per-row Select / Zoom / Fix buttons act on the doc
through pyrevit / Revit API calls; each Fix wraps in its own
transaction so partial fixes commit even if a later row fails.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Collections.Generic import List as _NetList  # noqa: E402
from System.Windows import (  # noqa: E402
    GridLength,
    GridUnitType,
    RoutedEventHandler,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (  # noqa: E402
    Button,
    CheckBox,
    ColumnDefinition,
    Grid,
    TextBlock,
)
from System.Windows.Media import Brushes  # noqa: E402

from Autodesk.Revit.DB import ElementId  # noqa: E402

import qaqc_workflow as _qa
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "QaqcWindow.xaml",
)


class QaqcController(object):

    def __init__(self, doc, profile_data, uidoc=None):
        self.doc = doc
        self.uidoc = uidoc
        self.profile_data = profile_data
        self._cat_filter_checks = {}
        self._handlers = []  # retain refs against pythonnet GC
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_category_filters()
        self._wire_events()
        self._set_status("Click Refresh to run the audit.")

    # ----------------------------------------------------------------
    # Setup
    # ----------------------------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.cat_filter_panel = f("CategoryFilterPanel")
        self.summary_label = f("SummaryLabel")
        self.refresh_btn = f("RefreshButton")
        self.findings_panel = f("FindingsPanel")
        self.status_label = f("StatusLabel")
        self.close_btn = f("CloseButton")

    def _populate_category_filters(self):
        self.cat_filter_panel.Children.Clear()
        self._cat_filter_checks = {}
        for cat in _qa.CAT_ALL:
            cb = CheckBox()
            cb.Content = _qa.CAT_LABELS[cat]
            cb.IsChecked = True
            cb.Margin = Thickness(0, 0, 12, 0)
            cb.VerticalAlignment = VerticalAlignment.Center
            self.cat_filter_panel.Children.Add(cb)
            self._cat_filter_checks[cat] = cb

    def _wire_events(self):
        self._h_refresh = self._delegate("refresh", lambda s, e: self._on_refresh(s, e))
        self._h_close = self._delegate("close", lambda s, e: self.window.Close())
        self.refresh_btn.Click += self._h_refresh
        self.close_btn.Click += self._h_close

    def _delegate(self, label, fn):
        def wrapped(s, e):
            try:
                fn(s, e)
            except Exception as exc:
                self._set_status("[{}] error: {}".format(label, exc))
                raise
        return RoutedEventHandler(wrapped)

    # ----------------------------------------------------------------
    # Refresh
    # ----------------------------------------------------------------

    def _selected_categories(self):
        return [c for c, cb in self._cat_filter_checks.items() if cb.IsChecked]

    def _on_refresh(self, sender, e):
        cats = self._selected_categories()
        if not cats:
            self._set_status("Select at least one category.")
            self.findings_panel.Children.Clear()
            self.summary_label.Text = ""
            return
        self._set_status("Auditing...")
        result = _qa.run_audit(self.doc, self.profile_data, categories=cats)
        self._render(result)
        parts = []
        for cat in _qa.CAT_ALL:
            n = result.counts.get(cat, 0)
            if n:
                parts.append("{}={}".format(cat, n))
        if parts:
            self.summary_label.Text = "Findings: " + ", ".join(parts)
            self._set_status(
                "{} finding(s). Use the row buttons to Select / Zoom / Fix.".format(
                    len(result.findings)
                )
            )
        else:
            self.summary_label.Text = "Clean — no findings."
            self._set_status("Audit complete; nothing to fix.")

    def _render(self, result):
        self.findings_panel.Children.Clear()
        # Group by category for visual order.
        by_cat = {}
        for f in result.findings:
            by_cat.setdefault(f.category, []).append(f)
        for cat in _qa.CAT_ALL:
            findings = by_cat.get(cat) or []
            if not findings:
                continue
            self.findings_panel.Children.Add(
                self._section_header(cat, len(findings))
            )
            for finding in findings:
                self.findings_panel.Children.Add(self._row(finding))

    def _section_header(self, cat, count):
        from System.Windows import FontWeights
        tb = TextBlock()
        tb.Text = "{}   ({} finding{})".format(
            _qa.CAT_LABELS[cat], count, "" if count == 1 else "s"
        )
        tb.FontWeight = FontWeights.Bold
        tb.Margin = Thickness(0, 8, 0, 4)
        return tb

    def _row(self, finding):
        grid = Grid()
        for w in (0.0, 2.5, 4.0, 0.0, 0.0, 0.0):
            col = ColumnDefinition()
            if w == 0.0:
                col.Width = GridLength(80)  # button columns get fixed width
            else:
                col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        target_tb = TextBlock()
        target_tb.Margin = Thickness(0, 4, 8, 4)
        target_tb.VerticalAlignment = VerticalAlignment.Center
        if finding.element_id is not None:
            target_tb.Text = "id {}  | {}  ({})".format(
                finding.element_id,
                finding.profile_name or "?",
                finding.led_label or finding.profile_id or "?",
            )
        else:
            target_tb.Text = "{}  ({})".format(
                finding.profile_name or "?", finding.profile_id or "?"
            )
        Grid.SetColumn(target_tb, 1)
        grid.Children.Add(target_tb)

        msg_tb = TextBlock()
        msg_tb.Text = finding.message
        msg_tb.Margin = Thickness(0, 4, 8, 4)
        msg_tb.VerticalAlignment = VerticalAlignment.Center
        msg_tb.Foreground = Brushes.Gray
        Grid.SetColumn(msg_tb, 2)
        grid.Children.Add(msg_tb)

        select_btn = self._small_button("Select")
        select_btn.IsEnabled = finding.element_id is not None
        Grid.SetColumn(select_btn, 3)
        grid.Children.Add(select_btn)

        zoom_btn = self._small_button("Zoom")
        zoom_btn.IsEnabled = finding.element_id is not None
        Grid.SetColumn(zoom_btn, 4)
        grid.Children.Add(zoom_btn)

        fix_btn = self._small_button("Fix")
        fix_btn.IsEnabled = finding.fix_kind != _qa.FIX_NONE
        Grid.SetColumn(fix_btn, 5)
        grid.Children.Add(fix_btn)

        # Closure-captured handlers — retain refs so pythonnet doesn't GC them.
        h_select = RoutedEventHandler(
            lambda s, e, fnd=finding: self._on_select(fnd)
        )
        h_zoom = RoutedEventHandler(
            lambda s, e, fnd=finding: self._on_zoom(fnd)
        )
        h_fix = RoutedEventHandler(
            lambda s, e, fnd=finding, btn=fix_btn: self._on_fix(fnd, btn)
        )
        select_btn.Click += h_select
        zoom_btn.Click += h_zoom
        fix_btn.Click += h_fix
        self._handlers.extend([h_select, h_zoom, h_fix])

        return grid

    def _small_button(self, text):
        btn = Button()
        btn.Content = text
        btn.Margin = Thickness(2, 2, 2, 2)
        btn.MinWidth = 70
        return btn

    # ----------------------------------------------------------------
    # Row actions
    # ----------------------------------------------------------------

    def _on_select(self, finding):
        if self.uidoc is None or finding.element_id is None:
            self._set_status("No active uidoc; cannot select.")
            return
        try:
            ids = _NetList[ElementId]()
            ids.Add(ElementId(int(finding.element_id)))
            self.uidoc.Selection.SetElementIds(ids)
            self._set_status("Selected element {}.".format(finding.element_id))
        except Exception as exc:
            self._set_status("Select failed: {}".format(exc))

    def _on_zoom(self, finding):
        if self.uidoc is None or finding.element_id is None:
            self._set_status("No active uidoc; cannot zoom.")
            return
        try:
            ids = _NetList[ElementId]()
            ids.Add(ElementId(int(finding.element_id)))
            self.uidoc.Selection.SetElementIds(ids)
            self.uidoc.ShowElements(ids)
            self._set_status("Zoomed to element {}.".format(finding.element_id))
        except Exception as exc:
            self._set_status("Zoom failed: {}".format(exc))

    def _on_fix(self, finding, btn):
        from pyrevit import revit
        if finding.fix_kind == _qa.FIX_NONE:
            self._set_status("No automated fix for this category.")
            return
        try:
            with revit.Transaction("QAQC fix {} ({})".format(
                    finding.category, finding.fix_kind), doc=self.doc):
                ok, msg = _qa.execute_fix(self.doc, self.profile_data, finding)
        except Exception as exc:
            self._set_status("Fix raised: {}".format(exc))
            return
        if ok:
            btn.IsEnabled = False
            btn.Content = "Fixed"
            self._set_status("[{}] {}".format(finding.category, msg))
        else:
            self._set_status("[{}] fix failed: {}".format(finding.category, msg))

    # ----------------------------------------------------------------
    # Misc
    # ----------------------------------------------------------------

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data, uidoc=None):
    return QaqcController(doc, profile_data, uidoc=uidoc).show()
