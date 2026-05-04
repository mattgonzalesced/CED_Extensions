# -*- coding: utf-8 -*-
"""
Modal preview UI for the Place Space Elements pushbutton.

Renders every placement plan from the workflow as a flat row, then
hands the lot to ``space_apply.apply_plans`` when the user clicks
*Place all*. Status (placed / failed / skipped) lands back into each
row so the user can see in-line which Family/Type names didn't
resolve.

Modal — NOT modeless. Spaces placement runs inside a single
transaction kicked off from a button click while the dialog still
holds the API context, so the ExternalEvent gateway used by
SuperCircuit isn't necessary here.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402

import wpf as _wpf  # noqa: E402
import space_placement_workflow as _spw  # noqa: E402


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "PlaceSpaceElementsWindow.xaml",
)


# ---------------------------------------------------------------------
# Row binding object
# ---------------------------------------------------------------------

class _PreviewRow(object):
    """One DataGrid row.

    Plain Python so WPF binds attributes directly. The ``plan`` and
    ``status`` fields are mutated by the controller during apply.
    """

    def __init__(self, plan):
        self.plan = plan
        self.Status = "Pending"

    @property
    def SpaceLabel(self):
        s = self.plan.space
        if s is None:
            return ""
        bits = []
        if s.number:
            bits.append(s.number)
        if s.name:
            bits.append(s.name)
        return " - ".join(bits) or "(unnamed)"

    @property
    def ProfileLabel(self):
        p = self.plan.profile
        if p is None:
            return ""
        return "{}  ({})".format(p.name or "(unnamed)", p.id or "??")

    @property
    def Label(self):
        return self.plan.label or ""

    @property
    def KindLabel(self):
        if self.plan.led is None:
            return ""
        return self.plan.led.placement_rule.kind

    @property
    def XText(self):
        return _fmt_float(self.plan.world_pt[0]) if self.plan.world_pt else ""

    @property
    def YText(self):
        return _fmt_float(self.plan.world_pt[1]) if self.plan.world_pt else ""

    @property
    def ZText(self):
        return _fmt_float(self.plan.world_pt[2]) if self.plan.world_pt else ""

    @property
    def RotText(self):
        return _fmt_float(self.plan.rotation_deg)


def _fmt_float(v):
    try:
        f = float(v)
    except (ValueError, TypeError):
        return ""
    return "{:.3f}".format(f)


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class PlaceSpaceElementsController(object):

    def __init__(self, doc, profile_data=None):
        self.doc = doc
        self.profile_data = profile_data or {}

        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = ObservableCollection[_NetObject]()
        self.run = _spw.SpacePlacementRun(doc=doc, profile_data=self.profile_data)

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

    # ----- pipeline ------------------------------------------------

    def _refresh(self):
        self._set_status("Collecting placement plans...")
        plans = self.run.collect()
        self._rows.Clear()
        for plan in plans:
            self._rows.Add(_PreviewRow(plan))
        n_plans = len(plans)
        n_warns = len(self.run.warnings)
        self.summary_label.Text = "{} planned placement(s); {} warning(s)".format(
            n_plans, n_warns,
        )
        if n_plans == 0:
            extra = "  (See output panel for details.)" if n_warns else ""
            self._set_status(
                "Nothing to place." + extra
            )
        else:
            self._set_status(
                "Ready. Click 'Place all' to commit. {} warning(s).".format(n_warns)
            )

    def _on_place(self):
        if not self._rows.Count:
            self._set_status("Nothing to place.")
            return
        self._set_status("Placing... (one transaction)")
        # Disable the buttons during the run to avoid a double-click.
        self.place_btn.IsEnabled = False
        self.refresh_btn.IsEnabled = False
        try:
            result = self.run.apply()
        finally:
            self.place_btn.IsEnabled = True
            self.refresh_btn.IsEnabled = True
        # Map plan -> row for status writeback.
        plan_to_row = {id(r.plan): r for r in self._rows}
        for plan, _elem in result.placed:
            row = plan_to_row.get(id(plan))
            if row is not None:
                row.Status = "Placed"
        for plan, status, info in result.failed:
            row = plan_to_row.get(id(plan))
            if row is None:
                continue
            if status == "family_missing":
                row.Status = "Family missing: {}".format(info.get("requested_family"))
            elif status == "type_missing":
                row.Status = "Type missing under {}".format(info.get("requested_family"))
            elif status == "no_label":
                row.Status = "No label"
            elif status == "create_failed":
                row.Status = "Create failed"
            elif status == "exception":
                row.Status = "Exception: {}".format(info.get("message", ""))
            else:
                row.Status = status

        self.preview_grid.Items.Refresh()
        self._set_status(
            "Done. Placed {} / Failed {}.".format(result.n_placed, result.n_failed)
        )


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

def show_modal(doc, profile_data=None):
    controller = PlaceSpaceElementsController(doc=doc, profile_data=profile_data)
    controller.window.ShowDialog()
    return controller
