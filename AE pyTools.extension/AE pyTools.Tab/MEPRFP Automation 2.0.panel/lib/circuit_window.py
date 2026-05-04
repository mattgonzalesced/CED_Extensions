# -*- coding: utf-8 -*-
"""
Modeless preview UI for SuperCircuit.

Stays open while the user can pan / pick / inspect in Revit. Edits to
panel / circuit / load apply via the *Apply edits* button — that
re-runs the grouping pipeline against the current items so the row
list reflects the new bucketing without re-walking the doc. *Run
circuits* hands the prepared groups to ``CircuitApplyGateway.request_apply``
which hops to Revit's main thread via ExternalEvent and creates the
systems inside one transaction.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.Generic import List as _NetList  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import (  # noqa: E402
    RoutedEventHandler,
)

from Autodesk.Revit.DB import ElementId  # noqa: E402

import circuit_grouping as _grouping
import circuit_workflow as _workflow
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "SuperCircuitWindow.xaml",
)


# ---------------------------------------------------------------------
# Row binding object
# ---------------------------------------------------------------------

class _PreviewRow(object):
    """One DataGrid row. Plain Python so WPF binding reads/writes via
    attribute access; the controller polls these on Apply edits."""

    def __init__(self, group, item, group_index, group_label,
                 panel_options, circuit_options, load_options):
        self._group = group
        self._item = item
        self.GroupIndex = group_index
        self.GroupLabel = group_label
        self.PanelName = item.effective_panel
        self.CircuitNumber = item.effective_circuit_token
        self.LoadName = item.effective_load_name
        self.FamilyType = "{} : {}".format(item.family_name, item.type_name).strip(" :")
        self.ElementIdText = str(item.element_id) if item.element_id is not None else ""
        self.PanelOptions = list(panel_options or [])
        self.CircuitOptions = list(circuit_options or [])
        self.LoadOptions = list(load_options or [])

    @property
    def item(self):
        return self._item

    @property
    def group(self):
        return self._group


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class SuperCircuitController(object):

    def __init__(self, doc, uidoc, client, scope=_workflow.SCOPE_ALL,
                 selected_element_ids=None, profile_data=None):
        self.doc = doc
        self.uidoc = uidoc
        self.run = _workflow.CircuitRun(
            doc=doc,
            client=client,
            scope=scope,
            selected_element_ids=selected_element_ids,
            profile_data=profile_data,
        )
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = ObservableCollection[_NetObject]()
        self._lookup_controls()
        self._wire_events()
        self._set_status(
            "Click Refresh from doc to collect items, or wait — auto-loading..."
        )
        self._populate_header()

    # ----- bootstrapping -------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.header_label = f("HeaderLabel")
        self.summary_label = f("SummaryLabel")
        self.keyword_combo = f("KeywordCombo")
        self.refresh_btn = f("RefreshButton")
        self.preview_grid = f("PreviewGrid")
        self.apply_edits_btn = f("ApplyEditsButton")
        self.run_btn = f("RunButton")
        self.close_btn = f("CloseButton")
        self.status_label = f("StatusLabel")
        self.preview_grid.ItemsSource = self._rows

    def _wire_events(self):
        # All retained as attributes so pythonnet doesn't GC the wrappers.
        self._h_refresh = self._delegate("refresh", self._on_refresh)
        self._h_apply = self._delegate("apply-edits", self._on_apply_edits)
        self._h_run = self._delegate("run", self._on_run)
        self._h_close = self._delegate("close", lambda s, e: self.window.Close())
        self._h_keyword = self._delegate("keyword", self._on_keyword_changed)
        self.refresh_btn.Click += self._h_refresh
        self.apply_edits_btn.Click += self._h_apply
        self.run_btn.Click += self._h_run
        self.close_btn.Click += self._h_close
        self.keyword_combo.SelectionChanged += self._h_keyword

    def _delegate(self, label, fn):
        def wrapped(s, e):
            try:
                fn(s, e)
            except Exception as exc:
                self._set_status("[{}] error: {}".format(label, exc))
                raise
        return RoutedEventHandler(wrapped)

    def _populate_header(self):
        client = self.run.client
        scope_text = "selection" if self.run.scope == _workflow.SCOPE_SELECTION else "all eligible"
        self.header_label.Text = "Client: {}   |   Scope: {}".format(
            client.display_name or client.key or "?", scope_text
        )

    # ----- pipeline -------------------------------------------------

    def _refresh(self):
        self._set_status("Collecting items...")
        self.run.collect()
        self.run.assemble()
        self._populate_keyword_combo()
        self._render_rows()
        n_items = len(self.run.items)
        n_groups = len(self.run.groups)
        if n_items == 0:
            self._set_status(
                "No eligible electrical / data / control fixtures found "
                "in this document. (Lighting fixtures and devices are "
                "intentionally skipped.)"
            )
        else:
            self._set_status(
                "{} item(s) -> {} group(s). Edit panel / circuit / load, "
                "then Apply edits + Run circuits.".format(n_items, n_groups)
            )

    def _on_refresh(self, sender, e):
        self._refresh()

    def _populate_keyword_combo(self):
        self.keyword_combo.Items.Clear()
        # Add an "(all)" sentinel so the user can clear the filter.
        self.keyword_combo.Items.Add("(all)")
        present = set()
        for g in self.run.groups:
            tok = (g.circuit_token or "").strip().upper()
            if tok:
                present.add(tok)
        options = self.run.client.run_keyword_options(present_in_groups=present)
        for opt in options:
            self.keyword_combo.Items.Add(opt)
        self.keyword_combo.SelectedIndex = 0

    def _on_keyword_changed(self, sender, e):
        self._render_rows()

    def _selected_keyword(self):
        item = self.keyword_combo.SelectedItem
        if item is None:
            return ""
        text = str(item).strip()
        if not text or text == "(all)":
            return ""
        return text.upper()

    def _render_rows(self):
        keyword = self._selected_keyword()
        groups = self.run.groups
        if keyword:
            groups = _workflow.filter_groups_by_keyword(
                groups, keyword, self.run.client
            )

        # Build fast option-lookup for the editable combos.
        panel_options = self._panel_options()
        circuit_options = self._circuit_options(groups)
        load_options = self._load_options(groups)

        self._rows.Clear()
        for idx, group in enumerate(groups, start=1):
            label = self._group_label(idx, group)
            for member in group.members:
                row = _PreviewRow(
                    group=group,
                    item=member,
                    group_index=idx,
                    group_label=label,
                    panel_options=panel_options,
                    circuit_options=circuit_options,
                    load_options=load_options,
                )
                self._rows.Add(row)

        n_rev = sum(1 for g in groups if g.needs_review)
        self.summary_label.Text = (
            "{} group(s), {} member(s); {} needs-review".format(
                len(groups), self._rows.Count, n_rev
            )
        )

    def _group_label(self, idx, group):
        bits = ["#{}".format(idx)]
        if group.bucket and group.bucket != _grouping.BUCKET_NORMAL:
            bits.append("[{}]".format(group.bucket.upper()))
        if group.panel_name:
            bits.append("panel={}".format(group.panel_name))
        if group.circuit_token:
            bits.append("ckt={}".format(group.circuit_token))
        if group.load_name:
            bits.append("load={}".format(group.load_name))
        bits.append("({} member{})".format(
            group.member_count, "" if group.member_count == 1 else "s"
        ))
        if group.needs_review:
            bits.append("[NEEDS REVIEW]")
        return "  ".join(bits)

    def _panel_options(self):
        return sorted({
            (panel_elem.Name or "")
            for key, panel_elem in (self.run.panel_index or {}).items()
            if panel_elem is not None
        }, key=lambda s: s.lower())

    def _circuit_options(self, groups):
        seen = set()
        for g in groups:
            tok = (g.circuit_token or "").strip()
            if tok:
                seen.add(tok)
        # Always include the universal tokens.
        for t in ("DEDICATED", "BYPARENT", "SECONDBYPARENT"):
            seen.add(t)
        # Numeric circuit numbers come first; alphabetic tokens after.
        return sorted(seen, key=lambda s: (not s.isdigit(), s))

    def _load_options(self, groups):
        seen = set()
        for g in groups:
            for m in g.members:
                ln = (m.effective_load_name or "").strip()
                if ln:
                    seen.add(ln)
        return sorted(seen, key=lambda s: s.lower())

    # ----- apply-edits / run --------------------------------------

    def _on_apply_edits(self, sender, e):
        try:
            self.preview_grid.CommitEdit()
            self.preview_grid.CommitEdit()
        except Exception:
            pass
        # Push row.PanelName / CircuitNumber / LoadName onto source items.
        for row in list(self._rows):
            item = row.item
            new_panel = (row.PanelName or "").strip()
            new_ckt = (row.CircuitNumber or "").strip()
            new_load = (row.LoadName or "").strip()
            item.user_panel = new_panel if new_panel != item.panel_name else None
            item.user_circuit_token = new_ckt if new_ckt != item.circuit_token else None
            item.user_load_name = new_load if new_load != item.load_name else None
            # Re-classify if the user changed the circuit token to a
            # bucket-driving keyword like DEDICATED.
            if item.user_circuit_token is not None:
                bucket, token = self.run.client.classify_circuit_token(item.user_circuit_token)
                item.bucket = bucket
                item.circuit_token = token
                item.user_circuit_token = token
            if item.user_panel is not None:
                item.panel_name = item.user_panel
                item.user_panel = None
            if item.user_load_name is not None:
                item.load_name = item.user_load_name
                item.user_load_name = None
        self.run.assemble()
        self._render_rows()
        self._set_status("Edits applied. {} group(s) ready.".format(len(self.run.groups)))

    def _on_run(self, sender, e):
        if not self.run.groups:
            self._set_status("No groups to run. Click Refresh from doc.")
            return
        # Honour the keyword filter — if the user picked a keyword,
        # apply only to the filtered groups, not to every group on
        # the run.
        keyword = self._selected_keyword()
        if keyword:
            target_groups = _workflow.filter_groups_by_keyword(
                self.run.groups, keyword, self.run.client
            )
            scope_msg = "{} group(s) with keyword {!r}".format(
                len(target_groups), keyword
            )
        else:
            target_groups = list(self.run.groups)
            scope_msg = "{} group(s)".format(len(target_groups))
        if not target_groups:
            self._set_status("No groups match the current filter.")
            return
        self._set_status(
            "Running circuits — Revit thread executing ({})...".format(scope_msg)
        )
        self.run_btn.IsEnabled = False
        self.apply_edits_btn.IsEnabled = False

        def _on_complete(result):
            try:
                self._on_apply_complete(result)
            finally:
                self.run_btn.IsEnabled = True
                self.apply_edits_btn.IsEnabled = True

        self.run.apply_async(groups=target_groups, on_complete=_on_complete)

    def _on_apply_complete(self, result):
        msg = "Created {} circuit(s); {} skipped (review); {} failed.".format(
            result.created_count, result.skipped_count, result.failed_count
        )
        if result.warnings:
            msg += " ({} warning(s))".format(len(result.warnings))
        self._set_status(msg)
        # Refresh from doc so the new systems show up in subsequent
        # passes and existing rows reflect any post-apply parameter
        # writes.
        self._refresh()

    # ----- misc ----------------------------------------------------

    def _set_status(self, text):
        try:
            self.status_label.Text = text or ""
        except Exception:
            pass

    # ----- entry points --------------------------------------------

    def show_modeless(self):
        # Auto-collect BEFORE Show() so the rows are present the first
        # frame the user sees. WPF Window controls loaded from XAML are
        # fully constructed; ItemsSource binds correctly before Show().
        # Using Window.Loaded with a transient lambda was unreliable —
        # the lambda got GC'd before the event fired.
        try:
            self._refresh()
        except Exception as exc:
            self._set_status("Initial refresh failed: {}".format(exc))
        # Defensive: if anything *did* defer to Loaded for any reason,
        # re-run there too. Handler retained so pythonnet keeps it.
        self._h_loaded = RoutedEventHandler(self._on_window_loaded)
        try:
            self.window.Loaded += self._h_loaded
        except Exception:
            pass
        self.window.Show()
        return self

    def _on_window_loaded(self, sender, e):
        # Only re-collect if the first sync pass produced no rows —
        # otherwise we'd flicker.
        try:
            if not list(self._rows):
                self._refresh()
        except Exception as exc:
            self._set_status("Loaded-event refresh failed: {}".format(exc))


def show_modeless(doc, uidoc, client, scope=_workflow.SCOPE_ALL,
                  selected_element_ids=None, profile_data=None):
    return SuperCircuitController(
        doc=doc,
        uidoc=uidoc,
        client=client,
        scope=scope,
        selected_element_ids=selected_element_ids,
        profile_data=profile_data,
    ).show_modeless()
