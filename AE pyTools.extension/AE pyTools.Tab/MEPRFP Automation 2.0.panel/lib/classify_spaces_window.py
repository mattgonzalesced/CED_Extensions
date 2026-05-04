# -*- coding: utf-8 -*-
"""
Modal UI for the Classify Spaces pushbutton.

The window walks every placed Space in the document, runs the auto-
classifier (case-insensitive substring match against
``space_buckets[*].classification_keywords``), and lets the user
override the assigned buckets per row before saving the result to the
project's Extensible Storage entity.

Multi-bucket assignment is intentional: a space can match more than one
bucket (e.g. RESTROOM + WOMEN'S) and the placement engine unions the
LEDs from every matching profile. The Edit... button per row pops a
``BucketPickerDialog`` checklist so the user can fine-tune the list
without typing bucket names.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402
from System.Windows.Controls import (  # noqa: E402
    Button as _WpfButton,
    SelectionChangedEventHandler,
)

from pyrevit import revit  # noqa: E402

import active_yaml as _active_yaml  # noqa: E402
import space_workflow as _workflow  # noqa: E402
import space_bucket_model as _bucket_model  # noqa: E402
import wpf as _wpf  # noqa: E402

try:
    import circuit_clients  # noqa: E402  -- reuse the same registry
    _HAS_CLIENTS = True
except Exception:  # pragma: no cover -- early-bring-up safety
    circuit_clients = None
    _HAS_CLIENTS = False


_RESOURCES = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources",
)
_MAIN_XAML = os.path.join(_RESOURCES, "ClassifySpacesWindow.xaml")
_PICKER_XAML = os.path.join(_RESOURCES, "BucketPickerDialog.xaml")

_NO_CLIENT_LABEL = "(no client / universal only)"


# ---------------------------------------------------------------------
# Row binding object
# ---------------------------------------------------------------------

class _SpaceRow(object):
    """One DataGrid row.

    Plain Python; WPF two-way bindings read/write attribute slots
    directly. ``assigned_ids`` is the list-of-bucket-ids backing the
    Assigned column; ``AssignedLabel`` is the comma-joined display.
    """

    def __init__(self, space, auto_ids, assigned_ids):
        self.space = space
        self.auto_ids = list(auto_ids or [])
        self.assigned_ids = list(assigned_ids or [])

    # Display columns (read by WPF) ----------------------------------

    @property
    def LevelName(self):
        return self.space.level_name or ""

    @property
    def Number(self):
        return self.space.number or ""

    @property
    def Name(self):
        return self.space.name or ""

    @property
    def ElementIdText(self):
        return str(self.space.element_id) if self.space.element_id is not None else ""

    @property
    def AutoLabel(self):
        return self._format_label(self.auto_ids)

    @property
    def AssignedLabel(self):
        return self._format_label(self.assigned_ids)

    # Mutators (called by the controller) -----------------------------

    def set_assigned(self, ids):
        self.assigned_ids = list(ids or [])

    def reset_to_auto(self):
        self.assigned_ids = list(self.auto_ids)

    # Internal --------------------------------------------------------

    def _format_label(self, ids):
        names = []
        controller = _SpaceRow._controller_ref
        if controller is None:
            return ", ".join(ids)
        for bid in ids:
            bucket = controller.bucket_by_id(bid)
            names.append(bucket.name if bucket else "({})".format(bid))
        return ", ".join(names) if names else "(unclassified)"


# Class-level back-reference used to resolve bucket-id -> name in the
# read-only AutoLabel / AssignedLabel properties. WPF reads those over
# and over so a per-row controller pointer is the cheapest way to
# resolve names without rebuilding the row each time.
_SpaceRow._controller_ref = None  # set by ClassifySpacesController.__init__


# ---------------------------------------------------------------------
# BucketPickerDialog row
# ---------------------------------------------------------------------

class _BucketCheckRow(object):
    """Row in the picker's checklist. WPF mutates ``IsChecked`` directly."""

    def __init__(self, bucket_id, label, tooltip, is_checked):
        self.BucketId = bucket_id
        self.Label = label
        self.ToolTip = tooltip or ""
        self.IsChecked = bool(is_checked)


class _BucketPickerDialog(object):
    """Modal sub-dialog: pick which buckets apply to one space."""

    def __init__(self, all_buckets, assigned_ids, auto_ids,
                 space_label="", header=""):
        self.window = _wpf.load_xaml(_PICKER_XAML)
        self._auto_ids = list(auto_ids or [])
        self._result = None
        self._cancelled = True

        f = self.window.FindName
        self.header_label = f("HeaderLabel")
        self.list_box = f("BucketList")
        self.ok_btn = f("OkButton")
        self.cancel_btn = f("CancelButton")
        self.reset_btn = f("ResetButton")

        self.header_label.Text = header or "Buckets for: {}".format(space_label)

        assigned_set = set(assigned_ids or [])
        self._items = []
        for bucket in all_buckets:
            label = bucket.name or "(unnamed)"
            kws = ", ".join(bucket.classification_keywords) or "no keywords"
            tooltip = "id={}  |  keywords: {}".format(
                bucket.id or "?", kws
            )
            self._items.append(_BucketCheckRow(
                bucket_id=bucket.id,
                label=label,
                tooltip=tooltip,
                is_checked=bucket.id in assigned_set,
            ))
        self.list_box.ItemsSource = self._items

        self._h_ok = self._delegate(self._on_ok)
        self._h_cancel = self._delegate(self._on_cancel)
        self._h_reset = self._delegate(self._on_reset)
        self.ok_btn.Click += self._h_ok
        self.cancel_btn.Click += self._h_cancel
        self.reset_btn.Click += self._h_reset

    @staticmethod
    def _delegate(fn):
        return RoutedEventHandler(lambda s, e: fn(s, e))

    def _on_ok(self, sender, e):
        self._result = [r.BucketId for r in self._items if r.IsChecked]
        self._cancelled = False
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._cancelled = True
        self.window.Close()

    def _on_reset(self, sender, e):
        target = set(self._auto_ids)
        for row in self._items:
            row.IsChecked = row.BucketId in target
        # Force WPF to redraw the list with the new checkbox states.
        self.list_box.Items.Refresh()

    def show_modal(self, owner=None):
        if owner is not None:
            try:
                self.window.Owner = owner
            except Exception:
                pass
        self.window.ShowDialog()
        if self._cancelled:
            return None
        return list(self._result or [])


# ---------------------------------------------------------------------
# Main controller
# ---------------------------------------------------------------------

class ClassifySpacesController(object):

    def __init__(self, doc, uidoc=None, profile_data=None):
        self.doc = doc
        self.uidoc = uidoc
        self.profile_data = profile_data or _active_yaml.load_active_data(doc) or {}

        self.window = _wpf.load_xaml(_MAIN_XAML)
        self._rows = ObservableCollection[_NetObject]()
        self._buckets_wrapped = []  # current client-filtered buckets
        self._all_buckets = []      # unfiltered, used for already-assigned IDs
        self._bucket_index_by_id = {}
        self._lookup_controls()
        self._wire_events()
        self._populate_client_combo()

        _SpaceRow._controller_ref = self

        self._set_status("Loading spaces...")
        self._refresh()

    # ----- bootstrapping -------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.client_combo = f("ClientCombo")
        self.summary_label = f("SummaryLabel")
        self.spaces_grid = f("SpacesGrid")
        self.refresh_btn = f("RefreshButton")
        self.reset_all_btn = f("ResetAllButton")
        self.save_btn = f("SaveButton")
        self.close_btn = f("CloseButton")
        self.status_label = f("StatusLabel")
        self.spaces_grid.ItemsSource = self._rows

    def _wire_events(self):
        # Retain handlers as attributes so pythonnet doesn't GC them.
        self._h_refresh = RoutedEventHandler(
            lambda s, e: self._safe(self._refresh, "refresh")
        )
        self._h_reset_all = RoutedEventHandler(
            lambda s, e: self._safe(self._on_reset_all, "reset-all")
        )
        self._h_save = RoutedEventHandler(
            lambda s, e: self._safe(self._on_save, "save")
        )
        self._h_close = RoutedEventHandler(
            lambda s, e: self.window.Close()
        )
        self._h_edit_row = RoutedEventHandler(
            lambda s, e: self._safe_with(s, e, self._on_edit_row, "edit-row")
        )
        self._h_client = RoutedEventHandler(
            lambda s, e: self._safe(self._on_client_changed, "client-changed")
        )

        self.refresh_btn.Click += self._h_refresh
        self.reset_all_btn.Click += self._h_reset_all
        self.save_btn.Click += self._h_save
        self.close_btn.Click += self._h_close

        self._h_client_sc = SelectionChangedEventHandler(
            lambda s, e: self._safe(self._on_client_changed, "client-changed")
        )
        self.client_combo.SelectionChanged += self._h_client_sc

        # Per-row Edit... buttons live inside the DataGrid's
        # DataTemplate, so we can't grab them by name. Register a
        # bubbling Click handler on the window itself: every Button
        # click bubbles up here, and ``_on_edit_row`` filters by Tag
        # so unrelated clicks (Save / Refresh / Close / etc) are
        # ignored cheaply.
        self.window.AddHandler(_WpfButton.ClickEvent, self._h_edit_row)

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _safe_with(self, sender, e, fn, label):
        try:
            fn(sender, e)
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _set_status(self, text):
        self.status_label.Text = text or ""

    # ----- client picker -------------------------------------------

    def _populate_client_combo(self):
        self.client_combo.Items.Clear()
        self.client_combo.Items.Add(_NO_CLIENT_LABEL)
        if _HAS_CLIENTS:
            for client in circuit_clients.all_clients():
                label = "{} ({})".format(client.display_name, client.key)
                self.client_combo.Items.Add(label)
        self.client_combo.SelectedIndex = 0

    def _selected_client_key(self):
        item = self.client_combo.SelectedItem
        if item is None:
            return None
        text = str(item).strip()
        if not text or text == _NO_CLIENT_LABEL:
            return None
        # Format: "<DisplayName> (<key>)"
        if text.endswith(")") and "(" in text:
            return text.rsplit("(", 1)[1].rstrip(")").strip()
        return text.lower()

    def _on_client_changed(self):
        # Re-classify with the new client filter (saved overrides preserved).
        self._refresh(preserve_assigned=True)

    # ----- pipeline ------------------------------------------------

    def _refresh(self, preserve_assigned=False):
        client_key = self._selected_client_key()

        # Keep a copy of current per-row assignments so a client switch
        # doesn't blow away unsaved edits.
        prior_assigned = {}
        if preserve_assigned:
            for row in self._rows:
                if row.space.element_id is not None:
                    prior_assigned[row.space.element_id] = list(row.assigned_ids)

        spaces = _workflow.collect_spaces(self.doc)
        all_buckets_raw = self.profile_data.get("space_buckets") or []
        self._all_buckets = _bucket_model.wrap_buckets(all_buckets_raw)
        self._bucket_index_by_id = {
            b.id: b for b in self._all_buckets if b.id
        }
        self._buckets_wrapped = _bucket_model.filter_buckets_for_client(
            self._all_buckets, client_key,
        )

        auto_pairs = _workflow.auto_classify(
            spaces, self._buckets_wrapped, client_key=client_key,
        )

        if preserve_assigned:
            saved_index = {
                sid: ids for sid, ids in prior_assigned.items()
            }
        else:
            saved_index = _workflow.load_classifications_indexed(self.doc)

        merged = _workflow.merge_with_saved(auto_pairs, saved_index)

        self._rows.Clear()
        for space, assigned_ids, auto_ids in merged:
            self._rows.Add(_SpaceRow(space, auto_ids, assigned_ids))

        n_total = len(self._rows)
        n_classified = sum(1 for r in self._rows if r.assigned_ids)
        n_universal = sum(1 for b in self._buckets_wrapped if b.is_universal)
        n_client = len(self._buckets_wrapped) - n_universal
        client_label = client_key or "(none)"
        self.summary_label.Text = (
            "{} space(s); {} classified, {} unclassified.   "
            "Buckets in scope: {} ({} universal + {} client). Client: {}"
        ).format(
            n_total, n_classified, n_total - n_classified,
            len(self._buckets_wrapped), n_universal, n_client, client_label,
        )
        if not all_buckets_raw:
            self._set_status(
                "No space_buckets defined in the active YAML. Use Manage Space "
                "Profiles (coming in batch 4) or hand-edit the YAML to add some."
            )
        elif n_total == 0:
            self._set_status("No placed MEP Spaces in this document.")
        else:
            self._set_status(
                "Edit per-row buckets, then Save classifications. "
                "Multi-bucket rows stack their assigned profiles on apply."
            )

    def _on_reset_all(self):
        for row in self._rows:
            row.reset_to_auto()
        self.spaces_grid.Items.Refresh()
        self._set_status("All rows reset to auto-detected buckets.")

    # ----- per-row Edit... -----------------------------------------

    def _on_edit_row(self, sender, e):
        # Bubbled Button.Click. ``e.Source`` is normally the Button
        # itself (Button.OnClick raises the event with Source=this).
        # The per-row Edit... button's Tag is bound to the SpaceRow,
        # so this filter implicitly ignores Save / Refresh / Close,
        # which carry no row tag.
        source = getattr(e, "Source", None) or getattr(e, "OriginalSource", None)
        tag = getattr(source, "Tag", None) if source is not None else None
        if not isinstance(tag, _SpaceRow):
            return

        row = tag
        space = row.space
        space_label = "{} {}".format(space.number, space.name).strip() or "(unnamed space)"
        header = "Edit buckets for: {}".format(space_label)
        picker = _BucketPickerDialog(
            all_buckets=self._all_buckets,  # show all buckets, not just client-filtered
            assigned_ids=row.assigned_ids,
            auto_ids=row.auto_ids,
            space_label=space_label,
            header=header,
        )
        result = picker.show_modal(owner=self.window)
        if result is None:
            return  # cancelled
        row.set_assigned(result)
        # Force WPF to re-read AssignedLabel for the affected row.
        self.spaces_grid.Items.Refresh()
        self._set_status("Updated buckets for {}.".format(space_label))

    # ----- save ----------------------------------------------------

    def _on_save(self):
        assignments = [(row.space, list(row.assigned_ids)) for row in self._rows]
        records = _workflow.payload_from_assignments(assignments)

        with revit.Transaction("Save Space Classifications", doc=self.doc):
            _active_yaml.save_classifications(self.doc, records)

        n_total = len(assignments)
        n_classified = sum(1 for _, ids in assignments if ids)
        self._set_status(
            "Saved. {} space(s) total; {} classified, {} unclassified.".format(
                n_total, n_classified, n_total - n_classified,
            )
        )

    # ----- helpers used by row labels -----------------------------

    def bucket_by_id(self, bucket_id):
        return self._bucket_index_by_id.get(bucket_id)


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

def show_modal(doc, profile_data=None, uidoc=None):
    """Open the Classify Spaces modal. Blocks until closed."""
    controller = ClassifySpacesController(
        doc=doc, uidoc=uidoc, profile_data=profile_data,
    )
    controller.window.ShowDialog()
    return controller
