# -*- coding: utf-8 -*-
"""
Modal editor for ``space_buckets[*]`` in the active YAML payload.

Each row is one bucket: name, client_keys (comma-separated free text),
classification_keywords (comma-separated free text). The flat shape
fits a single DataGrid — no nested-list editor needed because both
list-typed fields are short (typically 1-5 entries).

The controller mutates ``profile_data["space_buckets"]`` in place; the
calling pushbutton script saves via ``active_yaml.save_active_data``
inside a Revit transaction (mirroring Manage Space Profiles).
"""

import os
import uuid

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402

import wpf as _wpf  # noqa: E402


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "ManageSpaceBucketsWindow.xaml",
)


def _new_id():
    return "SBKT-{}".format(uuid.uuid4().hex[:8].upper())


def _split_csv(text):
    if not text:
        return []
    return [t.strip() for t in str(text).split(",") if t.strip()]


def _join_csv(items):
    if not items:
        return ""
    return ", ".join(str(i) for i in items if str(i).strip())


# ---------------------------------------------------------------------
# Row binding object
# ---------------------------------------------------------------------

class _BucketRow(object):
    """One DataGrid row backed by a ``space_buckets[*]`` dict."""

    def __init__(self, bucket_dict):
        self._data = bucket_dict
        # Shadow text fields — the underlying YAML stores lists, but
        # the grid edits free-text. Push back to lists in setters.

    @property
    def BucketId(self):
        return self._data.get("id") or ""

    @property
    def Name(self):
        return self._data.get("name") or ""

    @Name.setter
    def Name(self, value):
        self._data["name"] = (value or "").strip()

    @property
    def ClientKeysText(self):
        raw = self._data.get("client_keys") or []
        return _join_csv(raw)

    @ClientKeysText.setter
    def ClientKeysText(self, value):
        self._data["client_keys"] = _split_csv(value)

    @property
    def KeywordsText(self):
        raw = self._data.get("classification_keywords") or []
        return _join_csv(raw)

    @KeywordsText.setter
    def KeywordsText(self, value):
        self._data["classification_keywords"] = _split_csv(value)


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class ManageSpaceBucketsController(object):

    def __init__(self, profile_data, doc=None):
        self.profile_data = profile_data
        self.doc = doc
        self.dirty = False

        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = ObservableCollection[_NetObject]()

        self._lookup_controls()
        self._wire_events()
        self._reload()
        self._set_status("Ready.")

    def _lookup_controls(self):
        f = self.window.FindName
        self.bucket_grid = f("BucketGrid")
        self.summary_label = f("SummaryLabel")
        self.new_btn = f("NewButton")
        self.duplicate_btn = f("DuplicateButton")
        self.delete_btn = f("DeleteButton")
        self.save_btn = f("SaveButton")
        self.close_btn = f("CloseButton")
        self.status_label = f("StatusLabel")
        self.bucket_grid.ItemsSource = self._rows

    def _wire_events(self):
        self._h_new = RoutedEventHandler(
            lambda s, e: self._safe(self._on_new, "new")
        )
        self._h_dup = RoutedEventHandler(
            lambda s, e: self._safe(self._on_duplicate, "duplicate")
        )
        self._h_del = RoutedEventHandler(
            lambda s, e: self._safe(self._on_delete, "delete")
        )
        self._h_save = RoutedEventHandler(
            lambda s, e: self._safe(self._on_save, "save")
        )
        self._h_close = RoutedEventHandler(
            lambda s, e: self.window.Close()
        )
        self.new_btn.Click += self._h_new
        self.duplicate_btn.Click += self._h_dup
        self.delete_btn.Click += self._h_del
        self.save_btn.Click += self._h_save
        self.close_btn.Click += self._h_close

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _set_status(self, text):
        self.status_label.Text = text or ""

    # ----- list ----------------------------------------------------

    def _buckets(self):
        raw = self.profile_data.setdefault("space_buckets", [])
        if not isinstance(raw, list):
            raw = []
            self.profile_data["space_buckets"] = raw
        return raw

    def _reload(self):
        self._rows.Clear()
        for b in self._buckets():
            if isinstance(b, dict):
                # Make sure every persisted bucket has a stable id.
                if not b.get("id"):
                    b["id"] = _new_id()
                self._rows.Add(_BucketRow(b))
        self.summary_label.Text = "{} bucket(s)".format(self._rows.Count)

    def _selected_row(self):
        sel = self.bucket_grid.SelectedItem
        return sel if isinstance(sel, _BucketRow) else None

    # ----- actions -------------------------------------------------

    def _on_new(self):
        new = {
            "id": _new_id(),
            "name": "NEW BUCKET",
            "client_keys": [],
            "classification_keywords": [],
        }
        self._buckets().append(new)
        self._rows.Add(_BucketRow(new))
        self.summary_label.Text = "{} bucket(s)".format(self._rows.Count)
        self.bucket_grid.SelectedItem = self._rows[self._rows.Count - 1]
        self.dirty = True
        self._set_status("New bucket created.")

    def _on_duplicate(self):
        row = self._selected_row()
        if row is None:
            self._set_status("Pick a row to duplicate.")
            return
        clone = {
            "id": _new_id(),
            "name": "{} (copy)".format(row._data.get("name") or "Bucket"),
            "client_keys": list(row._data.get("client_keys") or []),
            "classification_keywords": list(row._data.get("classification_keywords") or []),
        }
        self._buckets().append(clone)
        self._rows.Add(_BucketRow(clone))
        self.summary_label.Text = "{} bucket(s)".format(self._rows.Count)
        self.dirty = True
        self._set_status("Bucket duplicated.")

    def _on_delete(self):
        row = self._selected_row()
        if row is None:
            self._set_status("Pick a row to delete.")
            return
        try:
            self._buckets().remove(row._data)
        except ValueError:
            pass
        self._rows.Remove(row)
        self.summary_label.Text = "{} bucket(s)".format(self._rows.Count)
        self.dirty = True
        self._set_status("Bucket deleted.")

    def _on_save(self):
        # Force any in-flight cell edit to commit.
        try:
            self.bucket_grid.CommitEdit()
            self.bucket_grid.CommitEdit()
        except Exception:
            pass
        self.dirty = True
        self._set_status("Edits flushed. Click Close to save & dismiss.")

    def show(self):
        self.window.ShowDialog()


# ---------------------------------------------------------------------
# Convenience entry point
# ---------------------------------------------------------------------

def show_modal(profile_data, doc=None):
    controller = ManageSpaceBucketsController(profile_data=profile_data, doc=doc)
    controller.show()
    return controller
