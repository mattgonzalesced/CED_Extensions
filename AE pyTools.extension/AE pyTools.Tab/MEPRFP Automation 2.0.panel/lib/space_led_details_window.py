# -*- coding: utf-8 -*-
"""
Modal sub-dialog for per-LED detail editing on a Space profile.

Three tabs in one dialog:

  * **Parameters** — flat key/value DataGrid bound to ``led.parameters``.
  * **Offsets**    — DataGrid (X / Y / Z / rotation) bound to
                     ``led.offsets``.
  * **Annotations** — DataGrid of annotation rows (kind, label, family,
                      type, offset). A per-row "Params..." button pops
                      a generic key/value editor for the annotation's
                      own ``parameters`` dict.

All edits land on the in-memory dicts of the calling Manage Space
Profiles editor; the parent window is responsible for persisting them
to the active YAML payload on save. OK commits in-flight cell edits
and closes; Cancel discards the current cell-in-progress.
"""

import copy
import os
import uuid

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402
from System.Windows.Controls import Button as _WpfButton  # noqa: E402

import wpf as _wpf  # noqa: E402


_RESOURCES = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources",
)
_DETAILS_XAML = os.path.join(_RESOURCES, "SpaceLedDetailsDialog.xaml")
_KV_XAML = os.path.join(_RESOURCES, "KeyValueDialog.xaml")


_ANNOTATION_KINDS = ("tag", "keynote", "text_note")


# ---------------------------------------------------------------------
# Shared row classes
# ---------------------------------------------------------------------

class _ParamRow(object):
    """Two-column key/value row backed by a host dict."""

    def __init__(self, name="", value=""):
        self.Name = "" if name is None else str(name)
        self.Value = "" if value is None else _coerce_to_text(value)


def _coerce_to_text(value):
    """Stringify a parameter value for display.

    Dict / list values come from directives or structured params and
    can't be edited inline; display them but mark them protected via
    repr() so the user understands they're not plain text.
    """
    if isinstance(value, (dict, list)):
        return repr(value)
    return str(value)


def _coerce_from_text(name, text, original_value):
    """Best-effort parse of edited text back into the YAML value type.

    Numeric-looking strings stay text — Revit's parameter writer
    handles unit conversion. We only special-case directive dicts
    (``{...}`` literal) and lists by parsing them with ``ast.literal_eval``;
    plain strings are returned as-is.
    """
    if text is None:
        return None
    s = str(text)
    if not s.strip():
        return ""
    # If the original was a dict / list, try to round-trip through repr.
    if isinstance(original_value, (dict, list)) and (s.startswith("{") or s.startswith("[")):
        try:
            import ast
            return ast.literal_eval(s)
        except Exception:
            return s
    return s


def _to_float(text, default=0.0):
    if text is None:
        return default
    s = str(text).strip()
    if not s:
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _fmt_float(value):
    if value is None:
        return ""
    try:
        v = float(value)
    except (ValueError, TypeError):
        return ""
    if v == int(v):
        return str(int(v))
    return "{:.4f}".format(v).rstrip("0").rstrip(".")


# ---------------------------------------------------------------------
# Offset row
# ---------------------------------------------------------------------

class _OffsetRow(object):
    """One ``led.offsets[*]`` entry."""

    def __init__(self, data):
        self._data = data

    @property
    def XText(self):
        return _fmt_float(self._data.get("x_inches"))

    @XText.setter
    def XText(self, value):
        self._data["x_inches"] = _to_float(value, 0.0)

    @property
    def YText(self):
        return _fmt_float(self._data.get("y_inches"))

    @YText.setter
    def YText(self, value):
        self._data["y_inches"] = _to_float(value, 0.0)

    @property
    def ZText(self):
        return _fmt_float(self._data.get("z_inches"))

    @ZText.setter
    def ZText(self, value):
        self._data["z_inches"] = _to_float(value, 0.0)

    @property
    def RotText(self):
        return _fmt_float(self._data.get("rotation_deg"))

    @RotText.setter
    def RotText(self, value):
        self._data["rotation_deg"] = _to_float(value, 0.0)


# ---------------------------------------------------------------------
# Annotation row
# ---------------------------------------------------------------------

class _AnnotationRow(object):
    """One ``led.annotations[*]`` entry."""

    def __init__(self, data):
        self._data = data
        self.KindOptions = list(_ANNOTATION_KINDS)

    @property
    def Kind(self):
        return self._data.get("kind") or "tag"

    @Kind.setter
    def Kind(self, value):
        if value in _ANNOTATION_KINDS:
            self._data["kind"] = value

    @property
    def Label(self):
        return self._data.get("label") or ""

    @Label.setter
    def Label(self, value):
        self._data["label"] = (value or "").strip()

    @property
    def FamilyName(self):
        return self._data.get("family_name") or ""

    @FamilyName.setter
    def FamilyName(self, value):
        self._data["family_name"] = (value or "").strip()

    @property
    def TypeName(self):
        return self._data.get("type_name") or ""

    @TypeName.setter
    def TypeName(self, value):
        self._data["type_name"] = (value or "").strip()

    def _offset_dict(self):
        # Annotation offsets are a single dict, not a list.
        d = self._data.setdefault("offsets", {})
        if isinstance(d, list):
            d = d[0] if d else {}
            self._data["offsets"] = d
        if not isinstance(d, dict):
            d = {}
            self._data["offsets"] = d
        return d

    @property
    def OffsetXText(self):
        return _fmt_float(self._offset_dict().get("x_inches"))

    @OffsetXText.setter
    def OffsetXText(self, value):
        self._offset_dict()["x_inches"] = _to_float(value, 0.0)

    @property
    def OffsetYText(self):
        return _fmt_float(self._offset_dict().get("y_inches"))

    @OffsetYText.setter
    def OffsetYText(self, value):
        self._offset_dict()["y_inches"] = _to_float(value, 0.0)

    @property
    def OffsetZText(self):
        return _fmt_float(self._offset_dict().get("z_inches"))

    @OffsetZText.setter
    def OffsetZText(self, value):
        self._offset_dict()["z_inches"] = _to_float(value, 0.0)

    @property
    def OffsetRotText(self):
        return _fmt_float(self._offset_dict().get("rotation_deg"))

    @OffsetRotText.setter
    def OffsetRotText(self, value):
        self._offset_dict()["rotation_deg"] = _to_float(value, 0.0)


# ---------------------------------------------------------------------
# Generic key/value sub-dialog
# ---------------------------------------------------------------------

class KeyValueDialog(object):
    """Modal: edit a flat string -> string dict."""

    def __init__(self, params_dict, header="Edit Parameters"):
        self._params = params_dict if isinstance(params_dict, dict) else {}
        self.window = _wpf.load_xaml(_KV_XAML)
        self._rows = ObservableCollection[_NetObject]()
        self._committed = False

        f = self.window.FindName
        self.header_label = f("HeaderLabel")
        self.grid = f("ParamGrid")
        self.add_btn = f("AddRowButton")
        self.del_btn = f("DeleteRowButton")
        self.ok_btn = f("OkButton")
        self.cancel_btn = f("CancelButton")
        self.header_label.Text = header
        self.grid.ItemsSource = self._rows
        self._snapshot = copy.deepcopy(self._params)

        for k, v in self._params.items():
            self._rows.Add(_ParamRow(k, v))

        self._h_add = RoutedEventHandler(lambda s, e: self._on_add())
        self._h_del = RoutedEventHandler(lambda s, e: self._on_delete())
        self._h_ok = RoutedEventHandler(lambda s, e: self._on_ok())
        self._h_cancel = RoutedEventHandler(lambda s, e: self._on_cancel())
        self.add_btn.Click += self._h_add
        self.del_btn.Click += self._h_del
        self.ok_btn.Click += self._h_ok
        self.cancel_btn.Click += self._h_cancel

    def _on_add(self):
        self._rows.Add(_ParamRow("", ""))
        self.grid.SelectedItem = self._rows[self._rows.Count - 1]

    def _on_delete(self):
        sel = self.grid.SelectedItem
        if isinstance(sel, _ParamRow):
            self._rows.Remove(sel)

    def _on_ok(self):
        try:
            self.grid.CommitEdit()
            self.grid.CommitEdit()
        except Exception:
            pass
        # Rebuild the source dict from the rows. Preserve order.
        new_data = {}
        for row in self._rows:
            name = (row.Name or "").strip()
            if not name:
                continue
            original = self._snapshot.get(name)
            new_data[name] = _coerce_from_text(name, row.Value, original)
        # Mutate the caller's dict in place (so references survive).
        self._params.clear()
        self._params.update(new_data)
        self._committed = True
        self.window.Close()

    def _on_cancel(self):
        # Restore the snapshot — caller's dict is mutated only on OK.
        self._params.clear()
        self._params.update(self._snapshot)
        self._committed = False
        self.window.Close()

    def show_modal(self, owner=None):
        if owner is not None:
            try:
                self.window.Owner = owner
            except Exception:
                pass
        self.window.ShowDialog()
        return self._committed


# ---------------------------------------------------------------------
# Main details dialog
# ---------------------------------------------------------------------

class SpaceLedDetailsController(object):
    """Edit one LED's parameters / offsets / annotations dicts.

    Mutates the passed-in LED dict in place. Parent caller (Manage
    Space Profiles) decides whether to persist to YAML.
    """

    def __init__(self, led_dict, header=""):
        self._led = led_dict if isinstance(led_dict, dict) else {}
        # Take a deep snapshot so Cancel can fully restore.
        self._snapshot = copy.deepcopy(self._led)
        self._committed = False

        self.window = _wpf.load_xaml(_DETAILS_XAML)
        self._param_rows = ObservableCollection[_NetObject]()
        self._offset_rows = ObservableCollection[_NetObject]()
        self._ann_rows = ObservableCollection[_NetObject]()

        self._lookup_controls()
        self._wire_events()
        self.header_label.Text = header or "Edit LED: {} ({})".format(
            self._led.get("label") or "(no label)",
            self._led.get("id") or "?",
        )
        self._reload_all_tabs()
        self._set_status("Ready.")

    # ----- bootstrapping -------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.header_label = f("HeaderLabel")
        self.tabs = f("DetailTabs")
        self.status_label = f("StatusLabel")
        self.ok_btn = f("OkButton")
        self.cancel_btn = f("CancelButton")

        # Parameters tab
        self.param_grid = f("ParamGrid")
        self.param_add_btn = f("ParamAddButton")
        self.param_del_btn = f("ParamDeleteButton")
        self.param_grid.ItemsSource = self._param_rows

        # Offsets tab
        self.offset_grid = f("OffsetGrid")
        self.offset_add_btn = f("OffsetAddButton")
        self.offset_del_btn = f("OffsetDeleteButton")
        self.offset_grid.ItemsSource = self._offset_rows

        # Annotations tab
        self.ann_grid = f("AnnGrid")
        self.ann_add_btn = f("AnnAddButton")
        self.ann_del_btn = f("AnnDeleteButton")
        self.ann_grid.ItemsSource = self._ann_rows

    def _wire_events(self):
        self._h_param_add = RoutedEventHandler(lambda s, e: self._safe(self._on_param_add, "param-add"))
        self._h_param_del = RoutedEventHandler(lambda s, e: self._safe(self._on_param_delete, "param-del"))
        self._h_offset_add = RoutedEventHandler(lambda s, e: self._safe(self._on_offset_add, "offset-add"))
        self._h_offset_del = RoutedEventHandler(lambda s, e: self._safe(self._on_offset_delete, "offset-del"))
        self._h_ann_add = RoutedEventHandler(lambda s, e: self._safe(self._on_ann_add, "ann-add"))
        self._h_ann_del = RoutedEventHandler(lambda s, e: self._safe(self._on_ann_delete, "ann-del"))
        self._h_ok = RoutedEventHandler(lambda s, e: self._safe(self._on_ok, "ok"))
        self._h_cancel = RoutedEventHandler(lambda s, e: self._safe(self._on_cancel, "cancel"))

        self.param_add_btn.Click += self._h_param_add
        self.param_del_btn.Click += self._h_param_del
        self.offset_add_btn.Click += self._h_offset_add
        self.offset_del_btn.Click += self._h_offset_del
        self.ann_add_btn.Click += self._h_ann_add
        self.ann_del_btn.Click += self._h_ann_del
        self.ok_btn.Click += self._h_ok
        self.cancel_btn.Click += self._h_cancel

        # Bubbled Click handler for the per-row "Params..." button on
        # annotation rows. The button's Tag is bound to the
        # _AnnotationRow; non-row clicks have no tag and are ignored.
        self._h_row_click = RoutedEventHandler(
            lambda s, e: self._safe_with(s, e, self._on_row_button_click, "row-click")
        )
        self.window.AddHandler(_WpfButton.ClickEvent, self._h_row_click)

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

    # ----- reload --------------------------------------------------

    def _reload_all_tabs(self):
        self._reload_params()
        self._reload_offsets()
        self._reload_annotations()

    def _reload_params(self):
        self._param_rows.Clear()
        params = self._led.setdefault("parameters", {})
        if not isinstance(params, dict):
            params = {}
            self._led["parameters"] = params
        for k, v in params.items():
            self._param_rows.Add(_ParamRow(k, v))

    def _reload_offsets(self):
        self._offset_rows.Clear()
        offsets = self._led.setdefault("offsets", [])
        if not isinstance(offsets, list):
            offsets = []
            self._led["offsets"] = offsets
        for o in offsets:
            if isinstance(o, dict):
                self._offset_rows.Add(_OffsetRow(o))

    def _reload_annotations(self):
        self._ann_rows.Clear()
        anns = self._led.setdefault("annotations", [])
        if not isinstance(anns, list):
            anns = []
            self._led["annotations"] = anns
        for a in anns:
            if isinstance(a, dict):
                a.setdefault("kind", "tag")
                a.setdefault("id", _new_id("ANN"))
                self._ann_rows.Add(_AnnotationRow(a))

    # ----- parameter actions ---------------------------------------

    def _on_param_add(self):
        self._param_rows.Add(_ParamRow("", ""))
        self.param_grid.SelectedItem = self._param_rows[self._param_rows.Count - 1]

    def _on_param_delete(self):
        sel = self.param_grid.SelectedItem
        if isinstance(sel, _ParamRow):
            self._param_rows.Remove(sel)

    # ----- offset actions ------------------------------------------

    def _on_offset_add(self):
        new = {
            "x_inches": 0.0, "y_inches": 0.0,
            "z_inches": 0.0, "rotation_deg": 0.0,
        }
        offsets = self._led.setdefault("offsets", [])
        offsets.append(new)
        self._offset_rows.Add(_OffsetRow(new))
        self.offset_grid.SelectedItem = self._offset_rows[self._offset_rows.Count - 1]

    def _on_offset_delete(self):
        sel = self.offset_grid.SelectedItem
        if not isinstance(sel, _OffsetRow):
            return
        try:
            self._led.get("offsets", []).remove(sel._data)
        except ValueError:
            pass
        self._offset_rows.Remove(sel)

    # ----- annotation actions --------------------------------------

    def _on_ann_add(self):
        new = {
            "id": _new_id("ANN"),
            "kind": "tag",
            "label": "",
            "family_name": "",
            "type_name": "",
            "parameters": {},
            "offsets": {"x_inches": 0.0, "y_inches": 0.0,
                        "z_inches": 0.0, "rotation_deg": 0.0},
        }
        anns = self._led.setdefault("annotations", [])
        anns.append(new)
        self._ann_rows.Add(_AnnotationRow(new))
        self.ann_grid.SelectedItem = self._ann_rows[self._ann_rows.Count - 1]

    def _on_ann_delete(self):
        sel = self.ann_grid.SelectedItem
        if not isinstance(sel, _AnnotationRow):
            return
        try:
            self._led.get("annotations", []).remove(sel._data)
        except ValueError:
            pass
        self._ann_rows.Remove(sel)

    def _on_row_button_click(self, sender, e):
        # Bubbled Click. e.Source = clicked Button (when it raised the event).
        source = getattr(e, "Source", None) or getattr(e, "OriginalSource", None)
        tag = getattr(source, "Tag", None) if source is not None else None
        if not isinstance(tag, _AnnotationRow):
            return
        row = tag
        params = row._data.setdefault("parameters", {})
        if not isinstance(params, dict):
            params = {}
            row._data["parameters"] = params
        dialog = KeyValueDialog(
            params,
            header="Annotation parameters: {} [{}]".format(
                row._data.get("label") or row._data.get("type_name") or "(no label)",
                row._data.get("kind") or "?",
            ),
        )
        dialog.show_modal(owner=self.window)
        # Force the grid to redraw in case the user added a Notes-like
        # parameter that may shift display elsewhere.
        self.ann_grid.Items.Refresh()

    # ----- OK / Cancel ---------------------------------------------

    def _on_ok(self):
        # Commit any in-flight DataGrid edit so the last-typed cell makes it.
        for grid in (self.param_grid, self.offset_grid, self.ann_grid):
            try:
                grid.CommitEdit()
                grid.CommitEdit()
            except Exception:
                pass

        # Rebuild parameters from rows.
        params_out = {}
        original_params = self._snapshot.get("parameters") or {}
        for row in self._param_rows:
            name = (row.Name or "").strip()
            if not name:
                continue
            params_out[name] = _coerce_from_text(
                name, row.Value, original_params.get(name),
            )
        self._led["parameters"] = params_out

        # Offsets are already mutated in place via the row setters,
        # but rebuild the list from the current rows (handles deletes
        # and any reorder if we ever add it).
        self._led["offsets"] = [r._data for r in self._offset_rows]

        # Annotations same — mutated in place via setters, just rebuild
        # the list from current rows.
        self._led["annotations"] = [r._data for r in self._ann_rows]

        self._committed = True
        self.window.Close()

    def _on_cancel(self):
        # Restore the snapshot fully — every nested dict / list.
        self._led.clear()
        self._led.update(copy.deepcopy(self._snapshot))
        self._committed = False
        self.window.Close()

    # ----- entry point --------------------------------------------

    def show_modal(self, owner=None):
        if owner is not None:
            try:
                self.window.Owner = owner
            except Exception:
                pass
        self.window.ShowDialog()
        return self._committed


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _new_id(prefix):
    return "{}-{}".format(prefix, uuid.uuid4().hex[:8].upper())


def show_modal(led_dict, header="", owner=None):
    """Open the Details dialog for ``led_dict``. Returns True on OK."""
    controller = SpaceLedDetailsController(led_dict=led_dict, header=header)
    return controller.show_modal(owner=owner)
