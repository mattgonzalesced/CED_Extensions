# -*- coding: utf-8 -*-
"""
WPF window for resolving parent parameter conflicts after sync.
Supports modal and modeless usage.
"""

from pyrevit import forms, revit
from System import TimeSpan
from System.Windows import Thickness, GridLength, GridUnitType, FontWeights, Visibility
from System.Windows.Controls import (
    Grid,
    RowDefinition,
    ColumnDefinition,
    TextBlock,
    ComboBox,
    StackPanel,
    Orientation,
    Button,
)
from System.Windows.Threading import DispatcherTimer
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

ACTION_UPDATE_CHILD = "update_child"
ACTION_UPDATE_PARENT = "update_parent"
ACTION_SKIP = "skip"

LABEL_UPDATE_CHILD = "Update child from parent"
LABEL_UPDATE_PARENT = "Update parent from child"
LABEL_SKIP = "Skip"
LABEL_NO_OVERRIDE = "<no override>"


class _ExternalActionHandler(IExternalEventHandler):
    def __init__(self, owner, name, executor):
        self._owner = owner
        self._name = name
        self._executor = executor
        self._payload = None

    def set_payload(self, payload):
        self._payload = payload

    def Execute(self, uiapp):
        payload = self._payload
        self._payload = None
        try:
            self._executor(uiapp, payload)
        except Exception:
            pass
        finally:
            try:
                self._owner._external_busy = False
            except Exception:
                pass

    def GetName(self):
        return self._name


class ParentParamConflictsWindow(forms.WPFWindow):
    def __init__(
        self,
        xaml_path,
        conflicts,
        param_keys,
        modeless=False,
        refresh_callback=None,
        apply_callback=None,
        select_callback=None,
        close_callback=None,
    ):
        forms.WPFWindow.__init__(self, xaml_path)
        self._conflicts = conflicts or []
        self._param_keys = param_keys or []
        self._combo_meta = {}
        self._combos_by_param = {}
        self._default_meta = {}
        self._select_meta = {}
        self.decisions = {}

        self._modeless = bool(modeless)
        self._refresh_callback = refresh_callback
        self._apply_callback = apply_callback
        self._select_callback = select_callback
        self._close_callback = close_callback

        self._selection_signature = None
        self._conflict_signature = self._build_conflict_signature(self._conflicts)
        self._external_busy = False
        self._refresh_timer = None

        self._refresh_handler = None
        self._refresh_event = None
        self._apply_handler = None
        self._apply_event = None
        self._select_handler = None
        self._select_event = None

        if self._modeless:
            self._setup_external_events()
            self._setup_refresh_timer()
            self.Closed += self._on_window_closed

        header = self.FindName("HeaderText")
        if header is not None:
            if self._modeless:
                header.Text = (
                    "Modeless conflict checker. Use Select to jump to elements. "
                    "The grid refreshes when selection changes and model data differs."
                )
            else:
                header.Text = "Resolve parent/sibling parameter discrepancies detected after sync."

        self._rebuild_defaults_panel()
        self._build_grid()

        refresh_btn = self.FindName("RefreshButton")
        apply_btn = self.FindName("ApplyButton")
        cancel_btn = self.FindName("CancelButton")
        if refresh_btn is not None:
            refresh_btn.Click += self._on_refresh
            if not self._modeless:
                refresh_btn.Visibility = Visibility.Collapsed
        if apply_btn is not None:
            apply_btn.Click += self._on_apply
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel
            if self._modeless:
                cancel_btn.Content = "Close"

    def _setup_external_events(self):
        self._refresh_handler = _ExternalActionHandler(self, "ParentConflictsRefresh", self._execute_refresh)
        self._refresh_event = ExternalEvent.Create(self._refresh_handler)
        self._apply_handler = _ExternalActionHandler(self, "ParentConflictsApply", self._execute_apply)
        self._apply_event = ExternalEvent.Create(self._apply_handler)
        self._select_handler = _ExternalActionHandler(self, "ParentConflictsSelect", self._execute_select)
        self._select_event = ExternalEvent.Create(self._select_handler)

    def _setup_refresh_timer(self):
        try:
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(800)
            timer.Tick += self._on_refresh_timer_tick
            timer.Start()
            self._refresh_timer = timer
        except Exception:
            self._refresh_timer = None

    def _on_window_closed(self, sender, args):
        try:
            if self._refresh_timer is not None:
                self._refresh_timer.Stop()
        except Exception:
            pass
        self._refresh_timer = None
        callback = self._close_callback
        if callback:
            try:
                callback()
            except Exception:
                pass

    def _on_refresh_timer_tick(self, sender, args):
        if not self._modeless:
            return
        self._request_refresh(force=False)

    def _raise_external(self, event_obj, handler, payload):
        if not self._modeless:
            return False
        if self._external_busy:
            return False
        if event_obj is None or handler is None:
            return False
        self._external_busy = True
        try:
            handler.set_payload(payload or {})
            event_obj.Raise()
            return True
        except Exception:
            self._external_busy = False
            return False

    def _request_refresh(self, force=False):
        if not self._modeless:
            return
        self._raise_external(self._refresh_event, self._refresh_handler, {"force": bool(force)})

    def request_refresh(self, force=False):
        self._request_refresh(force=force)

    def _build_conflict_signature(self, conflicts):
        rows = []
        for item in conflicts or []:
            rows.append((
                item.get("id"),
                item.get("parent_id"),
                item.get("child_id"),
                item.get("param_key"),
                item.get("parent_display"),
                item.get("child_display"),
                bool(item.get("allow_update_child")),
                bool(item.get("allow_update_parent")),
            ))
        rows.sort()
        return tuple(rows)

    def _collect_selection_signature(self, uidoc):
        if uidoc is None:
            return ()
        try:
            elem_ids = list(uidoc.Selection.GetElementIds())
        except Exception:
            return ()
        values = []
        for eid in elem_ids:
            try:
                values.append(int(eid.IntegerValue))
            except Exception:
                continue
        values.sort()
        return tuple(values)

    def _collect_current_decisions(self):
        decisions = {}
        for combo, meta in self._combo_meta.items():
            conflict_id = meta.get("conflict_id")
            if not conflict_id:
                continue
            selection = combo.SelectedItem
            if isinstance(selection, list):
                selection = selection[0] if selection else None
            if not selection:
                continue
            decisions[conflict_id] = selection
        return decisions

    def _refresh_from_doc(self, doc, uidoc=None, force=False):
        if doc is None or self._refresh_callback is None:
            return
        if uidoc is not None:
            current_sel_sig = self._collect_selection_signature(uidoc)
            if (not force) and current_sel_sig == self._selection_signature:
                return
            self._selection_signature = current_sel_sig
        prev_decisions = self._collect_current_decisions()
        try:
            conflicts, param_keys = self._refresh_callback(doc)
        except Exception:
            return
        conflicts = conflicts or []
        param_keys = param_keys or []
        new_sig = self._build_conflict_signature(conflicts)
        if new_sig == self._conflict_signature and list(param_keys) == list(self._param_keys):
            return
        self._conflicts = conflicts
        self._param_keys = param_keys
        self._conflict_signature = new_sig
        self._rebuild_defaults_panel()
        self._build_grid(initial_decisions=prev_decisions)

    def _execute_refresh(self, uiapp, payload):
        uidoc = getattr(uiapp, "ActiveUIDocument", None) if uiapp is not None else None
        doc = getattr(uidoc, "Document", None) if uidoc is not None else None
        force = bool((payload or {}).get("force"))
        self._refresh_from_doc(doc, uidoc=uidoc, force=force)

    def _execute_apply(self, uiapp, payload):
        uidoc = getattr(uiapp, "ActiveUIDocument", None) if uiapp is not None else None
        doc = getattr(uidoc, "Document", None) if uidoc is not None else None
        if doc is None or self._apply_callback is None:
            return
        decisions = (payload or {}).get("decisions") or {}
        if not decisions:
            return
        try:
            self._apply_callback(doc, list(self._conflicts), decisions)
        except Exception:
            return
        self._refresh_from_doc(doc, uidoc=uidoc, force=True)

    def _select_in_doc(self, uidoc, conflict, target):
        if uidoc is None or conflict is None:
            return
        elem_key = "parent_id" if target == "parent" else "child_id"
        raw_id = conflict.get(elem_key)
        if raw_id in (None, ""):
            return
        try:
            elem_id = ElementId(int(raw_id))
        except Exception:
            return
        ids = List[ElementId]()
        ids.Add(elem_id)
        try:
            uidoc.Selection.SetElementIds(ids)
        except Exception:
            return
        try:
            uidoc.ShowElements(ids)
        except Exception:
            pass

    def _execute_select(self, uiapp, payload):
        uidoc = getattr(uiapp, "ActiveUIDocument", None) if uiapp is not None else None
        doc = getattr(uidoc, "Document", None) if uidoc is not None else None
        conflict = (payload or {}).get("conflict")
        target = (payload or {}).get("target")
        if self._select_callback is not None and doc is not None:
            try:
                self._select_callback(doc, conflict, target)
                return
            except Exception:
                pass
        self._select_in_doc(uidoc, conflict, target)

    def _rebuild_defaults_panel(self):
        panel = self.FindName("ParamDefaultsPanel")
        if panel is None:
            return
        panel.Children.Clear()
        self._default_meta = {}
        if not self._param_keys:
            return

        header = TextBlock()
        header.Text = "Apply to all (per parameter):"
        header.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(header)

        for param_key in self._param_keys:
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 0, 0, 4)

            label = TextBlock()
            label.Text = param_key
            label.Width = 360
            label.Margin = Thickness(0, 0, 8, 0)
            row.Children.Add(label)

            combo = ComboBox()
            combo.Width = 220
            combo.Items.Add(LABEL_NO_OVERRIDE)
            combo.Items.Add(LABEL_UPDATE_CHILD)
            combo.Items.Add(LABEL_UPDATE_PARENT)
            combo.Items.Add(LABEL_SKIP)
            combo.SelectedIndex = 0
            combo.SelectionChanged += self._on_default_changed

            self._default_meta[combo] = param_key
            row.Children.Add(combo)
            panel.Children.Add(row)

    def _build_grid(self, initial_decisions=None):
        grid = self.FindName("ConflictGrid")
        if grid is None:
            raise Exception("ConflictGrid not found in XAML.")
        grid.Children.Clear()
        grid.RowDefinitions.Clear()
        grid.ColumnDefinitions.Clear()
        self._combo_meta = {}
        self._combos_by_param = {}
        self._select_meta = {}

        columns = [
            "Parent",
            "Select",
            "Child",
            "Select",
            "Parameter",
            "Parent Value",
            "Child Value",
            "Action",
        ]
        widths = [250, 80, 250, 80, 300, 170, 170, 220]
        for idx, _ in enumerate(columns):
            col_def = ColumnDefinition()
            col_def.Width = GridLength(widths[idx], GridUnitType.Pixel)
            grid.ColumnDefinitions.Add(col_def)

        header_row = RowDefinition()
        header_row.Height = GridLength(28, GridUnitType.Pixel)
        grid.RowDefinitions.Add(header_row)

        for col_idx, title in enumerate(columns):
            cell = TextBlock()
            cell.Text = title
            cell.FontWeight = FontWeights.Bold
            cell.Margin = Thickness(0, 0, 6, 4)
            Grid.SetRow(cell, 0)
            Grid.SetColumn(cell, col_idx)
            grid.Children.Add(cell)

        row_idx = 1
        initial_decisions = initial_decisions or {}
        for conflict in self._conflicts:
            row_def = RowDefinition()
            row_def.Height = GridLength(30, GridUnitType.Pixel)
            grid.RowDefinitions.Add(row_def)

            parent_cell = TextBlock()
            parent_cell.Text = conflict.get("parent_label") or ""
            parent_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(parent_cell, row_idx)
            Grid.SetColumn(parent_cell, 0)
            grid.Children.Add(parent_cell)

            parent_btn = Button()
            parent_btn.Content = "Select"
            parent_btn.Width = 70
            parent_btn.Margin = Thickness(0, 0, 6, 2)
            parent_btn.Click += self._on_select_click
            self._select_meta[parent_btn] = {"conflict": conflict, "target": "parent"}
            Grid.SetRow(parent_btn, row_idx)
            Grid.SetColumn(parent_btn, 1)
            grid.Children.Add(parent_btn)

            child_cell = TextBlock()
            child_cell.Text = conflict.get("child_label") or ""
            child_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(child_cell, row_idx)
            Grid.SetColumn(child_cell, 2)
            grid.Children.Add(child_cell)

            child_btn = Button()
            child_btn.Content = "Select"
            child_btn.Width = 70
            child_btn.Margin = Thickness(0, 0, 6, 2)
            child_btn.Click += self._on_select_click
            self._select_meta[child_btn] = {"conflict": conflict, "target": "child"}
            Grid.SetRow(child_btn, row_idx)
            Grid.SetColumn(child_btn, 3)
            grid.Children.Add(child_btn)

            param_cell = TextBlock()
            param_cell.Text = conflict.get("param_key") or ""
            param_cell.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(param_cell, row_idx)
            Grid.SetColumn(param_cell, 4)
            grid.Children.Add(param_cell)

            parent_val = TextBlock()
            parent_val.Text = conflict.get("parent_display") or ""
            parent_val.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(parent_val, row_idx)
            Grid.SetColumn(parent_val, 5)
            grid.Children.Add(parent_val)

            child_val = TextBlock()
            child_val.Text = conflict.get("child_display") or ""
            child_val.Margin = Thickness(0, 0, 6, 2)
            Grid.SetRow(child_val, row_idx)
            Grid.SetColumn(child_val, 6)
            grid.Children.Add(child_val)

            combo = ComboBox()
            options = []
            label_map = {}
            if conflict.get("allow_update_child"):
                options.append(LABEL_UPDATE_CHILD)
                label_map[LABEL_UPDATE_CHILD] = ACTION_UPDATE_CHILD
            if conflict.get("allow_update_parent"):
                options.append(LABEL_UPDATE_PARENT)
                label_map[LABEL_UPDATE_PARENT] = ACTION_UPDATE_PARENT
            options.append(LABEL_SKIP)
            label_map[LABEL_SKIP] = ACTION_SKIP
            for opt in options:
                combo.Items.Add(opt)
            default_choice = initial_decisions.get(conflict.get("id"))
            if default_choice in options:
                combo.SelectedItem = default_choice
            else:
                combo.SelectedItem = LABEL_SKIP
            if len(options) == 1:
                combo.IsEnabled = False
            Grid.SetRow(combo, row_idx)
            Grid.SetColumn(combo, 7)
            grid.Children.Add(combo)

            param_key = conflict.get("param_key") or ""
            self._combo_meta[combo] = {
                "conflict_id": conflict.get("id"),
                "param_key": param_key,
                "label_map": label_map,
            }
            if param_key:
                self._combos_by_param.setdefault(param_key, []).append(combo)
            row_idx += 1

    def _on_default_changed(self, sender, args):
        param_key = self._default_meta.get(sender)
        if not param_key:
            return
        selection = sender.SelectedItem
        if selection in (None, LABEL_NO_OVERRIDE):
            return
        combos = self._combos_by_param.get(param_key) or []
        for combo in combos:
            if selection in list(combo.Items):
                combo.SelectedItem = selection

    def _on_select_click(self, sender, args):
        meta = self._select_meta.get(sender) or {}
        conflict = meta.get("conflict")
        target = meta.get("target")
        if conflict is None or target not in ("parent", "child"):
            return
        if self._modeless:
            self._raise_external(
                self._select_event,
                self._select_handler,
                {"conflict": conflict, "target": target},
            )
            return
        uidoc = getattr(revit, "uidoc", None)
        self._select_in_doc(uidoc, conflict, target)

    def _on_refresh(self, sender, args):
        if self._modeless:
            self._request_refresh(force=True)

    def _on_apply(self, sender, args):
        decisions = {}
        for combo, meta in self._combo_meta.items():
            selection = combo.SelectedItem
            if isinstance(selection, list):
                selection = selection[0] if selection else None
            if not selection:
                continue
            action = meta["label_map"].get(selection, ACTION_SKIP)
            conflict_id = meta.get("conflict_id")
            if conflict_id:
                decisions[conflict_id] = action

        if not self._modeless:
            self.decisions = decisions
            self.DialogResult = True
            self.Close()
            return

        self._raise_external(
            self._apply_event,
            self._apply_handler,
            {"decisions": decisions},
        )

    def _on_cancel(self, sender, args):
        if self._modeless:
            self.Close()
            return
        self.DialogResult = False
        self.Close()
