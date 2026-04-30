# -*- coding: utf-8 -*-
"""
Tree-view UI for the synced-relationship audit (option C).

Hierarchy: Profile -> LED -> Conflict. Each conflict row carries a
ComboBox per ``Conflict.UPDATE_CHILD / UPDATE_PARENT / SKIP`` choice.
On Apply, the controller calls ``sync_audit.apply_resolution`` for each
non-skip choice inside one Revit transaction.
"""

import os

import clr  # noqa: F401

from System.Windows.Controls import (  # noqa: E402
    ComboBox,
    ComboBoxItem,
    Grid,
    ColumnDefinition,
    StackPanel,
    TextBlock,
    TreeViewItem,
)
from System.Windows import GridLength, GridUnitType, Thickness  # noqa: E402

import sync_audit
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "SyncAuditWindow.xaml"
)


# Maps display label -> action constant.
_ACTION_LABELS = [
    ("Update child", sync_audit.Conflict.UPDATE_CHILD),
    ("Update parent", sync_audit.Conflict.UPDATE_PARENT),
    ("Skip", sync_audit.Conflict.SKIP),
]


class _ConflictRow(object):
    """One leaf in the tree. Holds the ComboBox so we can read its choice on Apply."""
    def __init__(self, conflict, combo):
        self.conflict = conflict
        self.combo = combo

    @property
    def chosen_action(self):
        item = self.combo.SelectedItem
        if item is None:
            return sync_audit.Conflict.SKIP
        return item.Tag


class SyncAuditController(object):

    def __init__(self, doc, profile_data, transaction_factory):
        """``transaction_factory`` is a callable returning a context
        manager (e.g. ``lambda: revit.Transaction("...", doc=doc)``).
        We let the caller own transaction semantics.
        """
        self.doc = doc
        self.profile_data = profile_data
        self.transaction_factory = transaction_factory
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = []
        self._wire_controls()
        self.refresh()

    # ---- wiring ------------------------------------------------------

    def _wire_controls(self):
        f = self.window.FindName
        self.tree = f("ConflictTree")
        self.summary = f("SummaryLabel")
        self.bulk_combo = f("BulkActionCombo")
        self.detail_box = f("DetailBox")
        # Retain handler refs so pythonnet's GC can't drop them.
        self._h_refresh = lambda s, e: self._on_refresh(s, e)
        self._h_apply = lambda s, e: self._on_apply(s, e)
        self._h_close = lambda s, e: self._on_close(s, e)
        self._h_bulk = lambda s, e: self._on_bulk(s, e)
        self._h_tree_select = lambda s, e: self._on_tree_select(s, e)
        f("RefreshButton").Click += self._h_refresh
        f("ApplyButton").Click += self._h_apply
        f("CloseButton").Click += self._h_close
        self.bulk_combo.SelectionChanged += self._h_bulk
        self.tree.SelectedItemChanged += self._h_tree_select

    # ---- public ------------------------------------------------------

    def show(self):
        self.window.ShowDialog()

    def refresh(self):
        conflicts = sync_audit.detect_conflicts(self.doc, self.profile_data)
        self._populate_tree(conflicts)
        self.summary.Text = "{} conflict(s) across {} profile(s)".format(
            len(conflicts),
            len({c.profile_id for c in conflicts}),
        )

    # ---- tree --------------------------------------------------------

    def _populate_tree(self, conflicts):
        self.tree.Items.Clear()
        self._rows = []
        # Group: profile -> led -> [conflicts]
        by_profile = {}
        for c in conflicts:
            by_profile.setdefault(
                (c.profile_id, c.profile_name), {}
            ).setdefault((c.led_id, c.led_label), []).append(c)

        for (profile_id, profile_name), led_groups in by_profile.items():
            profile_node = TreeViewItem()
            profile_node.Header = "{}  ({})  -  {} conflict(s)".format(
                profile_name, profile_id,
                sum(len(v) for v in led_groups.values()),
            )
            profile_node.IsExpanded = True
            for (led_id, led_label), led_conflicts in led_groups.items():
                led_node = TreeViewItem()
                led_node.Header = "{}  ({})  -  {} conflict(s)".format(
                    led_label, led_id, len(led_conflicts)
                )
                led_node.IsExpanded = True
                for conflict in led_conflicts:
                    led_node.Items.Add(self._build_conflict_node(conflict))
                profile_node.Items.Add(led_node)
            self.tree.Items.Add(profile_node)

    def _build_conflict_node(self, conflict):
        node = TreeViewItem()
        node.IsExpanded = False
        node.Tag = ("conflict", conflict)
        grid = Grid()
        for w in (3, 2, 2, 2):
            col = ColumnDefinition()
            col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        # Compose a label that already shows what the directive references,
        # so users see the source even before they click into the detail panel.
        if conflict.kind == "parent":
            ref_text = "(parent.{})".format(conflict.target_param_name or "?")
        elif conflict.kind == "sibling":
            ref_text = "(sibling.{})".format(conflict.target_param_name or "?")
        else:
            ref_text = ""
        param = TextBlock()
        param.Text = "{}  {}".format(conflict.parameter_name, ref_text)
        param.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(param, 0)
        grid.Children.Add(param)

        actual = TextBlock()
        actual.Text = "Actual: {}".format(_short(conflict.actual_value))
        actual.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(actual, 1)
        grid.Children.Add(actual)

        expected = TextBlock()
        expected.Text = "Expected: {}".format(_short(conflict.expected_value))
        expected.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(expected, 2)
        grid.Children.Add(expected)

        combo = ComboBox()
        combo.Margin = Thickness(0)
        for label, action in _ACTION_LABELS:
            item = ComboBoxItem()
            item.Content = label
            item.Tag = action
            combo.Items.Add(item)
        combo.SelectedIndex = 2  # default = skip
        Grid.SetColumn(combo, 3)
        grid.Children.Add(combo)

        node.Header = grid
        self._rows.append(_ConflictRow(conflict, combo))
        return node

    # ---- handlers ----------------------------------------------------

    def _on_refresh(self, sender, e):
        self.refresh()
        self._render_detail(None)

    def _on_close(self, sender, e):
        self.window.Close()

    def _on_tree_select(self, sender, e):
        item = self.tree.SelectedItem
        if item is None:
            self._render_detail(None)
            return
        tag = getattr(item, "Tag", None)
        if isinstance(tag, tuple) and len(tag) >= 2 and tag[0] == "conflict":
            self._render_detail(tag[1])
        else:
            self._render_detail(None)

    # ---- detail panel -----------------------------------------------

    def _render_detail(self, conflict):
        if not hasattr(self, "detail_box") or self.detail_box is None:
            return
        if conflict is None:
            self.detail_box.Text = "Select a conflict in the tree to see its comparison."
            return

        # Resolve the actual elements involved.
        from Autodesk.Revit.DB import ElementId
        try:
            child_id = ElementId(int(conflict.element_id)) if conflict.element_id else None
        except Exception:
            child_id = None
        try:
            target_id = ElementId(int(conflict.target_element_id)) if conflict.target_element_id else None
        except Exception:
            target_id = None

        child_elem = self.doc.GetElement(child_id) if child_id else None
        target_elem = self.doc.GetElement(target_id) if target_id else None

        kind = conflict.kind or "?"
        ref_label = "parent" if kind == "parent" else ("sibling" if kind == "sibling" else kind)

        lines = []
        lines.append("SELECTED CONFLICT")
        lines.append("=================")
        lines.append("Profile:        {}  ({})".format(
            conflict.profile_name or "?", conflict.profile_id or "?"))
        lines.append("LED:            {}  ({})".format(
            conflict.led_label or "?", conflict.led_id or "?"))
        lines.append("Parameter:      {}".format(conflict.parameter_name or "?"))
        lines.append("Directive kind: {}".format(kind))
        lines.append("Actual (child): {}".format(_short(conflict.actual_value)))
        lines.append("Expected:       {}".format(_short(conflict.expected_value)))
        if conflict.target_param_name:
            lines.append("Source:         {}.{}{}".format(
                ref_label,
                conflict.target_param_name,
                "  (id {})".format(conflict.target_element_id)
                if conflict.target_element_id is not None else "",
            ))
        lines.append("")

        # ---- child parameters ---------------------------------------
        lines.append("CHILD ELEMENT  (id {})".format(
            conflict.element_id if conflict.element_id is not None else "?"))
        lines.append("=" * 60)
        if child_elem is None:
            lines.append("  (element not found in the active document)")
        else:
            child_params = _collect_params(child_elem)
            if not child_params:
                lines.append("  (no parameters)")
            else:
                lines.append(_format_params(
                    child_params,
                    highlight={conflict.parameter_name: " *** mismatch ***"},
                ))
        lines.append("")

        # ---- referenced (parent / sibling) parameters ----------------
        ref_label_upper = "PARENT ELEMENT" if kind == "parent" else \
                          "SIBLING ELEMENT" if kind == "sibling" else "TARGET ELEMENT"
        lines.append("{}  (id {})".format(
            ref_label_upper,
            conflict.target_element_id
            if conflict.target_element_id is not None else "?",
        ))
        lines.append("=" * 60)
        if target_elem is None:
            lines.append("  (element not found in the active document)")
        else:
            target_params = _collect_params(target_elem)
            if not target_params:
                lines.append("  (no parameters)")
            else:
                lines.append(_format_params(
                    target_params,
                    highlight={
                        conflict.target_param_name:
                            " *** referenced by directive ***"
                    } if conflict.target_param_name else {},
                ))

        self.detail_box.Text = "\n".join(lines)

    def _on_bulk(self, sender, e):
        idx = self.bulk_combo.SelectedIndex
        # 0 = placeholder, 1 = update child, 2 = update parent, 3 = skip
        if idx <= 0:
            return
        target = (
            sync_audit.Conflict.UPDATE_CHILD,
            sync_audit.Conflict.UPDATE_PARENT,
            sync_audit.Conflict.SKIP,
        )[idx - 1]
        for row in self._rows:
            for i in range(row.combo.Items.Count):
                if row.combo.Items[i].Tag == target:
                    row.combo.SelectedIndex = i
                    break

    def _on_apply(self, sender, e):
        applied = 0
        skipped = 0
        with self.transaction_factory():
            for row in self._rows:
                ok = sync_audit.apply_resolution(
                    self.doc, row.conflict, row.chosen_action
                )
                if ok:
                    applied += 1
                else:
                    skipped += 1
        self.refresh()
        self.summary.Text = "Applied {}, skipped {}".format(applied, skipped)


def _short(value):
    if value is None:
        return "(empty)"
    s = str(value)
    if len(s) > 40:
        return s[:37] + "..."
    return s


def _collect_params(elem):
    """Return ``[(name, value_string), ...]`` for every parameter on an
    element. Sorted alphabetically by name. Includes empty-value
    parameters so the user sees the full picture."""
    out = []
    if elem is None:
        return out
    seen = set()
    try:
        params_iter = elem.Parameters
    except Exception:
        return out
    for p in params_iter:
        if p is None:
            continue
        try:
            name = p.Definition.Name
        except Exception:
            continue
        if not name or name in seen:
            continue
        seen.add(name)
        value = None
        try:
            value = p.AsValueString()
        except Exception:
            value = None
        if value is None:
            try:
                value = p.AsString()
            except Exception:
                value = None
        out.append((name, "" if value is None else str(value)))
    out.sort(key=lambda nv: nv[0].lower())
    return out


def _format_params(name_value_pairs, highlight=None):
    """Render ``[(name, value), ...]`` as aligned text. ``highlight`` is
    ``{name: marker_string}`` — names matching get an inline marker."""
    if not name_value_pairs:
        return "  (no parameters)"
    highlight = highlight or {}
    name_width = max(len(n) for n, _ in name_value_pairs)
    name_width = min(name_width, 40)  # don't run away on weird names
    lines = []
    for name, value in name_value_pairs:
        marker = highlight.get(name, "")
        lines.append("  {n:<{w}}  {v}{m}".format(
            n=name[:name_width],
            w=name_width,
            v=value if value else "(empty)",
            m=marker,
        ))
    return "\n".join(lines)


def show_modal(doc, profile_data, transaction_factory):
    SyncAuditController(doc, profile_data, transaction_factory).show()
