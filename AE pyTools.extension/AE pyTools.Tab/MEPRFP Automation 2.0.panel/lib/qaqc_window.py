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
    ComboBox,
    ComboBoxItem,
    Grid,
    TextBlock,
    TextBox,
)
from System.Windows.Media import Brushes  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    Reference,
    Transform,
    XYZ,
)

import placement as _placement
import qaqc_workflow as _qa
import wpf as _wpf


def _linked_element_host_bbox(doc, link_instance_id, linked_element_id):
    """Return ``(min_xyz, max_xyz)`` in HOST coordinates for the
    linked element, or ``None`` if any step fails.

    The linked element's ``get_BoundingBox`` returns coordinates in
    the linked doc's frame; we transform all 8 corners through the
    RevitLinkInstance's total transform to lift them into the host
    coordinate system, then take the axis-aligned union. That bbox
    is what ``UIView.ZoomAndCenterRectangle`` needs to frame the
    specific linked element regardless of how the link is rotated
    or offset relative to the host.
    """
    if link_instance_id is None or linked_element_id is None:
        return None
    try:
        link_inst = doc.GetElement(ElementId(int(link_instance_id)))
    except Exception:
        return None
    if link_inst is None:
        return None
    try:
        link_doc = link_inst.GetLinkDocument()
    except Exception:
        link_doc = None
    if link_doc is None:
        return None
    try:
        linked_elem = link_doc.GetElement(ElementId(int(linked_element_id)))
    except Exception:
        linked_elem = None
    if linked_elem is None:
        return None
    try:
        bbox = linked_elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None
    try:
        transform = link_inst.GetTotalTransform()
    except Exception:
        return None
    corners_in_link = (
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
    )
    try:
        host_corners = [transform.OfPoint(p) for p in corners_in_link]
    except Exception:
        return None
    xs = [c.X for c in host_corners]
    ys = [c.Y for c in host_corners]
    zs = [c.Z for c in host_corners]
    return XYZ(min(xs), min(ys), min(zs)), XYZ(max(xs), max(ys), max(zs))


def _zoom_active_uiview_to_rect(uidoc, min_pt, max_pt, pad_ft=4.0):
    """Zoom the *active* UIView to ``[min_pt..max_pt]`` (host coords),
    padded by ``pad_ft`` so the element isn't flush against the edge.

    Returns True on success. The caller should fall back to
    ``ShowElements`` if this returns False (no active UIView, or the
    view type doesn't support ``ZoomAndCenterRectangle``).
    """
    if uidoc is None or min_pt is None or max_pt is None:
        return False
    pad = float(pad_ft)
    padded_min = XYZ(min_pt.X - pad, min_pt.Y - pad, min_pt.Z - pad)
    padded_max = XYZ(max_pt.X + pad, max_pt.Y + pad, max_pt.Z + pad)
    try:
        active_view_id = uidoc.Document.ActiveView.Id
    except Exception:
        return False
    try:
        ui_views = uidoc.GetOpenUIViews()
    except Exception:
        return False
    for uiview in ui_views:
        try:
            if uiview.ViewId != active_view_id:
                continue
        except Exception:
            continue
        try:
            uiview.ZoomAndCenterRectangle(padded_min, padded_max)
            return True
        except Exception:
            return False
    return False


def _linked_reference(doc, link_instance_id, linked_element_id):
    """Build a host-doc ``Reference`` that points at a specific element
    inside a linked document.

    Two-step construction (matches Revit API requirement):
      1. Resolve the RevitLinkInstance in the host doc, get its
         linked Document.
      2. Make a plain ``Reference(linked_elem)`` in the linked doc,
         then call ``CreateLinkReference(link_instance)`` on it to
         lift it into host coordinates.

    Returns the resulting Reference, or ``None`` if any step fails
    (the caller should fall back to selecting the link instance
    instead).
    """
    if link_instance_id is None or linked_element_id is None:
        return None
    try:
        link_inst = doc.GetElement(ElementId(int(link_instance_id)))
    except Exception:
        return None
    if link_inst is None:
        return None
    try:
        link_doc = link_inst.GetLinkDocument()
    except Exception:
        link_doc = None
    if link_doc is None:
        return None
    try:
        linked_elem = link_doc.GetElement(ElementId(int(linked_element_id)))
    except Exception:
        linked_elem = None
    if linked_elem is None:
        return None
    try:
        ref = Reference(linked_elem)
    except Exception:
        return None
    try:
        return ref.CreateLinkReference(link_inst)
    except Exception:
        return None


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "QaqcWindow.xaml",
)


class _FilterRule(object):
    """One row of the multi-filter UI.

    ``operator`` is either ``"contains"`` or ``"does_not_contain"``.
    ``text`` is the substring to match (case-insensitive). An empty
    ``text`` makes the rule a no-op — the filter pass skips it rather
    than treating it as "hide everything".
    """

    __slots__ = ("operator", "text")

    def __init__(self, operator="contains", text=""):
        self.operator = operator
        self.text = text or ""


class QaqcController(object):

    def __init__(self, doc, profile_data, uidoc=None):
        self.doc = doc
        self.uidoc = uidoc
        self.profile_data = profile_data
        self._cat_filter_checks = {}
        self._handlers = []  # retain refs against pythonnet GC
        # Cache the last QaqcResult so the filter rules can re-render
        # without re-running the audit.
        self._last_result = None
        # List of ``_FilterRule`` instances. Empty = no filtering.
        # Multiple rules combine with AND: every "contains" rule must
        # match, every "does_not_contain" must not match.
        self._filter_rules = []
        self._filter_row_handlers = []
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_category_filters()
        # Start with one empty filter row so the user sees the shape.
        self._add_filter_rule()
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
        self.filter_rules_panel = f("FilterRulesPanel")
        self.add_filter_btn = f("AddFilterButton")
        self.clear_filters_btn = f("ClearFiltersButton")
        self.match_count_label = f("MatchCountLabel")

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
        self._h_add_filter = self._delegate(
            "add-filter", lambda s, e: self._on_add_filter(),
        )
        self._h_clear_filters = self._delegate(
            "clear-filters", lambda s, e: self._on_clear_filters(),
        )
        self.add_filter_btn.Click += self._h_add_filter
        self.clear_filters_btn.Click += self._h_clear_filters

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

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
            self.match_count_label.Text = ""
            self._last_result = None
            return
        self._set_status("Auditing...")
        result = _qa.run_audit(self.doc, self.profile_data, categories=cats)
        self._last_result = result
        self._render(result)
        parts = []
        for cat in _qa.CAT_ALL:
            n = result.counts.get(cat, 0)
            if n:
                parts.append("{}={}".format(cat, n))
        if parts:
            self.summary_label.Text = "Findings: " + ", ".join(parts)
            self._set_status(
                "{} finding(s). Use the row buttons to Select / Zoom / "
                "Fix. Cat G rows have a Place button.".format(
                    len(result.findings)
                )
            )
        else:
            self.summary_label.Text = "Clean — no findings."
            self._set_status("Audit complete; nothing to fix.")

    # ----- filter rules ---------------------------------------------
    #
    # A filter rule is one ``(operator, text)`` pair. Operators are
    # ``"contains"`` and ``"does_not_contain"`` — both substring
    # matches (case-insensitive) against the row's searchable text
    # fields. Empty text means the rule is ignored, so adding a row
    # and leaving it blank doesn't accidentally hide everything.
    #
    # Multiple rules combine with AND. Examples:
    #   * "Contains: HEB"               + "Does not contain: vendor"
    #     → rows that mention HEB but not "vendor"
    #   * "Contains: Soda_Merch"        + "Contains: id 92"
    #     → rows that mention Soda_Merch AND an id starting with 92

    def _on_add_filter(self):
        self._add_filter_rule()
        self._refilter("Added filter rule.")

    def _on_clear_filters(self):
        self._filter_rules = []
        # Keep one blank row visible so the panel doesn't look broken.
        self._add_filter_rule()
        self._refilter("Cleared filter rules.")

    def _add_filter_rule(self, operator="contains", text=""):
        rule = _FilterRule(operator, text)
        self._filter_rules.append(rule)
        self._render_filter_rules()

    def _remove_filter_rule(self, rule):
        if rule in self._filter_rules:
            self._filter_rules.remove(rule)
        if not self._filter_rules:
            # Always keep one row visible.
            self._add_filter_rule()
        else:
            self._render_filter_rules()
        self._refilter()

    def _render_filter_rules(self):
        """Rebuild the dynamic filter-rule rows from ``self._filter_rules``."""
        self.filter_rules_panel.Children.Clear()
        # Releasing prior row handlers lets pythonnet GC them.
        self._filter_row_handlers = []
        for rule in self._filter_rules:
            self.filter_rules_panel.Children.Add(self._build_filter_row(rule))

    def _build_filter_row(self, rule):
        grid = Grid()
        grid.Margin = Thickness(0, 2, 0, 2)
        # 3 columns: operator combobox (140), text box (stretch),
        # remove button (36).
        col_widths = (
            GridLength(140),
            GridLength(1.0, GridUnitType.Star),
            GridLength(36),
        )
        for w in col_widths:
            col = ColumnDefinition()
            col.Width = w
            grid.ColumnDefinitions.Add(col)

        combo = ComboBox()
        combo.Margin = Thickness(0, 0, 6, 0)
        item_contains = ComboBoxItem()
        item_contains.Content = "Contains"
        item_not = ComboBoxItem()
        item_not.Content = "Does not contain"
        combo.Items.Add(item_contains)
        combo.Items.Add(item_not)
        combo.SelectedIndex = 0 if rule.operator == "contains" else 1
        combo.ToolTip = (
            "Contains: row must match the text below.\n"
            "Does not contain: row must NOT match the text below."
        )
        Grid.SetColumn(combo, 0)
        grid.Children.Add(combo)

        tb = TextBox()
        tb.VerticalContentAlignment = VerticalAlignment.Center
        tb.Margin = Thickness(0, 0, 6, 0)
        tb.Text = rule.text or ""
        tb.ToolTip = (
            "Substring matched (case-insensitive) against the row's "
            "profile name, profile id, LED label, LED id, message, "
            "element id, and category."
        )
        Grid.SetColumn(tb, 1)
        grid.Children.Add(tb)

        remove_btn = Button()
        remove_btn.Content = "✕"  # ×
        remove_btn.Width = 28
        remove_btn.ToolTip = "Remove this filter rule."
        Grid.SetColumn(remove_btn, 2)
        grid.Children.Add(remove_btn)

        h_combo = (
            lambda s, e, r=rule, c=combo:
            self._safe(lambda: self._on_rule_combo(r, c), "filter-op")
        )
        h_text = (
            lambda s, e, r=rule, t=tb:
            self._safe(lambda: self._on_rule_text(r, t), "filter-text")
        )
        h_remove = RoutedEventHandler(
            lambda s, e, r=rule: self._safe(
                lambda: self._remove_filter_rule(r), "filter-remove",
            )
        )
        combo.SelectionChanged += h_combo
        tb.TextChanged += h_text
        remove_btn.Click += h_remove
        # Retain refs against pythonnet GC.
        self._filter_row_handlers.extend([h_combo, h_text, h_remove])

        return grid

    def _on_rule_combo(self, rule, combo):
        rule.operator = "contains" if combo.SelectedIndex == 0 else "does_not_contain"
        self._refilter()

    def _on_rule_text(self, rule, text_box):
        try:
            rule.text = text_box.Text or ""
        except Exception:
            rule.text = ""
        self._refilter()

    def _refilter(self, status_msg=None):
        if self._last_result is not None:
            self._render(self._last_result)
        if status_msg:
            self._set_status(status_msg)

    # ----- finding match check --------------------------------------

    def _finding_searchable_values(self, finding):
        """Return the lower-cased text fields a filter rule tests
        against. Concatenated only at the lookup site so each rule's
        substring search is O(field) rather than O(joined-string)."""
        return (
            (finding.profile_name or "").lower(),
            (finding.profile_id or "").lower(),
            (finding.led_label or "").lower(),
            (finding.led_id or "").lower(),
            (finding.message or "").lower(),
            (str(finding.element_id).lower()
             if finding.element_id is not None else ""),
            (finding.category or "").lower(),
            (finding.category_label or "").lower(),
        )

    def _finding_matches_filter(self, finding):
        """Apply every non-empty filter rule with AND combination.

        Empty-text rules are skipped (treating "Add filter" + leave-
        blank as a no-op rather than as "hide everything"). For each
        active rule we check whether any searchable field contains
        the rule's text. ``contains`` requires a hit; ``does_not_
        contain`` requires no hit. Failing either condition rejects
        the finding immediately — short-circuits on first miss.
        """
        active_rules = [
            r for r in self._filter_rules
            if (r.text or "").strip()
        ]
        if not active_rules:
            return True
        searchables = self._finding_searchable_values(finding)
        for rule in active_rules:
            needle = rule.text.strip().lower()
            matched = any(needle in v for v in searchables)
            if rule.operator == "contains":
                if not matched:
                    return False
            else:  # does_not_contain
                if matched:
                    return False
        return True

    def _render(self, result):
        self.findings_panel.Children.Clear()
        # Drop stale row-button handlers — they captured findings from
        # the prior render and prevent pythonnet from GCing the old
        # rows otherwise.
        self._handlers = []
        # Group by category for visual order.
        by_cat = {}
        total_shown = 0
        for f in result.findings:
            if not self._finding_matches_filter(f):
                continue
            by_cat.setdefault(f.category, []).append(f)
            total_shown += 1
        for cat in _qa.CAT_ALL:
            findings = by_cat.get(cat) or []
            if not findings:
                continue
            self.findings_panel.Children.Add(
                self._section_header(cat, len(findings))
            )
            for finding in findings:
                self.findings_panel.Children.Add(self._row(finding))
        active_n = sum(
            1 for r in self._filter_rules if (r.text or "").strip()
        )
        if active_n:
            self.match_count_label.Text = "{} of {} finding(s) match.".format(
                total_shown, len(result.findings),
            )
        else:
            self.match_count_label.Text = (
                "{} finding(s).".format(len(result.findings))
                if result.findings else ""
            )
        if active_n and total_shown == 0:
            tb = TextBlock()
            tb.Text = "No findings match the current filter rules."
            tb.Margin = Thickness(0, 8, 0, 4)
            tb.Foreground = Brushes.Gray
            self.findings_panel.Children.Add(tb)

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
        # 7 columns total: 1 left spacer + 2 stretched text + 4 fixed
        # buttons (Select, Zoom, Fix, Place). Place is only meaningful
        # for Cat G findings; we render it on every row but enable it
        # selectively to keep button alignment stable across categories.
        for w in (0.0, 2.5, 4.0, 0.0, 0.0, 0.0, 0.0):
            col = ColumnDefinition()
            if w == 0.0:
                col.Width = GridLength(80)
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

        place_btn = self._small_button("Place")
        # Only Cat G findings get a meaningful Place action — they
        # carry the parent element id + the matching profile id, which
        # is exactly what ``placement.execute_placement`` needs.
        place_btn.IsEnabled = (
            finding.category == _qa.CAT_G
            and finding.profile_id is not None
            and (
                finding.element_id is not None
                or getattr(finding, "linked_element_id", None) is not None
            )
        )
        place_btn.ToolTip = (
            "Place the matching profile against this parent element."
            if place_btn.IsEnabled else
            "Place is only available for Cat G findings."
        )
        Grid.SetColumn(place_btn, 6)
        grid.Children.Add(place_btn)

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
        h_place = RoutedEventHandler(
            lambda s, e, fnd=finding, btn=place_btn: self._on_place_profile(fnd, btn)
        )
        select_btn.Click += h_select
        zoom_btn.Click += h_zoom
        fix_btn.Click += h_fix
        place_btn.Click += h_place
        self._handlers.extend([h_select, h_zoom, h_fix, h_place])

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

    def _select_finding(self, finding):
        """Set the host doc's selection to the finding's target.

        For host-doc targets this is just SetElementIds. For linked
        targets we prefer the host-coord ``Reference`` produced by
        ``Reference.CreateLinkReference`` so the specific linked
        element gets highlighted, not the whole RevitLinkInstance.
        Falls back to selecting the link instance if the reference
        construction fails.

        Returns ``True`` on success.
        """
        link_inst_id = getattr(finding, "link_instance_id", None)
        linked_elem_id = getattr(finding, "linked_element_id", None)
        if link_inst_id is not None and linked_elem_id is not None:
            ref = _linked_reference(self.doc, link_inst_id, linked_elem_id)
            if ref is not None:
                refs = _NetList[Reference]()
                refs.Add(ref)
                self.uidoc.Selection.SetReferences(refs)
                return True
        if finding.element_id is None:
            return False
        ids = _NetList[ElementId]()
        ids.Add(ElementId(int(finding.element_id)))
        self.uidoc.Selection.SetElementIds(ids)
        return True

    def _on_select(self, finding):
        if self.uidoc is None or (
            finding.element_id is None
            and getattr(finding, "linked_element_id", None) is None
        ):
            self._set_status("No active uidoc; cannot select.")
            return
        try:
            if not self._select_finding(finding):
                self._set_status("Select failed: nothing to target.")
                return
            link_inst_id = getattr(finding, "link_instance_id", None)
            if link_inst_id is not None:
                self._set_status(
                    "Selected linked element {} (in link {}).".format(
                        getattr(finding, "linked_element_id", "?"),
                        link_inst_id,
                    )
                )
            else:
                self._set_status("Selected element {}.".format(finding.element_id))
        except Exception as exc:
            self._set_status("Select failed: {}".format(exc))

    def _on_zoom(self, finding):
        if self.uidoc is None or (
            finding.element_id is None
            and getattr(finding, "linked_element_id", None) is None
        ):
            self._set_status("No active uidoc; cannot zoom.")
            return
        try:
            if not self._select_finding(finding):
                self._set_status("Zoom failed: nothing to target.")
                return

            link_inst_id = getattr(finding, "link_instance_id", None)
            linked_elem_id = getattr(finding, "linked_element_id", None)

            # Linked target — zoom to the linked element's
            # host-coord bbox, NOT the link instance's overall
            # extent. ``ShowElements`` only takes host-doc element
            # ids and would frame the whole link, so we compute
            # the linked element's transformed bbox and feed it to
            # the active UIView's ``ZoomAndCenterRectangle``. Falls
            # back to ShowElements(link_inst.Id) if the bbox
            # computation or UIView lookup fails (e.g., view type
            # doesn't support rectangle zoom).
            if link_inst_id is not None and linked_elem_id is not None:
                bbox = _linked_element_host_bbox(
                    self.doc, link_inst_id, linked_elem_id,
                )
                zoomed = False
                if bbox is not None:
                    zoomed = _zoom_active_uiview_to_rect(
                        self.uidoc, bbox[0], bbox[1],
                    )
                if not zoomed:
                    ids = _NetList[ElementId]()
                    ids.Add(ElementId(int(link_inst_id)))
                    self.uidoc.ShowElements(ids)
                self._set_status(
                    "Zoomed to linked element {} (within link {}).".format(
                        linked_elem_id, link_inst_id,
                    )
                )
                return

            # Host target — keep the existing ShowElements path.
            ids = _NetList[ElementId]()
            ids.Add(ElementId(int(finding.element_id)))
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
    # Cat G: place the matching profile against the missing parent.
    # The finding carries the parent's element id (host) or
    # (link_instance_id, linked_element_id) pair (linked), plus the
    # matched profile_id. We resolve the parent, lift its location +
    # rotation into host coordinates, build a placement.Target /
    # Match, and run execute_placement in its own transaction.
    # ----------------------------------------------------------------

    def _resolve_finding_parent(self, finding):
        """Return ``(elem, transform_or_None, source, link_inst,
        link_elem_id_int)`` for the parent referenced by a Cat G
        finding, or ``None`` if anything along the way fails."""
        link_inst_id = getattr(finding, "link_instance_id", None)
        linked_elem_id = getattr(finding, "linked_element_id", None)
        if link_inst_id is not None and linked_elem_id is not None:
            try:
                link_inst = self.doc.GetElement(ElementId(int(link_inst_id)))
            except Exception:
                return None
            if link_inst is None:
                return None
            try:
                link_doc = link_inst.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                return None
            try:
                elem = link_doc.GetElement(ElementId(int(linked_elem_id)))
            except Exception:
                elem = None
            if elem is None:
                return None
            try:
                transform = link_inst.GetTotalTransform()
            except Exception:
                transform = Transform.Identity
            return (
                elem, transform, _placement.SOURCE_LINKED_REVIT,
                link_inst, int(linked_elem_id),
            )
        # Host-doc parent.
        if finding.element_id is None:
            return None
        try:
            elem = self.doc.GetElement(ElementId(int(finding.element_id)))
        except Exception:
            return None
        if elem is None:
            return None
        return (
            elem, Transform.Identity, _placement.SOURCE_HOST_MODEL,
            None, int(finding.element_id),
        )

    def _find_profile_by_id(self, profile_id):
        if not profile_id:
            return None
        for p in self.profile_data.get("equipment_definitions") or []:
            if isinstance(p, dict) and p.get("id") == profile_id:
                return p
        return None

    def _on_place_profile(self, finding, btn):
        from pyrevit import revit
        if finding.category != _qa.CAT_G:
            self._set_status("Place is only available for Cat G findings.")
            return
        profile = self._find_profile_by_id(finding.profile_id)
        if profile is None:
            self._set_status(
                "Place failed: profile {!r} not found in active YAML.".format(
                    finding.profile_id
                )
            )
            return
        resolved = self._resolve_finding_parent(finding)
        if resolved is None:
            self._set_status(
                "Place failed: could not resolve parent element {}.".format(
                    finding.element_id
                )
            )
            return
        elem, transform, source, link_inst, link_elem_id_int = resolved
        family_name = _placement._element_family_name(elem)
        local_pt = _placement._element_location_point(elem)
        if local_pt is None:
            self._set_status(
                "Place failed: parent has no resolvable location point."
            )
            return
        try:
            world_pt = transform.OfPoint(local_pt)
        except Exception:
            world_pt = local_pt
        try:
            rot_deg = _placement._element_rotation_deg(elem, transform)
        except Exception:
            rot_deg = 0.0
        target = _placement.Target(
            source=source,
            name=family_name,
            world_pt=(world_pt.X, world_pt.Y, world_pt.Z),
            rotation_deg=rot_deg,
            link_inst=link_inst,
            link_elem_id=link_elem_id_int,
        )
        match = _placement.Match(target, profile)
        options = _placement.PlacementOptions(
            skip_already_placed=True,
            transaction_action="QAQC Cat G place ({})".format(
                finding.profile_name or "?"
            ),
        )
        try:
            with revit.Transaction(
                "QAQC Cat G place ({})".format(finding.profile_name or "?"),
                doc=self.doc,
            ):
                result = _placement.execute_placement(
                    self.doc, [match], options,
                )
        except Exception as exc:
            self._set_status("Place failed: {}".format(exc))
            return
        if result.placed_fixture_count == 0 and result.placed_annotation_count == 0:
            self._set_status(
                "[G] Nothing placed against parent {} — "
                "see pyRevit output for any LED-level skips.".format(
                    link_elem_id_int
                )
            )
            return
        btn.IsEnabled = False
        btn.Content = "Placed"
        self._set_status(
            "[G] Placed {} fixture(s){} against parent {}.".format(
                result.placed_fixture_count,
                ", {} annotation(s)".format(result.placed_annotation_count)
                if result.placed_annotation_count else "",
                link_elem_id_int,
            )
        )

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
