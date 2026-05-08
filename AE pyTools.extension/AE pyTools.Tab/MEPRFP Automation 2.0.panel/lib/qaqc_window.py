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

from Autodesk.Revit.DB import ElementId, Reference, XYZ  # noqa: E402

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
    # Misc
    # ----------------------------------------------------------------

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data, uidoc=None):
    return QaqcController(doc, profile_data, uidoc=uidoc).show()
