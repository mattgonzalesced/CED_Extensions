# -*- coding: utf-8 -*-
"""Modal dialog for Hide Existing Profiles."""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import RoutedEventHandler  # noqa: E402

import hide_profiles_workflow as _hp
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "HideProfilesWindow.xaml",
)


class HideProfilesController(object):

    def __init__(self, doc, profile_data):
        self.doc = doc
        self.profile_data = profile_data
        self.profiles = list(profile_data.get("equipment_definitions") or [])
        self.committed = False
        self._last_result = None
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_filters()
        self._wire_events()
        self._set_status("Pick filters, then 'Duplicate view + hide'.")

    def _lookup_controls(self):
        f = self.window.FindName
        self.category_list = f("CategoryList")
        self.profile_list = f("ProfileList")
        self.hide_btn = f("HideButton")
        self.close_btn = f("CloseButton")
        self.status_label = f("StatusLabel")

    def _wire_events(self):
        def safe(label, fn):
            def wrapped(s, e):
                try:
                    self._set_status("[{}] running...".format(label))
                    fn(s, e)
                except Exception as exc:
                    self._set_status("[{}] error: {}".format(label, exc))
                    raise
            return RoutedEventHandler(wrapped)

        self._h_hide = safe("hide", lambda s, e: self._on_hide(s, e))
        self._h_close = safe("close", lambda s, e: self.window.Close())
        self.hide_btn.Click += self._h_hide
        self.close_btn.Click += self._h_close

    def _populate_filters(self):
        cats = sorted({
            (p.get("parent_filter") or {}).get("category") or ""
            for p in self.profiles if isinstance(p, dict)
        })
        cats = [c for c in cats if c]
        self.category_list.Items.Clear()
        for c in cats:
            self.category_list.Items.Add(c)
        self.profile_list.Items.Clear()
        for p in self.profiles:
            self.profile_list.Items.Add(
                "{}  ({})".format(p.get("name") or "(unnamed)", p.get("id") or "?")
            )

    def _selected_profile_ids(self):
        ids = set()
        for label in self.profile_list.SelectedItems:
            label = str(label)
            if "(" in label and label.endswith(")"):
                ids.add(label.rsplit("(", 1)[1].rstrip(")"))
        return ids or None

    def _selected_categories(self):
        out = {str(item) for item in self.category_list.SelectedItems}
        return out or None

    def _on_hide(self, sender, e):
        from pyrevit import revit
        active_view = self.doc.ActiveView
        if active_view is None:
            self._set_status("No active view")
            return

        profile_ids = self._selected_profile_ids()
        categories = self._selected_categories()
        link_pairs, host_ids = _hp.collect_targets(
            self.doc, self.profile_data,
            profile_ids=profile_ids, categories=categories,
        )
        if not link_pairs and not host_ids:
            self._set_status("Nothing matched the filters.")
            return

        with revit.Transaction("Hide Existing Profiles (MEPRFP 2.0)", doc=self.doc):
            new_view = _hp.duplicate_active_view(self.doc, active_view)
            host_count, link_count, warnings = _hp.hide_in_view(
                self.doc, new_view, link_pairs, host_ids
            )

        # Switch the user to the new view.
        try:
            uidoc = revit.uidoc
            uidoc.ActiveView = new_view
        except Exception:
            pass

        self.committed = True
        self._last_result = {
            "view_name": new_view.Name,
            "host_count": host_count,
            "link_count": link_count,
            "warnings": warnings,
        }
        self._set_status(
            "Done. Switched to '{}'. Host hidden: {}, link hidden: {}.".format(
                new_view.Name, host_count, link_count
            )
        )
        self.window.Close()

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show(self):
        self.window.ShowDialog()
        return self


def show_modal(doc, profile_data):
    return HideProfilesController(doc, profile_data).show()
