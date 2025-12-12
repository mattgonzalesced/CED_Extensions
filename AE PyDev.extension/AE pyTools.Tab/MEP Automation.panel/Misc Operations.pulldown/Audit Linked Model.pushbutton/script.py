# -*- coding: utf-8 -*-
"""
Audit Linked Model
------------------
Counts every placed linked element (instances with Element_Linker metadata) in
the active model and shows a searchable, modeless list that can remain open
while the user runs other commands.
"""

from __future__ import division, absolute_import, print_function

import os
import sys
from collections import Counter

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance, Group, RevitLinkInstance

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Audit Linked Model"
LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
XAML_PATH = script.get_bundle_file("AuditWindow.xaml")
LOG = script.get_logger()
_AUDIT_WINDOW = None


def _get_selected_link_title():
    try:
        selection = revit.get_selection()
        elems = list(selection.elements)
    except Exception:
        elems = []
    for elem in elems or []:
        if isinstance(elem, RevitLinkInstance):
            link_doc = None
            try:
                link_doc = elem.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc:
                title = getattr(link_doc, "Title", None)
                if title:
                    return title
            name = getattr(elem, "Name", None)
            if name:
                return name
    return None


def _get_element_linker_payload(elem):
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        try:
            value = param.AsString()
        except Exception:
            value = None
        if value:
            text = value.strip()
            if text:
                return text
    return ""


def _extract_led_id(payload):
    if not payload:
        return ""
    target_key = "linked element definition id"
    for raw_line in payload.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, _, remainder = line.partition(":")
        if key.strip().lower() == target_key:
            return remainder.strip()
    return ""


def _load_led_metadata():
    data_path, data = load_active_yaml_data()
    lookup = {}
    for eq in data.get("equipment_definitions") or []:
        eq_name = (eq.get("name") or eq.get("id") or "").strip()
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_id = (led.get("id") or led.get("led_id") or "").strip()
                if not led_id:
                    continue
                label = (led.get("label") or led.get("name") or "").strip()
                if not label and led.get("parameters"):
                    label = eq_name
                lookup[led_id.lower()] = {
                    "led_id": led_id,
                    "label": label or led_id,
                    "equipment": eq_name,
                }
    return data_path, lookup


def _iter_source_documents(doc):
    seen = set()

    def _mark(target):
        if target is None:
            return False
        try:
            key = target.GetHashCode()
        except Exception:
            key = id(target)
        if key in seen:
            return False
        seen.add(key)
        return True

    if _mark(doc):
        yield doc
    if doc is None:
        return
    try:
        link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    except Exception:
        link_instances = []
    for link in link_instances:
        try:
            link_doc = link.GetLinkDocument()
        except Exception:
            link_doc = None
        if _mark(link_doc):
            yield link_doc


def _iter_host_candidates(doc):
    if doc is None:
        return
    collectors = []
    try:
        collectors.append(FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType())
    except Exception:
        pass
    try:
        collectors.append(FilteredElementCollector(doc).OfClass(Group).WhereElementIsNotElementType())
    except Exception:
        pass
    for collector in collectors:
        for elem in collector:
            yield elem


def _build_label(elem):
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            return u"{} : {}".format(fam_name, type_name)
        if type_name:
            return type_name
        if fam_name:
            return fam_name
    try:
        name = getattr(elem, "Name", None)
        if name:
            return name
    except Exception:
        pass
    return "Unnamed Element"


def _collect_instance_counts(doc):
    doc_counts = {}
    details = {}
    for source_doc in _iter_source_documents(doc):
        doc_title = getattr(source_doc, "Title", None) or "Active Document"
        counter = doc_counts.setdefault(doc_title, Counter())
        for elem in _iter_host_candidates(source_doc):
            label = _build_label(elem)
            payload = _get_element_linker_payload(elem)
            led_id = (_extract_led_id(payload) or "").strip() if payload else ""
            key = u"{}||{}".format(label, led_id.lower())
            counter[key] += 1
            entry = details.setdefault(
                key,
                {"led_id": led_id, "label": label, "doc_titles": set()},
            )
            entry["doc_titles"].add(doc_title)
    return doc_counts, details


class AuditWindow(forms.WPFWindow):
    def __init__(self):
        super(AuditWindow, self).__init__(XAML_PATH)
        self.doc = revit.doc
        self.entries = []
        self.filtered = []
        self.total_instances = 0
        self.current_doc_total = 0
        self._doc_counts = {}
        self._details = {}
        self._led_lookup = {}
        self._selected_doc_name = _get_selected_link_title()
        self.Loaded += self._on_loaded
        self.Closed += self._on_closed
        self.SearchButton.Click += self._apply_filter
        self.RefreshButton.Click += self._refresh_entries
        self._doc_populating = False
        try:
            self.DocCombo.SelectionChanged += self._on_doc_selection_changed
        except Exception:
            pass
        try:
            self.set_modeless()
        except Exception:
            pass
        self._refresh_entries()

    def _sanitize_count_text(self):
        text = (self.CountBox.Text or "").strip()
        cleaned = "".join(ch for ch in text if ch.isdigit())
        if cleaned != text:
            try:
                caret = self.CountBox.SelectionStart
            except Exception:
                caret = None
            self.CountBox.Text = cleaned
            if caret is not None:
                try:
                    self.CountBox.SelectionStart = min(caret, len(cleaned))
                except Exception:
                    pass
        return cleaned

    def _on_loaded(self, sender, args):
        try:
            if self.DocCombo.Items.Count > 0 and self.DocCombo.SelectedIndex < 0:
                self.DocCombo.SelectedIndex = 0
            self.SearchBox.Focus()
        except Exception:
            pass

    def _on_doc_selection_changed(self, sender=None, args=None):
        if getattr(self, "_doc_populating", False):
            return
        try:
            name = self.DocCombo.SelectedItem
        except Exception:
            name = None
        text = str(name or "").strip()
        if text and text.lower() != "all documents":
            self._selected_doc_name = text
        else:
            self._selected_doc_name = None
        self._update_entries_for_selection()
        self._apply_filter()

    def _refresh_entries(self, sender=None, args=None):
        if not self.doc:
            forms.alert("No active document.", title=TITLE)
            self.Close()
            return
        self.SearchButton.IsEnabled = False
        self.SummaryLabel.Text = "Scanning linked elements..."
        try:
            try:
                _, led_lookup = _load_led_metadata()
            except RuntimeError:
                led_lookup = {}
            doc_counts, details = _collect_instance_counts(self.doc)
            self._led_lookup = led_lookup
            self._doc_counts = doc_counts
            self._details = details
            self._populate_doc_combo(doc_counts)
            self._update_entries_for_selection()
            self._apply_filter()
        except RuntimeError as exc:
            self.entries = []
            self.total_instances = 0
            self.ResultsList.ItemsSource = []
            self.SummaryLabel.Text = str(exc)
        except Exception as exc:  # pragma: no cover - best effort logging
            LOG.error("Failed to build audit list: %s", exc)
            forms.alert("Failed to audit linked model.\n{}".format(exc), title=TITLE)
            self.Close()
        finally:
            if getattr(self, "SearchButton", None):
                self.SearchButton.IsEnabled = True

    def _populate_doc_combo(self, doc_counts):
        combo = getattr(self, "DocCombo", None)
        if combo is None:
            return
        items = ["All Documents"] + sorted(doc_counts.keys())
        self._doc_populating = True
        try:
            try:
                combo.ItemsSource = items
            except Exception:
                combo.Items.Clear()
                for entry in items:
                    combo.Items.Add(entry)
            preferred = _get_selected_link_title() or self._selected_doc_name
            if preferred and preferred in doc_counts:
                self._selected_doc_name = preferred
                target = preferred
            else:
                self._selected_doc_name = None
                target = "All Documents"
            try:
                combo.SelectedItem = target
            except Exception:
                combo.SelectedIndex = 0
        finally:
            self._doc_populating = False

    def _update_entries_for_selection(self):
        doc_counts = self._doc_counts or {}
        if self._selected_doc_name and self._selected_doc_name in doc_counts:
            counts = doc_counts[self._selected_doc_name]
            self.current_doc_total = sum(counts.values())
        else:
            combined = Counter()
            for counter in doc_counts.values():
                combined.update(counter)
            counts = combined
            self.current_doc_total = sum(counts.values())
        self.entries = self._build_entries_from_counts(counts)
        self.total_instances = self.current_doc_total

    def _build_entries_from_counts(self, counts):
        details = getattr(self, "_details", {}) or {}
        led_lookup = getattr(self, "_led_lookup", {}) or {}
        entries = []
        sort_key = []
        for key in counts.keys():
            info = details.get(key, {})
            sort_key.append((info.get("label") or key, key))
        for _, key in sorted(sort_key, key=lambda pair: pair[0].lower()):
            qty = counts.get(key, 0)
            if qty <= 0:
                continue
            info = details.get(key, {})
            led_id = info.get("led_id") or ""
            lookup = led_lookup.get(led_id.lower()) if led_id else None
            label = lookup.get("label") if lookup and lookup.get("label") else info.get("label")
            eq_name = lookup.get("equipment") if lookup else ""
            display_label = label or eq_name or led_id or info.get("label") or "Linked Element"
            display = display_label
            if eq_name and eq_name not in display_label:
                display = u"{} ({})".format(display_label, eq_name)
            if led_id:
                display = u"{}  [{}]".format(display, led_id)
            doc_titles = ", ".join(sorted(info.get("doc_titles") or []))
            entries.append({
                "display": display,
                "count": qty,
                "search": u"{} {} {} {}".format(
                    (display_label or "").lower(),
                    (led_id or "").lower(),
                    (eq_name or "").lower(),
                    doc_titles.lower(),
                ),
            })
        return entries

    def _get_min_count(self):
        raw = (self.CountBox.Text or "").strip()
        if not raw:
            return 0
        try:
            value = int(float(raw))
        except Exception:
            value = 0
        if value < 0:
            value = 0
        return value

    def _apply_filter(self, sender=None, args=None):
        self._sanitize_count_text()
        query = (self.SearchBox.Text or "").strip().lower()
        min_count = self._get_min_count()
        if not self.entries:
            self.ResultsList.ItemsSource = []
            doc_label = getattr(self, "_selected_doc_name", "All Documents")
            if doc_label and doc_label != "All Documents":
                self.SummaryLabel.Text = "No linked elements were found in '{}'.".format(doc_label)
            else:
                self.SummaryLabel.Text = "No linked elements were found."
            return
        def _matches(entry):
            if entry.get("count", 0) < min_count:
                return False
            if not query:
                return True
            return query in entry.get("search", "")

        results = [entry for entry in self.entries if _matches(entry)]
        self.ResultsList.ItemsSource = results
        doc_total = getattr(self, "current_doc_total", sum(entry["count"] for entry in self.entries))
        filtered_total = sum(entry["count"] for entry in results)
        doc_label = getattr(self, "_selected_doc_name", None) or "All Documents"
        self.SummaryLabel.Text = "Document: {} | Showing {} of {} types | Filtered instances: {} | Document total: {}".format(
            doc_label,
            len(results),
            len(self.entries),
            filtered_total,
            doc_total,
        )

    def _on_closed(self, sender, args):
        global _AUDIT_WINDOW
        if _AUDIT_WINDOW is self:
            _AUDIT_WINDOW = None


def main():
    if not revit.doc:
        forms.alert("No active Revit document.", title=TITLE)
        return
    if not XAML_PATH or not os.path.exists(XAML_PATH):
        forms.alert("Missing window definition (AuditWindow.xaml).", title=TITLE)
        return
    global _AUDIT_WINDOW
    existing = _AUDIT_WINDOW
    try:
        if existing and getattr(existing, "IsVisible", False):
            existing.Activate()
            existing._refresh_entries()
            return
    except Exception:
        pass
    window = AuditWindow()
    _AUDIT_WINDOW = window
    window.show()


if __name__ == "__main__":
    main()
