# -*- coding: utf-8 -*-
"""
Classify Spaces
---------------
Classifies MEP spaces into operational buckets, lets the user adjust the
results in a XAML UI, and stores the final mapping in Extensible Storage.
"""

import os
import re
import sys
from collections import OrderedDict
from datetime import datetime

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FilteredElementCollector

output = script.get_output()
output.close_others()

TITLE = "Classify Spaces"
STORAGE_ID = "space_operations.classifications.v1"
SPACE_CLASSIFICATION_SCHEMA_VERSION = 1

BUCKETS = [
    "Restrooms",
    "Offices",
    "Sales Floor",
    "Freezers",
    "Coolers",
    "Receiving",
    "Break",
    "Food Prep",
    "Utility",
    "Storage",
    "Other",
]

FILTER_ALL_LABEL = "All Suggested Buckets"


def _resolve_lib_root():
    cursor = os.path.abspath(os.path.dirname(__file__))
    for _ in range(12):
        candidate = os.path.join(cursor, "CEDLib.lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if not parent or parent == cursor:
            break
        cursor = parent
    return None


LIB_ROOT = _resolve_lib_root()
if LIB_ROOT and LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

try:
    from ExtensibleStorage import ExtensibleStorage  # noqa: E402
except Exception:
    ExtensibleStorage = None


try:
    basestring
except NameError:  # pragma: no cover
    basestring = str


def _element_id_value(elem_id, default=""):
    if elem_id is None:
        return default
    for attr in ("IntegerValue", "Value"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return str(int(value))
        except Exception:
            try:
                return str(value)
            except Exception:
                continue
    return default


def _param_text(element, built_in_param):
    if element is None:
        return ""
    try:
        param = element.get_Parameter(built_in_param)
    except Exception:
        param = None
    if not param:
        return ""
    for getter_name in ("AsString", "AsValueString"):
        try:
            getter = getattr(param, getter_name)
            value = getter()
        except Exception:
            value = None
        if value:
            text = str(value).strip()
            if text:
                return text
    return ""


def _space_name(space):
    name = _param_text(space, BuiltInParameter.ROOM_NAME)
    if name:
        return name
    try:
        value = getattr(space, "Name", None)
    except Exception:
        value = None
    return str(value).strip() if value else ""


def _space_number(space):
    return _param_text(space, BuiltInParameter.ROOM_NUMBER)


def _collect_spaces(doc):
    try:
        collector = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_MEPSpaces)
            .WhereElementIsNotElementType()
        )
        spaces = list(collector)
    except Exception:
        spaces = []
    return spaces


def _normalize_text(value):
    raw = (value or "").strip().lower()
    collapsed = re.sub(r"[^a-z0-9]+", " ", raw).strip()
    tokens = set(collapsed.split()) if collapsed else set()
    return raw, collapsed, tokens


def _contains_phrase(collapsed_text, phrase):
    if not collapsed_text or not phrase:
        return False
    hay = " {} ".format(collapsed_text)
    needle = " {} ".format((phrase or "").strip().lower())
    return needle in hay


def _first_substring(raw_text, values):
    if not raw_text:
        return None
    for value in values or []:
        token = (value or "").strip().lower()
        if token and token in raw_text:
            return token
    return None


def _first_token(tokens, values):
    if not tokens:
        return None
    for value in values or []:
        token = (value or "").strip().lower()
        if token and token in tokens:
            return token
    return None


def _first_prefix(tokens, prefixes):
    if not tokens:
        return None
    for token in tokens:
        for prefix in prefixes or []:
            candidate = (prefix or "").strip().lower()
            if candidate and token.startswith(candidate):
                return token
    return None


def _classify_space_name(space_name):
    raw, collapsed, tokens = _normalize_text(space_name)

    restroom_hit = _first_substring(
        raw,
        ["restroom", "bathroom", "washroom", "toilet", "lavatory", "lav"],
    )
    if restroom_hit:
        return "Restrooms", "keyword '{}'".format(restroom_hit)
    if "men's" in raw or "womens" in raw or "women's" in raw or "mens" in raw:
        return "Restrooms", "keyword men's/women's"
    if "men" in tokens or "women" in tokens:
        return "Restrooms", "keyword men/women"

    freezer_hit = _first_substring(raw, ["freezer"]) or _first_token(tokens, ["freezer"])
    if freezer_hit:
        return "Freezers", "keyword '{}'".format(freezer_hit)

    cooler_hit = _first_substring(raw, ["cooler"]) or _first_token(tokens, ["cooler"])
    if cooler_hit:
        return "Coolers", "keyword '{}'".format(cooler_hit)

    receiving_hit = _first_substring(raw, ["receiving"]) or _first_token(tokens, ["receiving"])
    if receiving_hit:
        return "Receiving", "keyword '{}'".format(receiving_hit)

    food_prep_hit = _first_substring(raw, ["food prep", "prep"]) or _first_token(tokens, ["prep"])
    if food_prep_hit:
        return "Food Prep", "keyword '{}'".format(food_prep_hit)

    break_hit = _first_substring(raw, ["break"]) or _first_token(tokens, ["break"])
    if break_hit:
        return "Break", "keyword '{}'".format(break_hit)

    utility_hit = _first_substring(raw, ["electrical", "mechanical", "utility"])
    utility_token = _first_token(tokens, ["electrical", "mechanical", "utility", "elec", "mech"])
    if utility_hit or utility_token:
        return "Utility", "keyword '{}'".format(utility_hit or utility_token)

    if _contains_phrase(collapsed, "sales floor") or "salesfloor" in raw:
        return "Sales Floor", "phrase 'sales floor'"
    sales_hit = _first_substring(raw, ["selling"]) or _first_token(tokens, ["selling", "sell", "sales"])
    if sales_hit:
        return "Sales Floor", "keyword '{}'".format(sales_hit)

    office_hit = _first_substring(raw, ["office", "admin", "desk", "administration"]) or _first_token(
        tokens,
        ["office", "admin", "desk", "administration"],
    )
    if office_hit:
        return "Offices", "keyword '{}'".format(office_hit)

    storage_hit = _first_substring(raw, ["storage", "janitor", "ware"])
    storage_prefix = _first_prefix(tokens, ["ware"])
    if storage_hit or storage_prefix:
        return "Storage", "keyword '{}'".format(storage_hit or storage_prefix)

    return "Other", "no keyword match"


def _load_saved_bucket_maps(doc):
    by_id = {}
    by_unique_id = {}

    if ExtensibleStorage is None:
        return by_id, by_unique_id

    stored = ExtensibleStorage.get_project_data(doc, STORAGE_ID, default=None)
    if not isinstance(stored, dict):
        return by_id, by_unique_id

    assignments = stored.get("space_assignments") or {}
    if not isinstance(assignments, dict):
        return by_id, by_unique_id

    for key, value in assignments.items():
        entry = value if isinstance(value, dict) else {"bucket": value}
        bucket = (entry.get("bucket") or "").strip()
        if bucket not in BUCKETS:
            continue

        space_id = entry.get("space_id")
        if space_id in (None, ""):
            space_id = key
        space_id = str(space_id).strip() if space_id is not None else ""

        unique_id = entry.get("unique_id")
        unique_id = str(unique_id).strip() if unique_id else ""

        if space_id:
            by_id[space_id] = bucket
        if unique_id:
            by_unique_id[unique_id] = bucket

    return by_id, by_unique_id


class SpaceClassificationRow(object):
    def __init__(
        self,
        space_id,
        unique_id,
        space_number,
        space_name,
        suggested_bucket,
        bucket,
        reason,
    ):
        self.space_id = str(space_id or "")
        self.unique_id = str(unique_id or "")
        self.space_number = space_number or ""
        self.space_name = space_name or ""
        self.suggested_bucket = suggested_bucket or "Other"
        self.bucket = bucket or self.suggested_bucket
        self.reason = reason or ""
        self.bucket_options = list(BUCKETS)


class SpaceClassificationWindow(forms.WPFWindow):
    def __init__(self, xaml_path, rows):
        forms.WPFWindow.__init__(self, xaml_path)
        self.rows = list(rows or [])
        self.accepted = False

        self._grid = self.FindName("ClassificationsGrid")
        self._suggested_filter_combo = self.FindName("SuggestedFilterCombo")

        if self._suggested_filter_combo is not None:
            self._suggested_filter_combo.ItemsSource = [FILTER_ALL_LABEL] + list(BUCKETS)
            self._suggested_filter_combo.SelectedItem = FILTER_ALL_LABEL

        self._apply_filter()

    def _current_suggested_filter(self):
        if self._suggested_filter_combo is None:
            return None
        selected = getattr(self._suggested_filter_combo, "SelectedItem", None)
        if selected in BUCKETS:
            return selected
        return None

    def _filtered_rows(self):
        bucket_filter = self._current_suggested_filter()
        if not bucket_filter:
            return list(self.rows)
        return [row for row in self.rows if row.suggested_bucket == bucket_filter]

    def _apply_filter(self):
        if self._grid is not None:
            self._grid.ItemsSource = self._filtered_rows()
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass
        self._refresh_summary()

    def _refresh_summary(self):
        summary = self.FindName("SummaryText")
        if summary is None:
            return

        counts = OrderedDict((bucket, 0) for bucket in BUCKETS)
        for row in self.rows:
            bucket = row.bucket if row.bucket in counts else "Other"
            counts[bucket] += 1

        bucket_filter = self._current_suggested_filter()
        visible_count = len(self._filtered_rows())

        lines = ["Total spaces: {}".format(len(self.rows))]
        lines.append("Visible rows: {}".format(visible_count))
        lines.append("Suggested filter: {}".format(bucket_filter or "All"))
        for bucket, count in counts.items():
            if count <= 0:
                continue
            lines.append("{:<12} {}".format(bucket + ":", count))
        summary.Text = "\n".join(lines)

    def OnSuggestedFilterChanged(self, sender, args):
        self._apply_filter()

    def OnBucketChanged(self, sender, args):
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        selected = getattr(sender, "SelectedItem", None)
        if selected in BUCKETS:
            row.bucket = selected
        self._refresh_summary()

    def OnSaveClicked(self, sender, args):
        if self._grid is not None:
            try:
                self._grid.CommitEdit()
                self._grid.CommitEdit()
            except Exception:
                pass
        self.accepted = True
        self.Close()

    def OnCancelClicked(self, sender, args):
        self.accepted = False
        self.Close()


def _build_rows(doc):
    saved_by_id, saved_by_unique_id = _load_saved_bucket_maps(doc)
    rows = []


    for space in _collect_spaces(doc):
        space_id = _element_id_value(getattr(space, "Id", None), default="")
        unique_id = ""
        try:
            unique_id = str(getattr(space, "UniqueId", "") or "").strip()
        except Exception:
            unique_id = ""

        space_number = _space_number(space)
        space_name = _space_name(space) or "<Unnamed Space>"
        suggested_bucket, reason = _classify_space_name(space_name)

        saved_bucket = ""
        if unique_id and unique_id in saved_by_unique_id:
            saved_bucket = saved_by_unique_id.get(unique_id) or ""
        elif space_id and space_id in saved_by_id:
            saved_bucket = saved_by_id.get(space_id) or ""

        bucket = saved_bucket if saved_bucket in BUCKETS else suggested_bucket
        if saved_bucket and saved_bucket in BUCKETS and saved_bucket != suggested_bucket:
            reason = "{} | saved override '{}'".format(reason, saved_bucket)

        rows.append(
            SpaceClassificationRow(
                space_id=space_id,
                unique_id=unique_id,
                space_number=space_number,
                space_name=space_name,
                suggested_bucket=suggested_bucket,
                bucket=bucket,
                reason=reason,
            )
        )

    rows.sort(key=lambda row: ((row.space_number or "").lower(), (row.space_name or "").lower()))
    return rows


def _build_payload(doc, rows, existing_payload=None):
    counts = OrderedDict((bucket, 0) for bucket in BUCKETS)
    assignments = OrderedDict()

    for row in rows:
        bucket = row.bucket if row.bucket in counts else "Other"
        counts[bucket] += 1

        assignments[row.space_id or row.unique_id] = {
            "space_id": row.space_id,
            "unique_id": row.unique_id,
            "space_number": row.space_number,
            "space_name": row.space_name,
            "bucket": bucket,
            "suggested_bucket": row.suggested_bucket,
            "reason": row.reason,
        }

    try:
        username = doc.Application.Username
    except Exception:
        username = ""

    payload = dict(existing_payload) if isinstance(existing_payload, dict) else {}
    payload.update({
        "storage_id": STORAGE_ID,
        "schema_version": SPACE_CLASSIFICATION_SCHEMA_VERSION,
        "saved_utc": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "saved_by": username or "unknown",
        "total_spaces": len(rows),
        "bucket_counts": dict(counts),
        "space_assignments": assignments,
    })
    return payload


def _summary_lines(payload):
    counts = payload.get("bucket_counts") or {}
    lines = [
        "Saved classifications for {} spaces.".format(payload.get("total_spaces", 0)),
        "Data storage ID: {}".format(STORAGE_ID),
        "",
        "Bucket totals:",
    ]
    for bucket in BUCKETS:
        count = counts.get(bucket, 0)
        if count <= 0:
            continue
        lines.append(" - {}: {}".format(bucket, count))
    return lines


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    if ExtensibleStorage is None:
        forms.alert("Failed to load ExtensibleStorage library from CEDLib.lib.", title=TITLE)
        return

    rows = _build_rows(doc)
    if not rows:
        forms.alert("No MEP spaces found in the active model.", title=TITLE)
        return

    xaml_path = os.path.join(os.path.dirname(__file__), "SpaceClassificationWindow.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("SpaceClassificationWindow.xaml is missing.", title=TITLE)
        return

    window = SpaceClassificationWindow(xaml_path, rows)
    window.ShowDialog()
    if not window.accepted:
        return

    existing_payload = ExtensibleStorage.get_project_data(doc, STORAGE_ID, default=None)
    payload = _build_payload(doc, rows, existing_payload=existing_payload)

    try:
        saved = ExtensibleStorage.set_project_data(
            doc,
            STORAGE_ID,
            payload,
            transaction_name="{} Save".format(TITLE),
        )
    except Exception as exc:
        forms.alert("Failed to save space classifications:\n\n{}".format(exc), title=TITLE)
        return

    if not saved:
        forms.alert("Space classifications were not saved.", title=TITLE)
        return

    forms.alert("\n".join(_summary_lines(payload)), title=TITLE)


if __name__ == "__main__":
    main()









