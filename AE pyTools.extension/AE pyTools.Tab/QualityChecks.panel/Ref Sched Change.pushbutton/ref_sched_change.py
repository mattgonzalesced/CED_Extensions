# -*- coding: utf-8 -*-
"""
After-sync check: notify when refrigeration schedule sheets change.
"""

import os
import sys
import time

from pyrevit import DB, forms, revit, script

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage import ExtensibleStorage  # noqa: E402

SETTING_KEY = "ref_sched_change_check"
PHRASE_A = "REFRIGERATION"
PHRASE_B = "SCHEDULE"
CONFIG_KEY = "ref_sched_change_config"
SNAPSHOT_KEY = "ref_sched_change_snapshot"


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def get_setting(default=True, doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return bool(default)
    value = ExtensibleStorage.get_user_setting(doc, SETTING_KEY, default=None)
    if value is None:
        return bool(default)
    return bool(value)


def set_setting(value, doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return False
    return ExtensibleStorage.set_user_setting(doc, SETTING_KEY, bool(value))


def _doc_key(doc):
    if doc is None:
        return "<none>"
    try:
        return doc.PathName or doc.Title or "<unnamed>"
    except Exception:
        return "<unknown>"


def _sheet_matches(sheet):
    if sheet is None:
        return False
    name = ""
    try:
        name = sheet.Name or ""
    except Exception:
        name = ""
    upper = name.upper()
    return PHRASE_A in upper and PHRASE_B in upper


def _sheet_label(sheet):
    if sheet is None:
        return "<sheet>"
    try:
        number = sheet.SheetNumber or ""
    except Exception:
        number = ""
    try:
        name = sheet.Name or ""
    except Exception:
        name = ""
    if number and name:
        return "{} - {}".format(number, name)
    return name or number or "<sheet>"


def _sheet_from_element(doc, elem):
    if doc is None or elem is None:
        return None
    if isinstance(elem, DB.ViewSheet):
        return elem
    sheet_id = None
    try:
        sheet_id = getattr(elem, "SheetId", None)
    except Exception:
        sheet_id = None
    if sheet_id:
        try:
            sheet = doc.GetElement(sheet_id)
            if isinstance(sheet, DB.ViewSheet):
                return sheet
        except Exception:
            pass
    owner_id = None
    try:
        owner_id = getattr(elem, "OwnerViewId", None)
    except Exception:
        owner_id = None
    if owner_id and owner_id != DB.ElementId.InvalidElementId:
        try:
            owner = doc.GetElement(owner_id)
            if isinstance(owner, DB.ViewSheet):
                return owner
        except Exception:
            pass
    return None


def _extract_changed_ids(args):
    if args is None:
        return []
    ids = []
    for name in ("GetModifiedElementIds", "GetAddedElementIds", "GetDeletedElementIds"):
        getter = getattr(args, name, None)
        if not callable(getter):
            continue
        try:
            batch = list(getter() or [])
        except Exception:
            batch = []
        for elem_id in batch:
            if elem_id is None:
                continue
            ids.append(elem_id)
    return ids


def _collect_refrigeration_sheets(doc):
    try:
        sheets = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.ViewSheet)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []
    return [sheet for sheet in sheets if _sheet_matches(sheet)]


def _snapshot_sheets(doc):
    snapshot = {
        "timestamp": time.time(),
        "sheets": {},
    }
    sheets = _collect_refrigeration_sheets(doc)
    for sheet in sheets:
        try:
            sheet_id = sheet.Id.IntegerValue
        except Exception:
            continue
        try:
            owned = (
                DB.FilteredElementCollector(doc, sheet.Id)
                .WhereElementIsNotElementType()
                .ToElementIds()
            )
        except Exception:
            owned = []
        owned_ids = []
        for elem_id in owned or []:
            try:
                owned_ids.append(elem_id.IntegerValue)
            except Exception:
                pass
        owned_ids.sort()
        snapshot["sheets"][str(sheet_id)] = {
            "label": _sheet_label(sheet),
            "name": sheet.Name,
            "number": sheet.SheetNumber,
            "owned_ids": owned_ids,
        }
    return snapshot


def _load_snapshot(doc):
    cfg = script.get_config(CONFIG_KEY)
    payload = getattr(cfg, SNAPSHOT_KEY, None)
    if not isinstance(payload, dict):
        return None
    return payload.get(_doc_key(doc))


def _save_snapshot(doc, snapshot):
    cfg = script.get_config(CONFIG_KEY)
    payload = getattr(cfg, SNAPSHOT_KEY, None)
    if not isinstance(payload, dict):
        payload = {}
    payload[_doc_key(doc)] = snapshot or {}
    setattr(cfg, SNAPSHOT_KEY, payload)
    script.save_config()


def _detect_changes_by_snapshot(doc):
    previous = _load_snapshot(doc)
    current = _snapshot_sheets(doc)
    changed = []
    if previous is None:
        _save_snapshot(doc, current)
        return changed
    prev_sheets = previous.get("sheets") or {}
    curr_sheets = current.get("sheets") or {}
    if set(prev_sheets.keys()) != set(curr_sheets.keys()):
        changed.extend([data.get("label") for data in curr_sheets.values() if data.get("label")])
        _save_snapshot(doc, current)
        return list(dict.fromkeys(changed))
    for sheet_id, curr_data in curr_sheets.items():
        prev_data = prev_sheets.get(sheet_id) or {}
        if (
            curr_data.get("name") != prev_data.get("name")
            or curr_data.get("number") != prev_data.get("number")
            or curr_data.get("owned_ids") != prev_data.get("owned_ids")
        ):
            label = curr_data.get("label")
            if label:
                changed.append(label)
    _save_snapshot(doc, current)
    return list(dict.fromkeys(changed))


def _detect_changes_by_args(doc, args):
    changed = []
    ids = _extract_changed_ids(args)
    if not ids:
        return changed
    for elem_id in ids:
        elem = None
        try:
            elem = doc.GetElement(elem_id)
        except Exception:
            elem = None
        sheet = _sheet_from_element(doc, elem)
        if sheet and _sheet_matches(sheet):
            label = _sheet_label(sheet)
            if label:
                changed.append(label)
    return list(dict.fromkeys(changed))


def _notify(changed, show_empty=False):
    if not changed:
        if show_empty:
            forms.alert("No refrigeration schedule sheet changes detected.", title="Ref Sched Change")
        return
    output = script.get_output()
    output.print_md("# Refrigeration Schedule Sheet Changes")
    output.print_md("The following sheet(s) were modified since the last sync:")
    for label in changed:
        output.print_md("- {}".format(label))
    forms.alert(
        "Refrigeration schedule sheet changes detected.\n\n"
        "Please communicate with the refrigeration team.\n\n"
        "See the output panel for sheet details.",
        title="Ref Sched Change",
    )


def run_check(doc=None, args=None, show_ui=True, show_empty=False):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    changed = _detect_changes_by_args(doc, args)
    if not changed:
        changed = _detect_changes_by_snapshot(doc)
    if show_ui:
        _notify(changed, show_empty=show_empty)
    return changed


def run_sync_check(doc, args=None):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return
    if not get_setting(default=True, doc=doc):
        return
    run_check(doc, args=args, show_ui=True, show_empty=False)


__all__ = [
    "get_setting",
    "set_setting",
    "run_check",
    "run_sync_check",
]
