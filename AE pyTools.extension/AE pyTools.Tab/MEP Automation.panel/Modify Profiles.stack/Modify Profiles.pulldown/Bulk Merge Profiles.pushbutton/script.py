# -*- coding: utf-8 -*-
"""
Bulk Merge Profiles
-------------------
Read an Excel/CSV worksheet and bulk-create missing profile names by ID group.

Expected columns:
    A = Shared ID
    B-E = Count columns (merge is only eligible when any of B/C/D >= 1)
    H-P = Profile names to compare against active YAML profile names
"""

import copy
import csv
import io
import os
import re
import sys

import System
from System import Activator, Type
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
from System.Windows.Forms import DialogResult, OpenFileDialog

from pyrevit import forms, script
output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Bulk Merge Profiles"
TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"
NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")


try:
    basestring
except NameError:
    basestring = str


class _MergeChoice(object):
    def __init__(self, label, data=None, checked=True):
        self.label = label
        self.data = data
        self.checked = checked

    def __str__(self):
        return self.label


def _args_array(*args):
    return System.Array[System.Object](list(args))


def _set(obj, prop, val):
    obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(val))


def _get(obj, prop):
    try:
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except Exception:
        return None


def _call(obj, name, *args):
    obj_type = obj.GetType()
    try:
        return obj_type.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
    except Exception:
        try:
            arg_array = _args_array(*args) if args else None
            return obj_type.InvokeMember(name, BindingFlags.GetProperty, None, obj, arg_array)
        except Exception:
            return None


def _excel_cell(cells, row, col):
    item = _call(cells, "Item", row, col)
    value = _get(item, "Value2")
    return ("" if value is None else str(value)).strip()


def _read_excel_matrix(path):
    xl = wb = ws = used = cells = rows_prop = cols_prop = None
    matrix = []
    try:
        excel_type = Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise RuntimeError("Excel is not available. Save the sheet as CSV and try again.")
        xl = Activator.CreateInstance(excel_type)
        _set(xl, "Visible", False)
        _set(xl, "DisplayAlerts", False)
        wb = _call(_get(xl, "Workbooks"), "Open", path)
        ws = _call(_get(wb, "Worksheets"), "Item", 1)
        used = _get(ws, "UsedRange")
        cells = _get(used, "Cells")
        rows_prop = _get(used, "Rows")
        cols_prop = _get(used, "Columns")
        nrows = int(_get(rows_prop, "Count") or 0)
        ncols = int(_get(cols_prop, "Count") or 0)
        for row_index in range(1, nrows + 1):
            row = [_excel_cell(cells, row_index, col_index) for col_index in range(1, ncols + 1)]
            if any(cell for cell in row):
                matrix.append(row)
    finally:
        try:
            if wb:
                _call(wb, "Close", False)
            if xl:
                _call(xl, "Quit")
        except Exception:
            pass
        try:
            if rows_prop:
                Marshal.ReleaseComObject(rows_prop)
            if cols_prop:
                Marshal.ReleaseComObject(cols_prop)
            if used:
                Marshal.ReleaseComObject(used)
            if ws:
                Marshal.ReleaseComObject(ws)
            if wb:
                Marshal.ReleaseComObject(wb)
            if xl:
                Marshal.ReleaseComObject(xl)
        except Exception:
            pass
    return matrix


def _read_csv_matrix(path):
    errors = []
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            matrix = []
            with io.open(path, "r", encoding=encoding) as handle:
                reader = csv.reader(handle)
                for row in reader:
                    cleaned = [(cell or "").strip() for cell in row]
                    if any(cleaned):
                        matrix.append(cleaned)
            return matrix
        except Exception as exc:
            errors.append("{}: {}".format(encoding, exc))
    raise RuntimeError("Failed to read CSV:\n{}".format("\n".join(errors)))


def _pick_source_file():
    dialog = OpenFileDialog()
    dialog.Title = "Select Bulk Merge Source Sheet"
    dialog.Filter = (
        "Excel Workbook (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|"
        "CSV File (*.csv)|*.csv|"
        "All Files (*.*)|*.*"
    )
    dialog.Multiselect = False
    dialog.CheckFileExists = True
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None


def _cell(row, index):
    if index < 0 or index >= len(row):
        return ""
    return (row[index] or "").strip()


def _as_number(text):
    raw = (text or "").strip()
    if not raw:
        return 0.0
    lowered = raw.lower()
    if lowered in ("yes", "y", "true"):
        return 1.0
    if lowered in ("no", "n", "false"):
        return 0.0
    try:
        return float(raw.replace(",", ""))
    except Exception:
        return 0.0


def _normalize_name(value):
    return " ".join((value or "").strip().lower().split())


def _normalize_compact(value):
    normalized = _normalize_name(value)
    if not normalized:
        return ""
    return NON_ALNUM_RE.sub("", normalized)


def _iter_name_aliases(value):
    text = (value or "").strip()
    if not text:
        return []
    aliases = []

    def _add(alias):
        item = (alias or "").strip()
        if item and item not in aliases:
            aliases.append(item)

    normalized = _normalize_name(text)
    compact = _normalize_compact(text)
    _add(normalized)
    _add(compact)

    if ":" in text:
        left, right = text.split(":", 1)
        left_norm = _normalize_name(left)
        right_norm = _normalize_name(right)
        _add(left_norm)
        _add(_normalize_compact(left_norm))
        _add(right_norm)
        _add(_normalize_compact(right_norm))

    if normalized.endswith(" types"):
        head = normalized[:-6].strip()
        _add(head)
        _add(_normalize_compact(head))

    return aliases


def _normalize_id(value):
    raw = (value or "").strip()
    if not raw:
        return ""
    try:
        numeric = float(raw.replace(",", ""))
        as_int = int(numeric)
        if abs(numeric - as_int) < 1e-9:
            return str(as_int)
    except Exception:
        pass
    return raw


def _load_matrix(path):
    extension = os.path.splitext(path)[1].lower()
    if extension == ".csv":
        return _read_csv_matrix(path)
    return _read_excel_matrix(path)


def _load_sheet_groups(path):
    matrix = _load_matrix(path)
    groups = {}
    for row in matrix or []:
        row_id = _normalize_id(_cell(row, 0))
        if not row_id:
            continue

        group = groups.get(row_id)
        if group is None:
            group = {
                "id": row_id,
                "counts": {"B": 0.0, "C": 0.0, "D": 0.0, "E": 0.0},
                "names": [],
                "seen_names": set(),
            }
            groups[row_id] = group

        b_val = _as_number(_cell(row, 1))
        c_val = _as_number(_cell(row, 2))
        d_val = _as_number(_cell(row, 3))
        e_val = _as_number(_cell(row, 4))
        group["counts"]["B"] += b_val
        group["counts"]["C"] += c_val
        group["counts"]["D"] += d_val
        group["counts"]["E"] += e_val

        for col_index in range(7, 16):  # H-P
            name = _cell(row, col_index)
            if not name:
                continue
            normalized = _normalize_name(name)
            if not normalized or normalized in group["seen_names"]:
                continue
            group["seen_names"].add(normalized)
            group["names"].append(name)
    return groups


def _next_eq_number(equipment_defs):
    max_id = 0
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        if not eq_id:
            continue
        suffix = eq_id.split("-")[-1]
        try:
            num = int(suffix)
        except Exception:
            continue
        if num > max_id:
            max_id = num
    return max_id + 1


def _register_alias(alias_map, alias, source_name):
    key = (alias or "").strip()
    value = (source_name or "").strip()
    if not key or not value:
        return
    names = alias_map.get(key)
    if names is None:
        names = set()
        alias_map[key] = names
    names.add(value)


def _build_definition_maps(equipment_defs):
    by_name = {}
    alias_to_source_names = {}
    normalized_name_to_source = {}
    for entry in equipment_defs or []:
        if not isinstance(entry, dict):
            continue
        name = (entry.get("name") or entry.get("id") or "").strip()
        if not name:
            continue
        by_name[name] = entry
        normalized_name_to_source[_normalize_name(name)] = name
        for alias in _iter_name_aliases(name):
            _register_alias(alias_to_source_names, alias, name)

        eq_id = (entry.get("id") or "").strip()
        for alias in _iter_name_aliases(eq_id):
            _register_alias(alias_to_source_names, alias, name)

        for linked_set in entry.get("linked_sets") or []:
            if not isinstance(linked_set, dict):
                continue
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_name = (led.get("label") or led.get("name") or "").strip()
                if not led_name:
                    continue
                for alias in _iter_name_aliases(led_name):
                    _register_alias(alias_to_source_names, alias, name)

    return by_name, alias_to_source_names, normalized_name_to_source


def _resolve_source_name(raw_name, alias_to_source_names):
    raw = (raw_name or "").strip()
    if not raw:
        return None
    raw_norm = _normalize_name(raw)
    for alias in _iter_name_aliases(raw):
        candidates = alias_to_source_names.get(alias)
        if not candidates:
            continue
        if len(candidates) == 1:
            return list(candidates)[0]
        exact = [name for name in candidates if _normalize_name(name) == raw_norm]
        if exact:
            return sorted(exact, key=lambda v: v.lower())[0]
        return sorted(candidates, key=lambda v: v.lower())[0]
    return None


def _copy_fields(source_entry, target_entry):
    keep_keys = {"name", "id"}
    for key in list(target_entry.keys()):
        if key in keep_keys:
            continue
        target_entry.pop(key, None)
    for key, value in source_entry.items():
        if key in keep_keys:
            continue
        target_entry[key] = copy.deepcopy(value)


def _ensure_truth_source(entry):
    if not isinstance(entry, dict):
        return None, None
    eq_id = (entry.get("id") or entry.get("name") or "").strip()
    eq_name = (entry.get("name") or entry.get("id") or "").strip()
    if not eq_id:
        return None, None
    entry[TRUTH_SOURCE_ID_KEY] = eq_id
    if eq_name:
        entry[TRUTH_SOURCE_NAME_KEY] = eq_name
    return eq_id, eq_name


def _repoint_truth_children(equipment_defs, old_source_id, new_source_id, new_source_name):
    old_id = (old_source_id or "").strip()
    new_id = (new_source_id or "").strip()
    if not old_id or not new_id or old_id == new_id:
        return
    for entry in equipment_defs or []:
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id and entry.get("id"):
            source_id = str(entry.get("id") or "").strip()
        if source_id == old_id:
            entry[TRUTH_SOURCE_ID_KEY] = new_id
            if new_source_name:
                entry[TRUTH_SOURCE_NAME_KEY] = new_source_name


def _id_sort_key(value):
    text = (value or "").strip()
    try:
        return 0, int(text)
    except Exception:
        return 1, text.lower()


def _preview_names(values, limit):
    if not values:
        return ""
    if len(values) <= limit:
        return ", ".join(values)
    return "{}, +{} more".format(", ".join(values[:limit]), len(values) - limit)


def _resolve_exact_name(raw_name, normalized_name_to_source):
    normalized = _normalize_name(raw_name)
    if not normalized:
        return None
    return normalized_name_to_source.get(normalized)


def _target_exists_exact(raw_name, normalized_name_to_source):
    return bool(_resolve_exact_name(raw_name, normalized_name_to_source))


def _build_merge_candidates(groups, alias_to_source_names, normalized_name_to_source):
    opportunities = []
    stats = {
        "total_ids": 0,
        "eligible_ids": 0,
        "skipped_zero": 0,
        "skipped_no_names": 0,
        "skipped_no_source": 0,
        "already_complete": 0,
        "no_source_examples": [],
    }

    for row_id in sorted(groups.keys(), key=_id_sort_key):
        stats["total_ids"] += 1
        group = groups[row_id]
        counts = group.get("counts") or {}
        names = group.get("names") or []

        if not ((counts.get("B", 0) >= 1) or (counts.get("C", 0) >= 1) or (counts.get("D", 0) >= 1)):
            stats["skipped_zero"] += 1
            continue
        stats["eligible_ids"] += 1

        if not names:
            stats["skipped_no_names"] += 1
            continue

        existing_exact_names = []
        existing_exact_seen = set()
        source_candidates = []
        source_seen = set()
        missing_names = []
        missing_seen = set()
        for raw_name in names:
            exact_name = _resolve_exact_name(raw_name, normalized_name_to_source)
            if exact_name:
                exact_key = _normalize_name(exact_name)
                if exact_key not in existing_exact_seen:
                    existing_exact_seen.add(exact_key)
                    existing_exact_names.append(exact_name)
                if exact_key not in source_seen:
                    source_seen.add(exact_key)
                    source_candidates.append(exact_name)
                continue

            source_name = _resolve_source_name(raw_name, alias_to_source_names)
            if source_name:
                source_key = _normalize_name(source_name)
                if source_key not in source_seen:
                    source_seen.add(source_key)
                    source_candidates.append(source_name)

            normalized = _normalize_name(raw_name)
            if not normalized:
                continue
            if normalized not in missing_seen:
                missing_seen.add(normalized)
                missing_names.append(raw_name)

        if not missing_names:
            stats["already_complete"] += 1
            continue

        if not source_candidates:
            stats["skipped_no_source"] += 1
            if len(stats["no_source_examples"]) < 8:
                stats["no_source_examples"].append(
                    "ID {}: {}".format(row_id, _preview_names(names, 4))
                )
            continue

        source_name = source_candidates[0]
        opportunities.append({
            "id": row_id,
            "source_name": source_name,
            "alt_existing": existing_exact_names[1:],
            "missing_names": missing_names,
            "all_names": names,
            "counts": counts,
        })

    return opportunities, stats


def _build_choice_label(opportunity):
    row_id = opportunity["id"]
    source_name = opportunity["source_name"]
    missing_names = opportunity.get("missing_names") or []
    all_names = opportunity.get("all_names") or []
    alt_existing = opportunity.get("alt_existing") or []
    counts = opportunity.get("counts") or {}

    count_summary = "B:{:.0f} C:{:.0f} D:{:.0f} E:{:.0f}".format(
        counts.get("B", 0),
        counts.get("C", 0),
        counts.get("D", 0),
        counts.get("E", 0),
    )
    label = "ID {} | {} | Source: {} | Add: {} | H-P: {}".format(
        row_id,
        count_summary,
        source_name,
        _preview_names(missing_names, 3),
        _preview_names(all_names, 4),
    )
    if alt_existing:
        label += " | Also in YAML: {}".format(_preview_names(alt_existing, 2))
    return label


def _select_opportunities(opportunities):
    if not opportunities:
        return []

    choices = []
    for item in opportunities:
        choices.append(_MergeChoice(_build_choice_label(item), data=item, checked=True))

    selected = forms.SelectFromList.show(
        choices,
        title=TITLE,
        multiselect=True,
        button_name="Merge Selected",
        width=1500,
        height=700,
    )
    if not selected:
        return []
    if not isinstance(selected, list):
        selected = [selected]
    return [item.data for item in selected if isinstance(item, _MergeChoice) and item.data]


def main():
    source_path = _pick_source_file()
    if not source_path:
        return

    try:
        groups = _load_sheet_groups(source_path)
    except Exception as exc:
        forms.alert("Failed to read source file:\n\n{}".format(exc), title=TITLE)
        return

    if not groups:
        forms.alert("No ID rows were found in column A.", title=TITLE)
        return

    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    equipment_defs = data.get("equipment_definitions") or []
    if not equipment_defs:
        forms.alert("No equipment definitions are available in the active YAML.", title=TITLE)
        return

    yaml_label = get_yaml_display_name(yaml_path)
    def_by_name, alias_to_source_names, normalized_name_to_source = _build_definition_maps(equipment_defs)
    opportunities, stats = _build_merge_candidates(
        groups,
        alias_to_source_names,
        normalized_name_to_source,
    )

    if not opportunities:
        details = []
        if stats.get("no_source_examples"):
            details.append("")
            details.append("Sample unresolved H-P names:")
            for sample in stats["no_source_examples"]:
                details.append(" - {}".format(sample))
        forms.alert(
            "No mergeable ID groups were found.\n\n"
            "Checked IDs: {total}\n"
            "Eligible by B/C/D: {eligible}\n"
            "Skipped (B/C/D all zero): {zero}\n"
            "Skipped (no H-P names): {nonames}\n"
            "Skipped (no existing YAML source in H-P): {nosource}\n"
            "Already complete (nothing missing): {complete}{details}".format(
                total=stats["total_ids"],
                eligible=stats["eligible_ids"],
                zero=stats["skipped_zero"],
                nonames=stats["skipped_no_names"],
                nosource=stats["skipped_no_source"],
                complete=stats["already_complete"],
                details="\n".join(details),
            ),
            title=TITLE,
        )
        return

    selected = _select_opportunities(opportunities)
    if not selected:
        return

    next_eq_num = _next_eq_number(equipment_defs)
    applied_ids = []
    created_names = []
    created_summary = []

    for opportunity in selected:
        source_name = opportunity.get("source_name") or ""
        source_entry = def_by_name.get(source_name)
        if not source_entry:
            canonical_name = _resolve_source_name(source_name, alias_to_source_names)
            source_entry = def_by_name.get(canonical_name) if canonical_name else None
        if not source_entry:
            continue

        root_id, root_name = _ensure_truth_source(source_entry)
        if not root_id:
            root_id = (source_entry.get("id") or source_entry.get("name") or source_name).strip()
        if not root_name:
            root_name = (source_entry.get("name") or source_name).strip()

        id_created = []
        for target_name_raw in opportunity.get("missing_names") or []:
            target_name = (target_name_raw or "").strip()
            if not target_name:
                continue

            if _target_exists_exact(target_name, normalized_name_to_source):
                continue

            target_entry = {
                "id": "EQ-{:03d}".format(next_eq_num),
                "name": target_name,
            }
            next_eq_num += 1
            equipment_defs.append(target_entry)
            def_by_name[target_name] = target_entry
            normalized_name_to_source[_normalize_name(target_name)] = target_name
            for alias in _iter_name_aliases(target_name):
                _register_alias(alias_to_source_names, alias, target_name)
            for alias in _iter_name_aliases(target_entry.get("id")):
                _register_alias(alias_to_source_names, alias, target_name)

            _copy_fields(source_entry, target_entry)
            target_entry[TRUTH_SOURCE_ID_KEY] = root_id
            if root_name:
                target_entry[TRUTH_SOURCE_NAME_KEY] = root_name
            _repoint_truth_children(
                equipment_defs,
                target_entry.get("id"),
                root_id,
                root_name,
            )

            id_created.append(target_name)
            created_names.append(target_name)

        if id_created:
            applied_ids.append(opportunity.get("id"))
            created_summary.append("ID {} -> {}".format(opportunity.get("id"), ", ".join(id_created)))

    if not created_names:
        forms.alert("No profiles were created from the selected items.", title=TITLE)
        return

    action_description = "Bulk merged {} profile(s) from {}".format(
        len(created_names),
        os.path.basename(source_path),
    )
    save_active_yaml_data(
        None,
        data,
        TITLE,
        action_description,
    )

    lines = [
        "Bulk merge complete.",
        "Source sheet: {}".format(os.path.basename(source_path)),
        "Selected ID groups: {}".format(len(selected)),
        "ID groups changed: {}".format(len(applied_ids)),
        "Profiles created: {}".format(len(created_names)),
        "",
        "Updated data saved back to {}.".format(yaml_label),
    ]
    if created_summary:
        lines.append("")
        lines.append("Created profile names:")
        for summary_line in created_summary[:20]:
            lines.append(" - {}".format(summary_line))
        if len(created_summary) > 20:
            lines.append(" - ... +{} more ID groups".format(len(created_summary) - 20))
    forms.alert("\n".join(lines), title=TITLE)


if __name__ == "__main__":
    main()
