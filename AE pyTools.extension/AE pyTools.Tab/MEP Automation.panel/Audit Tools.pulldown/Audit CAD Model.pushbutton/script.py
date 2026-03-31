# -*- coding: utf-8 -*-
"""
Audit CAD Model
---------------
Scan CAD block names from an Excel or CSV sheet for names that do not have
profiles in the active YAML, but look like close name matches to existing
profiles.
"""

import copy
import csv
import io
import os
import re
import sys
from collections import Counter
from difflib import SequenceMatcher

import System
from System import Activator, Type
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
from System.Windows.Forms import DialogResult, OpenFileDialog

from pyrevit import forms, revit, script
output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Audit CAD Model"
LOG = script.get_logger()

DIRECTION_TOKENS = {
    "left", "right", "lh", "rh", "l", "r", "lhs", "rhs", "left-hand", "right-hand",
}
VERSION_TOKEN_RE = re.compile(r"^v\d+$", re.IGNORECASE)
HAS_DIGIT_TOKEN_RE = re.compile(r".*\d")
SEPARATORS_RE = re.compile(r"[_/\\\\]+")
NON_ALNUM_RE = re.compile(r"[^a-zA-Z0-9 -]+")
MATCH_STOP_TOKENS = {"name"}
MIN_PARTIAL_SCORE = 80.0
MIN_COMMON_TOKENS = 1
MIN_LARGER_TOKEN_COVERAGE = 0.75
TOKEN_LENGTH_GAP_PENALTY = 12.0


try:
    basestring
except NameError:
    basestring = str


class _MissingChoice(object):
    def __init__(self, label, data=None, checked=True):
        self.label = label
        self.data = data
        self.checked = checked

    def __str__(self):
        return self.label


class _ColumnChoice(object):
    def __init__(self, label, index):
        self.label = label
        self.index = index

    def __str__(self):
        return self.label


def _normalize_full_name(value):
    if not value:
        return ""
    text = SEPARATORS_RE.sub(" ", str(value))
    text = NON_ALNUM_RE.sub(" ", text)
    return " ".join(text.lower().split())


def _truncate_text(value, max_len):
    text = str(value or "")
    if max_len <= 3 or len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def _tokenize_name(value):
    normalized = _normalize_full_name(value)
    return [token for token in normalized.split() if token]


def _strip_direction_tokens(tokens):
    return [token for token in tokens if token not in DIRECTION_TOKENS]


def _strip_non_descriptor_tokens(tokens):
    filtered = []
    for token in tokens or []:
        if not token:
            continue
        if token in MATCH_STOP_TOKENS:
            continue
        if VERSION_TOKEN_RE.match(token):
            continue
        if HAS_DIGIT_TOKEN_RE.match(token):
            continue
        filtered.append(token)
    return filtered


def _collapse_acronym_tokens(tokens):
    collapsed = []
    letters = []
    for token in tokens or []:
        if len(token) == 1 and token.isalpha():
            letters.append(token)
            continue
        if letters:
            if len(letters) > 1:
                collapsed.append("".join(letters))
            else:
                collapsed.extend(letters)
            letters = []
        collapsed.append(token)
    if letters:
        if len(letters) > 1:
            collapsed.append("".join(letters))
        else:
            collapsed.extend(letters)
    return collapsed


def _base_key(value):
    if not value:
        return ""
    tokens = _match_tokens(value)
    base = " ".join(sorted(set(tokens)))
    return base or _normalize_full_name(value)


def token_set_ratio(a, b):
    a_tokens = set((a or "").lower().split())
    b_tokens = set((b or "").lower().split())
    if not a_tokens or not b_tokens:
        return 0.0
    common = " ".join(sorted(a_tokens & b_tokens))
    a_diff = " ".join(sorted(a_tokens - b_tokens))
    b_diff = " ".join(sorted(b_tokens - a_tokens))
    return max(
        SequenceMatcher(None, common, (common + " " + a_diff).strip()).ratio(),
        SequenceMatcher(None, common, (common + " " + b_diff).strip()).ratio()
    ) * 100.0


def _match_tokens(value):
    if not value:
        return []
    tokens = _tokenize_name(value)
    tokens = _strip_direction_tokens(tokens)
    tokens = _strip_non_descriptor_tokens(tokens)
    tokens = _collapse_acronym_tokens(tokens)
    return tokens


def _match_key(value):
    tokens = _match_tokens(value)
    if tokens:
        return " ".join(sorted(set(tokens)))
    return _normalize_full_name(value)


def _token_overlap_metrics(left, right):
    left_tokens = set(_match_tokens(left))
    right_tokens = set(_match_tokens(right))
    if not left_tokens or not right_tokens:
        return 0, 0.0, 0.0
    common = left_tokens & right_tokens
    common_count = len(common)
    if common_count == 0:
        return 0, 0.0, 0.0
    small = float(min(len(left_tokens), len(right_tokens)))
    large = float(max(len(left_tokens), len(right_tokens)))
    return common_count, (common_count / small), (common_count / large)


def _best_match_name(target, candidates):
    target_norm = _match_key(target)
    target_tokens = set(_match_tokens(target))
    target_len = len(target_tokens)
    best = None
    best_score = -1.0
    for name in candidates:
        candidate_key = _match_key(name)
        score = token_set_ratio(target_norm, candidate_key)
        candidate_tokens = set(_match_tokens(name))
        candidate_len = len(candidate_tokens)
        if target_len and candidate_len:
            length_gap = abs(target_len - candidate_len)
            score -= (TOKEN_LENGTH_GAP_PENALTY * float(length_gap))
        if score > best_score:
            best_score = score
            best = name
    return best, best_score


def _split_label(label):
    cleaned = (label or "").strip()
    if not cleaned:
        return "", ""
    if ":" in cleaned:
        fam_part, type_part = cleaned.split(":", 1)
        return fam_part.strip(), type_part.strip()
    return cleaned, ""


def _collect_profile_names(data):
    names = set()
    for eq in data.get("equipment_definitions") or []:
        raw = eq.get("name") or eq.get("id")
        if raw:
            names.add(str(raw).strip())
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                label = led.get("label") or led.get("name") or ""
                if label:
                    names.add(str(label).strip())
    return names


def _find_equipment_def_by_name(data, name):
    if not name:
        return None
    target = str(name).strip()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        eq_name = (eq.get("name") or eq.get("id") or "").strip()
        if eq_name == target:
            return eq
    return None


def _find_equipment_def_by_label(data, label):
    if not label:
        return None
    target = str(label).strip()
    if not target:
        return None
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_label = (led.get("label") or led.get("name") or "").strip()
                if led_label == target:
                    return eq
    return None


def _next_eq_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        eq_id = (eq.get("id") or "").strip()
        if eq_id.upper().startswith("EQ-"):
            try:
                num = int(eq_id.split("-")[-1])
            except Exception:
                continue
            if num > max_id:
                max_id = num
    return max_id + 1


def _next_set_number(data):
    max_id = 0
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            set_id = (linked_set.get("id") or "").strip()
            if not set_id.upper().startswith("SET-"):
                continue
            try:
                num = int(set_id.split("-")[-1])
            except Exception:
                continue
            if num > max_id:
                max_id = num
    return max_id + 1


def _rewrite_linker_payload(params, old_set_id, new_set_id, old_led_id, new_led_id):
    if not isinstance(params, dict):
        return
    for key in ("Element_Linker Parameter", "Element_Linker"):
        value = params.get(key)
        if not isinstance(value, basestring):
            continue
        updated = value
        if old_set_id:
            updated = updated.replace(old_set_id, new_set_id)
        if old_led_id:
            updated = updated.replace(old_led_id, new_led_id)
        if updated != value:
            params[key] = updated


def _clone_equipment_def(source_eq, new_name, new_eq_id, next_set_num):
    eq_copy = copy.deepcopy(source_eq)
    eq_copy["id"] = new_eq_id
    eq_copy["name"] = new_name
    truth_id = source_eq.get("ced_truth_source_id") or source_eq.get("id") or source_eq.get("name") or new_eq_id
    truth_name = source_eq.get("ced_truth_source_name") or source_eq.get("name") or new_name
    eq_copy["ced_truth_source_id"] = truth_id
    eq_copy["ced_truth_source_name"] = truth_name

    parent_filter = eq_copy.get("parent_filter")
    if isinstance(parent_filter, dict):
        family_name, type_name = _split_label(new_name)
        if family_name:
            parent_filter["family_name_pattern"] = family_name
        if ":" in (new_name or ""):
            parent_filter["type_name_pattern"] = type_name or "*"

    linked_sets = eq_copy.get("linked_sets") or []
    if not linked_sets:
        new_set_id = "SET-{:03d}".format(next_set_num)
        next_set_num += 1
        eq_copy["linked_sets"] = [{
            "id": new_set_id,
            "name": "{} Types".format(new_name),
            "linked_element_definitions": [],
        }]
        return eq_copy, next_set_num

    for idx, linked_set in enumerate(linked_sets):
        if not isinstance(linked_set, dict):
            continue
        old_set_id = linked_set.get("id")
        new_set_id = "SET-{:03d}".format(next_set_num)
        next_set_num += 1
        linked_set["id"] = new_set_id
        if idx == 0:
            linked_set["name"] = "{} Types".format(new_name)
        led_list = linked_set.get("linked_element_definitions") or []
        counter = 0
        for led in led_list:
            if not isinstance(led, dict):
                continue
            old_led_id = led.get("id")
            if led.get("is_parent_anchor"):
                new_led_id = "{}-LED-000".format(new_set_id)
            else:
                counter += 1
                new_led_id = "{}-LED-{:03d}".format(new_set_id, counter)
            led["id"] = new_led_id
            _rewrite_linker_payload(led.get("parameters"), old_set_id, new_set_id, old_led_id, new_led_id)
    return eq_copy, next_set_num


def _pick_source_file():
    dialog = OpenFileDialog()
    dialog.Title = "Select CAD Block Names Sheet"
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


def _column_letter(index):
    value = index + 1
    letters = []
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _header_score(value):
    normalized = _normalize_full_name(value)
    if not normalized:
        return 0
    compact = normalized.replace(" ", "")
    if "cadblockname" in compact:
        return 100
    tokens = set(normalized.split())
    if "cad" in tokens and "block" in tokens and "name" in tokens:
        return 100
    if "block" in tokens and "name" in tokens:
        return 90
    if "cad" in tokens and "block" in tokens:
        return 85
    if "name" in tokens:
        return 65
    if "block" in tokens:
        return 55
    return 0


def _column_sample(rows, index, limit=3):
    values = []
    for row in rows:
        if index < len(row):
            value = (row[index] or "").strip()
            if value:
                values.append(value)
        if len(values) >= limit:
            break
    return ", ".join(values)


def _choose_column(headers, rows, preferred_idx):
    choices = []
    for index, header in enumerate(headers):
        column_name = "Column {}".format(_column_letter(index))
        display_header = (header or "").strip() or column_name
        sample = _column_sample(rows, index)
        if sample:
            label = "{} [{}] - sample: {}".format(display_header, column_name, sample)
        else:
            label = "{} [{}]".format(display_header, column_name)
        choices.append(_ColumnChoice(label, index))

    if not choices:
        return None
    if len(choices) == 1:
        return choices[0].index

    default_choice = None
    for choice in choices:
        if choice.index == preferred_idx:
            default_choice = choice
            break
    if default_choice:
        ordered = [default_choice] + [choice for choice in choices if choice is not default_choice]
    else:
        ordered = choices

    selected = forms.SelectFromList.show(
        ordered,
        title="Select CAD Block Name Column",
        multiselect=False,
        button_name="Use Column",
        width=900,
        height=600,
    )
    if not selected:
        return None
    if isinstance(selected, list):
        selected = selected[0] if selected else None
    return selected.index if selected else None


def _load_cad_name_counts(path):
    extension = os.path.splitext(path)[1].lower()
    if extension == ".csv":
        matrix = _read_csv_matrix(path)
    else:
        matrix = _read_excel_matrix(path)
    if not matrix:
        return Counter(), ""

    ncols = max(len(row) for row in matrix)
    padded = []
    for row in matrix:
        if len(row) < ncols:
            row = row + [""] * (ncols - len(row))
        padded.append([(cell or "").strip() for cell in row])

    first_row = padded[0]
    scores = [_header_score(cell) for cell in first_row]
    has_header = max(scores) >= 80
    headers = first_row[:] if has_header else ["Column {}".format(_column_letter(i)) for i in range(ncols)]
    data_rows = padded[1:] if has_header else padded

    if not data_rows:
        return Counter(), ""

    preferred_index = 0
    if scores:
        best_score = max(scores)
        if best_score > 0:
            preferred_index = scores.index(best_score)
        else:
            counts = [sum(1 for row in data_rows if (row[i] or "").strip()) for i in range(ncols)]
            preferred_index = counts.index(max(counts)) if counts else 0

    if has_header and scores.count(max(scores)) == 1 and max(scores) >= 85:
        selected_index = preferred_index
    elif ncols == 1:
        selected_index = 0
    else:
        selected_index = _choose_column(headers, data_rows, preferred_index)
        if selected_index is None:
            return None, ""

    counts = Counter()
    for row in data_rows:
        if selected_index >= len(row):
            continue
        value = (row[selected_index] or "").strip()
        if not value:
            continue
        counts[value] += 1

    selected_header = headers[selected_index] if selected_index < len(headers) else "Column {}".format(_column_letter(selected_index))
    return counts, selected_header


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active Revit document.", title=TITLE)
        return

    source_path = _pick_source_file()
    if not source_path:
        return

    try:
        cad_counts, selected_column = _load_cad_name_counts(source_path)
    except Exception as exc:
        forms.alert("Failed to read source file:\n\n{}".format(exc), title=TITLE)
        return
    if cad_counts is None:
        return
    if not cad_counts:
        forms.alert("No CAD block names were found in the selected file.", title=TITLE)
        return

    try:
        _, data = load_active_yaml_data()
    except Exception as exc:
        forms.alert(str(exc), title=TITLE)
        return

    profile_names = _collect_profile_names(data)
    if not profile_names:
        forms.alert("No profiles found in the active YAML.", title=TITLE)
        return

    profile_norms = {_normalize_full_name(name) for name in profile_names if name}
    profile_base_map = {}
    profile_default_heads = set()
    for name in profile_names:
        base = _base_key(name)
        if not base:
            continue
        profile_base_map.setdefault(base, []).append(name)
        head, type_name = _split_label(name)
        type_norm = _normalize_full_name(type_name)
        if type_norm == "default":
            head_norm = _normalize_full_name(head)
            if head_norm:
                profile_default_heads.add(head_norm)

    missing = []
    for cad_name, count in cad_counts.items():
        norm = _normalize_full_name(cad_name)
        if norm in profile_norms:
            continue
        if ":" not in cad_name and norm in profile_default_heads:
            continue
        base = _base_key(cad_name)
        if not base:
            continue
        candidates = profile_base_map.get(base) or []
        if candidates:
            best, _ = _best_match_name(cad_name, candidates)
            if best:
                missing.append((cad_name, best, count))
            continue
        best, score = _best_match_name(cad_name, profile_names)
        if not best:
            continue
        if score < MIN_PARTIAL_SCORE:
            continue
        overlap_count, _small_cov, large_cov = _token_overlap_metrics(cad_name, best)
        if overlap_count < MIN_COMMON_TOKENS:
            continue
        if large_cov < MIN_LARGER_TOKEN_COVERAGE:
            continue
        missing.append((cad_name, best, count))

    if not missing:
        forms.alert("No close-name profile gaps were found in the selected CAD source.", title=TITLE)
        return

    missing.sort(key=lambda row: row[0].lower())
    items = []
    cad_header = "CAD Block Name"
    profile_header = "YAML Profile"
    count_header = "Count"
    cad_width = max(
        len(cad_header),
        max([len(str(row[0] or "")) for row in missing] or [0]),
    )
    profile_width = max(
        len(profile_header),
        max([len(str(row[1] or "")) for row in missing] or [0]),
    )
    cad_width = min(cad_width, 64)
    profile_width = min(profile_width, 64)
    header_label = "{:<{w1}} | {:<{w2}} | {:>5}".format(
        cad_header,
        profile_header,
        count_header,
        w1=cad_width,
        w2=profile_width,
    )
    divider_label = "{}-+-{}-+-{}".format("-" * cad_width, "-" * profile_width, "-" * 5)
    items.append(_MissingChoice(header_label, data=None, checked=False))
    items.append(_MissingChoice(divider_label, data=None, checked=False))
    for cad_name, best, count in missing:
        cad_cell = _truncate_text(cad_name, cad_width)
        profile_cell = _truncate_text(best, profile_width)
        if count > 1:
            count_cell = "x{}".format(count)
        else:
            count_cell = ""
        label = "{:<{w1}} | {:<{w2}} | {:>5}".format(
            cad_cell,
            profile_cell,
            count_cell,
            w1=cad_width,
            w2=profile_width,
        )
        items.append(_MissingChoice(label, (cad_name, best, count), checked=True))

    selected = forms.SelectFromList.show(
        items,
        title=TITLE,
        multiselect=True,
        button_name="Create Selected",
        width=1150,
        height=600,
    )
    if not selected:
        return
    missing = [
        item.data
        for item in selected
        if isinstance(item, _MissingChoice) and item.data
    ]
    if not missing:
        return

    equipment_defs = data.get("equipment_definitions") or []
    existing_norms = {
        _normalize_full_name(eq.get("name") or eq.get("id") or "")
        for eq in equipment_defs
        if isinstance(eq, dict)
    }
    next_eq_num = _next_eq_number(data)
    next_set_num = _next_set_number(data)
    created = []
    skipped_existing = []
    skipped_unresolved = []
    for cad_name, best, _ in missing:
        new_norm = _normalize_full_name(cad_name)
        if new_norm in existing_norms:
            skipped_existing.append(cad_name)
            continue
        source_eq = _find_equipment_def_by_name(data, best) or _find_equipment_def_by_label(data, best)
        if not source_eq:
            skipped_unresolved.append((cad_name, best))
            continue
        new_eq_id = "EQ-{:03d}".format(next_eq_num)
        next_eq_num += 1
        eq_copy, next_set_num = _clone_equipment_def(source_eq, cad_name, new_eq_id, next_set_num)
        equipment_defs.append(eq_copy)
        existing_norms.add(new_norm)
        created.append((cad_name, best))

    if not created:
        forms.alert("No new profiles were created.", title=TITLE)
        return

    try:
        save_active_yaml_data(
            None,
            data,
            "Audit CAD Model",
            "Created {} profile(s) from CAD audit".format(len(created)),
        )
    except Exception as exc:
        forms.alert("Failed to save updates:\n\n{}".format(exc), title=TITLE)
        return

    summary = [
        "Source file: {}".format(source_path),
        "Column used: {}".format(selected_column or "<unknown>"),
        "",
        "Created {} profile(s).".format(len(created)),
    ]
    summary.extend(" - {} -> {}".format(name, best) for name, best in created[:20])
    if len(created) > 20:
        summary.append(" (+{} more)".format(len(created) - 20))
    if skipped_existing:
        summary.append("")
        summary.append("Skipped existing:")
        summary.extend(" - {}".format(name) for name in skipped_existing[:10])
        if len(skipped_existing) > 10:
            summary.append(" (+{} more)".format(len(skipped_existing) - 10))
    if skipped_unresolved:
        summary.append("")
        summary.append("Skipped (no source profile found):")
        summary.extend(" - {} -> {}".format(name, best) for name, best in skipped_unresolved[:10])
        if len(skipped_unresolved) > 10:
            summary.append(" (+{} more)".format(len(skipped_unresolved) - 10))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
