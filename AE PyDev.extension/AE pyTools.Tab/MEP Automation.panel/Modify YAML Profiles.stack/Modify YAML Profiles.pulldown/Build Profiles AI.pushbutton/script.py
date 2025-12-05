# -*- coding: utf-8 -*-
"""
Build Profiles AI
-----------------
Reads an Excel workbook describing equipment definitions and their linked
element types, then builds those definitions directly inside the active YAML
store (Extensible Storage). Each Excel row becomes a linked element definition,
similar to running Add YAML Profiles manually.
"""

import math
import os
import clr

from pyrevit import forms, revit, script

clr.AddReference("System")
clr.AddReference("System.Reflection")
from System import Type, Activator, Array, Object  # noqa: E402
from System.Reflection import BindingFlags  # noqa: E402

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in os.sys.path:
    os.sys.path.append(LIB_ROOT)

from profile_schema import ensure_equipment_definition, get_type_set, next_led_id  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

LOG = script.get_logger()

RADIUS_INCHES = 6.0


def _args_array(*items):
    return Array[Object](list(items))


def _set(obj, prop, val):
    try:
        obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(val))
    except Exception:
        pass


def _get(obj, prop):
    try:
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except Exception:
        return None


def _call(obj, name, *args):
    t = obj.GetType()
    try:
        return t.InvokeMember(name, BindingFlags.GetProperty, None, obj, _args_array(*args))
    except Exception:
        try:
            return t.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
        except Exception:
            return None


def _cell(cells, r, c):
    item = _call(cells, "Item", r, c)
    val = _get(item, "Value2")
    if val is None:
        return ""
    return str(val).strip()


def read_excel_rows(path):
    LOG.info("Reading Excel file: %s", path)
    xl = wb = ws = used = cells = None
    rows = []
    try:
        t = Type.GetTypeFromProgID("Excel.Application")
        if t is None:
            raise RuntimeError("Excel is not installed (COM automation unavailable).")
        xl = Activator.CreateInstance(t)
        _set(xl, "Visible", False)
        _set(xl, "DisplayAlerts", False)
        workbooks = _get(xl, "Workbooks")
        wb = _call(workbooks, "Open", path)
        sheets = _get(wb, "Worksheets")
        ws = _call(sheets, "Item", 1)
        used = _get(ws, "UsedRange")
        cells = _get(used, "Cells")
        nrows = int(_get(_get(used, "Rows"), "Count") or 0)
        ncols = int(_get(_get(used, "Columns"), "Count") or 0)
        if nrows < 2 or ncols < 2:
            raise RuntimeError("Excel sheet is empty.")
        headers = [_cell(cells, 1, c) for c in range(1, ncols + 1)]
        header_map = {h.strip(): idx + 1 for idx, h in enumerate(headers) if h}
        if len(headers) < 2:
            raise RuntimeError("Excel sheet must include at least two columns (Equipment, Element Label).")
        for r in range(2, nrows + 1):
            row = {}
            for name, col in header_map.items():
                row[name] = _cell(cells, r, col)
            rows.append(row)
        return headers, rows
    finally:
        try:
            if wb:
                _call(wb, "Close", False)
        except Exception:
            pass
        try:
            if xl:
                _call(xl, "Quit")
        except Exception:
            pass


def _find_value(row, header_name):
    value = row.get(header_name, "")
    if value is None:
        return ""
    return str(value).strip()


def _parse_float(value, default=0.0):
    try:
        if value is None:
            return default
        s = str(value).strip()
        if not s:
            return default
        return float(s)
    except Exception:
        return default


def _parse_bool(value, default=False):
    if value is None:
        return default
    s = str(value).strip().lower()
    if s in ("1", "true", "yes", "y", "on"):
        return True
    if s in ("0", "false", "no", "n", "off"):
        return False
    return default


def _parse_parameter_pairs(raw_value):
    data = {}
    if not raw_value:
        return data
    for chunk in raw_value.split(","):
        part = chunk.strip()
        if not part:
            continue
        if "=" in part:
            key, val = part.split("=", 1)
        elif ":" in part:
            key, val = part.split(":", 1)
        else:
            continue
        key = key.strip()
        val = val.strip()
        if key:
            data[key] = val
    return data


def build_led_entries(row_info, equipment_def, type_set):
    led_list = type_set.setdefault("linked_element_definitions", [])
    label_block = row_info["label"]
    labels = []
    if label_block:
        for piece in label_block.split(";"):
            candidate = piece.strip()
            if candidate:
                labels.append(candidate)
    if not labels:
        fam = row_info.get("family")
        typ = row_info.get("type")
        if fam and typ:
            labels = [u"{} : {}".format(fam, typ)]
        elif typ:
            labels = [typ]
        elif fam:
            labels = [fam]
        else:
            raise RuntimeError("Row is missing an element label and family/type information.")
    entries = []
    for label in labels:
        led_id = next_led_id(type_set, equipment_def)
        raw_params = row_info.get("parameters_dict") or {}
        clean_params = {}
        for key, val in raw_params.items():
            if val is None:
                continue
            sval = str(val).strip()
            if not sval:
                continue
            clean_params[key] = sval
        entry = {
            "id": led_id,
            "label": label,
            "category": row_info.get("category") or "",
            "is_group": row_info.get("is_group", False),
            "offsets": [
                {
                    "x_inches": _parse_float(row_info.get("x"), 0.0),
                    "y_inches": _parse_float(row_info.get("y"), 0.0),
                    "z_inches": _parse_float(row_info.get("z"), 0.0),
                    "rotation_deg": _parse_float(row_info.get("rotation"), 0.0),
                }
            ],
            "parameters": clean_params,
            "tags": [],
        }
        led_list.append(entry)
        entries.append(entry)
    return entries


def ensure_parent_filter(equipment_def, row_info):
    parent = equipment_def.get("parent_filter") or {}
    if row_info.get("parent_category"):
        parent["category"] = row_info["parent_category"]
    if row_info.get("parent_family"):
        parent["family_name_pattern"] = row_info["parent_family"]
    if row_info.get("parent_type"):
        parent["type_name_pattern"] = row_info["parent_type"]
    equipment_def["parent_filter"] = parent


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title="Build Profiles AI")
        return
    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title="Build Profiles AI")
        return
    if "equipment_definitions" not in data or not isinstance(data.get("equipment_definitions"), list):
        data["equipment_definitions"] = []
    excel_path = forms.pick_file(file_ext="xlsx", title="Select equipment definition Excel")
    if not excel_path:
        return
    try:
        headers, raw_rows = read_excel_rows(excel_path)
    except Exception as exc:
        forms.alert("Failed to read Excel file:\n{}".format(exc), title="Build Profiles AI")
        return

    equipment_header = headers[0]
    label_header = headers[1]
    parameter_headers = headers[2:]

    processed_rows = []
    for row in raw_rows:
        eq_name = _find_value(row, equipment_header)
        lbl = _find_value(row, label_header)
        if not eq_name:
            continue
        info = {
            "equipment": eq_name,
            "label": lbl,
            "is_group": False,
            "category": "",
            "family": "",
            "type": "",
            "parameters_dict": {},
        }
        params = {}
        for header in parameter_headers:
            value = _find_value(row, header)
            if value:
                params[header] = value
        info["parameters_dict"] = params
        processed_rows.append(info)

    if not processed_rows:
        forms.alert("Excel file contained no usable rows.", title="Build Profiles AI")
        return

    replace_existing = forms.alert(
        "Replace linked element definitions for equipment already present?",
        title="Build Profiles AI",
        yes=True,
        no=True,
    )
    if replace_existing is None:
        return

    equipment_defs = data.get("equipment_definitions") or []
    defs_by_name = {d.get("name"): d for d in equipment_defs if isinstance(d, dict)}
    cleared = set()
    created = 0
    updated = 0
    grouped = {}
    for entry in processed_rows:
        grouped.setdefault(entry["equipment"], []).append(entry)

    for eq_name, rows in grouped.items():
        first_row = rows[0]
        sample_entry = {
            "label": first_row.get("label") or "",
            "category_name": first_row.get("category") or "Uncategorized",
        }
        equipment_def = defs_by_name.get(eq_name)
        if equipment_def is None:
            ensure_equipment_definition(data, eq_name, sample_entry)
            equipment_def = next(d for d in data["equipment_definitions"] if d.get("name") == eq_name)
            defs_by_name[eq_name] = equipment_def
            cleared.add(eq_name)  # ensures we don't attempt to clear again
            created += 1
        type_set = get_type_set(equipment_def)
        if replace_existing and eq_name not in cleared:
            type_set["linked_element_definitions"] = []
            cleared.add(eq_name)
        n_rows = len(rows)
        for idx, row in enumerate(rows):
            angle_deg = 0.0 if n_rows <= 1 else (360.0 / float(n_rows)) * idx
            angle_rad = math.radians(angle_deg)
            row["x"] = RADIUS_INCHES * math.cos(angle_rad)
            row["y"] = RADIUS_INCHES * math.sin(angle_rad)
            row["z"] = _parse_float(row.get("z"), 0.0)
            try:
                build_led_entries(row, equipment_def, type_set)
            except RuntimeError as ex:
                LOG.warning("Skipping row for '%s': %s", eq_name, ex)
                continue
        ensure_parent_filter(equipment_def, first_row)
        if eq_name not in cleared or replace_existing:
            updated += 1

    try:
        save_active_yaml_data(
            None,
            data,
            "Build Profiles AI",
            "Imported {} equipment definitions from Excel".format(len(grouped)),
        )
    except Exception as exc:
        forms.alert("Failed to save YAML data:\n{}".format(exc), title="Build Profiles AI")
        return

    forms.alert(
        "Processed {} equipment definitions ({} created, {} updated) from '{}'.\n"
        "Run QA/QC or Place Linked Elements to review the new profiles.".format(
            len(grouped), created, updated, os.path.basename(excel_path)
        ),
        title="Build Profiles AI",
    )


if __name__ == "__main__":
    main()
