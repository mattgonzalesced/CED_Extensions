# -*- coding: utf-8 -*-
import re
from collections import defaultdict

from pyrevit import DB, forms, revit, script
from pyrevit.interop import xl as pyxl

doc = revit.doc
logger = script.get_logger()
output = script.get_output()

try:
    text_type = unicode
except NameError:
    text_type = str


TAG_PARAM_CANDIDATES = [
    "Unit Tags",
    "Unit Tag",
    "Mark",
    "Identity Mark",
    "Tag",
]


PARAM_MAPPINGS = [
    {"excel": "Model Number", "revit": ["Model Number"], "convert": "text"},
    {"excel": "Tonnage", "revit": ["Tons"], "convert": "number"},
    {"excel": "Gross Total Capacity", "revit": ["Cooling Coil Total", "Cooling Coil Total - UNITS"], "convert": "mbh_to_btuh"},
    {"excel": "Gross Sensible Capacity", "revit": ["Cooling Coil Sensible", "Cooling Coil Sensible - UNITS"], "convert": "mbh_to_btuh"},
    {"excel": "Heating Input Capacity", "revit": ["Gas Heat Input Rate", "Gas Heat Input Rate - READ ONLY TYPE param"], "convert": "number"},
    {"excel": "Output Heating Capacity", "revit": ["Gas Heat Output Capacity", "Gas Heat Output Capacity - READ ONLY TYPE param"], "convert": "number"},
    {"excel": "Design Airflow", "revit": ["Design Airflow Rate"], "convert": "number"},
    {"excel": "EER @ AHRI", "revit": ["Cooling EER"], "convert": "number"},
    {"excel": "IEER @ AHRI", "revit": ["Cooling IEER"], "convert": "number"},
    {"excel": "Supply Motor Horsepower", "revit": ["Supply Fan bhp"], "convert": "number"},
    {"excel": "Design ESP", "revit": ["Supply Fan ESP", "Supply Fan ESP - UNITS"], "convert": "esp_inh2o_to_wg"},
    {"excel": "Approx Installed Weight", "revit": ["Operating Weight"], "convert": "number"},
    {"excel": "Unit Voltage", "revit": ["Voltage_CED", "Voltage_CED - Type Param"], "convert": "text"},
    {"excel": "MCA", "revit": ["Circuit 1 FLA Input_CED", "Circuit 1 FLA Input_CED - Type Param"], "convert": "number"},
    {"excel": "MOP", "revit": ["Circuit 1 MOCP Input_CED", "Circuit 1 MOCP Input_CED - Type Param"], "convert": "number"},
]


def as_text(value):
    if value is None:
        return ""
    if isinstance(value, text_type):
        return value.strip()
    return text_type(value).strip()


def normalize_key(value):
    text = as_text(value)
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().upper()


def is_blank(value):
    return normalize_key(value) == ""


def get_cell(row, index):
    if row is None:
        return None
    if index >= len(row):
        return None
    return row[index]


def try_parse_number(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = as_text(value)
    if not text:
        return None
    text = text.replace(",", "")
    match = re.search(r"[-+]?\d*\.?\d+", text)
    if not match:
        return None
    try:
        return float(match.group(0))
    except Exception:
        return None


def convert_value(raw_value, mode):
    if mode == "text":
        return as_text(raw_value)
    if mode == "number":
        return try_parse_number(raw_value)
    if mode == "mbh_to_btuh":
        value = try_parse_number(raw_value)
        if value is None:
            return None
        return value * 1000.0
    if mode == "esp_inh2o_to_wg":
        # Inches of water (in H2O) and inches water gauge (in. w.g.) are equivalent.
        return try_parse_number(raw_value)
    return raw_value


def read_param_display_value(param):
    if not param:
        return None
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString() or param.AsValueString()
        if param.StorageType == DB.StorageType.Integer:
            return param.AsValueString() or text_type(param.AsInteger())
        if param.StorageType == DB.StorageType.Double:
            return param.AsValueString() or text_type(param.AsDouble())
        if param.StorageType == DB.StorageType.ElementId:
            return param.AsValueString()
    except Exception:
        return None
    return None


def locate_type_element(element):
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != DB.ElementId.InvalidElementId:
            return doc.GetElement(type_id)
    except Exception:
        return None
    return None


def set_parameter_value(param, value):
    if param is None:
        return False, "missing"
    if param.IsReadOnly:
        return False, "readonly"

    try:
        storage_type = param.StorageType
        if storage_type == DB.StorageType.String:
            param.Set(as_text(value))
            return True, "ok"

        if storage_type == DB.StorageType.Integer:
            number = try_parse_number(value)
            if number is None:
                return False, "invalid-number"
            param.Set(int(round(number)))
            return True, "ok"

        if storage_type == DB.StorageType.Double:
            number = try_parse_number(value)
            if number is None:
                return False, "invalid-number"
            try:
                param.SetValueString(str(number))
            except Exception:
                param.Set(float(number))
            return True, "ok"

        if storage_type == DB.StorageType.ElementId:
            if isinstance(value, DB.ElementId):
                param.Set(value)
                return True, "ok"
            return False, "unsupported-elementid"

        return False, "unsupported-storage"
    except Exception as ex:
        return False, "error: {}".format(ex)


def resolve_write_target(element, revit_param_names):
    type_element = locate_type_element(element)
    readonly_found = False

    for param_name in revit_param_names:
        instance_param = element.LookupParameter(param_name)
        if instance_param:
            if not instance_param.IsReadOnly:
                return "instance", element, instance_param, param_name
            readonly_found = True

        if type_element:
            type_param = type_element.LookupParameter(param_name)
            if type_param:
                if not type_param.IsReadOnly:
                    return "type", type_element, type_param, param_name
                readonly_found = True

    if readonly_found:
        return None, None, None, "readonly"
    return None, None, None, "missing"


def find_header_row(rows):
    for idx, row in enumerate(rows):
        for cell in row:
            key = normalize_key(cell)
            if key in ("UNIT TAGS", "UNIT TAG"):
                return idx
    return None


def build_headers(rows, header_row_idx):
    upper = rows[header_row_idx] if header_row_idx < len(rows) else []
    lower = rows[header_row_idx + 1] if (header_row_idx + 1) < len(rows) else []

    max_cols = max(len(upper), len(lower))
    headers = {}
    tag_col = None
    for col_idx in range(max_cols):
        top = as_text(get_cell(upper, col_idx))
        bottom = as_text(get_cell(lower, col_idx))
        header_text = top if top else bottom
        if not header_text:
            continue
        header_text = re.sub(r"\s+", " ", header_text).strip()
        headers[col_idx] = header_text
        if normalize_key(header_text) in ("UNIT TAGS", "UNIT TAG"):
            tag_col = col_idx
    return headers, tag_col


def pick_sheet_name(xldata):
    sheet_names = sorted(xldata.keys())
    if not sheet_names:
        forms.alert("No sheets were found in the selected workbook.", exitscript=True)
    if len(sheet_names) == 1:
        return sheet_names[0]
    selected = forms.SelectFromList.show(sheet_names, title="Select RTU Schedule Sheet", multiselect=False)
    if not selected:
        script.exit()
    return selected


def load_excel_rows():
    path = forms.pick_file(title="Select RTU Schedule Excel File", multi_file=False)
    if not path:
        script.exit()

    xldata = pyxl.load(path, headers=False)
    sheet_name = pick_sheet_name(xldata)
    rows = xldata[sheet_name].get("rows", [])
    if not rows:
        forms.alert("Selected sheet has no rows.", exitscript=True)

    header_row_idx = find_header_row(rows)
    if header_row_idx is None:
        forms.alert("Could not find a header row with 'Unit Tags'.", exitscript=True)

    headers, tag_col_idx = build_headers(rows, header_row_idx)
    if tag_col_idx is None:
        forms.alert("Could not locate 'Unit Tags' column.", exitscript=True)

    unit_tag_header = headers[tag_col_idx]
    parsed_rows = []
    for row_idx in range(header_row_idx + 1, len(rows)):
        row = rows[row_idx]
        tag_value = get_cell(row, tag_col_idx)
        if is_blank(tag_value):
            continue

        if normalize_key(tag_value) in ("UNIT TAGS", "UNIT TAG"):
            continue

        row_dict = {}
        for col_idx, header_name in headers.items():
            row_dict[header_name] = get_cell(row, col_idx)
        parsed_rows.append(row_dict)

    if not parsed_rows:
        forms.alert("No RTU data rows found under 'Unit Tags'.", exitscript=True)

    return path, sheet_name, unit_tag_header, parsed_rows


def collect_family_instances():
    def _collect(collector):
        result = []
        for elem in collector:
            try:
                if elem.ViewSpecific:
                    continue
            except Exception:
                pass
            result.append(elem)
        return result

    mech_instances = _collect(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.FamilyInstance)
        .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if mech_instances:
        return mech_instances

    return _collect(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.FamilyInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def build_tag_lookup(elements):
    lookup = defaultdict(list)
    for elem in elements:
        seen_for_element = set()
        for priority, param_name in enumerate(TAG_PARAM_CANDIDATES):
            value = read_param_display_value(elem.LookupParameter(param_name))
            key = normalize_key(value)
            if not key or key in seen_for_element:
                continue
            lookup[key].append((priority, elem))
            seen_for_element.add(key)
    return lookup


def choose_element_for_tag(matches):
    if not matches:
        return None
    sorted_matches = sorted(matches, key=lambda item: (item[0], item[1].Id.IntegerValue))
    return sorted_matches[0][1], sorted_matches


def apply_row_to_element(row, unit_tag, target_element):
    row_writes = 0
    row_issues = []

    for mapping in PARAM_MAPPINGS:
        excel_header = mapping["excel"]
        revit_names = mapping["revit"]
        raw_value = row.get(excel_header)

        if is_blank(raw_value):
            continue

        converted = convert_value(raw_value, mapping["convert"])
        if converted is None:
            row_issues.append("{} -> {}: invalid value '{}'".format(excel_header, revit_names[0], raw_value))
            continue

        target_kind, target_owner, param, resolved_name = resolve_write_target(target_element, revit_names)
        if param is None:
            if resolved_name == "readonly":
                row_issues.append("{} ({}) is read-only.".format(revit_names[0], unit_tag))
            else:
                row_issues.append("{} ({}) was not found.".format(revit_names[0], unit_tag))
            continue

        ok, detail = set_parameter_value(param, converted)
        if ok:
            row_writes += 1
        else:
            row_issues.append("{} ({}, {}): {}".format(resolved_name, unit_tag, target_kind, detail))

    return row_writes, row_issues


def main():
    output.close_others()
    output.show()
    output.set_title("Insert RTU Params")

    path, sheet_name, unit_tag_header, excel_rows = load_excel_rows()

    elements = collect_family_instances()
    if not elements:
        forms.alert("No family instances found in the model.", exitscript=True)

    tag_lookup = build_tag_lookup(elements)

    stats = {
        "rows_total": len(excel_rows),
        "rows_matched": 0,
        "params_written": 0,
        "issues": 0,
    }
    missing_tags = []
    duplicate_tag_matches = []
    issue_lines = []

    with DB.Transaction(doc, "Insert RTU Params from Excel") as tx:
        tx.Start()
        for row in excel_rows:
            unit_tag = as_text(row.get(unit_tag_header))
            unit_key = normalize_key(unit_tag)
            if not unit_key:
                continue

            matches = tag_lookup.get(unit_key, [])
            if not matches:
                missing_tags.append(unit_tag)
                continue

            target = choose_element_for_tag(matches)
            if not target:
                missing_tags.append(unit_tag)
                continue

            target_elem, sorted_matches = target
            if len(sorted_matches) > 1:
                duplicate_tag_matches.append(
                    "{} -> {}".format(
                        unit_tag,
                        ", ".join([str(m[1].Id.IntegerValue) for m in sorted_matches[:5]])
                    )
                )

            stats["rows_matched"] += 1
            writes, row_issues = apply_row_to_element(row, unit_tag, target_elem)
            stats["params_written"] += writes

            if row_issues:
                stats["issues"] += len(row_issues)
                issue_lines.extend(row_issues)
        tx.Commit()

    output.print_md("## Insert RTU Params")
    output.print_md("**Workbook:** `{}`".format(path))
    output.print_md("**Sheet:** `{}`".format(sheet_name))
    output.print_md("**Excel rows processed:** `{}`".format(stats["rows_total"]))
    output.print_md("**Rows matched to model elements:** `{}`".format(stats["rows_matched"]))
    output.print_md("**Parameters written:** `{}`".format(stats["params_written"]))

    if missing_tags:
        output.print_md("\n### Missing Unit Tags")
        for tag in sorted(set(missing_tags)):
            output.print_md("- `{}`".format(tag))

    if duplicate_tag_matches:
        output.print_md("\n### Duplicate Tag Matches (first few element ids shown)")
        for line in sorted(set(duplicate_tag_matches)):
            output.print_md("- {}".format(line))

    if issue_lines:
        output.print_md("\n### Write Issues")
        for line in issue_lines[:150]:
            output.print_md("- {}".format(line))
        if len(issue_lines) > 150:
            output.print_md("- ...and {} more".format(len(issue_lines) - 150))
    else:
        output.print_md("\nAll mapped parameters were written without reported issues.")


if __name__ == "__main__":
    main()
