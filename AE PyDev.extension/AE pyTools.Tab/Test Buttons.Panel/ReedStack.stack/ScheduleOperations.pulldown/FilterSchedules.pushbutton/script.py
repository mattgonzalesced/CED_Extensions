# -*- coding: utf-8 -*-
# Revit 2024 / IronPython 2.7 (pyRevit)
# Batch-apply schedule filters from a simple rules file to multiple schedules.
#
# Rules example:
# [Mark, Equals, 109];
# [Height, Equals, 7' - 0"];
# [Width,  = , 3' - 0"]
#
from __future__ import print_function

import re
import sys
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    Transaction,
    ScheduleFilter,
    ScheduleFilterType,
    ElementId,
)

# Revit 2021+ data type API (use if present)
try:
    from Autodesk.Revit.DB import SpecTypeId, ForgeTypeId
    HAS_SPECTYPE = True
except Exception:
    SpecTypeId = None
    ForgeTypeId = None
    HAS_SPECTYPE = False

from pyrevit import forms, revit, script

# ------------------ CONFIG ------------------
CLEAR_EXISTING = True  # set False to append filters instead of replacing
# --------------------------------------------

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

# ------------------ Helpers ------------------
def parse_feet_inches(text):
    s = text.strip().replace("’", "'").replace("“", '"').replace("”", '"')
    if "'" not in s and '"' not in s:
        try:
            return float(s)
        except:
            raise ValueError("Not a numeric length: {}".format(text))

    feet = 0.0
    inches_total = 0.0

    feet_match = re.search(r"(-?\d+(\.\d+)?)\s*'", s)
    if feet_match:
        feet = float(feet_match.group(1))

    inch_match = re.search(r"(-?\d+(?:\.\d+)?(?:\s+\d+/\d+)?)\s*\"", s)
    if inch_match:
        inch_str = inch_match.group(1).strip()
        if " " in inch_str:
            whole, frac = inch_str.split(" ", 1)
            inches_total = float(whole) + eval_fraction(frac)
        elif "/" in inch_str:
            inches_total = eval_fraction(inch_str)
        else:
            inches_total = float(inch_str)

    return feet + (inches_total / 12.0)

def eval_fraction(frac_text):
    num, den = frac_text.split("/", 1)
    return float(num) / float(den)

def normalize_op(op_raw):
    s = op_raw.strip().lower()
    mapping = {
        "equals": "equal", "equal": "equal", "=": "equal", "==": "equal",
        "not equals": "not_equal", "not equal": "not_equal", "!=": "not_equal",
        "contains": "contains", "does contain": "contains",
        "beginswith": "beginswith", "begins with": "beginswith",
        "startswith": "beginswith", "starts with": "beginswith",
        "endswith": "endswith", "ends with": "endswith",
        ">": "greater", "greater than": "greater",
        ">=": "greaterorequal", "greater or equal": "greaterorequal",
        "<": "less", "less than": "less",
        "<=": "lessorequal", "less or equal": "lessorequal",
    }
    return mapping.get(s, s)

def op_to_enum(op_norm):
    m = {
        "equal": ScheduleFilterType.Equal,
        "not_equal": ScheduleFilterType.NotEqual,
        "contains": ScheduleFilterType.Contains,
        "beginswith": ScheduleFilterType.BeginsWith,
        "endswith": ScheduleFilterType.EndsWith,
        "greater": ScheduleFilterType.GreaterThan,
        "greaterorequal": ScheduleFilterType.GreaterThanOrEqual,
        "less": ScheduleFilterType.LessThan,
        "lessorequal": ScheduleFilterType.LessThanOrEqual,
    }
    if op_norm not in m:
        raise ValueError("Unsupported operator: {}".format(op_norm))
    return m[op_norm]

def parse_rules(text_blob):
    blocks = re.findall(r"\[(.*?)\]", text_blob, flags=re.DOTALL)
    if not blocks:
        blocks = [seg for seg in text_blob.split(";") if seg.strip()]

    rules = []
    for blk in blocks:
        seg = blk if "[" not in blk else blk.strip("[]")
        parts = [p.strip() for p in seg.split(",")]
        parts = [p for p in parts if p != ""]
        if len(parts) != 3:
            parts = re.split(r"[\t,]+", seg)
            parts = [p.strip() for p in parts if p.strip()]
        if len(parts) != 3:
            raise ValueError("Could not parse rule (need 3 items): {}".format(seg))
        field, op, val = parts
        if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
            val = val[1:-1]
        rules.append({"field": field, "op": normalize_op(op), "value": val})
    return rules

def list_view_schedules(document):
    return [s for s in FilteredElementCollector(document).OfClass(ViewSchedule)]

def pick_schedules(schedules):
    items = [{"label": u"{0} [{1}]".format(s.Name, s.Id.IntegerValue), "id": s.Id}
             for s in schedules if not s.IsTemplate]
    if not items:
        forms.alert("No non-template schedules found.", exitscript=True)
    labels = sorted([it["label"] for it in items], key=lambda x: x.lower())
    sel = forms.SelectFromList.show(labels, title="Select Schedules", multiselect=True)
    if not sel:
        forms.alert("No schedules selected.", exitscript=True)
    label_to_id = {it["label"]: it["id"] for it in items}
    return [label_to_id[lbl] for lbl in sel]

def pick_rules_file():
    fp = forms.pick_file(file_ext="txt", title="Pick Filters Rules Text File")
    if not fp:
        forms.alert("No rules file selected.", exitscript=True)
    return fp

def get_fieldid_by_name(schedule, target_name):
    def norm(x): return (x or "").strip().lower()

    sd = schedule.Definition
    for i in range(sd.GetFieldCount()):
        f = sd.GetField(i)
        try:
            header = f.GetName()
        except:
            header = None
        if norm(header) == norm(target_name):
            return f.FieldId

    for i in range(sd.GetFieldCount()):
        f = sd.GetField(i)
        try:
            sfield = f.GetSchedulableField()
            sname = sfield.GetName() if sfield else None
            if norm(sname) == norm(target_name):
                return f.FieldId
        except:
            pass

    return None

# ---------- NEW: determine the field's true data type (if API supports it) ----------
def get_field_datatype(schedule, field_id):
    """
    Returns one of: 'string', 'integer', 'number', 'length', 'unknown'
    Uses SchedulableField.GetDataTypeId() when available (Revit 2021+).
    """
    try:
        sd = schedule.Definition
        for i in range(sd.GetFieldCount()):
            f = sd.GetField(i)
            if f.FieldId == field_id:
                sf = f.GetSchedulableField()
                if not (HAS_SPECTYPE and sf):
                    return 'unknown'
                dtid = sf.GetDataTypeId()
                # Compare against common spec types
                # String
                if dtid == SpecTypeId.String.Text:
                    return 'string'
                # Integer (e.g., number of poles, counts). There's no direct "int" spec;
                # most numeric non-unit params come back as Number.
                if dtid == SpecTypeId.Int.Integer:  # may not exist in all versions; guard:
                    return 'integer'
                # Length / Number
                if dtid == SpecTypeId.Length:
                    return 'length'
                if dtid == SpecTypeId.Number:
                    return 'number'
                # Fallback: provide string/number/length best-effort based on id string
                s = str(dtid) if dtid else ""
                s_low = s.lower()
                if "length" in s_low:
                    return 'length'
                if "string" in s_low or "text" in s_low:
                    return 'string'
                if "number" in s_low:
                    return 'number'
                return 'unknown'
    except Exception:
        return 'unknown'
    return 'unknown'

# ---------- NEW: generate candidate filters and try them in sensible order ----------
def build_filter_candidates(field_id, op_enum, raw_value, datatype_hint):
    """
    Returns a list of lambdas that, when called, create ScheduleFilter objects with different overloads.
    Order is chosen by datatype_hint.
    """
    def as_string():
        return ScheduleFilter(field_id, op_enum, raw_value)

    def as_double_feet():
        return ScheduleFilter(field_id, op_enum, float(parse_feet_inches(raw_value)))

    def as_double_plain():
        return ScheduleFilter(field_id, op_enum, float(raw_value))

    def as_int():
        return ScheduleFilter(field_id, op_enum, int(raw_value))

    looks_numeric = bool(re.search(r"[0-9]", raw_value)) and (("'" in raw_value) or ('"' in raw_value) or re.match(r"^\s*-?\d+(\.\d+)?\s*$", raw_value))

    # Base pools
    numeric_pool = [as_double_feet, as_double_plain, as_int, as_string]
    string_pool  = [as_string, as_int, as_double_feet, as_double_plain]

    if datatype_hint == 'string':
        return string_pool
    if datatype_hint in ('length', 'number', 'integer'):
        return numeric_pool

    # Unknown type: pick by heuristic
    return numeric_pool if looks_numeric else string_pool

def clear_filters(schedule):
    sd = schedule.Definition
    count = sd.GetFilterCount()
    for idx in reversed(range(count)):
        sd.RemoveFilter(idx)

# ------------------ Main ------------------
def main():
    all_schedules = list_view_schedules(doc)
    picked_ids = pick_schedules(all_schedules)
    rules_path = pick_rules_file()

    with open(rules_path, "r") as f:
        raw = f.read()

    rules = parse_rules(raw)
    logger.info("Parsed {} rule(s).".format(len(rules)))
    for r in rules:
        logger.info("  - [{}, {}, {}]".format(r["field"], r["op"], r["value"]))

    t = Transaction(doc, "Apply Schedule Filters (batch)")
    t.Start()
    try:
        for sid in picked_ids:
            vs = doc.GetElement(sid)
            sd = vs.Definition
            applied = 0

            if CLEAR_EXISTING:
                clear_filters(vs)

            for rule in rules:
                field_id = get_fieldid_by_name(vs, rule["field"])
                if not field_id:
                    logger.warning("[{}] → field not found in schedule '{}'; skipping."
                                   .format(rule["field"], vs.Name))
                    continue

                try:
                    op_enum = op_to_enum(rule["op"])
                except Exception:
                    logger.warning("[{}] → unsupported operator '{}'; skipping."
                                   .format(rule["field"], rule["op"]))
                    continue

                # NEW: detect datatype to choose best overload order,
                # then try candidates until AddFilter accepts one.
                dtype = get_field_datatype(vs, field_id)
                candidates = build_filter_candidates(field_id, op_enum, rule["value"], dtype)

                added = False
                last_err = None
                for maker in candidates:
                    try:
                        sf = maker()
                        sd.AddFilter(sf)  # validation happens here
                        added = True
                        break
                    except Exception as e:
                        last_err = e
                        continue

                if added:
                    applied += 1
                else:
                    logger.warning("[FilterSchedule] [{} {} {}] → failed: {}"
                                   .format(rule["field"], rule["op"], rule["value"], last_err))

            logger.info("Schedule '{}' → applied {} filter(s).".format(vs.Name, applied))

        t.Commit()
        forms.alert("Done. Check the Output panel for details.", title="Schedule Filters")
    except Exception as e:
        logger.error("Failed with exception:\n" + traceback.format_exc())
        try:
            t.RollBack()
        except:
            pass
        forms.alert("Error: {}".format(e), title="Schedule Filters – Error")

if __name__ == "__main__":
    main()
