# -*- coding: utf-8 -*-
"""
Write Element Linker
--------------------
Writes Element_Linker metadata to fixtures and model groups that are missing it,
using the active YAML profiles to match by label and parameter scoring.
"""

import math
import os
import re
import sys

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInParameter,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    GroupType,
    Transaction,
    XYZ,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402

try:
    basestring
except NameError:
    basestring = str

TITLE = "Write Element Linker"
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")


def _normalize_label(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _is_blank(value):
    if value is None:
        return True
    if isinstance(value, basestring):
        return not value.strip()
    return False


def _coerce_number(value):
    try:
        return float(str(value).strip())
    except Exception:
        return None


def _values_match(left, right):
    left_num = _coerce_number(left)
    right_num = _coerce_number(right)
    if left_num is not None and right_num is not None:
        return abs(left_num - right_num) <= 1e-6
    return _normalize_label(left) == _normalize_label(right)


def _parse_linker_payload(payload_text):
    if not payload_text:
        return {}
    text = str(payload_text)
    entries = {}
    if "\n" in text:
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, remainder = line.partition(":")
            entries[key.strip()] = remainder.strip()
    else:
        pattern = re.compile(
            r"(Linked Element Definition ID|Set Definition ID|Parent ElementId)\s*:\s*"
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")

    def _as_int(value):
        try:
            return int(value)
        except Exception:
            return None

    return {
        "led_id": (entries.get("Linked Element Definition ID", "") or "").strip(),
        "set_id": (entries.get("Set Definition ID", "") or "").strip(),
        "parent_element_id": _as_int(entries.get("Parent ElementId", "")),
    }


def _get_point(elem):
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    if hasattr(loc, "Point") and loc.Point:
        return loc.Point
    if hasattr(loc, "Curve") and loc.Curve:
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            return None
    return None


def _get_rotation_degrees(elem):
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            pass
    facing = getattr(elem, "FacingOrientation", None)
    if facing:
        try:
            return math.degrees(math.atan2(facing.Y, facing.X))
        except Exception:
            pass
    return 0.0


def _get_level_element_id(elem):
    try:
        lvl = getattr(elem, "LevelId", None)
        if lvl and getattr(lvl, "IntegerValue", -1) > 0:
            return lvl.IntegerValue
    except Exception:
        pass
    level_params = (
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_LEVEL_PARAM,
    )
    for bip in level_params:
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if not param:
            continue
        try:
            eid = param.AsElementId()
            if eid and getattr(eid, "IntegerValue", -1) > 0:
                return eid.IntegerValue
        except Exception:
            continue
    return None


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)


def _build_element_linker_payload(led_id, set_id, elem, host_point, parent_element_id=None):
    point = host_point or _get_point(elem)
    if point is None:
        return ""
    rotation_deg = _get_rotation_degrees(elem)
    level_id = _get_level_element_id(elem)
    try:
        elem_id = elem.Id.IntegerValue
    except Exception:
        elem_id = ""
    facing = getattr(elem, "FacingOrientation", None)
    lines = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Location XYZ (ft): {}".format(_format_xyz(point)),
        "Rotation (deg): {:.6f}".format(rotation_deg),
        "Parent ElementId: {}".format(parent_element_id if parent_element_id is not None else ""),
        "LevelId: {}".format(level_id if level_id is not None else ""),
        "ElementId: {}".format(elem_id),
        "FacingOrientation: {}".format(_format_xyz(facing)),
    ]
    return "\n".join(lines).strip()


def _set_element_linker_parameter(elem, value):
    if not elem or value is None:
        return False
    for name in ELEMENT_LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(value)
            return True
        except Exception:
            continue
    return False


def _get_existing_linker_text(elem):
    if not elem:
        return ""
    for name in ELEMENT_LINKER_PARAM_NAMES:
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
        if not value:
            try:
                value = param.AsValueString()
            except Exception:
                value = None
        if value and str(value).strip():
            return str(value)
    return ""


def _element_label(elem):
    if isinstance(elem, (Group, GroupType)):
        name = getattr(elem, "Name", None) or ""
        return name.strip()
    fam_name = None
    type_name = None
    try:
        sym = getattr(elem, "Symbol", None) or getattr(elem, "GroupType", None)
        if sym:
            fam = getattr(sym, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else None
            type_name = getattr(sym, "Name", None)
            if not type_name and hasattr(sym, "get_Parameter"):
                tparam = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if tparam:
                    type_name = tparam.AsString()
    except Exception:
        pass
    if fam_name and type_name:
        return u"{} : {}".format(fam_name, type_name)
    return (type_name or fam_name or "").strip()


def _element_category(elem):
    try:
        cat = getattr(elem, "Category", None)
        return getattr(cat, "Name", None) if cat else None
    except Exception:
        return None


def _feet_to_inches(value):
    try:
        return float(value) * 12.0
    except Exception:
        return 0.0


def _collect_params(elem):
    def _convert_collected_double(target_key, param_obj, raw_value):
        if target_key not in ("Apparent Load Input_CED", "Voltage_CED"):
            return raw_value
        get_unit = getattr(param_obj, "GetUnitTypeId", None)
        if not callable(get_unit):
            return raw_value
        try:
            unit_id = get_unit()
        except Exception:
            unit_id = None
        if not unit_id:
            return raw_value
        try:
            from Autodesk.Revit.DB import UnitUtils
            return UnitUtils.ConvertFromInternalUnits(raw_value, unit_id)
        except Exception:
            return raw_value

    try:
        cat = getattr(elem, "Category", None)
        cat_name = getattr(cat, "Name", "") if cat else ""
    except Exception:
        cat_name = ""
    cat_l = (cat_name or "").lower()
    is_electrical = ("electrical" in cat_l) or ("lighting" in cat_l) or ("data" in cat_l)

    base_targets = {
        "dev-Group ID": ["dev-Group ID", "dev_Group ID"],
        "Number of Poles_CED": ["Number of Poles_CED", "Number of Poles_CEDT"],
        "Apparent Load Input_CED": ["Apparent Load Input_CED", "Apparent Load Input_CEDT"],
        "Voltage_CED": ["Voltage_CED", "Voltage_CEDT"],
    }
    electrical_targets = {
        "CKT_Rating_CED": ["CKT_Rating_CED"],
        "CKT_Panel_CEDT": ["CKT_Panel_CED", "CKT_Panel_CEDT"],
        "CKT_Schedule Notes_CEDT": ["CKT_Schedule Notes_CED", "CKT_Schedule Notes_CEDT"],
        "CKT_Circuit Number_CEDT": ["CKT_Circuit Number_CED", "CKT_Circuit Number_CEDT"],
        "CKT_Load Name_CEDT": ["CKT_Load Name_CED", "CKT_Load Name_CEDT"],
    }

    targets = dict(base_targets)
    if is_electrical:
        targets.update(electrical_targets)

    found = {k: "" for k in targets.keys()}
    for param in getattr(elem, "Parameters", []) or []:
        try:
            name = param.Definition.Name
        except Exception:
            continue
        target_key = None
        for out_key, aliases in targets.items():
            if name in aliases:
                target_key = out_key
                break
        if not target_key:
            continue
        try:
            st = param.StorageType.ToString()
        except Exception:
            st = ""
        try:
            if st == "String":
                found[target_key] = param.AsString() or ""
            elif st == "Double":
                found[target_key] = _convert_collected_double(target_key, param, param.AsDouble())
            elif st == "Integer":
                found[target_key] = param.AsInteger()
            else:
                found[target_key] = param.AsValueString() or ""
        except Exception:
            continue

    if not found.get("Voltage_CED"):
        found["Voltage_CED"] = 120
    return found


def _iter_led_entries(data):
    for eq in data.get("equipment_definitions") or []:
        if not isinstance(eq, dict):
            continue
        linked_sets = eq.get("linked_sets") or []
        if not linked_sets and isinstance(eq.get("linked_element_definitions"), list):
            linked_sets = [{
                "id": eq.get("id"),
                "linked_element_definitions": eq.get("linked_element_definitions"),
            }]
        for linked_set in linked_sets or []:
            set_id = linked_set.get("id") or eq.get("id") or ""
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                if led.get("is_parent_anchor"):
                    continue
                label = (led.get("label") or led.get("id") or "").strip()
                if not label:
                    continue
                params = led.get("parameters") or {}
                if not isinstance(params, dict):
                    params = {}
                payload = None
                for key in ELEMENT_LINKER_PARAM_NAMES:
                    value = params.get(key)
                    if value not in (None, ""):
                        payload = value
                        break
                parsed = _parse_linker_payload(payload) if payload else {}
                yield {
                    "label": label,
                    "label_key": _normalize_label(label),
                    "set_id": set_id,
                    "led_id": led.get("id") or "",
                    "category": led.get("category") or led.get("category_name") or "",
                    "is_group": bool(led.get("is_group")),
                    "parameters": params,
                    "parent_element_id": parsed.get("parent_element_id"),
                }


def _build_led_map(data):
    led_map = {}
    for entry in _iter_led_entries(data):
        key = entry["label_key"]
        if not key:
            continue
        led_map.setdefault(key, []).append(entry)
    return led_map


def _has_scored_params(params):
    for key, value in (params or {}).items():
        if key in ELEMENT_LINKER_PARAM_NAMES:
            continue
        if not _is_blank(value):
            return True
    return False


def _score_candidate(candidate, elem_params):
    score = 0
    for key, value in (candidate.get("parameters") or {}).items():
        if key in ELEMENT_LINKER_PARAM_NAMES:
            continue
        if _is_blank(value):
            continue
        if key not in elem_params:
            continue
        if _values_match(value, elem_params.get(key)):
            score += 1
    return score


def _filter_candidates(candidates, is_group, category_name):
    if not candidates:
        return []
    filtered = []
    for entry in candidates:
        if bool(entry.get("is_group")) != bool(is_group):
            continue
        led_cat = (entry.get("category") or "").strip().lower()
        if led_cat and category_name:
            if led_cat != category_name.strip().lower():
                continue
        filtered.append(entry)
    if filtered:
        return filtered
    return [entry for entry in candidates if bool(entry.get("is_group")) == bool(is_group)]


def _filter_candidates_by_parent(candidates, parent_element_id):
    if parent_element_id is None:
        return list(candidates or [])
    return [entry for entry in candidates if entry.get("parent_element_id") == parent_element_id]


def _select_best_candidate(candidates, elem_params):
    if not candidates:
        return None, 0, "no_match"
    scored = []
    for entry in candidates:
        scored.append((entry, _score_candidate(entry, elem_params)))
    if len(scored) == 1:
        entry, score = scored[0]
        if score > 0 or not _has_scored_params(entry.get("parameters")):
            return entry, score, "unique"
        return None, score, "no_match"
    max_score = max(score for _, score in scored)
    if max_score <= 0:
        return None, max_score, "no_match"
    best = [entry for entry, score in scored if score == max_score]
    if len(best) > 1:
        return None, max_score, "ambiguous"
    return best[0], max_score, "matched"


def _collect_target_elements(doc):
    selection = revit.get_selection()
    selected = list(getattr(selection, "elements", []) or []) if selection else []
    if selected:
        elements = [elem for elem in selected if isinstance(elem, (FamilyInstance, Group))]
        return elements, True

    elements = []
    for cls in (FamilyInstance, Group):
        try:
            collector = FilteredElementCollector(doc).OfClass(cls).WhereElementIsNotElementType()
            elements.extend(list(collector))
        except Exception:
            continue
    return elements, False


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)

    led_map = _build_led_map(data)
    if not led_map:
        forms.alert("No linked element definitions found in {}.".format(yaml_label), title=TITLE)
        return

    elements, from_selection = _collect_target_elements(doc)
    if not elements:
        forms.alert("No fixture or model group elements found.", title=TITLE)
        return

    stats = {
        "scanned": 0,
        "already_linked": 0,
        "missing_linker": 0,
        "matched": 0,
        "written": 0,
        "no_label": 0,
        "no_profile": 0,
        "no_match": 0,
        "ambiguous": 0,
        "parent_filtered": 0,
        "parent_mismatch": 0,
        "no_location": 0,
        "write_failed": 0,
    }
    ambiguous_labels = []
    unmatched_labels = []
    to_update = []

    for elem in elements:
        stats["scanned"] += 1
        existing_payload = _get_existing_linker_text(elem)
        if existing_payload:
            stats["already_linked"] += 1
        else:
            stats["missing_linker"] += 1
        label = _element_label(elem)
        if not label:
            stats["no_label"] += 1
            continue
        label_key = _normalize_label(label)
        candidates = led_map.get(label_key, [])
        if not candidates:
            stats["no_profile"] += 1
            unmatched_labels.append(label)
            continue
        category_name = _element_category(elem)
        candidates = _filter_candidates(candidates, isinstance(elem, Group), category_name)
        if not candidates:
            stats["no_profile"] += 1
            unmatched_labels.append(label)
            continue
        parent_id = None
        if existing_payload:
            parent_id = _parse_linker_payload(existing_payload).get("parent_element_id")
        if parent_id is not None:
            stats["parent_filtered"] += 1
            candidates = _filter_candidates_by_parent(candidates, parent_id)
            if not candidates:
                stats["parent_mismatch"] += 1
                unmatched_labels.append(label)
                continue
        elem_params = _collect_params(elem)
        best, score, reason = _select_best_candidate(candidates, elem_params)
        if best is None:
            if reason == "ambiguous":
                stats["ambiguous"] += 1
                ambiguous_labels.append(label)
            else:
                stats["no_match"] += 1
                unmatched_labels.append(label)
            continue
        host_point = _get_point(elem)
        payload = _build_element_linker_payload(
            best.get("led_id"),
            best.get("set_id"),
            elem,
            host_point,
            parent_element_id=best.get("parent_element_id"),
        )
        if not payload:
            stats["no_location"] += 1
            continue
        to_update.append((elem, payload))
        stats["matched"] += 1

    if not to_update:
        summary = [
            "No Element_Linker payloads were written to {}.".format(yaml_label),
            "",
            "Elements scanned: {}".format(stats["scanned"]),
            "Missing Element_Linker: {}".format(stats["missing_linker"]),
            "Already had Element_Linker: {}".format(stats["already_linked"]),
            "Matched profiles: {}".format(stats["matched"]),
            "Skipped (no profile): {}".format(stats["no_profile"]),
            "Skipped (ambiguous match): {}".format(stats["ambiguous"]),
            "Skipped (no param match): {}".format(stats["no_match"]),
            "Skipped (parent mismatch): {}".format(stats["parent_mismatch"]),
        ]
        forms.alert("\n".join(summary), title=TITLE)
        return

    txn = Transaction(doc, "Write Element_Linker metadata")
    try:
        txn.Start()
        for elem, payload in to_update:
            if _set_element_linker_parameter(elem, payload):
                stats["written"] += 1
            else:
                stats["write_failed"] += 1
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass
        forms.alert("Failed to write Element_Linker metadata.", title=TITLE)
        return

    summary = [
        "Write Element Linker completed for {}.".format(yaml_label),
        "",
        "Elements scanned: {}".format(stats["scanned"]),
        "Missing Element_Linker: {}".format(stats["missing_linker"]),
        "Already had Element_Linker: {}".format(stats["already_linked"]),
        "Matched profiles: {}".format(stats["matched"]),
        "Element_Linker written: {}".format(stats["written"]),
        "Skipped (no profile): {}".format(stats["no_profile"]),
        "Skipped (ambiguous match): {}".format(stats["ambiguous"]),
        "Skipped (no param match): {}".format(stats["no_match"]),
        "Skipped (parent mismatch): {}".format(stats["parent_mismatch"]),
    ]
    if stats["no_location"]:
        summary.append("Skipped (no location): {}".format(stats["no_location"]))
    if stats["write_failed"]:
        summary.append("Failed writes: {}".format(stats["write_failed"]))
    if from_selection:
        summary.append("")
        summary.append("Scope: current selection")
    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
