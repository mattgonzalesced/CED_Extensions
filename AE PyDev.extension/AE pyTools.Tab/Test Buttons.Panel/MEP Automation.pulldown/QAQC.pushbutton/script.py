# -*- coding: utf-8 -*-
"""
QA/QC report for the active YAML data stored in Extensible Storage.

Summarises, per equipment definition, how many matching host elements were found
in the model (or its links) and how many Revit elements have been placed using
the automation (detected via the Element_Linker metadata written during
placement). Results are printed to the pyRevit output panel so we can use
Markdown for emphasis (bold = has placements, italics = none placed).
"""

from __future__ import print_function

import os
import sys
from collections import defaultdict

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FamilyInstance,
    FilteredElementCollector,
    Group,
    RevitLinkInstance,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "QA/QC Equipment Coverage"
LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _name_variants(elem):
    variants = set()
    try:
        raw = getattr(elem, "Name", None)
        if raw:
            variants.add(_normalize_name(raw))
    except Exception:
        pass
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            variants.add(_normalize_name(u"{} : {}".format(fam_name, type_name)))
        if fam_name:
            variants.add(_normalize_name(fam_name))
        if type_name:
            variants.add(_normalize_name(type_name))
    elif isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        group_name = getattr(gtype, "Name", None) if gtype else None
        if group_name:
            variants.add(_normalize_name(group_name))
    return {name for name in variants if name}


def _iter_host_candidates(doc):
    if doc is None:
        return
    collectors = (
        FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType(),
        FilteredElementCollector(doc).OfClass(Group).WhereElementIsNotElementType(),
    )
    for collector in collectors:
        for elem in collector:
            yield elem


def _collect_placeholder_counts(doc, target_map):
    """Count how many host elements exist whose names match equipment definitions."""
    counts = defaultdict(int)
    if not target_map:
        return counts

    def register(match_key):
        eq_names = target_map.get(match_key)
        if not eq_names:
            return
        for eq in eq_names:
            counts[eq] += 1

    for elem in _iter_host_candidates(doc):
        for variant in _name_variants(elem):
            if variant in target_map:
                register(variant)

    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        linked = FilteredElementCollector(link_doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
        for elem in linked:
            for variant in _name_variants(elem):
                if variant in target_map:
                    register(variant)
    return counts


def _extract_led_id(payload):
    if not payload:
        return ""
    for raw_line in payload.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, _, remainder = line.partition(":")
        if key.strip().lower() == "linked element definition id":
            return remainder.strip()
    return ""


def _get_linker_payload(elem):
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
            return value
    return ""


def _collect_placed_counts(doc, led_to_equipment):
    eq_counts = defaultdict(int)
    led_counts = defaultdict(int)
    if not led_to_equipment:
        return eq_counts, led_counts
    elems = _iter_host_candidates(doc)
    for elem in elems:
        payload = _get_linker_payload(elem)
        if not payload:
            continue
        led_id = _extract_led_id(payload)
        if not led_id:
            continue
        eq_name = led_to_equipment.get(led_id)
        if not eq_name:
            continue
        eq_counts[eq_name] += 1
        led_counts[led_id] += 1
    return eq_counts, led_counts


def _build_led_map(data):
    eq_mapping = {}
    led_metadata = {}
    order = []
    for eq in data.get("equipment_definitions") or []:
        if not isinstance(eq, dict):
            continue
        name = (eq.get("name") or eq.get("id") or "").strip()
        if not name:
            continue
        order.append(name)
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                led_id = (led.get("id") or "").strip()
                if not led_id:
                    continue
                if led_id not in eq_mapping:
                    eq_mapping[led_id] = name
                label = (led.get("label") or led_id).strip()
                led_metadata[led_id] = {
                    "equipment": name,
                    "label": label or led_id,
                }
    return eq_mapping, led_metadata, order


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    try:
        yaml_path, yaml_data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(yaml_path)

    equipment_names = sorted({
        (entry.get("name") or entry.get("id") or "").strip()
        for entry in (yaml_data.get("equipment_definitions") or [])
        if isinstance(entry, dict)
    })
    if not equipment_names:
        forms.alert("No equipment definitions found in {}.".format(yaml_label), title=TITLE)
        return

    led_map, led_metadata, eq_order = _build_led_map(yaml_data)
    if eq_order:
        seen = set()
        ordered_names = []
        for name in eq_order:
            if name in equipment_names and name not in seen:
                ordered_names.append(name)
                seen.add(name)
        remaining = [name for name in equipment_names if name not in seen]
        equipment_names = ordered_names + sorted(remaining)

    target_map = defaultdict(list)
    for name in equipment_names:
        normalized = _normalize_name(name)
        if normalized:
            target_map[normalized].append(name)
    host_counts = _collect_placeholder_counts(doc, target_map)
    placed_counts, led_type_counts = _collect_placed_counts(doc, led_map)

    output = script.get_output()
    output.print_md("### QA/QC Report – {}".format(yaml_label))
    total_found = 0
    total_placed = 0
    for name in equipment_names:
        found = host_counts.get(name, 0)
        placed = placed_counts.get(name, 0)
        total_found += found
        total_placed += placed
        label = "**{}**".format(name) if placed > 0 else "_{}_".format(name)
        delta = found - placed
        note = ""
        if delta > 0:
            note = " ({} awaiting placement)".format(delta)
        elif found == 0 and placed == 0:
            note = " (no hosts detected yet)"
        output.print_md(
            "{} — placed `{}` / hosts `{}`{}".format(label, placed, found, note)
        )
    output.print_md("")
    output.print_md(
        "**Totals:** placed `{}` elements across `{}` host matches."
        .format(total_placed, total_found)
    )

    type_rows = []
    label_totals = defaultdict(int)
    for led_id, count in led_type_counts.items():
        if count <= 0:
            continue
        meta = led_metadata.get(led_id) or {}
        eq_name = meta.get("equipment") or led_map.get(led_id) or "<Unknown>"
        label = (meta.get("label") or led_id).strip() or led_id
        type_rows.append((eq_name, label, count))
        label_totals[label] += count
    if type_rows:
        type_rows.sort(key=lambda row: (row[0], row[1]))
        output.print_md("")
        output.print_md("#### Placed Type Totals")
        for eq_name, label, count in type_rows:
            output.print_md("* {} – `{}` placed for '{}'".format(label, count, eq_name))
    if label_totals:
        output.print_md("")
        output.print_md("#### Grand Totals by Type")
        for label in sorted(label_totals.keys()):
            output.print_md("* {} – `{}` placed in project".format(label, label_totals[label]))

    forms.alert(
        "QA/QC summary sent to the pyRevit output panel for {}.\n"
        "Placed entries are bold; entries with zero placements are italic.".format(yaml_label),
        title=TITLE,
    )


if __name__ == "__main__":
    main()
