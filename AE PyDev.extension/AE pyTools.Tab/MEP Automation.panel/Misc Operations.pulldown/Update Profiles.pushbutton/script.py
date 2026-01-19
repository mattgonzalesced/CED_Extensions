# -*- coding: utf-8 -*-
"""
Update Profiles
---------------
Scans elements with Element_Linker metadata and merges hosted annotations
(tags/keynotes/text notes) plus parameter changes back into the active YAML.
"""

import importlib.util
import os
import re
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import FamilyInstance, FilteredElementCollector, Group

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import (  # noqa: E402
    load_active_yaml_data,
    save_active_yaml_data,
)

LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
TITLE = "Update Profiles"

_MANAGE_MODULE = None


def _manage_profiles_module():
    global _MANAGE_MODULE
    if _MANAGE_MODULE is not None:
        return _MANAGE_MODULE
    module_path = os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "..",
            "..",
            "Modify Profiles.stack",
            "Modify Profiles.pulldown",
            "Manage Profiles.pushbutton",
            "script.py",
        )
    )
    if not os.path.exists(module_path):
        forms.alert("Manage Profiles script not found at:\n{}".format(module_path), title=TITLE)
        return None
    spec = importlib.util.spec_from_file_location("ced_manage_profiles", module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    _MANAGE_MODULE = module
    return module


def _get_linker_text(elem):
    if elem is None:
        return ""
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        text = None
        try:
            text = param.AsString()
        except Exception:
            text = None
        if not text:
            try:
                text = param.AsValueString()
            except Exception:
                text = None
        if text and str(text).strip():
            return str(text)
    return ""


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
            r"(Linked Element Definition ID|Set Definition ID|Location XYZ \(ft\)|"
            r"Rotation \(deg\)|Parent Rotation \(deg\)|Parent ElementId|LevelId|"
            r"ElementId|FacingOrientation)\s*:\s*"
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(text) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")
    return {
        "led_id": entries.get("Linked Element Definition ID", "").strip(),
        "set_id": entries.get("Set Definition ID", "").strip(),
    }


def _normalize_key(value):
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().split())


def _tag_key(entry):
    if not isinstance(entry, dict):
        return None
    fam = entry.get("family_name") or entry.get("family") or ""
    typ = entry.get("type_name") or entry.get("type") or ""
    cat = entry.get("category_name") or entry.get("category") or ""
    return (_normalize_key(fam), _normalize_key(typ), _normalize_key(cat))


def _text_note_key(entry):
    if not isinstance(entry, dict):
        return None
    text_value = entry.get("text") or ""
    type_name = entry.get("type_name") or ""
    return (_normalize_key(text_value), _normalize_key(type_name))


def _update_tag_entry(existing, incoming):
    changed = False
    for key in ("family_name", "type_name", "category_name"):
        incoming_value = incoming.get(key)
        if incoming_value is not None and existing.get(key) != incoming_value:
            existing[key] = incoming_value
            changed = True
    if "offsets" in incoming and existing.get("offsets") != incoming.get("offsets"):
        existing["offsets"] = incoming.get("offsets") or {}
        changed = True
    if "parameters" in incoming and existing.get("parameters") != incoming.get("parameters"):
        existing["parameters"] = incoming.get("parameters") or {}
        changed = True
    return changed


def _update_text_note_entry(existing, incoming):
    changed = False
    for key in ("text", "type_name", "width_inches"):
        incoming_value = incoming.get(key)
        if incoming_value is not None and existing.get(key) != incoming_value:
            existing[key] = incoming_value
            changed = True
    if "offsets" in incoming and existing.get("offsets") != incoming.get("offsets"):
        existing["offsets"] = incoming.get("offsets") or {}
        changed = True
    if "leaders" in incoming and existing.get("leaders") != incoming.get("leaders"):
        existing["leaders"] = incoming.get("leaders") or []
        changed = True
    return changed


def _merge_entries(existing, incoming, key_func, update_func):
    changed = False
    added = 0
    updated = 0
    existing_map = {}
    for idx, entry in enumerate(existing):
        key = key_func(entry)
        if key is None:
            continue
        existing_map.setdefault(key, []).append(idx)
    used_indices = set()
    for entry in incoming:
        key = key_func(entry)
        if key is None:
            continue
        match_idx = None
        for idx in existing_map.get(key, []):
            if idx not in used_indices:
                match_idx = idx
                used_indices.add(idx)
                break
        if match_idx is None:
            existing.append(entry)
            added += 1
            changed = True
            continue
        if update_func(existing[match_idx], entry):
            updated += 1
            changed = True
    return changed, added, updated


def _merge_tag_entries(led, tags):
    existing = led.get("tags")
    if not isinstance(existing, list):
        existing = []
        led["tags"] = existing
    return _merge_entries(existing, tags, _tag_key, _update_tag_entry)


def _merge_text_note_entries(led, notes):
    existing = led.get("text_notes")
    if not isinstance(existing, list):
        existing = []
        led["text_notes"] = existing
    return _merge_entries(existing, notes, _text_note_key, _update_text_note_entry)


def _merge_params(led, led_id, params, stats, param_cache):
    if not params:
        return False
    led_params = led.get("parameters")
    if not isinstance(led_params, dict):
        led_params = {}
        led["parameters"] = led_params
    changed = False
    cache = param_cache.setdefault(led_id, {})
    for key, value in params.items():
        if key in cache and cache[key] != value:
            stats["param_conflicts"] += 1
        cache[key] = value
        if led_params.get(key) != value:
            led_params[key] = value
            stats["params_updated"] += 1
            changed = True
    return changed


def _index_leds(data):
    index = {}
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if led.get("is_parent_anchor"):
                    continue
                led_id = led.get("id")
                if led_id:
                    index[led_id] = led
    return index


def _collect_candidate_elements(doc):
    elements = []
    seen = set()
    for cls in (FamilyInstance, Group):
        try:
            collector = FilteredElementCollector(doc).OfClass(cls).WhereElementIsNotElementType()
            for elem in collector:
                try:
                    elem_id = elem.Id.IntegerValue
                except Exception:
                    elem_id = None
                if elem_id is None or elem_id in seen:
                    continue
                seen.add(elem_id)
                elements.append(elem)
        except Exception:
            continue
    return elements


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    manage = _manage_profiles_module()
    if manage is None:
        return

    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)

    led_index = _index_leds(data)
    if not led_index:
        forms.alert("No linked element definitions found in {}.".format(yaml_label), title=TITLE)
        return

    elements = _collect_candidate_elements(doc)
    if not elements:
        forms.alert("No candidate elements found in the current model.", title=TITLE)
        return

    stats = {
        "elements_scanned": 0,
        "elements_with_linker": 0,
        "definitions_updated": set(),
        "definitions_found": set(),
        "params_updated": 0,
        "param_conflicts": 0,
        "tags_added": 0,
        "tags_updated": 0,
        "notes_added": 0,
        "notes_updated": 0,
        "missing_defs": set(),
        "missing_led": 0,
    }

    param_cache = {}

    for elem in elements:
        stats["elements_scanned"] += 1
        linker_text = _get_linker_text(elem)
        if not linker_text:
            continue
        stats["elements_with_linker"] += 1
        payload = _parse_linker_payload(linker_text)
        led_id = payload.get("led_id")
        if not led_id:
            stats["missing_led"] += 1
            continue
        led = led_index.get(led_id)
        if led is None:
            stats["missing_defs"].add(led_id)
            continue
        stats["definitions_found"].add(led_id)

        params = manage._collect_params(elem)
        if _merge_params(led, led_id, params, stats, param_cache):
            stats["definitions_updated"].add(led_id)

        host_point = manage._get_point(elem)
        tags, text_notes = manage._collect_hosted_tags(elem, host_point)
        if tags:
            changed, added, updated = _merge_tag_entries(led, tags)
            stats["tags_added"] += added
            stats["tags_updated"] += updated
            if changed:
                stats["definitions_updated"].add(led_id)
        if text_notes:
            changed, added, updated = _merge_text_note_entries(led, text_notes)
            stats["notes_added"] += added
            stats["notes_updated"] += updated
            if changed:
                stats["definitions_updated"].add(led_id)

    if not stats["definitions_updated"]:
        summary = [
            "No profile updates detected in {}.".format(yaml_label),
            "",
            "Elements scanned: {}".format(stats["elements_scanned"]),
            "Elements with Element_Linker: {}".format(stats["elements_with_linker"]),
        ]
        if stats["missing_defs"]:
            summary.append("")
            summary.append("Missing definitions for LED IDs:")
            sample = sorted(stats["missing_defs"])[:6]
            summary.extend(" - {}".format(item) for item in sample)
            if len(stats["missing_defs"]) > len(sample):
                summary.append("   (+{} more)".format(len(stats["missing_defs"]) - len(sample)))
        forms.alert("\n".join(summary), title=TITLE)
        return

    try:
        save_active_yaml_data(doc, data, "Update Profiles", "Merged model changes into profiles")
    except Exception as exc:
        forms.alert("Failed to save updated profiles:\n\n{}".format(exc), title=TITLE)
        return

    summary = [
        "Updated profiles in {}.".format(yaml_label),
        "",
        "Elements scanned: {} (with Element_Linker: {})".format(
            stats["elements_scanned"],
            stats["elements_with_linker"],
        ),
        "Definitions updated: {}".format(len(stats["definitions_updated"])),
        "Parameter values updated: {}".format(stats["params_updated"]),
        "Tags added/updated: {} / {}".format(stats["tags_added"], stats["tags_updated"]),
        "Text notes added/updated: {} / {}".format(stats["notes_added"], stats["notes_updated"]),
    ]
    if stats["param_conflicts"]:
        summary.append("Parameter conflicts detected: {}".format(stats["param_conflicts"]))
    if stats["missing_defs"]:
        summary.append("")
        summary.append("Missing definitions for LED IDs:")
        sample = sorted(stats["missing_defs"])[:6]
        summary.extend(" - {}".format(item) for item in sample)
        if len(stats["missing_defs"]) > len(sample):
            summary.append("   (+{} more)".format(len(stats["missing_defs"]) - len(sample)))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
