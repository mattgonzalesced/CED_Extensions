# -*- coding: utf-8 -*-
"""
Update Profiles
---------------
Scans elements with Element_Linker metadata and merges hosted annotations
(tags/keynotes/text notes) plus parameter changes back into the active YAML.
Also captures nearby keynotes and text notes within a 5 ft radius.
"""

import imp
import math
import os
import re
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import FamilyInstance, FilteredElementCollector, Group, TextNote

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
PROXIMITY_RADIUS_FT = 5.0
PROXIMITY_CELL_SIZE_FT = 5.0
NOTE_KEY_PRECISION = 3
TIE_DISTANCE_TOLERANCE_FT = 1e-4

_MANAGE_MODULE = None
_UI_MODULE = None


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
    try:
        module = imp.load_source("ced_manage_profiles", module_path)
    except Exception as exc:
        forms.alert("Failed to load Manage Profiles script:\n{}\n\n{}".format(module_path, exc), title=TITLE)
        return None
    _MANAGE_MODULE = module
    return module


def _update_profiles_ui():
    global _UI_MODULE
    if _UI_MODULE is not None:
        return _UI_MODULE
    module_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "UpdateProfilesUI.py"))
    if not os.path.exists(module_path):
        forms.alert("Update Profiles UI not found at:\n{}".format(module_path), title=TITLE)
        return None
    try:
        module = imp.load_source("ced_update_profiles_ui", module_path)
    except Exception as exc:
        forms.alert("Failed to load Update Profiles UI:\n{}\n\n{}".format(module_path, exc), title=TITLE)
        return None
    _UI_MODULE = module
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
            r"ElementId|FacingOrientation|CKT_Circuit Number_CEDT|CKT_Panel_CEDT)\s*:\s*"
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")
    return {
        "led_id": entries.get("Linked Element Definition ID", "").strip(),
        "set_id": entries.get("Set Definition ID", "").strip(),
        "CKT_Circuit Number_CEDT": entries.get("CKT_Circuit Number_CEDT", "").strip(),
        "CKT_Panel_CEDT": entries.get("CKT_Panel_CEDT", "").strip(),
    }


def _apply_linker_params(params, payload):
    if not isinstance(params, dict):
        params = {}
    updated = dict(params)
    for key in ("CKT_Circuit Number_CEDT", "CKT_Panel_CEDT"):
        updated.pop(key, None)
        value = payload.get(key) if isinstance(payload, dict) else None
        if value not in (None, ""):
            updated[key] = value
    return updated


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


def _replace_entries(current, incoming):
    existing = list(current) if isinstance(current, list) else []
    new_entries = list(incoming or [])
    if existing == new_entries:
        return False, 0, 0, new_entries
    added = max(len(new_entries) - len(existing), 0)
    updated = len(new_entries) if len(new_entries) == len(existing) else min(len(existing), len(new_entries))
    return True, added, updated, new_entries


def _normalize_keynote_family(value):
    if not value:
        return ""
    text = str(value)
    if ":" in text:
        text = text.split(":", 1)[0]
    return "".join([ch for ch in text.lower() if ch.isalnum()])


def _is_ga_keynote_symbol(family_name):
    return _normalize_keynote_family(family_name) == "gakeynotesymbolced"


def _is_builtin_keynote_tag(tag_entry):
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
        category = tag_entry.get("category_name") or tag_entry.get("category") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
        category = getattr(tag_entry, "category_name", None) or getattr(tag_entry, "category", None) or ""
    if _is_ga_keynote_symbol(family):
        return False
    fam_text = (family or "").lower()
    cat_text = (category or "").lower()
    if "keynote tags" in cat_text:
        return True
    if "keynote tag" in fam_text:
        return True
    return False


def _is_keynote_entry(tag_entry):
    if not tag_entry:
        return False
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
    return _is_ga_keynote_symbol(family)


def _split_keynote_entries(entries):
    normal = []
    keynotes = []
    for entry in entries or []:
        if _is_builtin_keynote_tag(entry):
            continue
        if _is_keynote_entry(entry):
            keynotes.append(entry)
        else:
            normal.append(entry)
    return normal, keynotes


def _partition_tag_records(records):
    tag_records = []
    keynote_records = []
    for record in records or []:
        if not isinstance(record, dict):
            continue
        entry = record.get("entry")
        if not entry:
            continue
        if _is_builtin_keynote_tag(entry):
            continue
        if _is_keynote_entry(entry):
            keynote_records.append(record)
        else:
            tag_records.append(record)
    return tag_records, keynote_records


def _entries_from_records(records, instances, mode):
    entries = []
    for record in records or []:
        if not isinstance(record, dict):
            continue
        if mode == "common" and record.get("count", 0) != instances:
            continue
        entry = record.get("entry")
        if entry:
            entries.append(entry)
    return entries


def _merge_tag_and_keynote_entries(led, tags, keynotes):
    existing = led.get("tags")
    if not isinstance(existing, list):
        existing = []
        led["tags"] = existing
    changed, added, updated = _merge_entries(existing, tags, _tag_key, _update_tag_entry)
    existing_keynotes = led.get("keynotes")
    if not isinstance(existing_keynotes, list):
        existing_keynotes = []
        led["keynotes"] = existing_keynotes
    changed_kn, added_kn, updated_kn = _merge_entries(existing_keynotes, keynotes, _tag_key, _update_tag_entry)
    return (changed or changed_kn), (added + added_kn), (updated + updated_kn)


def _set_tag_and_keynote_entries(led, tags, keynotes):
    changed, added, updated, new_entries = _replace_entries(led.get("tags"), tags)
    led["tags"] = new_entries
    changed_kn, added_kn, updated_kn, new_keynotes = _replace_entries(led.get("keynotes"), keynotes)
    led["keynotes"] = new_keynotes
    return (changed or changed_kn), (added + added_kn), (updated + updated_kn)


def _merge_tag_entries(led, tags):
    normal_tags, keynotes = _split_keynote_entries(tags)
    existing = led.get("tags")
    if not isinstance(existing, list):
        existing = []
        led["tags"] = existing
    changed, added, updated = _merge_entries(existing, normal_tags, _tag_key, _update_tag_entry)
    existing_keynotes = led.get("keynotes")
    if not isinstance(existing_keynotes, list):
        existing_keynotes = []
        led["keynotes"] = existing_keynotes
    changed_kn, added_kn, updated_kn = _merge_entries(existing_keynotes, keynotes, _tag_key, _update_tag_entry)
    return (changed or changed_kn), (added + added_kn), (updated + updated_kn)


def _merge_text_note_entries(led, notes):
    existing = led.get("text_notes")
    if not isinstance(existing, list):
        existing = []
        led["text_notes"] = existing
    return _merge_entries(existing, notes, _text_note_key, _update_text_note_entry)


def _set_tag_entries(led, tags):
    normal_tags, keynotes = _split_keynote_entries(tags)
    changed, added, updated, new_entries = _replace_entries(led.get("tags"), normal_tags)
    led["tags"] = new_entries
    changed_kn, added_kn, updated_kn, new_keynotes = _replace_entries(led.get("keynotes"), keynotes)
    led["keynotes"] = new_keynotes
    return (changed or changed_kn), (added + added_kn), (updated + updated_kn)


def _set_text_note_entries(led, notes):
    changed, added, updated, new_entries = _replace_entries(led.get("text_notes"), notes)
    led["text_notes"] = new_entries
    return changed, added, updated


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


def _track_param_value(obs, key, value):
    bucket = obs["params"].setdefault(key, {})
    token = repr(value)
    entry = bucket.get(token)
    if not entry:
        entry = {"value": value, "count": 0}
        bucket[token] = entry
    entry["count"] += 1


def _track_annotation(obs, store_key, entry, key_func):
    key = key_func(entry)
    if key is None:
        return
    bucket = obs[store_key]
    record = bucket.get(key)
    if not record:
        record = {"entry": entry, "count": 0}
        bucket[key] = record
    record["count"] += 1


def _shorten(text, limit=60):
    text = (text or "").replace("\r", " ").replace("\n", " ").strip()
    if not text:
        return "<none>"
    if limit and len(text) > limit:
        return text[: limit - 3] + "..."
    return text


def _format_tag_entry(entry):
    if not isinstance(entry, dict):
        return "<tag>"
    fam = entry.get("family_name") or entry.get("family") or "<Family?>"
    typ = entry.get("type_name") or entry.get("type") or "<Type?>"
    cat = entry.get("category_name") or entry.get("category") or ""
    label = u"{} : {}".format(fam, typ)
    if cat:
        label = u"{} ({})".format(label, cat)
    return label


def _format_note_entry(entry):
    if not isinstance(entry, dict):
        return "<text note>"
    text_value = _shorten(entry.get("text") or "")
    type_name = entry.get("type_name") or "<Type?>"
    return u"\"{}\" ({})".format(text_value, type_name)


def _index_leds(data):
    index = {}
    meta = {}
    for eq in data.get("equipment_definitions") or []:
        if not isinstance(eq, dict):
            continue
        profile_name = eq.get("name") or eq.get("id") or "<profile>"
        for linked_set in eq.get("linked_sets") or []:
            if not isinstance(linked_set, dict):
                continue
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict) or led.get("is_parent_anchor"):
                    continue
                led_id = led.get("id")
                if not led_id:
                    continue
                index[led_id] = led
                meta_entry = meta.setdefault(led_id, {})
                meta_entry.setdefault("profile_name", profile_name)
                meta_entry.setdefault("type_label", led.get("label") or led_id)
    return index, meta


def _collect_candidate_elements(doc, elements=None):
    if elements is not None:
        filtered = []
        seen = set()
        for elem in elements:
            if not isinstance(elem, (FamilyInstance, Group)):
                continue
            try:
                elem_id = elem.Id.IntegerValue
            except Exception:
                elem_id = None
            if elem_id is None or elem_id in seen:
                continue
            seen.add(elem_id)
            filtered.append(elem)
        return filtered

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


def _note_location_key(point, kind):
    if point is None:
        return None
    try:
        return (
            kind,
            round(float(point.X), NOTE_KEY_PRECISION),
            round(float(point.Y), NOTE_KEY_PRECISION),
            round(float(point.Z), NOTE_KEY_PRECISION),
        )
    except Exception:
        return None


def _entry_world_point(host_point, entry):
    if host_point is None or not isinstance(entry, dict):
        return None
    offsets = entry.get("offsets") or {}
    try:
        x = float(offsets.get("x_inches", 0.0) or 0.0) / 12.0
        y = float(offsets.get("y_inches", 0.0) or 0.0) / 12.0
        z = float(offsets.get("z_inches", 0.0) or 0.0) / 12.0
        return host_point.__class__(host_point.X + x, host_point.Y + y, host_point.Z + z)
    except Exception:
        return None


def _host_cell_key(point, cell_size):
    if point is None or not cell_size:
        return None
    try:
        return (
            int(math.floor(point.X / cell_size)),
            int(math.floor(point.Y / cell_size)),
        )
    except Exception:
        return None


def _build_host_spatial_index(host_records, cell_size):
    index = {}
    for record in host_records or []:
        host_point = record.get("host_point")
        cell = _host_cell_key(host_point, cell_size)
        if cell is None:
            continue
        index.setdefault(cell, []).append(record)
    return index


def _distance_sq(a, b):
    if a is None or b is None:
        return None
    try:
        dx = a.X - b.X
        dy = a.Y - b.Y
        return (dx * dx) + (dy * dy)
    except Exception:
        return None


def _candidate_hosts(note_point, index, cell_size):
    cell = _host_cell_key(note_point, cell_size)
    if cell is None:
        return []
    cx, cy = cell
    hosts = []
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            hosts.extend(index.get((cx + dx, cy + dy), []))
    return hosts


def _select_host_for_note(note_elem, candidates, manage, kind):
    if not candidates:
        return None
    preview = ""
    if kind == "text_note":
        try:
            preview = manage._text_note_preview(note_elem)
        except Exception:
            preview = ""
    else:
        try:
            params = manage._collect_keynote_parameters(note_elem)
            preview = params.get("Keynote Value") or params.get("Key Value") or ""
        except Exception:
            preview = ""
        if not preview:
            try:
                fam, typ = manage._annotation_family_type(note_elem)
                if fam and typ:
                    preview = u"{} : {}".format(fam, typ)
                else:
                    preview = fam or typ or ""
            except Exception:
                preview = ""
    title = "Select host for note"
    if preview:
        title = "Select host for note: {}".format(preview)

    option_map = {}
    options = []
    for idx, record in enumerate(candidates):
        profile_name = record.get("profile_name") or record.get("led_id") or "Profile"
        type_label = record.get("type_label") or record.get("led_id") or "Type"
        elem = record.get("element")
        elem_id = getattr(elem, "Id", None)
        elem_id_val = getattr(elem_id, "IntegerValue", None) if elem_id else None
        label = u"{} / {}".format(profile_name, type_label)
        if elem_id_val is not None:
            label = u"{} (Id: {})".format(label, elem_id_val)
        option = u"{:02d}. {}".format(idx + 1, label)
        option_map[option] = record
        options.append(option)
    try:
        selection = forms.SelectFromList.show(
            options,
            title=title,
            button_name="Assign",
            multiselect=False,
        )
    except Exception:
        selection = None
    if selection:
        chosen = selection[0] if isinstance(selection, list) else selection
        return option_map.get(chosen)
    return None


def _assign_proximity_notes(doc, host_records, assigned_keys, observations, led_index, led_meta, manage):
    if doc is None or not host_records:
        return
    index = _build_host_spatial_index(host_records, PROXIMITY_CELL_SIZE_FT)
    radius_sq = PROXIMITY_RADIUS_FT * PROXIMITY_RADIUS_FT

    try:
        keynote_candidates = list(
            FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
        )
    except Exception:
        keynote_candidates = []
    keynotes = [elem for elem in keynote_candidates if manage._is_ga_keynote_symbol_element(elem)]

    try:
        text_notes = list(FilteredElementCollector(doc).OfClass(TextNote))
    except Exception:
        text_notes = []

    def _ensure_obs(led_id):
        obs = observations.get(led_id)
        if obs:
            return obs
        led = led_index.get(led_id)
        meta = led_meta.get(led_id, {})
        type_label = meta.get("type_label") or (led.get("label") if led else None) or led_id
        obs = {
            "label": type_label,
            "profile_name": meta.get("profile_name"),
            "type_label": type_label,
            "instances": 0,
            "params": {},
            "tags": {},
            "notes": {},
        }
        observations[led_id] = obs
        return obs

    def _assign_note(note_elem, kind):
        if note_elem is None:
            return
        try:
            note_point = getattr(note_elem, "Coord", None)
        except Exception:
            note_point = None
        if note_point is None:
            try:
                note_point = manage._get_point(note_elem)
            except Exception:
                note_point = None
        key = _note_location_key(note_point, kind)
        if key is None or key in assigned_keys:
            return
        candidates = []
        for record in _candidate_hosts(note_point, index, PROXIMITY_CELL_SIZE_FT):
            host_point = record.get("host_point")
            dist_sq = _distance_sq(note_point, host_point)
            if dist_sq is None or dist_sq > radius_sq:
                continue
            try:
                dist = math.sqrt(dist_sq)
            except Exception:
                dist = None
            if dist is None:
                continue
            candidates.append((record, dist))
        if not candidates:
            return
        candidates.sort(key=lambda item: item[1])
        min_dist = candidates[0][1]
        tie_candidates = [
            rec for rec, dist in candidates
            if abs(dist - min_dist) <= TIE_DISTANCE_TOLERANCE_FT
        ]
        if len(tie_candidates) > 1:
            chosen = _select_host_for_note(note_elem, tie_candidates, manage, kind)
        else:
            chosen = candidates[0][0]
        if not chosen:
            return
        host_point = chosen.get("host_point")
        led_id = chosen.get("led_id")
        if host_point is None or not led_id:
            return
        if kind == "text_note":
            note_entry = manage._build_text_note_entry(note_elem, host_point)
            if not note_entry:
                return
            obs = _ensure_obs(led_id)
            _track_annotation(obs, "notes", note_entry, _text_note_key)
        else:
            note_entry = manage._build_annotation_tag_entry(note_elem, host_point)
            if not note_entry or not _is_keynote_entry(note_entry):
                return
            offsets = note_entry.get("offsets") or {}
            offsets["rotation_deg"] = 0.0
            note_entry["offsets"] = offsets
            obs = _ensure_obs(led_id)
            _track_annotation(obs, "tags", note_entry, _tag_key)
        assigned_keys.add(key)

    for note_elem in keynotes:
        _assign_note(note_elem, "keynote")
    for note_elem in text_notes:
        _assign_note(note_elem, "text_note")


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

    led_index, led_meta = _index_leds(data)
    if not led_index:
        forms.alert("No linked element definitions found in {}.".format(yaml_label), title=TITLE)
        return

    selection = revit.get_selection()
    selected_elements = list(getattr(selection, "elements", []) or []) if selection else []
    if selected_elements:
        elements = _collect_candidate_elements(doc, selected_elements)
        if not elements:
            forms.alert("No candidate elements found in the current selection.", title=TITLE)
            return
    else:
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
        "missing_linker_ckts": 0,
    }

    observations = {}
    host_records = []
    assigned_note_keys = set()

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

        meta = led_meta.get(led_id, {})
        profile_name = meta.get("profile_name")
        type_label = meta.get("type_label") or led.get("label") or led_id
        obs = observations.get(led_id)
        if not obs:
            obs = {
                "label": type_label,
                "profile_name": profile_name,
                "type_label": type_label,
                "instances": 0,
                "params": {},
                "tags": {},
                "notes": {},
            }
            observations[led_id] = obs
        obs["instances"] += 1

        host_point = manage._get_point(elem)
        if host_point is not None:
            host_records.append({
                "led_id": led_id,
                "element": elem,
                "host_point": host_point,
                "profile_name": profile_name,
                "type_label": type_label,
            })
        params = manage._collect_params(elem)
        if isinstance(payload, dict):
            if not payload.get("CKT_Circuit Number_CEDT") or not payload.get("CKT_Panel_CEDT"):
                stats["missing_linker_ckts"] += 1
        params = _apply_linker_params(params, payload)
        for key, value in (params or {}).items():
            _track_param_value(obs, key, value)

        tags, keynotes, text_notes = manage._collect_hosted_tags(elem, host_point)
        for tag in (tags or []) + (keynotes or []):
            _track_annotation(obs, "tags", tag, _tag_key)
        for note in text_notes or []:
            _track_annotation(obs, "notes", note, _text_note_key)
        for note in keynotes or []:
            note_point = _entry_world_point(host_point, note)
            note_key = _note_location_key(note_point, "keynote")
            if note_key is not None:
                assigned_note_keys.add(note_key)
        for note in text_notes or []:
            note_point = _entry_world_point(host_point, note)
            note_key = _note_location_key(note_point, "text_note")
            if note_key is not None:
                assigned_note_keys.add(note_key)

    _assign_proximity_notes(doc, host_records, assigned_note_keys, observations, led_index, led_meta, manage)

    discrepancies = []
    any_tag_conflicts = False
    any_note_conflicts = False
    param_conflict_names = set()
    stats["param_conflicts"] = 0

    for led_id, obs in observations.items():
        led = led_index.get(led_id)
        if led is None:
            continue
        led_label = obs.get("label") or led_id
        profile_name = obs.get("profile_name")
        type_label = obs.get("type_label") or led_label
        existing_params = led.get("parameters") or {}
        tags_map = obs.get("tags", {})
        notes_map = obs.get("notes", {})
        instances = max(obs.get("instances", 1), 1)
        missing_tags = [
            _format_tag_entry(record["entry"])
            for record in tags_map.values()
            if record["count"] < instances
        ]
        missing_notes = [
            _format_note_entry(record["entry"])
            for record in notes_map.values()
            if record["count"] < instances
        ]
        param_conflicts = {}
        for param_name, value_map in obs.get("params", {}).items():
            if value_map and len(value_map) > 1:
                param_conflicts[param_name] = value_map
        if missing_tags or missing_notes or param_conflicts:
            discrepancies.append({
                "led_id": led_id,
                "label": led_label,
                "profile_name": profile_name,
                "type_label": type_label,
                "missing_tags": missing_tags,
                "missing_notes": missing_notes,
                "param_conflicts": param_conflicts,
                "existing_params": existing_params,
            })
            if missing_tags:
                any_tag_conflicts = True
            if missing_notes:
                any_note_conflicts = True
            for name in param_conflicts.keys():
                param_conflict_names.add(name)
            stats["param_conflicts"] += len(param_conflicts)

    decisions = {}
    if discrepancies:
        ui_module = _update_profiles_ui()
        if ui_module is None:
            return
        xaml_path = os.path.join(os.path.dirname(__file__), "UpdateProfilesUI.xaml")
        window = ui_module.UpdateProfilesWindow(
            xaml_path=xaml_path,
            discrepancies=discrepancies,
            param_names=sorted(param_conflict_names),
            include_tags=any_tag_conflicts,
            include_notes=any_note_conflicts,
        )
        result = window.show_dialog()
        if not result:
            return
        decisions = getattr(window, "decisions", {}) or {}

    global_settings = decisions.get("_global", {})
    replace_mode = bool(global_settings.get("replace_mode"))
    tag_set_mode = global_settings.get("tag_set") or "union"
    keynote_set_mode = global_settings.get("keynote_set") or "union"
    note_set_mode = global_settings.get("note_set") or "union"
    if not replace_mode:
        tag_set_mode = "union"
        keynote_set_mode = "union"
        note_set_mode = "union"

    for led_id, obs in observations.items():
        led = led_index.get(led_id)
        if led is None:
            continue
        existing_params = led.get("parameters") or {}
        chosen_params = {}
        param_choices = decisions.get(led_id, {}).get("params", {})

        for param_name, value_map in obs.get("params", {}).items():
            if not value_map:
                continue
            existing_value = existing_params.get(param_name)
            if len(value_map) == 1:
                value = next(iter(value_map.values()))["value"]
                if existing_value != value:
                    chosen_params[param_name] = value
                continue
            choice = param_choices.get(param_name)
            if not choice:
                continue
            action, value = choice
            if action in ("keep", "skip"):
                continue
            if existing_value != value:
                chosen_params[param_name] = value

        if chosen_params:
            if not isinstance(existing_params, dict):
                existing_params = {}
                led["parameters"] = existing_params
            for key, value in chosen_params.items():
                existing_params[key] = value
            stats["params_updated"] += 1
            stats["definitions_updated"].add(led_id)

        tags_map = obs.get("tags", {})
        notes_map = obs.get("notes", {})
        instances = max(obs.get("instances", 1), 1)
        decision_tags = decisions.get(led_id, {}).get("tags")
        decision_notes = decisions.get(led_id, {}).get("notes")

        if decision_tags == "skip":
            changed = False
            added = updated = 0
        else:
            tag_records, keynote_records = _partition_tag_records(tags_map.values())
            tag_entries = _entries_from_records(tag_records, instances, tag_set_mode)
            keynote_entries = _entries_from_records(keynote_records, instances, keynote_set_mode)
            if replace_mode:
                changed, added, updated = _set_tag_and_keynote_entries(led, tag_entries, keynote_entries)
            else:
                changed, added, updated = _merge_tag_and_keynote_entries(led, tag_entries, keynote_entries)
        if changed:
            stats["tags_added"] += added
            stats["tags_updated"] += updated
            stats["definitions_updated"].add(led_id)

        if decision_notes == "skip":
            changed = False
            added = updated = 0
        else:
            note_entries = _entries_from_records(notes_map.values(), instances, note_set_mode)
            if replace_mode:
                changed, added, updated = _set_text_note_entries(led, note_entries)
            else:
                changed, added, updated = _merge_text_note_entries(led, note_entries)
        if changed:
            stats["notes_added"] += added
            stats["notes_updated"] += updated
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
        "Annotation update mode: {}".format("Replace" if replace_mode else "Add-only"),
    ]
    if replace_mode:
        summary.append(
            "Replacement sets: tags={}, keynotes={}, text notes={}".format(
                tag_set_mode, keynote_set_mode, note_set_mode
            )
        )
    summary.extend([
        "Elements scanned: {} (with Element_Linker: {})".format(
            stats["elements_scanned"],
            stats["elements_with_linker"],
        ),
        "Definitions updated: {}".format(len(stats["definitions_updated"])),
        "Parameter values updated: {}".format(stats["params_updated"]),
        "Tags added/updated: {} / {}".format(stats["tags_added"], stats["tags_updated"]),
        "Text notes added/updated: {} / {}".format(stats["notes_added"], stats["notes_updated"]),
    ])
    if stats["missing_linker_ckts"]:
        summary.append(
            "Note: {} element(s) missing CKT values in Element_Linker payload.".format(
                stats["missing_linker_ckts"]
            )
        )
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
