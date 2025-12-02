# -*- coding: utf-8 -*-
"""
Delete YAML Profiles
--------------------
Select an equipment definition and remove linked element labels
from CEDLib.lib/profileData.yaml.
"""

from __future__ import print_function

import io
import os
import sys

from pyrevit import forms

LIB_ROOT = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "..",
        "..",
        "CEDLib.lib",
    )
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from profile_schema import load_data as load_profile_data, save_data as save_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")

try:
    basestring
except NameError:
    basestring = str

try:
    from System.Collections import IDictionary  # type: ignore
except Exception:  # pragma: no cover
    IDictionary = None


def _to_native(value):
    if value is None:
        return None
    if isinstance(value, dict):
        return {k: _to_native(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_to_native(v) for v in value]
    if IDictionary and isinstance(value, IDictionary):
        native = {}
        try:
            for key in value.Keys:
                native[str(key)] = _to_native(value[key])
            return native
        except Exception:
            pass
    keys_attr = getattr(value, "Keys", None)
    if keys_attr is not None:
        try:
            native = {}
            for key in list(value.Keys):
                native[str(key)] = _to_native(value[key])
            return native
        except Exception:
            pass
    if hasattr(value, "__iter__") and not isinstance(value, basestring):
        try:
            return [_to_native(v) for v in list(value)]
        except Exception:
            pass
    return value


def _parse_scalar(token):
    token = (token or "").strip()
    if not token:
        return ""
    if token in ("{}",):
        return {}
    if token in ("[]",):
        return []
    if token.startswith('"') and token.endswith('"'):
        return token[1:-1]
    lowered = token.lower()
    if lowered in ("true", "false"):
        return lowered == "true"
    if lowered == "null":
        return None
    try:
        if "." in token:
            return float(token)
        return int(token)
    except Exception:
        return token


def _simple_yaml_parse(text):
    lines = text.splitlines()

    def parse_block(start_idx, base_indent):
        idx = start_idx
        result = None
        while idx < len(lines):
            raw_line = lines[idx]
            stripped_line = raw_line.strip()
            if not stripped_line or stripped_line.startswith("#"):
                idx += 1
                continue
            indent = len(raw_line) - len(raw_line.lstrip(" "))
            if indent < base_indent:
                break
            if stripped_line.startswith("-"):
                if result is None:
                    result = []
                elif not isinstance(result, list):
                    break
                remainder = stripped_line[1:].strip()
                if remainder:
                    result.append(_parse_scalar(remainder))
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result.append(value)
            else:
                if result is None:
                    result = {}
                elif isinstance(result, list):
                    break
                key, _, remainder = stripped_line.partition(":")
                key = key.strip().strip('"')
                remainder = remainder.strip()
                if remainder:
                    result[key] = _parse_scalar(remainder)
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result[key] = value
        if result is None:
            result = {}
        return result, idx

    parsed, _ = parse_block(0, 0)
    return parsed if isinstance(parsed, dict) else {}


def _pick_profile_data_path():
    cached = get_cached_yaml_path()
    if cached and os.path.exists(cached):
        return cached
    path = forms.pick_file(
        file_ext="yaml",
        title="Select profileData YAML file",
    )
    if path:
        set_cached_yaml_path(path)
    return path


def _load_profile_store(data_path):
    data = load_profile_data(data_path)
    if data.get("equipment_definitions"):
        return data, False
    try:
        with io.open(data_path, "r", encoding="utf-8") as yaml_handle:
            fallback_data = _simple_yaml_parse(yaml_handle.read())
        if fallback_data.get("equipment_definitions"):
            return fallback_data, True
    except Exception:
        pass
    return data, False


def _normalize_name(value):
    if not value:
        return ""
    if isinstance(value, basestring):
        return value.strip()
    return str(value).strip()


def _build_definition_index(equipment_defs):
    index = {}
    for entry in equipment_defs or []:
        native = _to_native(entry) or {}
        display_name = _normalize_name(native.get("name") or native.get("id"))
        if not display_name:
            continue
        index[display_name] = native
    return index


def _collect_type_entries(equipment_def):
    entries = []
    seen_ids = set()
    for linked_set in equipment_def.get("linked_sets") or []:
        for linked_def in linked_set.get("linked_element_definitions") or []:
            if not isinstance(linked_def, dict):
                continue
            label = _normalize_name(linked_def.get("label"))
            led_id = _normalize_name(linked_def.get("id")) or label
            if not led_id or led_id in seen_ids:
                continue
            seen_ids.add(led_id)
            display = label or "<Unnamed>"
            entries.append({
                "id": led_id,
                "label": label or "<Unnamed>",
                "display": u"{}  [{}]".format(label or "<Unnamed>", led_id),
            })
    return entries


def _get_relations(entry):
    if not isinstance(entry, dict):
        return {"children": [], "parent": {}}
    relations = entry.get("linked_relations")
    if not isinstance(relations, dict):
        relations = {}
        entry["linked_relations"] = relations
    relations.setdefault("children", [])
    relations.setdefault("parent", {})
    return relations


def _pop_equipment_by_id(equipment_defs, target_id):
    target = (target_id or "").strip().lower()
    if not target:
        return None
    for entry in list(equipment_defs):
        current_id = (entry.get("id") or "").strip().lower()
        if current_id == target:
            try:
                equipment_defs.remove(entry)
            except ValueError:
                pass
            return entry
    return None


def _collect_children_refs(entry):
    relations = _get_relations(entry)
    child_refs = []
    for child in relations.get("children") or []:
        cid = (child.get("equipment_id") or "").strip()
        anchor_led = (child.get("anchor_led_id") or "").strip()
        if cid:
            child_refs.append((cid, anchor_led))
    return child_refs


def _cascade_remove_children(equipment_defs, initial_entries):
    removed_entries = []
    queue = list(initial_entries)
    processed = set()
    while queue:
        entry = queue.pop(0)
        children = _collect_children_refs(entry)
        for cid, _ in children:
            norm = cid.strip().lower()
            if not norm or norm in processed:
                continue
            child_entry = _pop_equipment_by_id(equipment_defs, cid)
            if child_entry:
                removed_entries.append(child_entry)
                queue.append(child_entry)
                processed.add(norm)
    return removed_entries


def _cleanup_relations(equipment_defs, removed_ids):
    removed_set = {(rid or "").strip().lower() for rid in removed_ids if rid}
    if not removed_set:
        return
    for entry in equipment_defs:
        relations = entry.get("linked_relations")
        if not isinstance(relations, dict):
            continue
        parent = relations.get("parent")
        if isinstance(parent, dict):
            pid = (parent.get("equipment_id") or "").strip().lower()
            if pid in removed_set:
                relations["parent"] = {}
        children = relations.get("children")
        if isinstance(children, list):
            relations["children"] = [
                child for child in children if (child.get("equipment_id") or "").strip().lower() not in removed_set
            ]


def _erase_entries(equipment_defs, definition_name, type_ids):
    if not equipment_defs:
        return False, []
    removed = False
    removed_entries = []
    for entry in list(equipment_defs):
        name = _normalize_name(entry.get("name") or entry.get("id"))
        if name != definition_name:
            continue
        linked_sets = entry.get("linked_sets") or []
        for linked_set in linked_sets:
            defs = linked_set.get("linked_element_definitions") or []
            filtered = [led for led in defs if _normalize_name(led.get("id")) not in type_ids]
            if len(filtered) != len(defs):
                removed = True
            linked_set["linked_element_definitions"] = filtered
        entry["linked_sets"] = [ls for ls in linked_sets if ls.get("linked_element_definitions")]
        if not entry["linked_sets"]:
            removed_entries.append(entry)
            try:
                equipment_defs.remove(entry)
            except ValueError:
                pass
        break
    return removed, removed_entries


def main():
    data_path = _pick_profile_data_path()
    if not data_path:
        return

    raw_data, _ = _load_profile_store(data_path)
    native_equipment_defs = [_to_native(entry) or {} for entry in raw_data.get("equipment_definitions") or []]
    definitions_by_name = _build_definition_index(native_equipment_defs)
    if not definitions_by_name:
        forms.alert("profileData.yaml currently contains no equipment definitions.", title="Delete YAML Profiles")
        return

    definition_choices = sorted(definitions_by_name.keys())

    definition_choice = forms.SelectFromList.show(
        definition_choices,
        title="Select equipment definition to delete linked types",
        multiselect=False,
        button_name="Select",
    )
    if not definition_choice:
        return

    definition = definitions_by_name.get(definition_choice)
    if not definition:
        forms.alert("Definition '{}' could not be loaded.".format(definition_choice), title="Delete YAML Profiles")
        return

    type_entries = _collect_type_entries(definition)
    if not type_entries:
        forms.alert("Definition '{}' has no linked element types to delete.".format(definition_choice), title="Delete YAML Profiles")
        return

    display_map = {entry["display"]: entry for entry in type_entries}
    picked = forms.SelectFromList.show(
        sorted(display_map.keys()),
        title="Select types to delete from '{}'".format(definition_choice),
        multiselect=True,
        button_name="Delete",
    )
    if not picked:
        return

    picked_entries = [display_map[name] for name in picked]
    picked_ids = {entry["id"] for entry in picked_entries}
    changed, removed_eq_entries = _erase_entries(native_equipment_defs, definition_choice, picked_ids)
    if not changed:
        return
    # Determine if individual LED deletions remove child relationships
    cascaded_entries = []
    for defn in native_equipment_defs:
        if (defn.get("name") or defn.get("id")) != definition_choice:
            continue
        relations = defn.get("linked_relations") or {}
        children = relations.get("children") or []
        led_by_id = {}
        for linked_set in defn.get("linked_sets") or []:
            for led_entry in linked_set.get("linked_element_definitions") or []:
                led_id = (led_entry.get("id") or "").strip()
                if led_id:
                    led_by_id[led_id] = led_entry
        affected_children = []
        for child in children:
            parent_led = (child.get("anchor_led_id") or "").strip()
            if not parent_led or parent_led.lower() not in {_normalize_name(tid) for tid in picked_ids}:
                continue
            child_id = (child.get("equipment_id") or "").strip()
            affected_children.append((child_id, child.get("anchor_led_id")))
        if affected_children:
            msg = [
                "Deleting those types will remove child links (shown below). Continue?",
                "",
            ]
            for cid, p_led in affected_children:
                msg.append(" - Child '{}', anchored to LED '{}'".format(cid or "<Unknown>", p_led or "<Unknown>"))
            if not forms.alert("\n".join(msg), title="Delete YAML Profiles", yes=True, no=True):
                return
            # also remove child entries referencing parent LEDs being deleted
            relations["children"] = [
                child
                for child in children
                if (child.get("anchor_led_id") or "").strip().lower() not in {_normalize_name(tid) for tid in picked_ids}
            ]
        break

    cascade_entries = _cascade_remove_children(native_equipment_defs, removed_eq_entries)
    all_removed_entries = list(removed_eq_entries) + list(cascade_entries)
    removed_ids = [(entry.get("id") or "").strip() for entry in all_removed_entries if isinstance(entry, dict)]

    _cleanup_relations(native_equipment_defs, removed_ids)

    save_profile_data(data_path, {"equipment_definitions": native_equipment_defs})
    summary = [
        "Deleted {} type(s) from definition '{}' and saved to profileData.yaml.".format(len(picked), definition_choice),
    ]
    if removed_ids:
        summary.append("Removed equipment definitions: {}".format(", ".join(sorted(set(removed_ids)))))
    summary.append("Reload Place Elements (YAML) to use updated data.")
    forms.alert("\n".join(summary), title="Delete YAML Profiles")


if __name__ == "__main__":
    main()
