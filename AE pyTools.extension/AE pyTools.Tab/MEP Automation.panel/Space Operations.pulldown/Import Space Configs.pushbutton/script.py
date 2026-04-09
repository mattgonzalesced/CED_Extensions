# -*- coding: utf-8 -*-
"""
Import Space Configs
--------------------
Import saved space-type templates from a YAML file.
"""

import io
import json
import os
import sys
from collections import OrderedDict
from datetime import datetime

from pyrevit import forms, revit, script

output = script.get_output()
output.close_others()

TITLE = "Import Space Configs"
CLASSIFICATION_STORAGE_ID = "space_operations.classifications.v1"
SPACE_PROFILE_SCHEMA_VERSION = 1
KEY_TYPE_ELEMENTS = "space_type_elements"

PLACEMENT_OPTIONS = [
    "Ceiling Corner Furthest from door",
    "One Foot off doorway wall",
    "Center of Furthest wall",
    "Center Ceiling",
    "Center Floor",
    "Center of Room",
    "Ceiling Corner Nearest Door",
]

DEFAULT_PLACEMENT_OPTION = "Center of Room"
BUCKETS = [
    "Restrooms",
    "Offices",
    "Sales Floor",
    "Freezers",
    "Coolers",
    "Receiving",
    "Break",
    "Food Prep",
    "Utility",
    "Storage",
    "Other",
]

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str

try:
    from System.Collections import IDictionary as NetIDictionary  # type: ignore
    from System.Collections import IEnumerable as NetIEnumerable  # type: ignore
except Exception:  # pragma: no cover
    NetIDictionary = None
    NetIEnumerable = None

try:
    import yaml  # type: ignore
except Exception:  # pragma: no cover
    try:
        from pyrevit.coreutils import yaml  # type: ignore
    except Exception:
        yaml = None


def _resolve_lib_root():
    cursor = os.path.abspath(os.path.dirname(__file__))
    for _ in range(12):
        candidate = os.path.join(cursor, "CEDLib.lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if not parent or parent == cursor:
            break
        cursor = parent
    return None


LIB_ROOT = _resolve_lib_root()
if LIB_ROOT and LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

try:
    from ExtensibleStorage import ExtensibleStorage  # noqa: E402
except Exception:
    ExtensibleStorage = None


def _to_python(value):
    if value is None:
        return None
    if isinstance(value, dict):
        return {k: _to_python(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_to_python(v) for v in value]
    if NetIDictionary and isinstance(value, NetIDictionary):
        py_dict = {}
        try:
            for key in value.Keys:
                py_dict[str(key)] = _to_python(value[key])
            return py_dict
        except Exception:
            return py_dict
    if NetIEnumerable and isinstance(value, NetIEnumerable) and not isinstance(value, basestring):
        try:
            return [_to_python(v) for v in list(value)]
        except Exception:
            pass
    keys_attr = getattr(value, "Keys", None)
    if keys_attr is not None:
        try:
            keys = list(value.Keys)
            return {str(k): _to_python(value[k]) for k in keys}
        except Exception:
            pass
    return value


def _sanitize_parameter_map(parameters):
    clean = OrderedDict()
    if not isinstance(parameters, dict):
        return clean

    for name, data in parameters.items():
        key = str(name or "").strip()
        if not key:
            continue
        if isinstance(data, dict):
            storage_type = str(data.get("storage_type") or "String")
            value = data.get("value")
            read_only = bool(data.get("read_only"))
        else:
            storage_type = "String"
            value = data
            read_only = False
        clean[key] = {
            "storage_type": storage_type,
            "value": "" if value is None else str(value),
            "read_only": read_only,
        }

    ordered = OrderedDict()
    for key in sorted(clean.keys(), key=lambda x: x.lower()):
        ordered[key] = clean[key]
    return ordered


def _valid_profile_duplicate_id(value):
    try:
        pid = int(str(value).strip())
    except Exception:
        return None
    if pid <= 0:
        return None
    return pid


def _sanitize_template_entry(entry):
    if not isinstance(entry, dict):
        return None

    entry_id = str(entry.get("id") or "").strip()
    kind = str(entry.get("kind") or "").strip().lower()
    if kind not in ("family_type", "model_group"):
        return None

    element_type_id = str(entry.get("element_type_id") or "").strip()
    if not element_type_id:
        element_type_id = entry_id.split(":", 1)[-1] if ":" in entry_id else ""

    if not entry_id and element_type_id:
        entry_id = "{}:{}".format(kind, element_type_id)

    if not entry_id:
        return None

    name = str(entry.get("name") or "").strip()
    if not name:
        name = "Family Type" if kind == "family_type" else "Model Group"

    placement_rule = str(entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION).strip()
    if placement_rule not in PLACEMENT_OPTIONS:
        placement_rule = DEFAULT_PLACEMENT_OPTION


    profile_duplicate_id = _valid_profile_duplicate_id(entry.get("profile_duplicate_id") or entry.get("duplicate_id"))

    return {
        "id": entry_id,
        "profile_duplicate_id": profile_duplicate_id,
        "kind": kind,
        "element_type_id": element_type_id,
        "name": name,
        "placement_rule": placement_rule,
        "parameters": _sanitize_parameter_map(entry.get("parameters") or {}),
    }


def _sanitize_template_list(raw_list):
    clean = []
    for raw in raw_list or []:
        entry = _sanitize_template_entry(raw)
        if not entry:
            continue
        clean.append(entry)
    return clean


def _sanitize_type_elements(raw_map):
    data = {}
    if not isinstance(raw_map, dict):
        raw_map = {}
    for bucket in BUCKETS:
        data[bucket] = _sanitize_template_list(raw_map.get(bucket) or [])
    return data


def _plain_parameter_map(parameters):
    out = {}
    if not isinstance(parameters, dict):
        return out
    for name, data in parameters.items():
        key = str(name or "").strip()
        if not key:
            continue
        if isinstance(data, dict):
            storage_type = str(data.get("storage_type") or "String")
            value = data.get("value")
            read_only = bool(data.get("read_only"))
        else:
            storage_type = "String"
            value = data
            read_only = False
        out[key] = {
            "storage_type": storage_type,
            "value": "" if value is None else str(value),
            "read_only": read_only,
        }
    return out


def _plain_template_entry(entry):
    return {
        "id": str(entry.get("id") or ""),
        "profile_duplicate_id": _valid_profile_duplicate_id(entry.get("profile_duplicate_id")),
        "kind": str(entry.get("kind") or ""),
        "element_type_id": str(entry.get("element_type_id") or ""),
        "name": str(entry.get("name") or ""),
        "placement_rule": str(entry.get("placement_rule") or DEFAULT_PLACEMENT_OPTION),
        "parameters": _plain_parameter_map(entry.get("parameters") or {}),
    }


def _plain_type_elements(type_elements):
    result = {}
    for bucket in BUCKETS:
        entries = type_elements.get(bucket) or []
        result[bucket] = [_plain_template_entry(entry) for entry in entries]
    return result


def _count_templates(type_elements):
    return sum(len(type_elements.get(bucket) or []) for bucket in BUCKETS)


def _load_yaml_mapping(path):
    with io.open(path, "r", encoding="utf-8") as handle:
        raw_text = handle.read()

    stripped = (raw_text or "").lstrip()
    if not stripped:
        return {}

    if stripped.startswith("{") or stripped.startswith("["):
        try:
            payload = json.loads(raw_text)
            if isinstance(payload, dict):
                return payload
        except Exception:
            pass

    if yaml:
        if hasattr(yaml, "load_as_dict"):
            loaded = yaml.load_as_dict(path) or {}
            payload = _to_python(loaded)
            if payload is None:
                return {}
            if not isinstance(payload, dict):
                raise ValueError("YAML root must be a map/object.")
            return payload

        loader = getattr(yaml, "safe_load", None) or getattr(yaml, "load", None)
        if loader is not None:
            try:
                loaded = loader(raw_text) or {}
            except TypeError:
                loader_cls = (
                    getattr(yaml, "SafeLoader", None)
                    or getattr(yaml, "CSafeLoader", None)
                    or getattr(yaml, "Loader", None)
                )
                if not loader_cls:
                    raise
                loaded = loader(raw_text, Loader=loader_cls) or {}
            payload = _to_python(loaded)
            if payload is None:
                return {}
            if not isinstance(payload, dict):
                raise ValueError("YAML root must be a map/object.")
            return payload

    payload = json.loads(raw_text)
    if not isinstance(payload, dict):
        raise ValueError("File root must be a map/object.")
    return payload


def _extract_type_elements_map(payload):
    if not isinstance(payload, dict):
        return None

    direct = payload.get(KEY_TYPE_ELEMENTS)
    if isinstance(direct, dict):
        return direct

    nested = payload.get("space_configs")
    if isinstance(nested, dict):
        nested_map = nested.get(KEY_TYPE_ELEMENTS)
        if isinstance(nested_map, dict):
            return nested_map

    if any(bucket in payload for bucket in BUCKETS):
        return payload

    return None


def _summary_lines(type_elements, source_path):
    lines = [
        "Imported space type templates.",
        "Storage ID: {}".format(CLASSIFICATION_STORAGE_ID),
        "Source file: {}".format(source_path),
        "",
        "Total templates: {}".format(_count_templates(type_elements)),
    ]
    for bucket in BUCKETS:
        lines.append("{}: {}".format(bucket, len(type_elements.get(bucket) or [])))
    lines.extend(
        [
            "",
            "Recommended workflow:",
            "1) Run Classify Spaces to map spaces into buckets.",
            "2) Open Manage Space Profiles for any per-space overrides.",
        ]
    )
    return lines


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    if ExtensibleStorage is None:
        forms.alert("Failed to load ExtensibleStorage library from CEDLib.lib.", title=TITLE)
        return

    source_path = forms.pick_file(
        file_ext="yaml",
        title=TITLE,
    )
    if not source_path:
        return

    try:
        loaded_payload = _load_yaml_mapping(source_path)
    except Exception as exc:
        forms.alert("Failed to parse space config YAML:\n\n{}".format(exc), title=TITLE)
        return

    raw_type_map = _extract_type_elements_map(loaded_payload)
    if not isinstance(raw_type_map, dict):
        forms.alert(
            "Selected file does not contain '{}' data.\n\n"
            "Expected either:\n"
            "- a top-level '{}' map, or\n"
            "- bucket keys like Restrooms/Offices/etc at the root.".format(KEY_TYPE_ELEMENTS, KEY_TYPE_ELEMENTS),
            title=TITLE,
        )
        return

    imported_type_elements = _sanitize_type_elements(raw_type_map)
    imported_total = _count_templates(imported_type_elements)

    existing_payload = ExtensibleStorage.get_project_data(doc, CLASSIFICATION_STORAGE_ID, default=None)
    if not isinstance(existing_payload, dict):
        existing_payload = {}

    existing_type_elements = _sanitize_type_elements(existing_payload.get(KEY_TYPE_ELEMENTS) or {})
    existing_total = _count_templates(existing_type_elements)

    if existing_total > 0:
        proceed = forms.alert(
            "This will replace currently saved space-type templates.\n\n"
            "Existing templates: {}\n"
            "Incoming templates: {}\n\n"
            "Continue?".format(existing_total, imported_total),
            title=TITLE,
            yes=True,
            no=True,
        )
        if not proceed:
            return

    payload = dict(existing_payload)
    payload["space_profile_schema_version"] = SPACE_PROFILE_SCHEMA_VERSION
    payload["space_profile_saved_utc"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    payload[KEY_TYPE_ELEMENTS] = _plain_type_elements(imported_type_elements)

    try:
        saved = ExtensibleStorage.set_project_data(
            doc,
            CLASSIFICATION_STORAGE_ID,
            payload,
            transaction_name="{} Save".format(TITLE),
        )
    except Exception as exc:
        forms.alert("Failed to import space configs:\n\n{}".format(exc), title=TITLE)
        return

    if not saved:
        forms.alert("Space configs were not saved.", title=TITLE)
        return

    forms.alert("\n".join(_summary_lines(imported_type_elements, source_path)), title=TITLE)


if __name__ == "__main__":
    main()
