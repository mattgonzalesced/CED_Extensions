# -*- coding: utf-8 -*-
"""
Export Space Configs
--------------------
Export saved space-type templates to a YAML file for reuse.
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

TITLE = "Export Space Configs"
CLASSIFICATION_STORAGE_ID = "space_operations.classifications.v1"
SPACE_CONFIG_SCHEMA_NAME = "space_operations.space_type_configs.v1"
SPACE_CONFIG_SCHEMA_VERSION = 1
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


def _to_builtin(value):
    if isinstance(value, OrderedDict):
        return {k: _to_builtin(v) for k, v in value.items()}
    if isinstance(value, dict):
        return {k: _to_builtin(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_to_builtin(v) for v in value]
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

    return {
        "id": entry_id,
        "kind": kind,
        "element_type_id": element_type_id,
        "name": name,
        "placement_rule": placement_rule,
        "parameters": _sanitize_parameter_map(entry.get("parameters") or {}),
    }


def _sanitize_template_list(raw_list):
    clean = []
    seen = set()
    for raw in raw_list or []:
        entry = _sanitize_template_entry(raw)
        if not entry:
            continue
        key = entry.get("id")
        if key in seen:
            continue
        seen.add(key)
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


def _build_export_payload(type_elements):
    payload = OrderedDict()
    payload["space_config_schema"] = SPACE_CONFIG_SCHEMA_NAME
    payload["schema_version"] = SPACE_CONFIG_SCHEMA_VERSION
    payload["storage_id"] = CLASSIFICATION_STORAGE_ID
    payload["exported_utc"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    payload[KEY_TYPE_ELEMENTS] = _plain_type_elements(type_elements)
    return payload


def _dump_yaml_to_path(path, data):
    payload = _to_builtin(_to_python(data))

    if yaml:
        if hasattr(yaml, "dump_dict") and not getattr(yaml, "safe_dump", None) and not getattr(yaml, "dump", None):
            try:
                yaml.dump_dict(payload, path)
                return True, None
            except Exception as exc:
                return False, exc

        dumper = getattr(yaml, "safe_dump", None) or getattr(yaml, "dump", None)
        if dumper is not None:
            try:
                with io.open(path, "w", encoding="utf-8") as stream:
                    try:
                        dumper(payload, stream, default_flow_style=False, sort_keys=False, allow_unicode=True)
                    except TypeError:
                        dumper(payload, stream)
                return True, None
            except Exception as exc:
                return False, exc

    try:
        with io.open(path, "w", encoding="utf-8") as stream:
            stream.write(json.dumps(payload, indent=2, ensure_ascii=False))
            stream.write("\n")
        return True, None
    except Exception as exc:
        return False, exc


def _summary_lines(type_elements, export_path):
    lines = [
        "Exported space type templates.",
        "Storage ID: {}".format(CLASSIFICATION_STORAGE_ID),
        "File: {}".format(export_path),
        "",
        "Total templates: {}".format(_count_templates(type_elements)),
    ]
    for bucket in BUCKETS:
        lines.append("{}: {}".format(bucket, len(type_elements.get(bucket) or [])))
    return lines


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    if ExtensibleStorage is None:
        forms.alert("Failed to load ExtensibleStorage library from CEDLib.lib.", title=TITLE)
        return

    payload = ExtensibleStorage.get_project_data(doc, CLASSIFICATION_STORAGE_ID, default=None)
    if not isinstance(payload, dict):
        payload = {}

    type_elements = _sanitize_type_elements(payload.get(KEY_TYPE_ELEMENTS) or {})

    save_path = forms.save_file(
        file_ext="yaml",
        title=TITLE,
        default_name="space_configs.yaml",
    )
    if not save_path:
        return

    export_payload = _build_export_payload(type_elements)
    saved, error = _dump_yaml_to_path(save_path, export_payload)
    if not saved:
        forms.alert("Failed to export space configs:\n\n{}".format(error), title=TITLE)
        return

    forms.alert("\n".join(_summary_lines(type_elements, save_path)), title=TITLE)


if __name__ == "__main__":
    main()
