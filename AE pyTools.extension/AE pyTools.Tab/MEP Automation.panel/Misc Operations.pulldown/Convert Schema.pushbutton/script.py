# -*- coding: utf-8 -*-
"""
Convert legacy YAML data into the current schema format using a template file.
"""

import copy
import io
import json
import os
import sys
try:
    from collections.abc import Mapping
except ImportError:
    from collections import Mapping

from pyrevit import forms, script
output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

import LogicClasses.profile_schema as profile_schema  # noqa: E402
from LogicClasses.profile_schema import load_data_from_text  # noqa: E402

TITLE = "Convert Schema"
BOOL_STRING_KEYS = {
    "allow_parentless",
    "allow_unmatched_parents",
    "prompt_on_parent_mismatch",
}
FORCE_EMPTY_DICT_KEYS = {"equipment_properties"}


def _read_text(path):
    with io.open(path, "r", encoding="utf-8") as handle:
        return handle.read()


def _normalize_schema_version(value):
    if value is None:
        return None
    try:
        return int(str(value).strip())
    except Exception:
        return None


def _collect_schema_versions(data):
    versions = []
    if isinstance(data, Mapping):
        if "schema_version" in data:
            versions.append(data.get("schema_version"))
        defs = data.get("equipment_definitions") or []
        if isinstance(defs, list):
            for eq in defs:
                if isinstance(eq, Mapping) and "schema_version" in eq:
                    versions.append(eq.get("schema_version"))
    normalized = []
    invalid = []
    for value in versions:
        normalized_value = _normalize_schema_version(value)
        if normalized_value is None:
            invalid.append(value)
        else:
            normalized.append(normalized_value)
    return normalized, invalid


def _coerce_bool_string(value):
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return "true" if value else "false"
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in ("true", "false"):
            return lowered
        return "true" if lowered else "false"
    return "true" if value else "false"


def _load_template(path):
    raw = _read_text(path)
    loader = getattr(profile_schema, "_yaml_load", None)
    if callable(loader):
        loaded, _err = loader(raw, path)
        if isinstance(loaded, Mapping):
            return loaded
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, Mapping):
            return parsed
    except Exception:
        pass
    fallback = getattr(profile_schema, "_simple_yaml_parse", None)
    if callable(fallback):
        try:
            parsed = fallback(raw)
            if isinstance(parsed, Mapping):
                return parsed
        except Exception:
            pass
    try:
        return load_data_from_text(raw, path)
    except Exception:
        return {}


def _merge_defaults(target, defaults):
    if not isinstance(target, Mapping) or not isinstance(defaults, Mapping):
        return target
    for key, value in defaults.items():
        if key not in target or target[key] is None:
            target[key] = copy.deepcopy(value)
            continue
        if isinstance(target[key], Mapping) and isinstance(value, Mapping):
            _merge_defaults(target[key], value)
    return target


def _is_empty_map_key(key):
    if isinstance(key, Mapping):
        return not key
    if isinstance(key, str):
        return key.strip() == "{}"
    return False


def _is_empty_list_key(key):
    if isinstance(key, list):
        return len(key) == 0
    if isinstance(key, str):
        return key.strip() == "[]"
    return False


def _is_empty_container(value):
    if isinstance(value, Mapping):
        return not value
    if isinstance(value, list):
        return len(value) == 0
    return False


def _normalize_empty_containers(value):
    if isinstance(value, Mapping):
        cleaned = {}
        for key, item in value.items():
            cleaned[key] = _normalize_empty_containers(item)
        if len(cleaned) == 1:
            key = list(cleaned.keys())[0]
            item = cleaned[key]
            if _is_empty_map_key(key) and _is_empty_container(item):
                return {}
            if _is_empty_list_key(key) and _is_empty_container(item):
                return []
        cleaned = {
            key: item
            for key, item in cleaned.items()
            if not (_is_empty_map_key(key) or _is_empty_list_key(key))
        }
        return cleaned
    if isinstance(value, list):
        return [_normalize_empty_containers(item) for item in value]
    return value


def _shape_to_template(value, template):
    if isinstance(template, Mapping):
        if not isinstance(value, Mapping):
            return copy.deepcopy(template)
        if not template:
            return _normalize_empty_containers(copy.deepcopy(value))
        shaped = {}
        for key in template.keys():
            if key in value:
                shaped[key] = _shape_to_template(value.get(key), template.get(key))
            else:
                shaped[key] = copy.deepcopy(template.get(key))
        for key in value.keys():
            if key in template:
                continue
            shaped[key] = _normalize_empty_containers(copy.deepcopy(value.get(key)))
        return shaped
    if isinstance(template, list):
        if not isinstance(value, list):
            return copy.deepcopy(template)
        if not template:
            return copy.deepcopy(value)
        item_template = template[0]
        return [_shape_to_template(item, item_template) for item in value]
    if value is None:
        return copy.deepcopy(template)
    return copy.deepcopy(value)


def _shape_to_template_with_key(value, template, key=None):
    if key in FORCE_EMPTY_DICT_KEYS and isinstance(template, Mapping) and not template:
        return {}
    if key in BOOL_STRING_KEYS:
        return _coerce_bool_string(value)
    if isinstance(template, Mapping):
        if not isinstance(value, Mapping):
            return copy.deepcopy(template)
        if not template:
            return _normalize_empty_containers(copy.deepcopy(value))
        shaped = {}
        for child_key in template.keys():
            if child_key in value:
                shaped[child_key] = _shape_to_template_with_key(
                    value.get(child_key), template.get(child_key), child_key
                )
            else:
                shaped[child_key] = copy.deepcopy(template.get(child_key))
        for child_key in value.keys():
            if child_key in template:
                continue
            shaped[child_key] = _normalize_empty_containers(copy.deepcopy(value.get(child_key)))
        return shaped
    if isinstance(template, list):
        if not isinstance(value, list):
            return copy.deepcopy(template)
        if not template:
            return _normalize_empty_containers(copy.deepcopy(value))
        item_template = template[0]
        return [
            _shape_to_template_with_key(item, item_template, key)
            for item in value
        ]
    if isinstance(template, str) and template.strip().lower() in ("true", "false"):
        return _coerce_bool_string(value)
    if value is None:
        return copy.deepcopy(template)
    return copy.deepcopy(value)


def _convert_defs(old_defs, template_defs, desired_version):
    template_defaults = None
    if isinstance(template_defs, list) and template_defs:
        if isinstance(template_defs[0], Mapping):
            template_defaults = template_defs[0]
    converted = []
    for entry in old_defs:
        if not isinstance(entry, Mapping):
            continue
        if template_defaults:
            updated = _shape_to_template_with_key(entry, template_defaults)
        else:
            updated = copy.deepcopy(entry)
        updated = _normalize_empty_containers(updated)
        if desired_version is not None:
            updated["schema_version"] = desired_version
        converted.append(updated)
    return converted


def _write_output(path, payload):
    dump_to_path = getattr(profile_schema, "_yaml_dump_to_path", None)
    if callable(dump_to_path):
        try:
            if dump_to_path(path, payload):
                return True
        except Exception:
            pass
    try:
        with io.open(path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
        return True
    except Exception:
        return False


def main():
    init_dir = LIB_ROOT if os.path.isdir(LIB_ROOT) else None
    old_path = forms.pick_file(
        file_ext="yaml",
        title="Select OLD versioned YAML",
        init_dir=init_dir,
    )
    if not old_path:
        return

    template_path = forms.pick_file(
        file_ext="yaml",
        title="Select CURRENT versioned YAML (schema template)",
        init_dir=os.path.dirname(old_path) or init_dir,
    )
    if not template_path:
        return

    try:
        old_raw = _read_text(old_path)
        old_data = load_data_from_text(old_raw, old_path)
    except Exception as exc:
        forms.alert("Failed to read OLD YAML:\n\n{}".format(exc), title=TITLE)
        return

    template_data = _load_template(template_path)
    normalized_versions, invalid_versions = _collect_schema_versions(template_data)
    if invalid_versions:
        forms.alert(
            "Template YAML has invalid schema_version values: {}.\n"
            "Conversion blocked.".format(", ".join([str(v) for v in invalid_versions])),
            title=TITLE,
        )
        return
    distinct_versions = sorted(set(normalized_versions))
    if not distinct_versions:
        forms.alert(
            "Template YAML is missing schema_version. Conversion blocked.",
            title=TITLE,
        )
        return
    if len(distinct_versions) > 1:
        forms.alert(
            "Template YAML has multiple schema_version values: {}.\n"
            "Conversion blocked.".format(", ".join([str(v) for v in distinct_versions])),
            title=TITLE,
        )
        return

    desired_version = distinct_versions[0]
    old_defs = list(old_data.get("equipment_definitions") or [])
    template_defs = []
    if isinstance(template_data, Mapping):
        template_defs = list(template_data.get("equipment_definitions") or [])

    converted_defs = _convert_defs(old_defs, template_defs, desired_version)

    if isinstance(template_data, Mapping) and "equipment_definitions" in template_data:
        output_payload = copy.deepcopy(template_data)
        output_payload["equipment_definitions"] = converted_defs
        if "schema_version" in output_payload:
            output_payload["schema_version"] = desired_version
    else:
        output_payload = {"equipment_definitions": converted_defs}
    output_payload = _normalize_empty_containers(output_payload)

    default_name = "converted_schema.yaml"
    save_path = forms.save_file(
        file_ext="yaml",
        title=TITLE,
        default_name=default_name,
    )
    if not save_path:
        return

    if not _write_output(save_path, output_payload):
        # Fallback to standard schema dump to preserve formatting expectations.
        try:
            text = profile_schema.dump_data_to_string({"equipment_definitions": converted_defs})
            with io.open(save_path, "w", encoding="utf-8") as handle:
                handle.write(text)
        except Exception:
            forms.alert("Failed to save converted YAML.", title=TITLE)
            return

    summary = [
        "Converted YAML saved to:",
        save_path,
        "",
        "Source entries: {}".format(len(old_defs)),
        "Converted entries: {}".format(len(converted_defs)),
        "Schema version applied: {}".format(desired_version),
    ]
    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
