# -*- coding: utf-8 -*-
"""
Helpers for working with the profileData.yaml equipment_definitions schema.
"""

import copy
import io
import json
import os
try:
    from collections.abc import Mapping
except ImportError:
    from collections import Mapping

try:
    from System.Collections import IDictionary as NetIDictionary  # type: ignore
    from System.Collections import IEnumerable as NetIEnumerable  # type: ignore
except Exception:  # pragma: no cover
    NetIDictionary = None
    NetIEnumerable = None

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str

SAFE_HASH = u"\uff03"
try:
    # Prefer system PyYAML if available
    import yaml  # type: ignore
except Exception:  # pragma: no cover - IronPython fallback
    try:
        from pyrevit.coreutils import yaml  # type: ignore
    except Exception:
        yaml = None

ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
ESCAPED_QUOTE_KEYS = ("label", "type_name")


class _ElementLinkerString(str):
    pass


def _wrap_element_linker_strings(value):
    if isinstance(value, Mapping):
        wrapped = {}
        for key, item in value.items():
            wrapped[key] = _wrap_element_linker_strings(item)
        return wrapped
    if isinstance(value, list):
        return [_wrap_element_linker_strings(item) for item in value]
    if isinstance(value, basestring):
        text = str(value)
        if "\r" in text:
            text = text.replace("\r\n", "\n").replace("\r", "\n")
        if "\n" in text:
            return _ElementLinkerString(text)
        return text
    return value


def _normalize_escaped_quotes(value):
    if isinstance(value, Mapping):
        cleaned = {}
        for key, item in value.items():
            if key in ESCAPED_QUOTE_KEYS and isinstance(item, basestring):
                cleaned[key] = str(item).replace('\\"', '"')
            else:
                cleaned[key] = _normalize_escaped_quotes(item)
        return cleaned
    if isinstance(value, list):
        return [_normalize_escaped_quotes(item) for item in value]
    return value


def _is_empty_map(value):
    return isinstance(value, Mapping) and not value


def _is_empty_map_key(key):
    if isinstance(key, Mapping):
        return not key
    if isinstance(key, basestring):
        return key.strip() == "{}"
    return False


def _cleanup_empty_maps(value):
    if isinstance(value, Mapping):
        empty_key_count = sum(1 for key in value.keys() if _is_empty_map_key(key))
        cleaned = {}
        for key, item in value.items():
            if _is_empty_map_key(key) and empty_key_count > 1:
                continue
            cleaned[key] = _cleanup_empty_maps(item)
        if not cleaned:
            return {}
        return cleaned
    if isinstance(value, list):
        cleaned_list = []
        for item in value:
            cleaned_item = _cleanup_empty_maps(item)
            if _is_empty_map(cleaned_item):
                continue
            cleaned_list.append(cleaned_item)
        return cleaned_list
    return value


def _prepare_dump_payload(payload):
    normalized = _normalize_escaped_quotes(payload)
    cleaned = _cleanup_empty_maps(normalized)
    return _wrap_element_linker_strings(cleaned)


def _build_element_linker_dumper():
    if not yaml:
        return None
    dumper_base = getattr(yaml, "SafeDumper", None)
    if dumper_base is None:
        return None

    class ElementLinkerDumper(dumper_base):
        pass

    def _represent_quoted_string(dumper, data):
        return dumper.represent_scalar("tag:yaml.org,2002:str", data, style='"')

    try:
        ElementLinkerDumper.add_representer(_ElementLinkerString, _represent_quoted_string)
    except Exception:
        return None
    return ElementLinkerDumper


_ELEMENT_LINKER_DUMPER = _build_element_linker_dumper()


def _to_python(value):
    """
    Convert .NET YamlDotNet objects or other generic containers into pure Python types.
    """
    if value is None:
        return None
    if isinstance(value, Mapping):
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
    # IronPython specific: objects with Keys attribute
    keys_attr = getattr(value, "Keys", None)
    if keys_attr is not None and callable(keys_attr):
        try:
            keys = list(value.Keys)
            return {str(k): _to_python(value[k]) for k in keys}
        except Exception:
            pass
    return value


def _yaml_load(raw_text, source_path):
    """
    Attempt to load YAML regardless of whether the module exposes safe_* helpers.
    Returns (data, error).
    """
    if not yaml:
        return None, None

    # pyRevit coreutils yaml exposes load_as_dict/dump_dict that take file paths.
    if hasattr(yaml, "load_as_dict"):
        try:
            loaded = yaml.load_as_dict(source_path) or {}
            return _to_python(loaded), None
        except Exception as ex:
            return None, ex

    loader = getattr(yaml, "safe_load", None) or getattr(yaml, "load", None)
    if loader is None:
        return None, AttributeError("yaml module does not expose load/safe_load")

    try:
        return loader(raw_text) or {}, None
    except TypeError:
        loader_cls = (
            getattr(yaml, "SafeLoader", None)
            or getattr(yaml, "CSafeLoader", None)
            or getattr(yaml, "Loader", None)
        )
        try:
            if loader_cls:
                return loader(raw_text, Loader=loader_cls) or {}, None
            return loader(raw_text) or {}, None
        except Exception as ex:
            return None, ex
    except Exception as ex:
        return None, ex


def _parse_scalar(token):
    token = token.strip()
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


def _simple_yaml_parse(raw_text):
    lines = raw_text.splitlines()

    def parse_block(start_idx, base_indent):
        result = None
        idx = start_idx

        while idx < len(lines):
            line = lines[idx]
            stripped_line = line.strip()
            if not stripped_line or stripped_line.startswith("#"):
                idx += 1
                continue

            indent = len(line) - len(line.lstrip(" "))
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
    if isinstance(parsed, dict):
        return parsed
    return {}


def _sanitize_hash_keys(raw_text):
    sanitized_lines = []
    for raw_line in raw_text.splitlines():
        line = raw_line
        if ":" in raw_line:
            prefix, suffix = raw_line.split(":", 1)
            if "#" in prefix and SAFE_HASH not in prefix:
                prefix = prefix.replace("#", SAFE_HASH)
                line = "{}:{}".format(prefix, suffix)
        sanitized_lines.append(line)
    return "\n".join(sanitized_lines)


def _yaml_dump_to_path(path, data):
    """
    Attempt to dump YAML regardless of whether safe_dump exists.
    Returns True if YAML dump succeeded, False otherwise.
    """
    if not yaml:
        return False

    if hasattr(yaml, "dump_dict") and not getattr(yaml, "safe_dump", None) and not getattr(yaml, "dump", None):
        try:
            yaml.dump_dict(data, path)
            return True
        except Exception:
            return False

    dumper = getattr(yaml, "safe_dump", None) or getattr(yaml, "dump", None)
    if dumper is None:
        return False

    prepared = _prepare_dump_payload(data)
    kwargs = {"default_flow_style": False, "sort_keys": False, "allow_unicode": True}
    if _ELEMENT_LINKER_DUMPER is not None:
        kwargs["Dumper"] = _ELEMENT_LINKER_DUMPER
    try:
        with io.open(path, "w", encoding="utf-8") as stream:
            dumper(prepared, stream, **kwargs)
        return True
    except TypeError:
        try:
            with io.open(path, "w", encoding="utf-8") as stream:
                dumper(prepared, stream)
            return True
        except Exception:
            return False
    except Exception:
        return False


def load_data(path):
    if not os.path.exists(path):
        return {"equipment_definitions": []}
    with io.open(path, "r", encoding="utf-8") as f:
        raw = _sanitize_hash_keys(f.read())
    return load_data_from_text(raw, path)


def load_data_from_text(raw_text, source_label="<memory>"):
    raw = _sanitize_hash_keys(raw_text or "")
    stripped = raw.lstrip()
    if stripped.startswith("{") or stripped.startswith("["):
        try:
            return _cleanup_empty_maps(_normalize_escaped_quotes(json.loads(raw or "{}")))
        except Exception:
            pass
    data = None
    yaml_error = None
    if yaml:
        loaded, yaml_error = _yaml_load(raw, source_label)
        if isinstance(loaded, Mapping):
            data = dict(loaded)

    if data is None:
        if yaml_error:
            # Attempt fallback parser before surfacing error
            try:
                alt_data = _simple_yaml_parse(raw)
                if isinstance(alt_data, Mapping):
                    data = dict(alt_data)
            except Exception:
                pass
            if data is None:
                raise yaml_error
        else:
            raise ValueError("profile data could not be parsed as YAML and does not look like JSON.")

    data = _cleanup_empty_maps(_normalize_escaped_quotes(data))

    if "equipment_definitions" in data:
        defs = data.get("equipment_definitions") or []
        data["equipment_definitions"] = [d for d in defs if isinstance(d, Mapping)]
        if data["equipment_definitions"]:
            return data
        diag_payload = {"error": None}
        try:
            alt_data = _simple_yaml_parse(raw)
            cleaned_defs = [
                entry for entry in (alt_data.get("equipment_definitions") or []) if isinstance(entry, Mapping)
            ]
            alt_data["equipment_definitions"] = cleaned_defs
            diag_payload["fallback_equipment_defs_length"] = len(cleaned_defs)
            diag_payload["fallback_sample"] = cleaned_defs[:1]
            if cleaned_defs:
                if source_label and os.path.exists(os.path.dirname(source_label)):
                    diag_path = os.path.join(os.path.dirname(source_label), "profileData_load_diag.json")
                    with io.open(diag_path, "w", encoding="utf-8") as diag_file:
                        json.dump(diag_payload, diag_file, indent=2)
                return _cleanup_empty_maps(_normalize_escaped_quotes(alt_data))
        except Exception as fallback_ex:
            diag_payload["error"] = str(fallback_ex)
        if source_label and os.path.exists(os.path.dirname(source_label)):
            diag_path = os.path.join(os.path.dirname(source_label), "profileData_load_diag.json")
            try:
                with io.open(diag_path, "w", encoding="utf-8") as diag_file:
                    json.dump(diag_payload, diag_file, indent=2)
            except Exception:
                pass
    profiles = data.get("profiles")
    if profiles:
        return {"equipment_definitions": legacy_to_equipment_defs(profiles, [])}
    return {"equipment_definitions": []}


def save_data(path, data):
    payload = {"equipment_definitions": data.get("equipment_definitions", [])}
    if not _yaml_dump_to_path(path, payload):
        with io.open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)


def dump_data_to_string(data):
    payload = {"equipment_definitions": data.get("equipment_definitions", [])}
    if yaml:
        dumper = getattr(yaml, "safe_dump", None) or getattr(yaml, "dump", None)
        if dumper:
            prepared = _prepare_dump_payload(payload)
            kwargs = {"default_flow_style": False, "sort_keys": False, "allow_unicode": True}
            try:
                if _ELEMENT_LINKER_DUMPER is not None:
                    kwargs["Dumper"] = _ELEMENT_LINKER_DUMPER
                return dumper(prepared, **kwargs)
            except TypeError:
                return dumper(prepared)
    return json.dumps(payload, indent=2)


def equipment_defs_to_legacy(equipment_defs):
    legacy = []
    for eq in equipment_defs or []:
        if not isinstance(eq, Mapping):
            continue
        cad_name = eq.get("name") or eq.get("id") or "Unknown"
        types = []
        for linked_set in eq.get("linked_sets") or []:
            set_id = linked_set.get("id")
            for led in linked_set.get("linked_element_definitions") or []:
                if led.get("is_parent_anchor"):
                    continue
                inst_cfg = {
                    "offsets": led.get("offsets") or [{}],
                    "parameters": led.get("parameters") or {},
                    "tags": led.get("tags") or [],
                    "keynotes": led.get("keynotes") or [],
                    "text_notes": led.get("text_notes") or [],
                }
                types.append({
                    "label": led.get("label"),
                    "category_name": led.get("category"),
                    "is_group": bool(led.get("is_group")),
                    "led_id": led.get("id"),
                    "set_id": set_id,
                    "instance_config": inst_cfg,
                })
        legacy.append({
            "cad_name": cad_name,
            "types": types,
        })
    return legacy


def legacy_to_equipment_defs(profiles, existing_defs=None):
    existing_defs = existing_defs or []
    existing_by_name = {(d.get("name") or d.get("id")): d for d in existing_defs}
    new_defs = []
    for idx, prof in enumerate(profiles or [], 1):
        cad = prof.get("cad_name") or "Unknown"
        base = existing_by_name.get(cad)
        if base:
            eq = copy.deepcopy(base)
        else:
            eq = _create_equipment_stub(cad, prof.get("types") or [], idx)
        eq["linked_sets"] = [_types_to_linked_set(eq, prof.get("types") or [], idx)]
        new_defs.append(eq)
    return new_defs


def ensure_equipment_definition(data, cad_name, sample_entry):
    defs = data.setdefault("equipment_definitions", [])
    for eq in defs:
        if (eq.get("name") or eq.get("id")) == cad_name:
            return eq
    stub = _create_equipment_stub(cad_name, [sample_entry], len(defs) + 1)
    defs.append(stub)
    return stub


def get_type_set(equipment_def):
    sets = equipment_def.setdefault("linked_sets", [])
    if sets:
        return sets[0]
    set_id = "{}-SET-001".format(equipment_def.get("id") or "SET")
    new_set = {
        "id": set_id,
        "name": "{} Types".format(equipment_def.get("name") or "Types"),
        "linked_element_definitions": [],
    }
    sets.append(new_set)
    return new_set


def next_led_id(type_set, equipment_def):
    base = type_set.get("id") or (equipment_def.get("id") or "LED")
    existing = {
        led.get("id")
        for led in type_set.get("linked_element_definitions") or []
        if not (led.get("is_parent_anchor") if isinstance(led, dict) else False)
    }
    counter = len(existing) + 1
    candidate = "{}-LED-{:03d}".format(base, counter)
    while candidate in existing:
        counter += 1
        candidate = "{}-LED-{:03d}".format(base, counter)
    return candidate


def _create_equipment_stub(cad_name, types, seq):
    label = ""
    category = "Uncategorized"
    if types:
        label = types[0].get("label") or ""
        category = types[0].get("category_name") or "Uncategorized"
    family, type_name = _split_label(label)
    eq_id = "EQ-{:03d}".format(seq)
    set_id = "SET-{:03d}".format(seq)
    return {
        "id": eq_id,
        "name": cad_name,
        "version": 1,
        "schema_version": 1,
        "allow_parentless": True,
        "allow_unmatched_parents": True,
        "prompt_on_parent_mismatch": False,
        "parent_filter": {
            "category": category,
            "family_name_pattern": family or "*",
            "type_name_pattern": type_name or "*",
            "parameter_filters": {},
        },
        "equipment_properties": {},
        "linked_sets": [
            {
                "id": set_id,
                "name": "{} Types".format(cad_name),
                "linked_element_definitions": [],
            }
        ],
    }


def _types_to_linked_set(equipment_def, types, seq):
    existing_sets = equipment_def.get("linked_sets") or []
    if existing_sets:
        base_set = copy.deepcopy(existing_sets[0])
        set_id = base_set.get("id") or "SET-{:03d}".format(seq)
        set_name = base_set.get("name") or "{} Types".format(equipment_def.get("name") or "Types")
    else:
        set_id = "SET-{:03d}".format(seq)
        set_name = "{} Types".format(equipment_def.get("name") or "Types")
    linked_defs = []
    for idx, t in enumerate(types or [], 1):
        inst = t.get("instance_config") or {}
        offsets = inst.get("offsets") or []
        params = inst.get("parameters") or {}
        tags = inst.get("tags") or []
        keynotes = inst.get("keynotes") or []
        led_id = t.get("led_id") or "{}-LED-{:03d}".format(set_id, idx)
        linked_defs.append({
            "id": led_id,
            "label": t.get("label"),
            "category": t.get("category_name"),
            "is_group": bool(t.get("is_group")),
            "offsets": offsets,
            "parameters": params,
            "tags": tags,
            "keynotes": keynotes,
            "text_notes": inst.get("text_notes") or [],
        })
    return {
        "id": set_id,
        "name": set_name,
        "linked_element_definitions": linked_defs,
    }


def _split_label(label):
    if not label or ":" not in label:
        lbl = (label or "").strip()
        return lbl or "*", "*"
    fam_part, type_part = label.split(":", 1)
    return fam_part.strip() or "*", type_part.strip() or "*"
