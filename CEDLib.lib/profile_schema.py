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

try:
    # Prefer system PyYAML if available
    import yaml  # type: ignore
except Exception:  # pragma: no cover - IronPython fallback
    try:
        from pyrevit.coreutils import yaml  # type: ignore
    except Exception:
        yaml = None


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


def _yaml_dump_to_path(path, data):
    """
    Attempt to dump YAML regardless of whether safe_dump exists.
    Returns True if YAML dump succeeded, False otherwise.
    """
    if not yaml:
        return False

    if hasattr(yaml, "dump_dict"):
        try:
            yaml.dump_dict(data, path)
            return True
        except Exception:
            return False

    dumper = getattr(yaml, "safe_dump", None) or getattr(yaml, "dump", None)
    if dumper is None:
        return False

    kwargs = {"default_flow_style": False, "sort_keys": False, "allow_unicode": True}
    try:
        with io.open(path, "w", encoding="utf-8") as stream:
            dumper(data, stream, **kwargs)
        return True
    except TypeError:
        try:
            with io.open(path, "w", encoding="utf-8") as stream:
                dumper(data, stream)
            return True
        except Exception:
            return False
    except Exception:
        return False


def load_data(path):
    if not os.path.exists(path):
        return {"equipment_definitions": []}
    with io.open(path, "r", encoding="utf-8") as f:
        raw = f.read()

    data = None
    yaml_error = None
    if yaml:
        loaded, yaml_error = _yaml_load(raw, path)
        if isinstance(loaded, Mapping):
            data = dict(loaded)

    if data is None:
        stripped = (raw or "").lstrip()
        if stripped.startswith("{") or stripped.startswith("["):
            data = json.loads(raw or "{}")
        elif yaml_error:
            raise yaml_error
        else:
            raise ValueError("profileData.yaml could not be parsed as YAML and does not look like JSON.")

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
                diag_path = os.path.join(os.path.dirname(path), "profileData_load_diag.json")
                with io.open(diag_path, "w", encoding="utf-8") as diag_file:
                    json.dump(diag_payload, diag_file, indent=2)
                return alt_data
        except Exception as fallback_ex:
            diag_payload["error"] = str(fallback_ex)
        diag_path = os.path.join(os.path.dirname(path), "profileData_load_diag.json")
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


def equipment_defs_to_legacy(equipment_defs):
    legacy = []
    for eq in equipment_defs or []:
        if not isinstance(eq, Mapping):
            continue
        cad_name = eq.get("name") or eq.get("id") or "Unknown"
        types = []
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                inst_cfg = {
                    "offsets": led.get("offsets") or [{}],
                    "parameters": led.get("parameters") or {},
                    "tags": led.get("tags") or [],
                }
                types.append({
                    "label": led.get("label"),
                    "category_name": led.get("category"),
                    "is_group": bool(led.get("is_group")),
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
    existing = {led.get("id") for led in type_set.get("linked_element_definitions") or []}
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
        led_id = "{}-LED-{:03d}".format(set_id, idx)
        linked_defs.append({
            "id": led_id,
            "label": t.get("label"),
            "category": t.get("category_name"),
            "is_group": bool(t.get("is_group")),
            "offsets": offsets,
            "parameters": params,
            "tags": tags,
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
