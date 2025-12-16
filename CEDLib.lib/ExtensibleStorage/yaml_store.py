# -*- coding: utf-8 -*-
"""
Helpers for working with the active YAML stored inside Extensible Storage.
"""

import json
import io

from pyrevit import revit, script

from LogicClasses.profile_schema import load_data_from_text, dump_data_to_string  # noqa: E402
from ExtensibleStorage import ExtensibleStorage  # noqa: E402

_ACTIVE_CACHE = None
_LED_DEBUG_ID = "SET-002-LED-002"
SAFE_HASH = u"\uff03"


def _extract_led_snippet(text):
    if not text:
        return ""
    target = _LED_DEBUG_ID
    idx = text.find(target)
    if idx == -1:
        idx = 0
    offset_idx = text.find('"x_inches"', idx)
    if offset_idx == -1:
        offset_idx = text.find('"x_inches"')
    if offset_idx == -1:
        return text[idx: idx + 120].replace("\n", "\\n")
    start = max(0, offset_idx - 40)
    end = min(len(text), offset_idx + 120)
    return text[start:end].replace("\n", "\\n")


def _sanitize_hash_keys(raw_text):
    sanitized_lines = []
    for raw_line in raw_text.splitlines():
        line = raw_line
        if ":" in raw_line:
            prefix, suffix = raw_line.split(":", 1)
            indent_len = len(prefix) - len(prefix.lstrip(" "))
            indent = prefix[:indent_len]
            key = prefix[indent_len:]
            key = key.replace(SAFE_HASH, "#")
            needs_quote = "#" in key and not (key.startswith('"') and key.endswith('"'))
            if needs_quote:
                escaped = key.replace('"', '\\"')
                line = '{}"{}":{}'.format(indent, escaped, suffix)
            else:
                line = "{}{}:{}".format(indent, key, suffix)
        sanitized_lines.append(line)
    return "\n".join(sanitized_lines)

def _get_doc(doc=None):
    if doc:
        return doc
    return getattr(revit, "doc", None)


def load_active_yaml_text(doc=None):
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    path, normalized, text = ExtensibleStorage.get_active_yaml(doc)
    if not path or text is None:
        raise RuntimeError("Select YAML first so the profile data is loaded into the project.")
    return path, text


def load_active_yaml_data(doc=None):
    global _ACTIVE_CACHE
    path, text = load_active_yaml_text(doc)
    sanitized_text = _sanitize_hash_keys(text or "")
    normalized = ExtensibleStorage._normalize_path(path)
    if _ACTIVE_CACHE and _ACTIVE_CACHE.get("normalized") == normalized:
        cached = _ACTIVE_CACHE.get("data")
        if cached:
            data = json.loads(json.dumps(cached))
            logger = script.get_logger()
            logger.info("[YAML Storage] loaded equipment definitions (cached): %s", [eq.get("name") or eq.get("id") for eq in data.get("equipment_definitions") or [] if isinstance(eq, dict)])
            return path, data
    try:
        data = load_data_from_text(sanitized_text, path)
    except Exception:
        try:
            diag_path = "{}.sanitized".format(path)
            with io.open(diag_path, "w", encoding="utf-8") as diag_handle:
                diag_handle.write(sanitized_text)
        except Exception:
            pass
        raise
    logger = script.get_logger()
    logger.info("[YAML Storage] loaded equipment definitions: %s | snippet=%s", [eq.get("name") or eq.get("id") for eq in data.get("equipment_definitions") or [] if isinstance(eq, dict)], _extract_led_snippet(text))
    return path, data


def save_active_yaml_data(doc, data, action, description):
    global _ACTIVE_CACHE
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    path, text = load_active_yaml_text(doc)
    new_text = dump_data_to_string(data)
    logger = script.get_logger()
    snippet = _extract_led_snippet(new_text)
    logger.info("[YAML Storage] saving action=%s len=%s snippet=%s", action, len(new_text or ""), snippet)
    if new_text == text:
        return
    ExtensibleStorage.update_active_yaml(doc, path, text, new_text, action, description)
    ExtensibleStorage.update_active_text_only(doc, path, new_text)
    _ACTIVE_CACHE = {
        "normalized": ExtensibleStorage._normalize_path(path),
        "data": json.loads(json.dumps(data)),
    }


def refresh_active_yaml_snapshot(doc, yaml_path, data):
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    new_text = dump_data_to_string(data)
    ExtensibleStorage.update_active_text_only(doc, yaml_path, new_text)


def seed_active_yaml(doc, yaml_path, raw_text):
    doc = _get_doc(doc)
    if doc is None:
        raise RuntimeError("No active document detected.")
    sanitized = _sanitize_hash_keys(raw_text or "")
    ExtensibleStorage.seed_active_yaml(doc, yaml_path, sanitized)
