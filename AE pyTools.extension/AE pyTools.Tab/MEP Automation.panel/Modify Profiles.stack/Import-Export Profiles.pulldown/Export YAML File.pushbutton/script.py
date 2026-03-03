# -*- coding: utf-8 -*-
"""
Export YAML File
----------------
Writes the currently active equipment-definition YAML (stored in Extensible Storage)
to a user-selected location so it can be versioned or shared. The export uses a
simple YAML serializer so the output resembles the original profile file.
"""

import io
import os
import re
from pyrevit import forms, script
output = script.get_output()
output.close_others()

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Export YAML File"
SAFE_HASH = u"\uff03"

BOOL_KEYS = {
    "allow_parentless",
    "allow_unmatched_parents",
    "prompt_on_parent_mismatch",
}
FLOAT_PATTERN = re.compile(r"^-?\d+\.\d+$")
NUMERIC_LIST_PATTERN = re.compile(r"^-?\d+(?:\.\d+)?(?:,-?\d+(?:\.\d+)?)+$")
YAML_SCHEMA_VERSION = 4


def _strip_control_chars(text):
    cleaned = []
    for ch in text:
        code = ord(ch)
        if code < 0x20 and code not in (0x09, 0x0A, 0x0D):
            cleaned.append(" ")
        else:
            cleaned.append(ch)
    return "".join(cleaned)


def _normalize_scalar_whitespace(text):
    if not text:
        return text
    text = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return text.replace("\t", " ")


def _strip_non_ascii(text):
    cleaned = []
    for ch in text:
        code = ord(ch)
        if code > 0x7E:
            cleaned.append(" ")
        else:
            cleaned.append(ch)
    return "".join(cleaned)


def _coerce_bool_strings(value, key=None):
    if isinstance(value, dict):
        return {k: _coerce_bool_strings(v, k) for k, v in value.items()}
    if isinstance(value, list):
        return [_coerce_bool_strings(item, key) for item in value]
    if key in BOOL_KEYS and isinstance(value, str):
        lowered = value.strip().lower()
        if lowered == "true":
            return True
        if lowered == "false":
            return False
    return value


def _with_schema_version(data):
    root = {"equipment_definitions": []}
    defs = []
    if isinstance(data, dict):
        defs = data.get("equipment_definitions") or []
    normalized_defs = []
    for entry in defs:
        if not isinstance(entry, dict):
            continue
        updated = dict(entry)
        updated["schema_version"] = YAML_SCHEMA_VERSION
        normalized_defs.append(updated)
    root["equipment_definitions"] = normalized_defs
    return root


try:
    basestring
except NameError:
    basestring = str


def _format_scalar(value, key=None):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    text = value if isinstance(value, basestring) else str(value)
    text = _normalize_scalar_whitespace(text)
    text = _strip_non_ascii(_strip_control_chars(text))
    if text == "":
        return "''"
    looks_like_numeric_list = bool(NUMERIC_LIST_PATTERN.match(text))
    needs_quotes = any(ch in text for ch in (":", "#", "{", "}", "[", "]", ",", "\n", "\r"))
    if key == "text":
        needs_quotes = True
    if looks_like_numeric_list:
        needs_quotes = False
    if text.lower() in ("true", "false", "null"):
        needs_quotes = True
    if needs_quotes:
        text = text.replace(SAFE_HASH, "#")
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        if "'" in text and '"' not in text:
            text = text.replace('"', '\\"')
            return '"' + text + '"'
        text = text.replace("'", "''")
        return "'" + text + "'"
    return text.replace(SAFE_HASH, "#")


def _dump_yaml_lines(value, indent=0):
    pad = " " * indent
    if isinstance(value, dict):
        lines = []
        for key, val in value.items():
            clean_key = (key or "").replace(SAFE_HASH, "#")
            if isinstance(val, dict):
                if not val:
                    lines.append("{}{}: {{}}".format(pad, clean_key))
                else:
                    lines.append("{}{}:".format(pad, clean_key))
                    lines.extend(_dump_yaml_lines(val, indent + 2))
            elif isinstance(val, list):
                if not val:
                    lines.append("{}{}: []".format(pad, clean_key))
                else:
                    lines.append("{}{}:".format(pad, clean_key))
                    lines.extend(_dump_yaml_lines(val, indent))
            else:
                lines.append("{}{}: {}".format(pad, clean_key, _format_scalar(val, key)))
        if not lines:
            lines.append("{}{{}}".format(pad))
        return lines
    if isinstance(value, list):
        lines = []
        if not value:
            lines.append("{}[]".format(pad))
            return lines
        for item in value:
            if isinstance(item, dict):
                if not item:
                    lines.append("{}- {{}}".format(pad))
                    continue
                items = list(item.items())
                first_key, first_val = items[0]
                clean_key = (first_key or "").replace(SAFE_HASH, "#")
                if isinstance(first_val, dict):
                    if not first_val:
                        lines.append("{}- {}: {{}}".format(pad, clean_key))
                    else:
                        lines.append("{}- {}:".format(pad, clean_key))
                        lines.extend(_dump_yaml_lines(first_val, indent + 2))
                elif isinstance(first_val, list):
                    if not first_val:
                        lines.append("{}- {}: []".format(pad, clean_key))
                    else:
                        lines.append("{}- {}:".format(pad, clean_key))
                        lines.extend(_dump_yaml_lines(first_val, indent + 2))
                else:
                    lines.append("{}- {}: {}".format(pad, clean_key, _format_scalar(first_val, first_key)))
                if len(items) > 1:
                    rest = dict(items[1:])
                    lines.extend(_dump_yaml_lines(rest, indent + 2))
            elif isinstance(item, list):
                lines.append("{}-".format(pad))
                lines.extend(_dump_yaml_lines(item, indent + 2))
            else:
                lines.append("{}- {}".format(pad, _format_scalar(item)))
        return lines
    return ["{}{}".format(pad, _format_scalar(value))]


def _dump_yaml_text(data):
    root = _with_schema_version(data)
    return "\n".join(_dump_yaml_lines(root)) + "\n"


def _sanitize_export_text(text):
    if not text:
        return text
    return _strip_non_ascii(_strip_control_chars(text))


def main():
    try:
        source_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    data = _coerce_bool_strings(data)
    yaml_text = _sanitize_export_text(_dump_yaml_text(data))
    default_name = os.path.basename(source_path) or "equipment_profiles.yaml"
    save_path = forms.save_file(
        file_ext="yaml",
        title=TITLE,
        default_name=default_name,
    )
    if not save_path:
        return

    try:
        with io.open(save_path, "w", encoding="utf-8") as handle:
            handle.write(yaml_text)
    except Exception as exc:
        forms.alert("Failed to export YAML:\n\n{}".format(exc), title=TITLE)
        return

    forms.alert(
        "Exported the active YAML to:\n{}".format(save_path),
        title=TITLE,
    )


if __name__ == "__main__":
    main()
