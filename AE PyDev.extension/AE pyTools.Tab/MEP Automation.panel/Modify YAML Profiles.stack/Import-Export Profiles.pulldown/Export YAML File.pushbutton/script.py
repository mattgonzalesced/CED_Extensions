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

from pyrevit import forms

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Export YAML File"

try:
    basestring
except NameError:
    basestring = str


def _format_scalar(value):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    text = value if isinstance(value, basestring) else str(value)
    if text == "":
        return '""'
    needs_quotes = any(ch in text for ch in (":", "#", "{", "}", "[", "]", ",", "\n", "\r"))
    if text.lower() in ("true", "false", "null"):
        needs_quotes = True
    if needs_quotes:
        return '"' + text.replace('"', '\\"') + '"'
    return text


def _emit_multiline_block(text, indent):
    pad = " " * indent
    lines = ["{}|".format(pad)]
    if text:
        for line in text.splitlines():
            lines.append("{}  {}".format(pad, line))
    else:
        lines.append("{}  ".format(pad))
    return lines


def _dump_yaml_lines(value, indent=0):
    pad = " " * indent
    if isinstance(value, dict):
        lines = []
        for key, val in value.items():
            if isinstance(val, (dict, list)):
                lines.append("{}{}:".format(pad, key))
                lines.extend(_dump_yaml_lines(val, indent + 2))
            elif isinstance(val, basestring) and ("\n" in val or "\r" in val):
                lines.append("{}{}: |".format(pad, key))
                body = val.splitlines()
                if not body:
                    lines.append("{}  ".format(pad))
                else:
                    for line in body:
                        lines.append("{}  {}".format(pad, line))
            else:
                lines.append("{}{}: {}".format(pad, key, _format_scalar(val)))
        if not lines:
            lines.append("{}{{}}".format(pad))
        return lines
    if isinstance(value, list):
        lines = []
        if not value:
            lines.append("{}[]".format(pad))
            return lines
        for item in value:
            if isinstance(item, (dict, list)):
                lines.append("{}-".format(pad))
                lines.extend(_dump_yaml_lines(item, indent + 2))
            elif isinstance(item, basestring) and ("\n" in item or "\r" in item):
                lines.append("{}- |".format(pad))
                body = item.splitlines()
                if not body:
                    lines.append("{}  ".format(pad))
                else:
                    for line in body:
                        lines.append("{}  {}".format(pad, line))
            else:
                lines.append("{}- {}".format(pad, _format_scalar(item)))
        return lines
    return ["{}{}".format(pad, _format_scalar(value))]


def _dump_yaml_text(data):
    root = {"equipment_definitions": data.get("equipment_definitions") or []}
    return "\n".join(_dump_yaml_lines(root)) + "\n"


def main():
    try:
        source_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    yaml_text = _dump_yaml_text(data)
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
