# -*- coding: utf-8 -*-
"""
Headless YAML import for MCP/AI automation.
Reads the YAML path from a sidecar file (yaml_path.txt) next to this script.
No UI dialogs — prints status to the pyRevit output window.
"""

import os
import sys
import io

from pyrevit import revit, script
output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.profile_schema import load_data_from_text, dump_data_to_string  # noqa: E402
from ExtensibleStorage import ExtensibleStorage  # noqa: E402
from ExtensibleStorage.yaml_store import seed_active_yaml  # noqa: E402

SAFE_HASH = u"\uff03"
SUPPORTED_SCHEMA_VERSIONS = [3, 4]
PATH_FILE = os.path.join(os.path.dirname(__file__), "yaml_path.txt")


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


def _normalize_schema_version(value):
    if value is None:
        return None
    try:
        return int(str(value).strip())
    except Exception:
        return None


def _collect_schema_versions(data):
    versions = []
    if isinstance(data, dict):
        if "schema_version" in data:
            versions.append(data.get("schema_version"))
        defs = data.get("equipment_definitions") or []
        if isinstance(defs, list):
            for eq in defs:
                if isinstance(eq, dict) and "schema_version" in eq:
                    versions.append(eq.get("schema_version"))
    return versions


def main():
    # Read YAML path from sidecar file
    if not os.path.isfile(PATH_FILE):
        print("ERROR: yaml_path.txt not found at: {}".format(PATH_FILE))
        print("Write the YAML file path to this file before invoking.")
        return

    with io.open(PATH_FILE, "r", encoding="utf-8") as f:
        picked = f.read().strip()

    if not picked or not os.path.isfile(picked):
        print("ERROR: YAML file not found: {}".format(picked))
        return

    print("Importing YAML: {}".format(picked))

    with io.open(picked, "r", encoding="utf-8") as handle:
        raw_text = handle.read()

    is_blank = not (raw_text or "").strip()
    sanitized_text = _sanitize_hash_keys(raw_text)

    try:
        data = load_data_from_text(sanitized_text, picked)
    except Exception as exc:
        print("Failed to parse YAML:\n\n{}".format(exc))
        return

    raw_versions = _collect_schema_versions(data)

    if not raw_versions and not is_blank:
        print(
            "Selected YAML is missing schema_version. Supported versions: {}.\n"
            "Import blocked.".format(", ".join([str(v) for v in SUPPORTED_SCHEMA_VERSIONS]))
        )
        return

    if not raw_versions and is_blank:
        normalized_text = dump_data_to_string(data)
        doc = getattr(revit, "doc", None)
        if doc is None:
            print("No active document detected; cannot store YAML in Extensible Storage.")
            return
        seed_active_yaml(doc, picked, normalized_text)
        print("Loaded '{}' into the project.".format(picked))
        return

    normalized_versions = []
    invalid_versions = []
    for value in raw_versions:
        normalized = _normalize_schema_version(value)
        if normalized is None:
            invalid_versions.append(value)
        else:
            normalized_versions.append(normalized)

    if invalid_versions:
        print(
            "Selected YAML has invalid schema_version values: {}.\n"
            "Supported versions: {}.\n"
            "Import blocked.".format(
                ", ".join([str(v) for v in invalid_versions]),
                ", ".join([str(v) for v in SUPPORTED_SCHEMA_VERSIONS]),
            )
        )
        return

    distinct_versions = sorted(set(normalized_versions))
    if any(v not in SUPPORTED_SCHEMA_VERSIONS for v in distinct_versions):
        print(
            "Selected YAML schema_version mismatch. Found: {}. Supported: {}.\n"
            "Import blocked.".format(
                ", ".join([str(v) for v in distinct_versions]),
                ", ".join([str(v) for v in SUPPORTED_SCHEMA_VERSIONS]),
            )
        )
        return

    normalized_text = dump_data_to_string(data)
    doc = getattr(revit, "doc", None)
    if doc is None:
        print("No active document detected; cannot store YAML in Extensible Storage.")
        return

    seed_active_yaml(doc, picked, normalized_text)
    print("Loaded '{}' into the project.".format(picked))


if __name__ == "__main__":
    main()
