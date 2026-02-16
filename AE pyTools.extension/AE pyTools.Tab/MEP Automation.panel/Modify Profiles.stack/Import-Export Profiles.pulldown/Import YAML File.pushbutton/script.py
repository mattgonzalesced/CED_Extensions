# -*- coding: utf-8 -*-
"""
Select default equipment-definition YAML for Let There Be YAML tools.
"""

import os
import sys
import io

from pyrevit import forms, revit, script
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

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")
SAFE_HASH = u"\uff03"
REQUIRED_SCHEMA_VERSION = 3


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
    active_path = None
    doc = getattr(revit, "doc", None)
    if doc is not None:
        try:
            active_path, _, _ = ExtensibleStorage.get_active_yaml(doc)
        except Exception:
            active_path = None
    init_dir = os.path.dirname(active_path) if active_path else os.path.dirname(DEFAULT_DATA_PATH)
    picked = forms.pick_file(
        file_ext="yaml",
        title="Select default equipment definition YAML",
        init_dir=init_dir,
    )
    if not picked:
        return
    with io.open(picked, "r", encoding="utf-8") as handle:
        raw_text = handle.read()
    is_blank = not (raw_text or "").strip()
    sanitized_text = _sanitize_hash_keys(raw_text)
    try:
        data = load_data_from_text(sanitized_text, picked)
    except Exception as exc:
        forms.alert("Failed to parse YAML:\n\n{}".format(exc), title="Select YAML")
        return
    raw_versions = _collect_schema_versions(data)
    if not raw_versions and not is_blank:
        forms.alert(
            "Selected YAML is missing schema_version. Expected schema_version: {}.\n"
            "Import blocked.".format(REQUIRED_SCHEMA_VERSION),
            title="Select YAML",
        )
        return
    if not raw_versions and is_blank:
        normalized_text = dump_data_to_string(data)
        doc = getattr(revit, "doc", None)
        if doc is None:
            forms.alert("No active document detected; cannot store YAML in Extensible Storage.", title="Select YAML")
            return
        seed_active_yaml(doc, picked, normalized_text)
        forms.alert(
            "Loaded '{}' into the project. All YAML operations now run from Extensible Storage.\n"
            "The original file will remain untouched until you export it again.".format(picked),
            title="Select YAML",
        )
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
        forms.alert(
            "Selected YAML has invalid schema_version values: {}.\n"
            "Expected schema_version: {}.\n"
            "Import blocked.".format(
                ", ".join([str(v) for v in invalid_versions]),
                REQUIRED_SCHEMA_VERSION,
            ),
            title="Select YAML",
        )
        return
    distinct_versions = sorted(set(normalized_versions))
    if distinct_versions != [REQUIRED_SCHEMA_VERSION]:
        forms.alert(
            "Selected YAML schema_version mismatch. Found: {}. Expected: {}.\n"
            "Import blocked.".format(
                ", ".join([str(v) for v in distinct_versions]),
                REQUIRED_SCHEMA_VERSION,
            ),
            title="Select YAML",
        )
        return
    normalized_text = dump_data_to_string(data)
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected; cannot store YAML in Extensible Storage.", title="Select YAML")
        return
    seed_active_yaml(doc, picked, normalized_text)
    forms.alert(
        "Loaded '{}' into the project. All YAML operations now run from Extensible Storage.\n"
        "The original file will remain untouched until you export it again.".format(picked),
        title="Select YAML",
    )


if __name__ == "__main__":
    main()
