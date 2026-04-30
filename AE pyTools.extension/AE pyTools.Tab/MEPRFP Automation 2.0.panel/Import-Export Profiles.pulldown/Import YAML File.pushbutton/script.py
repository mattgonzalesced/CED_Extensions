#! python3
# -*- coding: utf-8 -*-
"""
MEPRFP Automation 2.0 :: Import YAML File

Pick a YAML file, validate its schema_version, and persist the canonical
payload into Project Information Extensible Storage. The store is
independent from the legacy MEP Automation panel.
"""

import os
import sys

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

import forms_compat as forms
import active_yaml
import storage
import schema as _schema
import yaml_io


TITLE = "Import YAML File (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()

    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    init_dir = None
    existing = storage.read_payload(doc)
    if existing and existing.get("source_path"):
        try:
            init_dir = os.path.dirname(existing["source_path"]) or None
        except Exception:
            init_dir = None

    picked = forms.pick_file(
        file_ext="yaml",
        title="Select equipment definition YAML",
        init_dir=init_dir,
    )
    if not picked:
        return

    try:
        result = active_yaml.import_yaml_file(doc, picked)
    except _schema.SchemaVersionError as exc:
        forms.alert(
            "Schema version check failed:\n\n{}".format(exc),
            title=TITLE,
        )
        return
    except yaml_io.YamlError as exc:
        forms.alert(
            "Failed to parse YAML:\n\n{}".format(exc),
            title=TITLE,
        )
        return
    except (IOError, OSError) as exc:
        forms.alert(
            "Failed to read file:\n\n{}".format(exc),
            title=TITLE,
        )
        return
    except Exception as exc:
        forms.alert(
            "Unexpected error during import:\n\n{}".format(exc),
            title=TITLE,
            exitscript=False,
        )
        raise

    note = " (blank file - empty store created)" if result["blank"] else ""
    migrated = result["input_schema_version"] != result["stored_schema_version"]
    version_line = "`{}`".format(result["stored_schema_version"])
    if migrated:
        version_line = "`{}` (migrated from `{}`)".format(
            result["stored_schema_version"], result["input_schema_version"]
        )
    output.print_md(
        "**Import succeeded{}**\n\n"
        "- Source: `{}`\n"
        "- Schema version: {}\n"
        "- Stored bytes: `{}`\n".format(
            note,
            result["source_path"],
            version_line,
            result["byte_count"],
        )
    )


if __name__ == "__main__":
    main()
