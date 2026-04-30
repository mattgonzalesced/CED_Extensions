#! python3
# -*- coding: utf-8 -*-
"""
MEPRFP Automation 2.0 :: Export YAML File

Read the active YAML payload from Extensible Storage and write it to a
user-selected file. The exported text is byte-identical to what was last
stored, so an Import -> Export round-trip is lossless.
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


TITLE = "Export YAML File (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()

    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    payload = storage.read_payload(doc)
    if payload is None:
        forms.alert(
            "No active YAML in this project.\n\n"
            "Use 'Import YAML File' first.",
            title=TITLE,
        )
        return

    default_name = "equipment_profiles.yaml"
    if payload.get("source_path"):
        default_name = (
            os.path.basename(payload["source_path"]) or default_name
        )

    save_path = forms.save_file(
        file_ext="yaml",
        title=TITLE,
        default_name=default_name,
    )
    if not save_path:
        return

    try:
        result = active_yaml.export_yaml_file(doc, save_path)
    except storage.StorageError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    except (IOError, OSError) as exc:
        forms.alert(
            "Failed to write file:\n\n{}".format(exc),
            title=TITLE,
        )
        return
    except Exception as exc:
        forms.alert(
            "Unexpected error during export:\n\n{}".format(exc),
            title=TITLE,
            exitscript=False,
        )
        raise

    output.print_md(
        "**Export succeeded**\n\n"
        "- Wrote: `{}`\n"
        "- Bytes: `{}`\n"
        "- Schema version: `{}`\n"
        "- Original source: `{}`\n"
        "- Last modified (UTC): `{}`\n".format(
            result["save_path"],
            result["byte_count"],
            result.get("schema_version"),
            result.get("source_path") or "(none)",
            result.get("last_modified_utc") or "(unset)",
        )
    )


if __name__ == "__main__":
    main()
