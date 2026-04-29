#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Bulk Merge Profiles"""

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
import merge_workflow

TITLE = "Bulk Merge Profiles (MEPRFP 2.0)"

_HELP = (
    "CSV format:\n\n"
    "  - One header row with 'source' and 'target' columns.\n"
    "    Aliases also accepted: source_id, source_name, target_id, target_name.\n"
    "  - One merge per data row.\n"
    "  - Each cell value is matched against profile ids first, then names.\n\n"
    "Example:\n\n"
    "  source,target\n"
    "  EQ-001,EQ-002\n"
    "  Foo Bar : Master,Foo Bar : V1\n"
    "  Foo Bar : Master,Foo Bar : V2\n\n"
    "Rows that fail eligibility checks (already merged, source-of-others, "
    "self-merge) are skipped and reported, while remaining rows still apply."
)


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert(
            "No profiles in the active store.",
            title=TITLE,
        )
        return

    if not forms.confirm(_HELP + "\n\nProceed to pick the CSV?", title=TITLE):
        return

    csv_path = forms.pick_file(
        file_ext="csv",
        title="Pick bulk-merge CSV",
    )
    if not csv_path:
        return

    try:
        results = merge_workflow.bulk_merge_from_csv(profile_data, csv_path)
    except merge_workflow.MergeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    succeeded = [r for r in results if r.ok]
    failed = [r for r in results if not r.ok]

    if succeeded:
        with revit.Transaction("Bulk Merge Profiles (MEPRFP 2.0)", doc=doc):
            active_yaml.save_active_data(doc, profile_data, action="Bulk Merge Profiles")

    output.print_md(
        "**Bulk merge complete**\n\n"
        "- CSV: `{}`\n"
        "- Rows merged: {}\n"
        "- Rows failed:  {}\n".format(
            csv_path, len(succeeded), len(failed),
        )
    )
    if succeeded:
        output.print_md(
            "\n**Merged:**\n"
            + "\n".join(
                "- row {}: `{}` <- `{}`".format(
                    r.row_number, r.target_label, r.source_label
                )
                for r in succeeded
            )
        )
    if failed:
        output.print_md(
            "\n**Failures:**\n"
            + "\n".join(
                "- row {}: {}".format(r.row_number, r.message)
                for r in failed
            )
        )


if __name__ == "__main__":
    main()
