#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Unmerge Profile (alias removal)"""

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
import wpf_dialogs
import active_yaml
import merge_workflow

TITLE = "Unmerge Profile (MEPRFP 2.0)"


def _entry_label(source, alias):
    return "{}  ({})    ->    {}".format(
        source.get("name") or "(unnamed)",
        source.get("id") or "?",
        alias,
    )


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc)
    entries = merge_workflow.all_alias_entries(profile_data)
    if not entries:
        forms.alert(
            "No aliases in the active store — nothing to unmerge.",
            title=TITLE,
        )
        return

    chosen = wpf_dialogs.pick_from_list(
        entries,
        title=TITLE,
        prompt="Pick an alias to remove (source -> alias):",
        display_func=lambda pair: _entry_label(pair[0], pair[1]),
    )
    if chosen is None:
        return
    source, alias = chosen

    if not forms.confirm(
        "Remove alias\n    {}\nfrom source\n    {}?".format(
            alias,
            "{}  ({})".format(source.get("name") or "?", source.get("id") or "?"),
        ),
        title=TITLE,
    ):
        return

    if not merge_workflow.remove_alias(source, alias):
        forms.alert("Alias not found on the source — nothing changed.", title=TITLE)
        return

    with revit.Transaction("Unmerge Profile (MEPRFP 2.0)", doc=doc):
        active_yaml.save_active_data(doc, profile_data, action="Unmerge Profile")

    output.print_md(
        "**Alias removed**\n\n"
        "- Source: `{}` (`{}`)\n"
        "- Alias:  `{}`\n".format(
            source.get("name") or "?", source.get("id") or "?",
            alias,
        )
    )


if __name__ == "__main__":
    main()
