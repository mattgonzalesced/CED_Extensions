#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Unmerge Profile"""

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
import truth_groups

TITLE = "Unmerge Profile (MEPRFP 2.0)"


def _member_label(profile_data):
    profiles = profile_data.get("equipment_definitions") or []
    by_id = {
        p.get("id"): p for p in profiles
        if isinstance(p, dict) and p.get("id")
    }

    def label(member):
        mname = member.get("name") or "(unnamed)"
        mid = member.get("id") or "?"
        sid = truth_groups.truth_source_id(member) or "?"
        source = by_id.get(sid)
        if source is not None:
            sname = source.get("name") or "(unnamed)"
            return "{}  ({})    <-    {}  ({})".format(mname, mid, sname, sid)
        sname = truth_groups.truth_source_name(member) or "<source missing>"
        return "{}  ({})    <-    {}  ({})".format(mname, mid, sname, sid)

    return label


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc)
    profiles = profile_data.get("equipment_definitions") or []
    members = [p for p in profiles if isinstance(p, dict) and truth_groups.is_group_member(p)]
    if not members:
        forms.alert(
            "No merged profiles in the active store — nothing to unmerge.",
            title=TITLE,
        )
        return

    chosen = wpf_dialogs.pick_from_list(
        members,
        title=TITLE,
        prompt="Pick a merged profile to detach (member <- source):",
        display_func=_member_label(profile_data),
    )
    if chosen is None:
        return

    if not forms.confirm(
        "Detach\n"
        "    {}  ({})\n"
        "from group\n"
        "    {}  ({})?\n\n"
        "The profile's structural content stays as-is. Only the "
        "ced_truth_source_id / ced_truth_source_name tags are cleared.".format(
            chosen.get("name") or "?", chosen.get("id") or "?",
            truth_groups.truth_source_name(chosen) or "?",
            truth_groups.truth_source_id(chosen) or "?",
        ),
        title=TITLE,
    ):
        return

    try:
        merge_workflow.unmerge(profile_data, chosen)
    except merge_workflow.MergeError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    with revit.Transaction("Unmerge Profile (MEPRFP 2.0)", doc=doc):
        active_yaml.save_active_data(doc, profile_data, action="Unmerge Profile")

    output.print_md(
        "**Unmerged**\n\n"
        "- Profile: `{}` (`{}`)\n".format(
            chosen.get("name") or "?", chosen.get("id") or "?"
        )
    )


if __name__ == "__main__":
    main()
