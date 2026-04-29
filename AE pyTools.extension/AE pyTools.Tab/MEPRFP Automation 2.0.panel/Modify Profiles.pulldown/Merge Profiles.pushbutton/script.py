#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Merge Profiles"""

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

TITLE = "Merge Profiles (MEPRFP 2.0)"


def _source_label(profile_data):
    """Return a function for displaying source candidates."""
    members_by_source = truth_groups.group_members_by_source(
        profile_data.get("equipment_definitions") or []
    )

    def label(p):
        sid = p.get("id") or "?"
        n_members = len(members_by_source.get(sid, []))
        suffix = "  [{} member(s)]".format(n_members) if n_members else ""
        return "{}  ({}){}".format(p.get("name") or "(unnamed)", sid, suffix)

    return label


def _target_label(p):
    return "{}  ({})".format(p.get("name") or "(unnamed)", p.get("id") or "?")


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
            "No profiles in the active store. Use 'New Profile' or 'Import YAML File' first.",
            title=TITLE,
        )
        return

    # 1. Pick source.
    sources = merge_workflow.eligible_sources(profile_data)
    if not sources:
        forms.alert(
            "No profiles eligible to be a source. Every profile is "
            "currently a member of some other group; unmerge one first.",
            title=TITLE,
        )
        return
    source = wpf_dialogs.pick_from_list(
        sources,
        title=TITLE,
        prompt="Pick the SOURCE profile (truth):",
        display_func=_source_label(profile_data),
    )
    if source is None:
        return

    # 2. Pick targets.
    targets_pool = merge_workflow.eligible_targets(profile_data, source)
    if not targets_pool:
        forms.alert(
            "No eligible targets for source {!r}. Targets must not be "
            "the source itself, must not already be members of any "
            "group, and must not themselves be sources for other groups.".format(
                source.get("name") or source.get("id")
            ),
            title=TITLE,
        )
        return
    chosen = wpf_dialogs.multi_select_from_list(
        targets_pool,
        title=TITLE,
        prompt="Check target profiles to merge into {!r} (source):".format(
            source.get("name") or source.get("id") or "?"
        ),
        display_func=_target_label,
    )
    if chosen is None or len(chosen) == 0:
        return

    # 3. Confirm.
    target_summary = "\n".join(
        "  - {}  ({})".format(t.get("name") or "(unnamed)", t.get("id") or "?")
        for t in chosen
    )
    if not forms.confirm(
        "About to merge the following {} target(s) into source\n"
        "    {}  ({})\n\n"
        "Targets:\n{}\n\n"
        "Each target keeps its own id and name but inherits the source's "
        "structural content (parent_filter, linked_sets, equipment_properties, flags). "
        "SET / LED / ANN ids in the copy will be renumbered.\n\n"
        "Proceed?".format(
            len(chosen),
            source.get("name") or "(unnamed)",
            source.get("id") or "?",
            target_summary,
        ),
        title=TITLE,
    ):
        return

    # 4. Apply.
    succeeded, failed = merge_workflow.merge_many(profile_data, source, chosen)

    if succeeded:
        with revit.Transaction("Merge Profiles (MEPRFP 2.0)", doc=doc):
            active_yaml.save_active_data(doc, profile_data, action="Merge Profiles")

    output.print_md(
        "**Merge complete**\n\n"
        "- Source: `{}` (`{}`)\n"
        "- Targets succeeded: {}\n"
        "- Targets failed: {}\n".format(
            source.get("name") or "?", source.get("id") or "?",
            len(succeeded), len(failed),
        )
    )
    if failed:
        output.print_md(
            "\n**Failures:**\n"
            + "\n".join(
                "- {}  ({}): {}".format(
                    t.get("name") or "?", t.get("id") or "?", reason
                )
                for t, reason in failed
            )
        )


if __name__ == "__main__":
    main()
