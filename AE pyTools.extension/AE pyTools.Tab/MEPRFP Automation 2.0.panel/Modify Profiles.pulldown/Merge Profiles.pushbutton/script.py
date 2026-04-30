#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Merge Profiles (alias model)"""

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
import capture
import merge_workflow
import selection

TITLE = "Merge Profiles (MEPRFP 2.0)"

_MODE_HOST = "Pick host element(s)"
_MODE_LINKED = "Pick linked element(s)"
_MODE_PROFILE = "Pick existing profile(s)"
_MODE_TYPE = "Type alias manually"


def _profile_label(p):
    return "{}  ({})".format(p.get("name") or "(unnamed)", p.get("id") or "?")


def _source_label(profile_data):
    """Display label for source candidates — shows current alias count."""
    def label(p):
        n_aliases = len(merge_workflow.aliases(p))
        suffix = "  [{} alias(es)]".format(n_aliases) if n_aliases else ""
        return "{}{}".format(_profile_label(p), suffix)
    return label


def _maybe_run_legacy_migration(doc, profile_data, output):
    """First-run prompt: if any legacy ced_truth_source_id markers exist,
    offer to migrate them to the new alias model."""
    legacy = merge_workflow.collect_legacy_members(profile_data)
    if not legacy:
        return False
    n = len(legacy)
    if not forms.confirm(
        "Found {} legacy merged profile(s) using the old data-duplication model.\n\n"
        "Migrate to the new alias-only model now?\n\n"
        "  - Each member's name will be added as an alias on its source.\n"
        "  - Members will keep their data but no longer be flagged as merged.\n"
        "  - You'll be asked next whether to delete the now-redundant member profiles.\n\n"
        "Recommended.".format(n),
        title=TITLE,
    ):
        return False
    report = merge_workflow.migrate_legacy_members(profile_data)
    output.print_md(
        "**Legacy merge migration**\n\n"
        "- Aliases added:        {}\n"
        "- Members reset:        {}\n"
        "- Unresolved sources:   {}\n".format(
            report.aliases_added,
            report.members_cleared,
            len(report.unresolved_members),
        )
    )
    # Offer to delete the migrated member records.
    delete_candidates = [m for m in legacy if m.get("id")]
    if delete_candidates and forms.confirm(
        "Delete the {} migrated profile records now?\n\n"
        "Their structural data was a copy of the source's. After "
        "deletion, only the source profile + its aliases remain. "
        "Choose No to keep them as standalone profiles.".format(len(delete_candidates)),
        title=TITLE,
    ):
        ids = [m.get("id") for m in delete_candidates]
        removed = merge_workflow.delete_profiles_by_id(profile_data, ids)
        output.print_md(
            "Deleted {} legacy member profile(s).".format(removed)
        )
    with revit.Transaction("Migrate legacy merges (MEPRFP 2.0)", doc=doc):
        active_yaml.save_active_data(doc, profile_data, action="Migrate legacy merges")
    return True


def _add_aliases_via_host(uidoc, source):
    forms.alert(
        "Pick host-model element(s). Each picked element's "
        "'Family : Type' label will be added as an alias.\n\n"
        "Click 'Finish' in the ribbon when done; press Esc to cancel.",
        title=TITLE,
    )
    try:
        refs = selection.pick_children(uidoc, "Pick host element(s) for alias", from_linked=False)
    except selection.SelectionCancelled:
        return [], []
    return _refs_to_aliases(refs, source)


def _add_aliases_via_linked(uidoc, source):
    forms.alert(
        "Pick LINKED-model element(s). Each picked element's "
        "'Family : Type' label will be added as an alias.\n\n"
        "Click 'Finish' in the ribbon when done; press Esc to cancel.",
        title=TITLE,
    )
    try:
        refs = selection.pick_children(uidoc, "Pick linked element(s) for alias", from_linked=True)
    except selection.SelectionCancelled:
        return [], []
    return _refs_to_aliases(refs, source)


def _refs_to_aliases(refs, source):
    added, skipped = [], []
    for r in refs or []:
        label = capture.element_label(r.element)
        if not label:
            skipped.append("(unlabelled element)")
            continue
        if merge_workflow.add_alias(source, label):
            added.append(label)
        else:
            skipped.append(label + " (already an alias)")
    return added, skipped


def _add_aliases_via_existing_profiles(profile_data, source):
    """Multi-select existing profiles; each picked profile's name becomes
    an alias on the source. Picked profiles can optionally be deleted."""
    candidates = [
        p for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict) and p is not source
    ]
    if not candidates:
        forms.alert("No other profiles to alias.", title=TITLE)
        return [], [], []
    chosen = wpf_dialogs.multi_select_from_list(
        candidates,
        title=TITLE,
        prompt="Check profiles whose names should become aliases on:\n    {}".format(
            _profile_label(source)
        ),
        display_func=_profile_label,
    )
    if not chosen:
        return [], [], []
    added, skipped = [], []
    for p in chosen:
        name = (p.get("name") or "").strip()
        if not name:
            skipped.append("(unnamed profile)")
            continue
        if merge_workflow.add_alias(source, name):
            added.append(name)
        else:
            skipped.append(name + " (already an alias)")
    chosen_ids = [p.get("id") for p in chosen if p.get("id")]
    return added, skipped, chosen_ids


def _add_alias_via_typed_string(source):
    text = wpf_dialogs.prompt_for_string(
        "Type the alias to add to:\n    {}".format(_profile_label(source)),
        title=TITLE,
        default="",
    )
    if not text:
        return [], []
    if merge_workflow.add_alias(source, text):
        return [text], []
    return [], [text + " (already an alias)"]


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert(
            "No profiles in the active store. "
            "Use 'New Profile' or 'Import YAML File' first.",
            title=TITLE,
        )
        return

    # First-run migration prompt.
    _maybe_run_legacy_migration(doc, profile_data, output)

    # Pick source.
    profiles = profile_data.get("equipment_definitions") or []
    if not profiles:
        forms.alert("No profiles available after migration.", title=TITLE)
        return
    source = wpf_dialogs.pick_from_list(
        profiles,
        title=TITLE,
        prompt=(
            "Pick the MASTER profile (the one that survives the merge).\n\n"
            "Other profiles' names will become aliases on this one. Future "
            "placement runs that match those alias names will resolve to "
            "this master profile. Pick the most complete / canonical one — "
            "the others can optionally be deleted on the next prompt."
        ),
        display_func=_source_label(profile_data),
    )
    if source is None:
        return

    # Pick mode.
    mode = wpf_dialogs.pick_from_list(
        [_MODE_HOST, _MODE_LINKED, _MODE_PROFILE, _MODE_TYPE],
        title=TITLE,
        prompt="How do you want to add aliases to:\n    {}".format(_profile_label(source)),
    )
    if mode is None:
        return

    added, skipped, picked_profile_ids = [], [], []
    if mode == _MODE_HOST:
        added, skipped = _add_aliases_via_host(uidoc, source)
    elif mode == _MODE_LINKED:
        added, skipped = _add_aliases_via_linked(uidoc, source)
    elif mode == _MODE_PROFILE:
        added, skipped, picked_profile_ids = _add_aliases_via_existing_profiles(
            profile_data, source
        )
    elif mode == _MODE_TYPE:
        added, skipped = _add_alias_via_typed_string(source)

    if not added and not skipped:
        return

    # Optionally delete the picked-profile records (since they're now
    # redundant aliases on the source).
    deleted_count = 0
    if mode == _MODE_PROFILE and picked_profile_ids and added:
        if forms.confirm(
            "Delete the {} picked profile record(s) now that their names "
            "are aliases on the source? Their data will be lost.\n\n"
            "Choose No if you want to keep them as standalone profiles.".format(
                len(picked_profile_ids)
            ),
            title=TITLE,
        ):
            deleted_count = merge_workflow.delete_profiles_by_id(
                profile_data, picked_profile_ids
            )

    # Save.
    with revit.Transaction("Merge Profiles (MEPRFP 2.0)", doc=doc):
        active_yaml.save_active_data(doc, profile_data, action="Merge Profiles")

    output.print_md(
        "**Aliases added to {!r}**\n\n"
        "- Added:    {} ({} skipped)\n"
        "- Deleted picked-profile records:  {}\n".format(
            source.get("name") or "?",
            len(added), len(skipped), deleted_count,
        )
    )
    if added:
        output.print_md(
            "\n**Added aliases:**\n"
            + "\n".join("- `{}`".format(a) for a in added)
        )
    if skipped:
        output.print_md(
            "\n**Skipped:**\n"
            + "\n".join("- {}".format(s) for s in skipped)
        )


if __name__ == "__main__":
    main()
