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


_ALL_FAMILIES_KEY = "__ALL__"


def _collect_family_groups(profiles):
    """Group profiles by ``parent_filter.family_name_pattern`` (case-fold).

    Returns a list of dicts ``{name, profiles}`` sorted by name. Profiles
    with no family pattern fall under ``(no parent_filter family)`` so
    they're still pickable.
    """
    groups = {}
    for p in profiles:
        if not isinstance(p, dict):
            continue
        pf = p.get("parent_filter") or {}
        fam = (pf.get("family_name_pattern") or "").strip()
        key = fam.lower() or "__empty__"
        display = fam or "(no parent_filter family)"
        bucket = groups.setdefault(key, {"name": display, "profiles": []})
        bucket["profiles"].append(p)
    return sorted(groups.values(), key=lambda g: g["name"].lower())


def _pick_family_group(profiles):
    """Two-step disambiguation: pick a family group, then pick within it.

    Returns the list of profiles to pick the master from, or None if
    cancelled. Adds an "(All profiles)" escape hatch as the first row
    so cross-family merges are still possible.
    """
    families = _collect_family_groups(profiles)
    if not families:
        return None
    options = [
        {"name": "(All profiles — show every profile)",
         "profiles": list(profiles),
         "key": _ALL_FAMILIES_KEY}
    ] + families
    chosen = wpf_dialogs.pick_from_list(
        options,
        title=TITLE,
        prompt=(
            "Pick the family group to merge within. Profiles sharing the "
            "same parent_filter.family_name_pattern are bucketed together "
            "so you can scope the master pick to (e.g.) just the stinger "
            "carts or just the checkstands. Pick (All profiles) to merge "
            "across families."
        ),
        display_func=lambda g: "{}  ({} profile(s))".format(
            g["name"], len(g["profiles"])
        ),
    )
    if chosen is None:
        return None
    return chosen["profiles"]


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
    # Default to siblings — profiles sharing the master's
    # parent_filter.family_name_pattern. Falls back to "all profiles" if
    # that yields nothing useful (e.g. master has no family pattern).
    source_pf = (source.get("parent_filter") or {}) if isinstance(source, dict) else {}
    source_fam_key = (source_pf.get("family_name_pattern") or "").strip().lower()

    all_others = [
        p for p in profile_data.get("equipment_definitions") or []
        if isinstance(p, dict) and p is not source
    ]
    if not all_others:
        forms.alert("No other profiles to alias.", title=TITLE)
        return [], [], []

    siblings = []
    if source_fam_key:
        for p in all_others:
            pf = p.get("parent_filter") or {}
            fam = (pf.get("family_name_pattern") or "").strip().lower()
            if fam and fam == source_fam_key:
                siblings.append(p)

    if siblings:
        candidates = siblings
        prompt_extra = (
            "\n\nShowing only profiles in family '{}'. Cancel and use a "
            "different mode if you need to alias outside this family."
        ).format(source_pf.get("family_name_pattern") or "?")
    else:
        candidates = all_others
        prompt_extra = (
            "\n\nNo siblings found in the master's family — showing every "
            "other profile."
        )

    chosen = wpf_dialogs.multi_select_from_list(
        candidates,
        title=TITLE,
        prompt=(
            "Check profiles whose names should become aliases on:\n    {}{}"
        ).format(_profile_label(source), prompt_extra),
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

    # Pick the family group first, then the master within it.
    profiles = profile_data.get("equipment_definitions") or []
    if not profiles:
        forms.alert("No profiles available after migration.", title=TITLE)
        return
    family_profiles = _pick_family_group(profiles)
    if not family_profiles:
        return
    source = wpf_dialogs.pick_from_list(
        family_profiles,
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
