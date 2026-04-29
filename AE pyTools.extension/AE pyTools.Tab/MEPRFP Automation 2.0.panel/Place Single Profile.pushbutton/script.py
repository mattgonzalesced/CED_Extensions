#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place Single Profile

Pick one profile and place it at one or more points clicked in the model.
Reuses ``placement.execute_placement`` — each clicked point becomes a
``placement.Target`` (source = SOURCE_PICKED_POINT) and is matched
directly to the selected profile, so all the engine's offset / rotation
/ Element_Linker handling applies automatically.
"""

import os
import sys

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

from Autodesk.Revit.Exceptions import OperationCanceledException

import forms_compat as forms
import wpf_dialogs
import active_yaml
import placement
import shared_params

TITLE = "Place Single Profile (MEPRFP 2.0)"


def _ensure_param_bound(doc):
    if shared_params.is_element_linker_bound(doc):
        return True
    if not forms.confirm(
        "Bind the MEPRFP 2.0 Element_Linker shared parameter now?",
        title=TITLE,
    ):
        return False
    try:
        with revit.Transaction("Bind MEPRFP Element_Linker", doc=doc):
            shared_params.ensure_element_linker_bound(doc)
    except shared_params.SharedParamError as exc:
        forms.alert(
            "Failed to bind shared parameter:\n\n{}".format(exc),
            title=TITLE,
        )
        return False
    return True


def _is_independent(profile):
    return bool(profile.get("allow_parentless"))


def _profile_label(profile):
    name = profile.get("name") or "(unnamed)"
    pid = profile.get("id") or "?"
    suffix = "  [independent]" if _is_independent(profile) else ""
    return "{}  ({}){}".format(name, pid, suffix)


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return
    if not _ensure_param_bound(doc):
        return

    profile_data = active_yaml.load_active_data(doc)
    profiles = profile_data.get("equipment_definitions") or []
    if not profiles:
        forms.alert(
            "No profiles in the active store. "
            "Use 'New Profile' / 'Create Independent Profile' / 'Import YAML File' first.",
            title=TITLE,
        )
        return

    independent_only = forms.confirm(
        "Show only INDEPENDENT profiles?\n\n"
        "Yes  -> only profiles with no parent-anchor requirement.\n"
        "No   -> show all profiles.",
        title=TITLE,
    )
    candidates = (
        [p for p in profiles if _is_independent(p)]
        if independent_only else profiles
    )
    if not candidates:
        forms.alert(
            "No independent profiles in the active store. "
            "Create one via 'Create Independent Profile' first, or pick 'No' "
            "on the previous prompt to see all profiles.",
            title=TITLE,
        )
        return

    selected = wpf_dialogs.pick_from_list(
        candidates,
        title=TITLE,
        prompt="Pick the profile to place:",
        display_func=_profile_label,
    )
    if selected is None:
        return

    multiple = forms.confirm(
        "Place multiple instances?\n\n"
        "Yes -> click repeatedly; press Esc when finished.\n"
        "No  -> place once and stop.",
        title=TITLE,
    )

    profile_name = selected.get("name") or "?"
    forms.alert(
        "Click in the model to place {!r}.\n\n{}".format(
            profile_name,
            "Click each location, then press Esc to finish."
            if multiple else "Click once to place.",
        ),
        title=TITLE,
    )

    placed_fixture_total = 0
    placed_annotation_total = 0
    point_count = 0
    warnings = []
    options = placement.PlacementOptions(
        skip_already_placed=False,
        transaction_action="Place Single Profile (MEPRFP 2.0)",
    )

    while True:
        try:
            point = uidoc.Selection.PickPoint(
                "Place {!r}{} - Esc to finish".format(
                    profile_name,
                    "  (#{})".format(point_count + 1) if multiple else "",
                )
            )
        except OperationCanceledException:
            break
        if point is None:
            break

        target = placement.Target(
            source=placement.SOURCE_PICKED_POINT,
            name=profile_name,
            world_pt=(point.X, point.Y, point.Z),
            rotation_deg=0.0,
        )
        match = placement.Match(target, selected)
        with revit.Transaction("Place Single Profile (MEPRFP 2.0)", doc=doc):
            result = placement.execute_placement(doc, [match], options)
        placed_fixture_total += result.placed_fixture_count
        placed_annotation_total += result.placed_annotation_count
        warnings.extend(result.warnings)
        point_count += 1
        if not multiple:
            break

    output.print_md(
        "**Place Single Profile complete**\n\n"
        "- Profile: `{}`\n"
        "- Click points: {}\n"
        "- Fixtures placed: {}\n"
        "- Annotations placed: {}\n".format(
            profile_name, point_count,
            placed_fixture_total, placed_annotation_total,
        )
    )
    if warnings:
        output.print_md(
            "\n**Warnings:**\n"
            + "\n".join("- {}".format(w) for w in warnings[:50])
        )


if __name__ == "__main__":
    main()
