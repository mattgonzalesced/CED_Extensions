# -*- coding: utf-8 -*-
"""
Pre-modal door picker for Place Space Elements.

A WPF modal dialog blocks Revit input, so ``Selection.PickObject``
can't run while the placement preview is open. This module is invoked
by the Place Space Elements pushbutton script *before* the modal
opens: it scans for spaces that have multiple doors AND at least one
door-dependent LED, prompts the user to click a door in the model
for each, and returns a ``{space_element_id: door_anchor}`` dict
that can be threaded into ``SpacePlacementRun(door_choices=...)``.

Implementation notes:

  * ``ObjectType.LinkedElement`` lets the user click a door inside
    a ``RevitLinkInstance`` (the typical MEP setup, where doors
    live in the linked architectural model). Host doors are also
    handled — the same Reference flow falls through cleanly when
    ``LinkedElementId`` is invalid.
  * ``ISelectionFilter`` accepts only Door-category elements, so
    the user can't accidentally click something else.
  * Cancellation (Esc) is treated as "skip this space, fall back
    to the first door" rather than aborting the whole script —
    feels less destructive when the user has many spaces to pick
    through.
"""

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    ElementId,
    RevitLinkInstance,
)
from Autodesk.Revit.UI import (  # noqa: E402
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
)
from Autodesk.Revit.UI.Selection import (  # noqa: E402
    ISelectionFilter,
    ObjectType,
)
from Autodesk.Revit.Exceptions import (  # noqa: E402
    OperationCanceledException,
)

import space_placement as _placement
import space_placement_workflow as _spw
import space_profile_model as _profile_model
import space_door_filter as _door_filter


# ``_DoorOnlyFilter`` lives in ``space_door_filter`` (the only module
# in this subsystem that's excluded from ``_dev_reload.purge()``), so
# this module can be edited and hot-reloaded without re-registering
# the ISelectionFilter CLR type.
_DoorOnlyFilter = _door_filter.DoorOnlyFilter


def _is_door(elem):
    """True iff ``elem`` is a Door-category element."""
    if elem is None:
        return False
    try:
        cat = elem.Category
    except Exception:
        return False
    if cat is None:
        return False
    try:
        cat_id = cat.Id
        cat_int = getattr(cat_id, "Value", None) or getattr(
            cat_id, "IntegerValue", None,
        )
        return int(cat_int) == int(BuiltInCategory.OST_Doors)
    except Exception:
        return False


# ---------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------

def pre_pick_doors(uidoc, doc, profile_data, output=None, **_unused):
    """Walk the placement run, prompting the user to click a reference
    door for every multi-door space whose plans need one.

    Returns ``{space_element_id: (origin_xy, inward_xy)}``. Spaces
    skipped (cancelled) get no entry — the workflow's first-door
    fallback applies. ``uidoc`` is required; if absent, returns ``{}``
    so the caller can still proceed with default first-door picks.

    ``output`` is the pyrevit script output handle. When provided,
    diagnostic info is printed (helps the user see why no prompt
    appeared if they expected one).
    """
    if uidoc is None or doc is None:
        if output is not None:
            output.print_md("Door pre-pick: missing uidoc or doc — skipped.")
        return {}
    if not profile_data:
        return {}

    # Dry-run collect to identify ambiguous spaces. Use a no-op picker
    # so the run doesn't try to prompt for anything; we just want the
    # ``spaces_with_multiple_doors`` list populated.
    run = _spw.SpacePlacementRun(doc=doc, profile_data=profile_data)
    run.collect()

    if output is not None:
        _print_pick_diagnostics(output, doc, profile_data, run)

    if not run.spaces_with_multiple_doors:
        return {}

    # Filter to spaces that ACTUALLY have door-dependent LEDs (the
    # multi-door list includes spaces just for diagnostics).
    profiles = _profile_model.wrap_profiles(
        profile_data.get("space_profiles") or []
    )
    needs_pick = []
    from space_workflow import load_classifications_indexed
    classifications = load_classifications_indexed(doc)
    for space, doors in run.spaces_with_multiple_doors:
        bucket_ids = classifications.get(space.element_id) or []
        matching = _profile_model.profiles_for_buckets(profiles, bucket_ids)
        if _any_door_dependent(matching):
            needs_pick.append((space, doors))

    if not needs_pick:
        return {}

    # Stable ordering by space label so the prompts appear in a
    # predictable sequence rather than collection order.
    def _space_sort_key(entry):
        s, _doors = entry
        return ((s.number or "").lower(), (s.name or "").lower(),
                s.element_id or 0)
    needs_pick.sort(key=_space_sort_key)

    # Up-front explainer — the architectural model is usually linked,
    # and the user has to have "Select Links" enabled (the lock icon
    # at the bottom-right of the Revit window) for PickObject to let
    # them click anything inside a link. Surface that requirement
    # explicitly so users don't get stuck unable to click.
    if not _show_intro_dialog(needs_pick):
        return {}  # user cancelled the explainer

    selection = uidoc.Selection
    door_filter = _DoorOnlyFilter(doc)
    choices = {}

    for idx, (space, doors) in enumerate(needs_pick, start=1):
        space_label = "{} {}".format(
            space.number or "", space.name or "",
        ).strip() or "(unnamed)"
        prompt = "[{}/{}] Pick the door for the {} space".format(
            idx, len(needs_pick), space_label,
        )

        # Show a modal alert that names the target space, so the user
        # is told WHICH space they're picking for before the status-
        # bar prompt fires. Pick = proceed with PickObject. Skip =
        # auto-pick the space's first door. Cancel = abort the whole
        # picking sequence (uses first-door fallback for the rest).
        choice = _show_per_space_alert(idx, len(needs_pick), space_label,
                                       len(doors))
        if choice == "cancel":
            break
        if choice == "skip":
            continue

        # Try linked-element pick first — typical MEP setup. If the
        # user has Select Links off OR clicks a host-doc door, fall
        # back to plain Element pick.
        anchor = _try_pick(selection, ObjectType.LinkedElement,
                           door_filter, prompt, doc)
        if anchor is None:
            anchor = _try_pick(selection, ObjectType.Element,
                               door_filter, prompt, doc)
        if anchor is not None:
            choices[space.element_id] = anchor
    return choices


def _show_per_space_alert(idx, total, space_label, n_doors):
    """Modal TaskDialog announcing which space the next pick is for.

    Returns one of:
      * ``"pick"``   — user clicked OK; caller fires PickObject.
      * ``"skip"``   — user clicked the "Skip this space" button; the
                       caller falls back to the first-door default
                       for this space and moves to the next prompt.
      * ``"cancel"`` — user closed the dialog; caller aborts the
                       remaining picks and uses first-door defaults
                       for every remaining space.
    """
    td = TaskDialog("Pick a door — {}/{}".format(idx, total))
    td.MainInstruction = "Pick the door for the {} space".format(space_label)
    td.MainContent = (
        "This space has {} detected door(s). After you click 'Pick door', "
        "click the reference door in the model. The chosen door anchors "
        "every door-relative LED in this space.\n\n"
        "Select Links must be ON to click a door inside a linked "
        "architectural model.".format(n_doors)
    )
    # CommandLinks would be nicer (Pick / Skip / Cancel as big
    # tappable buttons) but the API surface for those is fiddly
    # under pythonnet; the basic Yes/No/Cancel set is reliable and
    # carries the same three intents.
    td.CommonButtons = (
        TaskDialogCommonButtons.Yes
        | TaskDialogCommonButtons.No
        | TaskDialogCommonButtons.Cancel
    )
    td.DefaultButton = TaskDialogResult.Yes
    # Rename the button captions via verification text / footer so the
    # user knows what each does. The CommonButtons captions themselves
    # are locale-fixed ("Yes" / "No" / "Cancel"), so we explain them
    # inline in MainContent for clarity.
    td.FooterText = (
        "Yes = Pick door  •  No = Skip (use first door)  •  Cancel = "
        "Stop picking and use defaults for the rest"
    )
    result = td.Show()
    if result == TaskDialogResult.Yes:
        return "pick"
    if result == TaskDialogResult.No:
        return "skip"
    return "cancel"


def _show_intro_dialog(needs_pick):
    """Up-front task dialog explaining what's about to happen.

    Returns True if the user wants to proceed with picking, False
    otherwise (script falls back to first-door defaults for every
    multi-door space).
    """
    n = len(needs_pick)
    space_lines = []
    for space, doors in needs_pick[:8]:  # cap the list
        label = "{} {}".format(
            space.number or "", space.name or "",
        ).strip() or "(unnamed)"
        space_lines.append("  - {}  ({} doors)".format(label, len(doors)))
    if n > 8:
        space_lines.append("  - ... and {} more".format(n - 8))

    td = TaskDialog("Pick reference doors")
    td.MainInstruction = (
        "{} space(s) have multiple doors and at least one "
        "door-relative LED. Pick a reference door for each."
    ).format(n)
    td.MainContent = (
        "After you click OK, you'll be prompted to pick a door in "
        "the active view, one space at a time. Each prompt names the "
        "space — e.g. 'Pick the door for the 10 Dairy Cooler space'. "
        "The chosen door is used as the reference for every "
        "door-relative anchor in that space (opposite / right / left "
        "wall, closest / furthest corner, and door_relative).\n\n"
        "Spaces requiring a pick:\n{}\n\n"
        "**Important:** if your architecture is in a linked model, "
        "you MUST have 'Select Links' enabled in Revit — that's "
        "the small lock icon at the bottom-right of the Revit "
        "window. With it OFF you won't be able to click any door "
        "inside the link.\n\n"
        "Press Esc on any individual prompt to skip that space "
        "(its first door is used as the default)."
    ).format("\n".join(space_lines))
    td.CommonButtons = (
        TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
    )
    td.DefaultButton = TaskDialogResult.Ok
    result = td.Show()
    return result == TaskDialogResult.Ok


def _print_pick_diagnostics(output, doc, profile_data, run):
    """Verbose summary of door discovery + which spaces need a pick.
    Surfaces the per-space data so a user who expected a prompt can
    see exactly why it was suppressed (no doors detected, no
    door-dependent LED, etc.)."""
    profiles = _profile_model.wrap_profiles(
        profile_data.get("space_profiles") or []
    )
    from space_workflow import (
        collect_spaces,
        load_classifications_indexed,
    )
    spaces = collect_spaces(doc)
    classifications = load_classifications_indexed(doc)

    n_classified = sum(1 for s in spaces
                       if s.element_id in classifications
                       and classifications[s.element_id])
    multi = run.spaces_with_multiple_doors
    needs = []
    rows = []
    for space in spaces:
        bucket_ids = classifications.get(space.element_id) or []
        if not bucket_ids:
            continue
        matching = _profile_model.profiles_for_buckets(profiles, bucket_ids)
        if not matching:
            continue
        # Door count comes from the geometry that the run already
        # built — but we don't store it on the run. Find from the
        # multi-door list when present, else count via build_space_geometry.
        n_doors = None
        for ms, doors in multi:
            if ms.element_id == space.element_id:
                n_doors = len(doors)
                break
        if n_doors is None:
            try:
                geom = _placement.build_space_geometry(doc, space.element)
                n_doors = len(geom.door_anchors) if geom else 0
            except Exception:
                n_doors = 0
        kinds = []
        for p in matching:
            for s in p.linked_sets:
                for led in s.leds:
                    kinds.append(led.placement_rule.kind)
        n_door_dep = sum(
            1 for k in kinds if _profile_model.is_door_dependent(k)
        )
        will_prompt = (n_doors > 1 and n_door_dep > 0)
        if will_prompt:
            needs.append(space)
        rows.append({
            "space": space,
            "n_doors": n_doors,
            "kinds": kinds,
            "n_door_dep": n_door_dep,
            "will_prompt": will_prompt,
        })

    output.print_md("**Door pre-pick analysis**\n")
    output.print_md(
        "- Classified spaces with matching profile(s): `{}`\n"
        "- Spaces with multiple doors detected: `{}`\n"
        "- Spaces that will prompt for a door pick: `{}`".format(
            n_classified, len(multi), len(needs),
        )
    )
    if rows:
        output.print_md("\n**Per-space breakdown:**")
        for r in rows:
            label = "{} {}".format(
                r["space"].number or "", r["space"].name or "",
            ).strip() or "(unnamed)"
            note = "WILL PROMPT" if r["will_prompt"] else "no prompt"
            output.print_md(
                "- `{}` (id {}) -- doors: `{}`, "
                "door-dependent LEDs: `{}/{}`, kinds: `{}` -- *{}*".format(
                    label, r["space"].element_id,
                    r["n_doors"], r["n_door_dep"], len(r["kinds"]),
                    r["kinds"], note,
                )
            )
    if not needs:
        output.print_md(
            "\n*No prompts will appear.* Common reasons:\n\n"
            "  1. No space has more than one detected door (check "
            "the per-space `doors` count above — if it's 0 or 1, "
            "the script can't prompt). Linked doors only count when "
            "the linked architectural model is present in the host "
            "doc as a RevitLinkInstance and the door's location is "
            "within ~1 ft of the space's boundary curve.\n"
            "  2. No matching profile has a door-dependent LED "
            "(`door_relative`, `wall_*_door`, `corner_*_door`). If "
            "your LED is on `center`, no door is needed.\n"
            "  3. The classified space's profile doesn't match any "
            "bucket the LEDs need."
        )


def _any_door_dependent(profiles):
    for profile in profiles or ():
        for s in profile.linked_sets:
            for led in s.leds:
                if _profile_model.is_door_dependent(led.placement_rule.kind):
                    return True
    return False


def _try_pick(selection, object_type, door_filter, prompt, doc):
    """One PickObject attempt. Returns the resolved anchor tuple or
    ``None`` on cancellation / failure."""
    try:
        ref = selection.PickObject(object_type, door_filter, prompt)
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    return _reference_to_anchor(doc, ref)


def _reference_to_anchor(doc, reference):
    """Resolve a PickObject Reference (host or linked door) to the
    ``(origin_xy, inward_xy)`` tuple shape the placement engine wants."""
    try:
        host_id = reference.ElementId
    except Exception:
        return None
    if host_id is None:
        return None

    # If the picked element lives inside a link, ``LinkedElementId``
    # is non-invalid and ``ElementId`` points at the RevitLinkInstance.
    linked_id = None
    try:
        linked_id = reference.LinkedElementId
    except Exception:
        linked_id = None

    is_linked = (
        linked_id is not None
        and linked_id != ElementId.InvalidElementId
    )

    if is_linked:
        link_inst = doc.GetElement(host_id)
        if link_inst is None or not isinstance(link_inst, RevitLinkInstance):
            return None
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            return None
        door = link_doc.GetElement(linked_id)
        try:
            transform = link_inst.GetTotalTransform()
        except Exception:
            transform = None
        return _placement.door_to_anchor(door, transform=transform)

    # Host-doc door.
    door = doc.GetElement(host_id)
    return _placement.door_to_anchor(door, transform=None)
