#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Capture Space Profile.

Step-by-step:
  1. User picks a placed Space in the active view.
  2. If the space has multiple detected doors, prompt for which door
     to use as the reference (single-door auto-resolves).
  3. User multi-picks child elements inside the space.
  4. ``space_capture_workflow.run_capture`` resolves each child to
     a door-relative wall + proportional position and builds a new
     ``space_profiles[*]`` entry.
  5. User confirms profile name + bucket; the result is persisted
     via ``active_yaml.save_active_data``.
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
import space_capture_workflow as _capture
import space_door_filter as _door_filter
import space_workflow as _space_workflow
import wpf_dialogs

from Autodesk.Revit.UI import (
    TaskDialog, TaskDialogCommonButtons, TaskDialogResult,
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException


TITLE = "Capture Space Profile (MEPRFP 2.0)"


def _alert(main_instruction, content="", title=None):
    """Modal TaskDialog popup announcing the next pick. Returns
    True when the user clicks OK, False on Cancel — caller stops
    the capture flow on False so the user can bail out cleanly.

    All Capture Space Profile prompts route through this so the user
    sees a proper popup instead of a status-bar hint that's easy to
    miss when the active view is large.
    """
    td = TaskDialog(title or TITLE)
    td.MainInstruction = main_instruction
    if content:
        td.MainContent = content
    td.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
    td.DefaultButton = TaskDialogResult.Ok
    return td.Show() == TaskDialogResult.Ok


# ---------------------------------------------------------------------
# Pick helpers
# ---------------------------------------------------------------------
# The ``SpaceOnlyFilter`` ISelectionFilter lives in
# ``space_door_filter`` (the only module in this subsystem excluded
# from ``_dev_reload.purge()``) so its CLR type stays registered
# across script reloads. Defining it at script top-level here would
# blow up with "Duplicate type name within an assembly" on every
# second run.

def _pick_space(uidoc, doc):
    """Prompt the user to click a placed Space. Returns the SpaceInfo
    wrapping the element, or None on cancel / failure."""
    if not _alert(
        "Pick the Space to capture",
        "Click the placed Space (not just a Room) in the active view. "
        "After you click OK, the cursor will switch to pick mode and "
        "the status bar will echo the prompt.",
    ):
        return None
    sel = uidoc.Selection
    space_filter = _door_filter.SpaceOnlyFilter()
    try:
        ref = sel.PickObject(
            ObjectType.Element, space_filter,
            "Pick the Space to capture",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    elem = doc.GetElement(ref.ElementId)
    if elem is None:
        return None
    # Wrap in a SpaceInfo so the workflow gets the same shape it'd
    # see from collect_spaces.
    eid = getattr(elem, "Id", None)
    return _space_workflow.SpaceInfo(
        element=elem,
        element_id=(
            getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
            if eid is not None else None
        ),
        unique_id=getattr(elem, "UniqueId", "") or "",
        name=getattr(elem, "Name", "") or "",
        number=str(getattr(elem, "Number", "") or ""),
    )


def _pick_reference_door(uidoc, doc, space, doors):
    """If multiple doors are present, prompt the user to click one.
    Returns the chosen ``(origin_xy, inward_xy)`` tuple, or the first
    door when there's only one. Esc returns None → caller falls back
    to the first door."""
    if not doors:
        return None
    if len(doors) == 1:
        return doors[0]
    space_label = "{} {}".format(
        space.number or "", space.name or "",
    ).strip() or "(unnamed)"
    if not _alert(
        "Pick the reference door for the {} space".format(space_label),
        "This space has {} doors. Click the door you want to use as "
        "the reference for wall labeling — the chosen door defines "
        "which wall counts as 'opposite_door', 'right_of_door', etc. "
        "Select Links must be ON to click a door inside a linked "
        "architectural model.".format(len(doors)),
    ):
        # User cancelled the door pick — fall back to the first door
        # rather than aborting the whole capture flow.
        return doors[0]
    sel = uidoc.Selection
    door_filter = _door_filter.DoorOnlyFilter(doc)
    prompt = "Pick the reference door for the {} space".format(space_label)
    for object_type in (ObjectType.LinkedElement, ObjectType.Element):
        try:
            ref = sel.PickObject(object_type, door_filter, prompt)
        except OperationCanceledException:
            return doors[0]
        except Exception:
            ref = None
        if ref is None:
            continue
        # Resolve to (origin_xy, inward_xy) via the same path the
        # placement-side picker uses.
        import space_door_picker as _picker
        anchor = _picker._reference_to_anchor(doc, ref)
        if anchor is not None:
            return anchor
    return doors[0]


def _pick_children(uidoc):
    """Multi-pick child elements in the HOST document.

    Captured children (receptacles, junction boxes, keynote symbols,
    text notes) live in the host doc, NOT in a linked architectural
    model — those would be impossible to write Element_Linker onto.
    So we open the pick in ``ObjectType.Element`` mode (host-only)
    rather than ``ObjectType.LinkedElement``. The workflow's
    ``_classify_child`` filters down to the categories we actually
    capture, so the pick mode itself is intentionally permissive.
    """
    if not _alert(
        "Pick the child elements inside this space",
        "Click each fixture, keynote symbol, or text note in the HOST "
        "document that should be captured against this space's walls. "
        "Click 'Finish' on the Revit options bar when done, or press "
        "Esc to cancel the capture entirely.\n\n"
        "Children must be host-doc elements — captured profile data "
        "stamps Element_Linker onto each, which isn't possible on "
        "elements that live inside a linked model.",
    ):
        return []
    sel = uidoc.Selection
    try:
        refs = sel.PickObjects(
            ObjectType.Element,
            "Pick the children inside this space (host doc; Finish when done).",
        )
    except OperationCanceledException:
        return []
    except Exception:
        return []
    return list(refs or [])


# ---------------------------------------------------------------------
# Bucket / name prompts
# ---------------------------------------------------------------------

def _prompt_for_bucket(profile_data):
    """Let the user pick a bucket from the YAML. Returns the
    bucket_id or "" on cancel."""
    buckets = list(profile_data.get("space_buckets") or [])
    if not buckets:
        return ""
    labels_by_id = {}
    options = []
    for b in buckets:
        if not isinstance(b, dict):
            continue
        bid = b.get("id") or ""
        name = b.get("name") or bid or "(unnamed)"
        if not bid:
            continue
        label = "{}  ({})".format(name, bid)
        labels_by_id[label] = bid
        options.append(label)
    if not options:
        return ""
    picked = wpf_dialogs.pick_from_list(
        sorted(options),
        title="Bucket for captured profile",
        prompt="Pick the bucket this captured profile targets:",
    )
    if not picked:
        return ""
    return labels_by_id.get(picked, "")


def _prompt_for_name(default=""):
    name = wpf_dialogs.prompt_for_string(
        prompt="Name for the captured space profile:",
        title=TITLE,
        default=default or "Captured space profile",
    )
    return (name or "").strip()


# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------

def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None or uidoc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc) or {}
    if not isinstance(profile_data.get("space_profiles"), list):
        profile_data["space_profiles"] = []

    space = _pick_space(uidoc, doc)
    if space is None or space.element is None:
        output.print_md("**Capture cancelled** — no space picked.")
        return

    # Build geometry to discover doors / dimensions.
    import space_placement as _placement
    geom = _placement.build_space_geometry(doc, space.element)
    if geom is None:
        forms.alert(
            "Space '{}' has no usable boundary. Confirm it's a placed "
            "Space (not just a Room or unplaced) and try again.".format(
                space.name or "?"
            ),
            title=TITLE,
        )
        return
    doors = list(geom.door_anchors or [])
    if not doors:
        forms.alert(
            "Space '{}' has no detected doors. Wall-relative capture "
            "needs at least one door to label walls by role "
            "(opposite_door / right_of_door / left_of_door / "
            "behind_door). Add a door to the space's boundary walls "
            "(host or linked architecture) and try again.".format(
                space.name or "?"
            ),
            title=TITLE,
        )
        return

    door_anchor = _pick_reference_door(uidoc, doc, space, doors)

    refs = _pick_children(uidoc)
    if not refs:
        output.print_md("**Capture cancelled** — no children picked.")
        return

    profile_name = _prompt_for_name(default="{} profile".format(space.name or "Space"))
    if not profile_name:
        output.print_md("**Capture cancelled** — no profile name supplied.")
        return
    bucket_id = _prompt_for_bucket(profile_data)

    request = _capture.CaptureRequest(
        space=space,
        door_anchor=door_anchor,
        picked_refs=refs,
        profile_name=profile_name,
        bucket_id=bucket_id,
    )
    result = _capture.run_capture(doc, request)

    if result.profile is None:
        output.print_md(
            "**Capture failed**\n\n"
            + ("\n".join("- {}".format(w) for w in result.warnings) or "- (no detail)")
        )
        return

    # Merge into profile_data (creates new or appends to same-named
    # existing profile, deduping LEDs at matching wall positions).
    action, target_profile, n_added, n_dup = _capture.commit_capture(
        profile_data, result,
    )
    if action == "noop":
        output.print_md("**Nothing to save** — capture produced no profile.")
        return

    try:
        with revit.Transaction(TITLE, doc=doc):
            active_yaml.save_active_data(doc, profile_data, action=TITLE)
    except Exception as exc:
        output.print_md(
            "**Save FAILED**\n\n"
            "- Error type: `{}`\n"
            "- Error: {}\n".format(type(exc).__name__, exc)
        )
        raise

    # Render every LED on the target profile so the user can see the
    # full state (including pre-existing entries when we appended).
    led_lines = []
    set0 = (target_profile.get("linked_sets") or [{}])[0]
    for led in set0.get("linked_element_definitions") or []:
        rule = led.get("placement_rule") or {}
        if rule.get("kind") == "space_anchored":
            pos_str = "fx=`{:.3f}` fy=`{:.3f}`".format(
                float(rule.get("x_fraction") or 0.0),
                float(rule.get("y_fraction") or 0.0),
            )
        else:
            pos_str = "pos=`{:.3f}`".format(
                float(rule.get("position_along_wall") or 0.0),
            )
        led_lines.append(
            "- `{}` {} — kind=`{}` wall=`{}` {} z=`{:.2f}\"`".format(
                led.get("id"),
                led.get("label"),
                rule.get("kind"),
                rule.get("wall_role"),
                pos_str,
                float((led.get("offsets") or [{}])[0].get("z_inches") or 0.0),
            )
        )

    if action == "created":
        headline = "**Created new space profile `{}`**".format(
            target_profile.get("name")
        )
    else:
        headline = (
            "**Appended to existing profile `{}`** (same name found in "
            "the YAML — merged instead of duplicating)".format(
                target_profile.get("name")
            )
        )
    output.print_md(
        "{}\n\n"
        "- LEDs added this run: {}\n"
        "- Duplicate LEDs skipped (matched existing position): {}\n"
        "- Captured from picks: {}\n"
        "- Skipped picks: {}\n"
        "- Warnings: {}\n\n"
        "Full LED list on the profile:\n{}\n".format(
            headline,
            n_added,
            n_dup,
            len(result.captured),
            len(result.skipped),
            len(result.warnings),
            "\n".join(led_lines) or "  (none)",
        )
    )
    if result.skipped:
        output.print_md(
            "\n**Skipped picks:**\n"
            + "\n".join("- {}".format(reason) for _ref, reason in result.skipped)
        )
    if result.warnings:
        output.print_md(
            "\n**Warnings:**\n"
            + "\n".join("- {}".format(w) for w in result.warnings)
        )


if __name__ == "__main__":
    main()
