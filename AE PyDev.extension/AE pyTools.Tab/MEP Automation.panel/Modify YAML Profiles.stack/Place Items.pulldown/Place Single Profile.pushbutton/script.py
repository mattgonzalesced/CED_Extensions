# -*- coding: utf-8 -*-
"""
Load Equipment Definition
-------------------------
Reads the active YAML stored in Extensible Storage, lets the user pick an
equipment definition, and places the definition (plus optional linked types)
at a picked point.
"""

import os
import re
import sys

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import XYZ
from System.Windows.Forms import (
    Button,
    CheckBox,
    DialogResult,
    Form,
    FormBorderStyle,
    FormStartPosition,
)

LIB_ROOT = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "..",
        "..",
        "CEDLib.lib",
    )
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.placement_engine import PlaceElementsEngine  # noqa: E402
from LogicClasses.profile_repository import ProfileRepository  # noqa: E402
from LogicClasses.linked_equipment import build_child_requests, find_equipment_by_name  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Load Equipment Definition"
LOG = script.get_logger()

try:
    basestring
except NameError:
    basestring = str


def _sanitize_equipment_definitions(equipment_defs):
    cleaned_defs = []
    for eq in equipment_defs or []:
        if not isinstance(eq, dict):
            continue
        sanitized = dict(eq)
        linked_sets = []
        for linked_set in sanitized.get("linked_sets") or []:
            if not isinstance(linked_set, dict):
                continue
            ls_copy = dict(linked_set)
            led_list = []
            for led in ls_copy.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                if led.get("is_parent_anchor"):
                    continue
                led_copy = dict(led)
                tags = led_copy.get("tags")
                if isinstance(tags, list):
                    led_copy["tags"] = [t if isinstance(t, dict) else {} for t in tags]
                else:
                    led_copy["tags"] = []
                offsets = led_copy.get("offsets")
                if isinstance(offsets, list):
                    led_copy["offsets"] = [o if isinstance(o, dict) else {} for o in offsets]
                else:
                    led_copy["offsets"] = [{}]
                led_list.append(led_copy)
            ls_copy["linked_element_definitions"] = led_list
            linked_sets.append(ls_copy)
        sanitized["linked_sets"] = linked_sets
        cleaned_defs.append(sanitized)
    return cleaned_defs


def _sanitize_profiles(profiles):
    cleaned = []
    for prof in profiles or []:
        if not isinstance(prof, dict):
            continue
        prof_copy = dict(prof)
        types = []
        for t in prof_copy.get("types") or []:
            if not isinstance(t, dict):
                continue
            t_copy = dict(t)
            inst_cfg = t_copy.get("instance_config")
            if not isinstance(inst_cfg, dict):
                inst_cfg = {}
            offsets = inst_cfg.get("offsets")
            if not isinstance(offsets, list) or not offsets:
                offsets = [{}]
            inst_cfg["offsets"] = [off if isinstance(off, dict) else {} for off in offsets]
            tags = inst_cfg.get("tags")
            if isinstance(tags, list):
                inst_cfg["tags"] = [tag if isinstance(tag, dict) else {} for tag in tags]
            else:
                inst_cfg["tags"] = []
            params = inst_cfg.get("parameters")
            if not isinstance(params, dict):
                params = {}
            inst_cfg["parameters"] = params
            t_copy["instance_config"] = inst_cfg
            types.append(t_copy)
        prof_copy["types"] = types
        cleaned.append(prof_copy)
    return cleaned


def _is_independent_name(cad_name):
    if not cad_name:
        return False
    trimmed = cad_name.strip()
    if re.match(r"^\d{3}", trimmed):
        return False
    return not trimmed.lower().startswith("heb")


def _group_truth_profile_choices(raw_data, available_cads, independent_only=False):
    """Collapse equipment definitions by truth-source metadata so only canonical profiles appear."""
    if independent_only:
        available = {(name or "").strip(): True for name in available_cads if _is_independent_name(name)}
    else:
        available = {(name or "").strip(): True for name in available_cads}
    groups = {}
    for eq_def in raw_data.get("equipment_definitions") or []:
        cad_name = (eq_def.get("name") or eq_def.get("id") or "").strip()
        if not cad_name or cad_name not in available:
            continue
        truth_id = (eq_def.get("ced_truth_source_id") or eq_def.get("id") or cad_name).strip()
        if not truth_id:
            truth_id = cad_name
        display_name = (eq_def.get("ced_truth_source_name") or cad_name).strip() or cad_name
        group = groups.setdefault(truth_id, {"display": display_name, "members": [], "primary": None})
        group["members"].append(cad_name)
        eq_id = (eq_def.get("id") or "").strip()
        if eq_id and eq_id == truth_id:
            group["primary"] = cad_name
    if not groups:
        return [{"label": name, "cad": name} for name in sorted(available_cads)]
    display_counts = {}
    for info in groups.values():
        label = info.get("display") or ""
        display_counts[label] = display_counts.get(label, 0) + 1
    options = []
    seen_cads = set()
    for truth_id in sorted(groups.keys()):
        info = groups[truth_id]
        cad = info.get("primary") or (info.get("members") or [None])[0]
        if not cad or cad not in available:
            continue
        label = info.get("display") or cad
        if display_counts.get(label, 0) > 1:
            label = u"{} [{}]".format(label, truth_id)
        options.append({"label": label, "cad": cad})
        seen_cads.add(cad)
    # Include any cad names not covered by truth metadata
    for cad in sorted(available_cads):
        cad = cad.strip()
        if cad and cad not in seen_cads and cad in available:
            options.append({"label": cad, "cad": cad})
    return options


def _ask_profile_filter():
    form = Form()
    form.Text = "Load Profiles"
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.StartPosition = FormStartPosition.CenterScreen
    form.MinimizeBox = False
    form.MaximizeBox = False
    form.ShowInTaskbar = False
    form.Width = 360
    form.Height = 140

    checkbox = CheckBox()
    checkbox.Text = "Show only Independent Profiles"
    checkbox.AutoSize = True
    checkbox.Left = 12
    checkbox.Top = 12

    ok_button = Button()
    ok_button.Text = "OK"
    ok_button.DialogResult = DialogResult.OK
    ok_button.Left = 180
    ok_button.Top = 60
    ok_button.Width = 70

    cancel_button = Button()
    cancel_button.Text = "Cancel"
    cancel_button.DialogResult = DialogResult.Cancel
    cancel_button.Left = 260
    cancel_button.Top = 60
    cancel_button.Width = 70

    form.Controls.Add(checkbox)
    form.Controls.Add(ok_button)
    form.Controls.Add(cancel_button)
    form.AcceptButton = ok_button
    form.CancelButton = cancel_button

    result = form.ShowDialog()
    if result != DialogResult.OK:
        return None
    return bool(checkbox.Checked)


def _build_repository(data):
    cleaned_defs = _sanitize_equipment_definitions(data.get("equipment_definitions") or [])
    legacy_profiles = equipment_defs_to_legacy(cleaned_defs)
    cleaned_profiles = _sanitize_profiles(legacy_profiles)
    eq_defs = ProfileRepository._parse_profiles(cleaned_profiles)
    repo = ProfileRepository(eq_defs)
    log = script.get_logger()
    try:
        for cad in repo.cad_names():
            defs = repo._label_map.get(cad) or {}
            log.info("[Load Equipment Definition] repo contains CAD '%s' with %d linked definitions", cad, len(defs))
            for label, linked_def in defs.items():
                placement = linked_def.get_placement()
                offsets = placement.get_offset_xyz() if placement else (0, 0, 0)
                log.info("    label='%s' offsets=%s rotation=%s", label, offsets, placement.get_rotation_degrees() if placement else 0)
    except Exception:
        pass
    return repo


def _place_child_requests(repo, child_requests):
    selection_map = {}
    rows = []
    for request in child_requests or []:
        cad_name = request.get("name")
        labels = request.get("labels")
        point = request.get("target_point")
        rotation = request.get("rotation")
        if not cad_name or not labels or point is None:
            continue
        selection_map[cad_name] = labels
        rows.append({
            "Name": cad_name,
            "Count": "1",
            "Position X": str(point.X * 12.0),
            "Position Y": str(point.Y * 12.0),
            "Position Z": str(point.Z * 12.0),
            "Rotation": str(rotation or 0.0),
        })
        try:
            LOG.info("[Load Equipment Definition] child request '%s' offsets=%s", cad_name, [
                req.get("offsets") for req in request.get("linked_element_definitions", [])
            ])
        except Exception:
            pass
    if not selection_map or not rows:
        return 0
    engine = PlaceElementsEngine(revit.doc, repo, allow_tags=False, transaction_name="Load Equipment Definition (Children)")
    try:
        results = engine.place_from_csv(rows, selection_map)
    except Exception as exc:
        forms.alert("Failed to place linked child equipment:\\n\\n{}".format(exc), title=TITLE)
        return 0
    return results.get("placed", 0)


def _gather_child_requests(parent_def, base_point, base_rotation, repo, data):
    requests = []
    if not parent_def:
        return requests
    for linked_set in parent_def.get("linked_sets") or []:
        for led_entry in linked_set.get("linked_element_definitions") or []:
            led_id = (led_entry.get("id") or "").strip()
            if not led_id:
                continue
            reqs = build_child_requests(repo, data, parent_def, base_point, base_rotation, led_id)
            if reqs:
                requests.extend(reqs)
    return requests


def main():
    try:
        data_path, raw_data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)
    repo = _build_repository(raw_data)
    cad_names = repo.cad_names()
    if not cad_names:
        forms.alert("No equipment definitions found in {}.".format(yaml_label), title=TITLE)
        return

    independent_only = _ask_profile_filter()
    if independent_only is None:
        return
    grouped_choices = _group_truth_profile_choices(raw_data, cad_names, independent_only=independent_only)
    option_labels = [entry["label"] for entry in grouped_choices]
    choice_map = {entry["label"]: entry["cad"] for entry in grouped_choices}

    cad_choice_label = forms.SelectFromList.show(
        option_labels,
        title="Select equipment definition to place",
        multiselect=False,
        button_name="Load",
    )
    if not cad_choice_label:
        return
    cad_choice_label = cad_choice_label if isinstance(cad_choice_label, basestring) else cad_choice_label[0]
    cad_choice = choice_map.get(cad_choice_label, cad_choice_label)

    labels = repo.labels_for_cad(cad_choice)
    if not labels:
        forms.alert("Equipment definition '{}' has no linked types.".format(cad_choice), title=TITLE)
        return

    try:
        base_pt = revit.pick_point(message="Pick base point for '{}'".format(cad_choice))
    except Exception:
        base_pt = None
    if not base_pt:
        return

    selection_map = {cad_choice: labels}
    rows = [{
        "Name": cad_choice,
        "Count": "1",
        "Position X": str(base_pt.X * 12.0),
        "Position Y": str(base_pt.Y * 12.0),
        "Position Z": str(base_pt.Z * 12.0),
        "Rotation": "0",
    }]
    try:
        LOG.info("[Load Equipment Definition] repo built cad=%s label_count=%s", cad_choice, len(labels))
    except Exception:
        pass

    parent_def = find_equipment_by_name(raw_data, cad_choice)
    if parent_def:
        try:
            LOG.info("[Load Equipment Definition] offsets for '%s':", cad_choice)
            for linked_set in parent_def.get("linked_sets") or []:
                for led_entry in linked_set.get("linked_element_definitions") or []:
                    led_id = led_entry.get("id")
                    label = led_entry.get("label")
                    offsets = led_entry.get("offsets") or []
                    offsets_desc = offsets[0] if offsets else {}
                    LOG.info("  LED=%s label=%s offsets=%s", led_id, label, offsets_desc)
        except Exception:
            pass
    if parent_def:
        child_requests = _gather_child_requests(parent_def, base_pt, 0.0, repo, raw_data)
        if child_requests:
            if forms.alert(
                "Load '{}' with {} linked child equipment definition(s)?".format(cad_choice, len(child_requests)),
                title=TITLE,
                yes=True,
                no=True,
            ):
                for request in child_requests:
                    name = request.get("name")
                    labels = request.get("labels")
                    point = request.get("target_point")
                    rotation = request.get("rotation")
                    if not name or not labels or point is None:
                        continue
                    selection_map[name] = labels
                    rows.append({
                        "Name": name,
                        "Count": "1",
                        "Position X": str(point.X * 12.0),
                        "Position Y": str(point.Y * 12.0),
                        "Position Z": str(point.Z * 12.0),
                        "Rotation": str(rotation or 0.0),
                    })

    engine = PlaceElementsEngine(revit.doc, repo, allow_tags=False, transaction_name="Load Equipment Definition")
    try:
        results = engine.place_from_csv(rows, selection_map)
    except Exception as exc:
        forms.alert("Error during placement:\n\n{}".format(exc), title=TITLE)
        return

    forms.alert("Placed {} element(s) for equipment definition '{}'.".format(results.get("placed", 0), cad_choice), title=TITLE)


if __name__ == "__main__":
    main()
