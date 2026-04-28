# -*- coding: utf-8 -*-
"""
Merge Linked Elements
---------------------
Allows the user to designate one equipment definition as the source of truth and
copy its configuration to one or more other definitions.  This lets similar
linked element names (e.g. "GO Checkstand 1" vs "GO Checkstand 2") share the
exact same placement data without editing the placer.
"""

import copy
import os
import sys

from pyrevit import forms, revit, script
output = script.get_output()
output.close_others()
from Autodesk.Revit.DB import ElementId, FamilyInstance, Group, RevitLinkInstance
from Autodesk.Revit.UI.Selection import ObjectType
from System.Drawing import Point, Size
from System.Windows.Forms import (
    BorderStyle,
    Button,
    DialogResult,
    Form,
    FormBorderStyle,
    FormStartPosition,
    Label,
    TextBox,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402

TITLE = "Merge Linked Elements"

TRUTH_SOURCE_ID_KEY = "ced_truth_source_id"
TRUTH_SOURCE_NAME_KEY = "ced_truth_source_name"


def _pick_merge_mode(source_name):
    """Use a WinForms dialog so users can explicitly pick merge mode."""
    try:
        dialog = Form()
        dialog.Text = "Select Merge Mode"
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog
        dialog.StartPosition = FormStartPosition.CenterScreen
        dialog.ClientSize = Size(780, 230)
        dialog.MaximizeBox = False
        dialog.MinimizeBox = False
        dialog.ShowInTaskbar = False

        header = Label()
        header.AutoSize = False
        header.Location = Point(12, 12)
        header.Size = Size(756, 18)
        header.Text = "Choose how to merge after selecting the source profile:"
        dialog.Controls.Add(header)

        source_label = Label()
        source_label.AutoSize = False
        source_label.Location = Point(12, 36)
        source_label.Size = Size(756, 18)
        source_label.Text = "Source profile (truth):"
        dialog.Controls.Add(source_label)

        source_box = TextBox()
        source_box.Location = Point(12, 58)
        source_box.Size = Size(756, 24)
        source_box.ReadOnly = True
        source_box.BorderStyle = BorderStyle.FixedSingle
        source_box.Text = source_name or ""
        source_box.TabStop = False
        dialog.Controls.Add(source_box)

        msg = Label()
        msg.AutoSize = False
        msg.Location = Point(12, 92)
        msg.Size = Size(756, 52)
        msg.Text = (
            "Use 'Merge existing profile(s)' to pick targets already in YAML.\n"
            "Use 'Add new parent' to use the selected host/linked element name (or pick one when prompted)."
        )
        dialog.Controls.Add(msg)

        merge_btn = Button()
        merge_btn.Text = "Merge existing profile(s)"
        merge_btn.Size = Size(160, 30)
        merge_btn.Location = Point(12, 170)
        merge_btn.DialogResult = DialogResult.Yes
        dialog.Controls.Add(merge_btn)

        add_parent_btn = Button()
        add_parent_btn.Text = "Add new parent"
        add_parent_btn.Size = Size(120, 30)
        add_parent_btn.Location = Point(185, 170)
        add_parent_btn.DialogResult = DialogResult.No
        dialog.Controls.Add(add_parent_btn)

        cancel_btn = Button()
        cancel_btn.Text = "Cancel"
        cancel_btn.Size = Size(90, 30)
        cancel_btn.Location = Point(317, 170)
        cancel_btn.DialogResult = DialogResult.Cancel
        dialog.Controls.Add(cancel_btn)

        dialog.AcceptButton = merge_btn
        dialog.CancelButton = cancel_btn
        result = dialog.ShowDialog()
        if result == DialogResult.Yes:
            return "Merge existing profile(s)"
        if result == DialogResult.No:
            return "Add new parent"
        return None
    except Exception:
        # Fallback in case WinForms fails in the current host.
        return forms.CommandSwitchWindow.show(
            ["Merge existing profile(s)", "Add new parent"],
            message="Choose merge mode after selecting source '{}':".format(source_name),
        )


def _build_selected_element_label(elem):
    if elem is None:
        return ""
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            return u"{} : {}".format(fam_name, type_name).strip()
        if type_name:
            return str(type_name).strip()
        if fam_name:
            return str(fam_name).strip()
    if isinstance(elem, Group):
        name = getattr(elem, "Name", None)
        if name:
            return str(name).strip()
    try:
        name = getattr(elem, "Name", None)
        if name:
            return str(name).strip()
    except Exception:
        pass
    return ""


def _selected_profile_name_from_revit_selection():
    selection = revit.get_selection()
    if not selection:
        return ""
    selected_elements = list(getattr(selection, "elements", []) or [])
    if not selected_elements:
        return ""
    for elem in selected_elements:
        if isinstance(elem, RevitLinkInstance):
            continue
        label = _build_selected_element_label(elem).strip()
        if label:
            return label
    return ""


def _linked_element_from_reference(doc, reference):
    linked_id = getattr(reference, "LinkedElementId", None)
    if isinstance(linked_id, ElementId) and linked_id != ElementId.InvalidElementId:
        try:
            host_elem = doc.GetElement(reference.ElementId)
        except Exception:
            host_elem = None
        if isinstance(host_elem, RevitLinkInstance):
            link_doc = host_elem.GetLinkDocument()
            if link_doc:
                try:
                    linked_elem = link_doc.GetElement(linked_id)
                except Exception:
                    linked_elem = None
                if linked_elem is not None:
                    return linked_elem
    try:
        return doc.GetElement(reference.ElementId)
    except Exception:
        return None


def _pick_profile_name_from_element(prompt):
    uidoc = getattr(revit, "uidoc", None)
    doc = getattr(revit, "doc", None)
    if uidoc is None or doc is None:
        return ""
    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, prompt)
    except Exception:
        return ""
    elem = _linked_element_from_reference(doc, reference)
    if isinstance(elem, RevitLinkInstance):
        forms.alert("Select the specific element inside the link after this message.", title=TITLE)
        try:
            link_ref = uidoc.Selection.PickObject(ObjectType.LinkedElement, prompt)
        except Exception:
            return ""
        elem = _linked_element_from_reference(doc, link_ref)
    return _build_selected_element_label(elem).strip()


def _next_eq_number(equipment_defs):
    max_id = 0
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        if not eq_id:
            continue
        suffix = eq_id.split("-")[-1]
        try:
            num = int(suffix)
        except Exception:
            continue
        if num > max_id:
            max_id = num
    return max_id + 1


def _build_definition_map(equipment_defs):
    mapping = {}
    ordered = []
    for entry in equipment_defs or []:
        if not isinstance(entry, dict):
            continue
        name = (entry.get("name") or entry.get("id") or "").strip()
        if not name:
            continue
        mapping[name] = entry
        ordered.append(name)
    ordered.sort(key=lambda val: val.lower())
    return mapping, ordered


def _truth_groups(equipment_defs):
    groups = {}
    id_to_entry = {}
    name_to_entry = {}
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        eq_name = (entry.get("name") or entry.get("id") or "").strip()
        if eq_id:
            id_to_entry[eq_id] = entry
        if eq_name:
            name_to_entry[eq_name] = entry
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        eq_name = (entry.get("name") or entry.get("id") or "").strip()
        if not eq_name:
            continue
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id:
            source_id = eq_id or eq_name
        display_name = (entry.get(TRUTH_SOURCE_NAME_KEY) or "").strip()
        if not display_name:
            display_name = eq_name
        group = groups.setdefault(source_id, {
            "display_name": display_name,
            "members": [],
            "source_entry": None,
            "source_profile_name": None,
        })
        group["members"].append(eq_name)
        if eq_id and eq_id == source_id:
            group["source_entry"] = entry
            group["source_profile_name"] = eq_name
    for source_id, data in groups.items():
        if not data.get("source_entry"):
            fallback = data["members"][0]
            entry = name_to_entry.get(fallback) or id_to_entry.get(source_id)
            data["source_entry"] = entry
            data["source_profile_name"] = fallback
        if not data.get("display_name"):
            data["display_name"] = data.get("source_profile_name") or source_id
    return groups


def _normalize_text(value):
    if value is None:
        return ""
    return " ".join(str(value).strip().split()).lower()


def _normalize_keynote_family(value):
    if not value:
        return ""
    text = str(value)
    if ":" in text:
        text = text.split(":", 1)[0]
    return "".join([ch for ch in text.lower() if ch.isalnum()])


def _is_ga_keynote_entry(entry):
    if not isinstance(entry, dict):
        return False
    family = entry.get("family_name") or entry.get("family") or ""
    return _normalize_keynote_family(family) == "gakeynotesymbolced"


def _normalize_keynote_params(params):
    if not isinstance(params, dict):
        return {}
    out = dict(params)
    if out.get("Keynote Value") not in (None, ""):
        out.pop("Key Value", None)
        out.pop("Keynote", None)
    elif out.get("Key Value") not in (None, ""):
        out["Keynote Value"] = out.pop("Key Value")
        out.pop("Keynote", None)
    elif out.get("Keynote") not in (None, ""):
        out["Keynote Value"] = out.pop("Keynote")
    if out.get("Keynote Description") not in (None, ""):
        out.pop("Keynote Text", None)
    elif out.get("Keynote Text") not in (None, ""):
        out["Keynote Description"] = out.pop("Keynote Text")
    return out


def _coerce_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return float(default)


def _keynote_identity(entry):
    family = entry.get("family_name") or entry.get("family") or ""
    type_name = entry.get("type_name") or entry.get("type") or ""
    category = entry.get("category_name") or entry.get("category") or ""
    offsets = entry.get("offsets") or {}
    return (
        _normalize_text(family),
        _normalize_text(type_name),
        _normalize_text(category),
        round(_coerce_float(offsets.get("x_inches", 0.0)), 4),
        round(_coerce_float(offsets.get("y_inches", 0.0)), 4),
        round(_coerce_float(offsets.get("z_inches", 0.0)), 4),
        round(_coerce_float(offsets.get("rotation_deg", 0.0)), 4),
    )


def _iter_led_entries(eq_entry):
    if not isinstance(eq_entry, dict):
        return
    for linked_set in eq_entry.get("linked_sets") or []:
        if not isinstance(linked_set, dict):
            continue
        for led in linked_set.get("linked_element_definitions") or []:
            if isinstance(led, dict):
                yield led


def _build_keynote_map(eq_entry):
    by_led = {}
    for led in _iter_led_entries(eq_entry):
        led_id = (led.get("id") or "").strip()
        if not led_id:
            continue
        items = {}
        for entry in led.get("keynotes") or []:
            if not _is_ga_keynote_entry(entry):
                continue
            params = _normalize_keynote_params(entry.get("parameters") or {})
            cloned = copy.deepcopy(entry)
            cloned["parameters"] = dict(params)
            items[_keynote_identity(entry)] = {
                "params": params,
                "entry": cloned,
            }
        if items:
            by_led[led_id] = items
    return by_led


def _preserve_keynote_params(target_before_copy, copied_target):
    before_map = _build_keynote_map(target_before_copy)
    if not before_map:
        return
    for led in _iter_led_entries(copied_target):
        led_id = (led.get("id") or "").strip()
        if not led_id:
            continue
        before_items = before_map.get(led_id) or {}
        if not before_items:
            continue

        current_keynotes = led.get("keynotes")
        if not isinstance(current_keynotes, list) or not current_keynotes:
            restored = [copy.deepcopy(data.get("entry")) for data in before_items.values() if isinstance(data, dict)]
            # Only restore full list when source had none for this LED.
            if restored:
                led["keynotes"] = restored
            continue

        for keynote in current_keynotes:
            if not _is_ga_keynote_entry(keynote):
                continue
            match_data = before_items.get(_keynote_identity(keynote))
            match = (match_data or {}).get("params") if isinstance(match_data, dict) else None
            if not match:
                continue
            incoming = _normalize_keynote_params(keynote.get("parameters") or {})
            if incoming.get("Keynote Value") in (None, "") and match.get("Keynote Value") not in (None, ""):
                incoming["Keynote Value"] = match.get("Keynote Value")
            if incoming.get("Keynote Description") in (None, "") and match.get("Keynote Description") not in (None, ""):
                incoming["Keynote Description"] = match.get("Keynote Description")
            keynote["parameters"] = _normalize_keynote_params(incoming)


def _copy_fields(source_entry, target_entry):
    """Copy everything except identifying fields (name, id)."""
    target_before_copy = copy.deepcopy(target_entry)
    keep_keys = {"name", "id"}
    for key, value in list(target_entry.items()):
        if key in keep_keys:
            continue
        target_entry.pop(key, None)
    for key, value in source_entry.items():
        if key in keep_keys:
            continue
        target_entry[key] = copy.deepcopy(value)
    _preserve_keynote_params(target_before_copy, target_entry)


def _ensure_truth_source(entry):
    if not isinstance(entry, dict):
        return None, None
    eq_id = (entry.get("id") or entry.get("name") or "").strip()
    eq_name = (entry.get("name") or entry.get("id") or "").strip()
    if not eq_id:
        return None, None
    entry[TRUTH_SOURCE_ID_KEY] = eq_id
    if eq_name:
        entry[TRUTH_SOURCE_NAME_KEY] = eq_name
    return eq_id, eq_name


def _repoint_truth_children(equipment_defs, old_source_id, new_source_id, new_source_name):
    old_id = (old_source_id or "").strip()
    new_id = (new_source_id or "").strip()
    if not old_id or not new_id or old_id == new_id:
        return
    for entry in equipment_defs or []:
        source_id = (entry.get(TRUTH_SOURCE_ID_KEY) or "").strip()
        if not source_id and entry.get("id"):
            source_id = str(entry.get("id") or "").strip()
        if source_id == old_id:
            entry[TRUTH_SOURCE_ID_KEY] = new_id
            if new_source_name:
                entry[TRUTH_SOURCE_NAME_KEY] = new_source_name


def main():
    try:
        yaml_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    equipment_defs = data.get("equipment_definitions") or []
    def_map, ordered_names = _build_definition_map(equipment_defs)
    truth_groups = _truth_groups(equipment_defs)
    if not def_map:
        forms.alert("No equipment definitions are available to merge.", title=TITLE)
        return

    yaml_label = get_yaml_display_name(yaml_path)

    if truth_groups:
        sorted_keys = sorted(truth_groups.keys(), key=lambda k: (truth_groups[k]["display_name"] or k).lower())
        label_counts = {}
        base_info = []
        for key in sorted_keys:
            group = truth_groups[key]
            base_label = group.get("display_name") or group.get("source_profile_name") or key
            member_count = len(group.get("members") or [])
            if member_count > 1:
                base_label = u"{} ({} profiles)".format(base_label, member_count)
            base_info.append((key, base_label))
            label_counts[base_label] = label_counts.get(base_label, 0) + 1
        source_items = []
        label_to_key = {}
        for key, base_label in base_info:
            label = base_label
            if label_counts.get(base_label, 0) > 1:
                label = u"{} [{}]".format(base_label, key)
            source_items.append(label)
            label_to_key[label] = key
        source_choice = forms.SelectFromList.show(
            source_items,
            title="Select source definition (truth)",
            multiselect=False,
            button_name="Select",
        )
        if not source_choice:
            return
        source_label = source_choice if isinstance(source_choice, basestring) else source_choice[0]
        source_key = label_to_key.get(source_label)
        group = truth_groups.get(source_key or "")
        source_entry = group.get("source_entry") if group else None
        source_name = group.get("source_profile_name") if group else None
        if not source_entry or not source_name:
            forms.alert("Could not resolve the selected source definition.", title=TITLE)
            return
        source_display = group.get("display_name") or source_name
    else:
        source_choice = forms.SelectFromList.show(
            ordered_names,
            title="Select source definition (truth)",
            multiselect=False,
            button_name="Select",
        )
        if not source_choice:
            return
        source_name = source_choice if isinstance(source_choice, basestring) else source_choice[0]
        source_entry = def_map.get(source_name)
        source_display = source_name
        if not source_entry:
            forms.alert("Could not resolve the selected source definition.", title=TITLE)
            return

    selected_profile_name = _selected_profile_name_from_revit_selection()
    action_choice = _pick_merge_mode(source_name)
    if not action_choice:
        return

    root_id, root_name = _ensure_truth_source(source_entry)
    if not root_id:
        root_id = (source_entry.get("id") or source_entry.get("name") or source_name).strip()
    if not root_name:
        root_name = (source_entry.get("name") or source_display or source_name).strip()

    merged = []
    created = []
    if action_choice == "Add new parent":
        target_name = selected_profile_name.strip() if selected_profile_name else ""
        if not target_name:
            target_name = _pick_profile_name_from_element("Select parent element for merged profile")
        if not target_name:
            forms.alert(
                "No selected element name could be resolved.\n"
                "Select a host or linked Revit element with a usable name.",
                title=TITLE,
            )
            return
        if target_name == source_name:
            forms.alert(
                "Selected element name matches the source profile.\n"
                "Choose a different element or use Merge existing profile(s).",
                title=TITLE,
            )
            return
        target_entry = def_map.get(target_name)
        if not target_entry:
            seq = _next_eq_number(equipment_defs)
            target_entry = {
                "id": "EQ-{:03d}".format(seq),
                "name": target_name,
            }
            equipment_defs.append(target_entry)
            def_map[target_name] = target_entry
            created.append(target_name)
        _copy_fields(source_entry, target_entry)
        target_entry[TRUTH_SOURCE_ID_KEY] = root_id
        if root_name:
            target_entry[TRUTH_SOURCE_NAME_KEY] = root_name
        _repoint_truth_children(
            equipment_defs,
            target_entry.get("id"),
            root_id,
            root_name,
        )
        merged.append(target_name)
    else:
        target_candidates = [name for name in ordered_names if name != source_name]
        if not target_candidates:
            forms.alert("There are no other definitions to merge into.", title=TITLE)
            return
        target_choices = forms.SelectFromList.show(
            target_candidates,
            title="Select definition(s) to merge into '{}'".format(source_name),
            multiselect=True,
            button_name="Merge",
        )
        if not target_choices:
            return
        for target_name in target_choices:
            target_entry = def_map.get(target_name)
            if not target_entry or target_entry is source_entry:
                continue
            _copy_fields(source_entry, target_entry)
            target_entry[TRUTH_SOURCE_ID_KEY] = root_id
            if root_name:
                target_entry[TRUTH_SOURCE_NAME_KEY] = root_name
            _repoint_truth_children(
                equipment_defs,
                target_entry.get("id"),
                root_id,
                root_name,
            )
            merged.append(target_name)

    if not merged:
        forms.alert("No definitions were merged.", title=TITLE)
        return

    save_active_yaml_data(
        None,
        data,
        "Merge Linked Elements",
        "Merged definitions {} into '{}'".format(", ".join(merged), source_name),
    )

    lines = [
        "Merged linked elements successfully.",
        "Source definition: {}".format(source_display or source_name),
        "Merged into source: {}".format(", ".join(merged)),
        "",
        "Updated data saved back to {}.".format(yaml_label),
    ]
    if created:
        lines.insert(3, "Created new parent profile(s): {}".format(", ".join(created)))
        lines.insert(4, "")
    forms.alert("\n".join(lines), title=TITLE)


if __name__ == "__main__":
    try:
        basestring
    except NameError:
        basestring = str
    main()
