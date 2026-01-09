# -*- coding: utf-8 -*-
"""
Post Audit Compare
------------------
Compare a Post Audit output file against the current model and report moved,
added, and removed IDs based on (Parent ElementId, Parent_location) pairs.
"""

import ast
import io
import re

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    Transaction,
    XYZ,
)

TITLE = "Post Audit Compare"
PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")


def _load_dict(path):
    with io.open(path, "r", encoding="utf-8") as handle:
        raw = handle.read()
    data = ast.literal_eval(raw)
    if not isinstance(data, dict):
        raise ValueError("File does not contain a dictionary.")
    return data


def _normalize_locations(value):
    if value is None:
        return set()
    if isinstance(value, (list, tuple, set)):
        items = list(value)
    else:
        items = [value]
    normalized = set()
    for item in items:
        if isinstance(item, str) and item.strip().lower() == "not found":
            normalized.add("Not found")
            continue
        if isinstance(item, (list, tuple)) and len(item) == 3:
            try:
                normalized.add((float(item[0]), float(item[1]), float(item[2])))
            except Exception:
                normalized.add("Not found")
            continue
        normalized.add("Not found")
    return normalized


def _build_map(data):
    result = {}
    for raw_id, raw_locations in data.items():
        try:
            parent_id = int(raw_id)
        except Exception:
            continue
        result[parent_id] = _normalize_locations(raw_locations)
    return result


def _parse_int(value):
    try:
        return int(value)
    except Exception:
        try:
            return int(float(value))
        except Exception:
            return None


def _parse_payload(text):
    if not text:
        return None
    parent_id = None
    parent_location = None
    match = re.search(r"Parent\s*Element\s*Id\s*:\s*([-\d\.]+)", text, re.IGNORECASE)
    if match:
        parent_id = _parse_int(match.group(1))
    if re.search(r"Parent[_\s]*location\s*:\s*Not found", text, re.IGNORECASE):
        parent_location = "Not found"
    else:
        match = re.search(
            r"Parent[_\s]*location\s*:\s*([-\d\.]+)\s*,\s*([-\d\.]+)\s*,\s*([-\d\.]+)",
            text,
            re.IGNORECASE,
        )
        if match:
            parent_location = (
                float(match.group(1)),
                float(match.group(2)),
                float(match.group(3)),
            )
    if parent_id is None and parent_location is None:
        return None
    if parent_location is None:
        parent_location = "Not found"
    return {
        "parent_element_id": parent_id,
        "parent_location": parent_location,
    }


def _get_param_text(param):
    if not param:
        return None
    text = None
    try:
        text = param.AsString()
    except Exception:
        text = None
    if not text:
        try:
            text = param.AsValueString()
        except Exception:
            text = None
    return text


def _lookup_param(elem, names):
    for name in names:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if param:
            return param
    return None


def _collect_current_data(doc):
    results = {}
    children = {}
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for elem in collector:
        param = _lookup_param(elem, PARAM_NAMES)
        text = _get_param_text(param)
        if not text:
            try:
                type_id = elem.GetTypeId()
            except Exception:
                type_id = None
            if type_id:
                try:
                    type_elem = doc.GetElement(type_id)
                except Exception:
                    type_elem = None
                if type_elem:
                    type_param = _lookup_param(type_elem, PARAM_NAMES)
                    text = _get_param_text(type_param)
        if not text:
            continue
        payload = _parse_payload(text)
        if not payload:
            continue
        parent_id = payload.get("parent_element_id")
        if parent_id is None:
            continue
        parent_location = payload.get("parent_location", "Not found")
        results.setdefault(parent_id, set()).add(parent_location)
        children.setdefault(parent_id, []).append(elem.Id.IntegerValue)
    return results, children


def _single_location(locations):
    if not locations:
        return None
    clean = [loc for loc in locations if isinstance(loc, tuple) and len(loc) == 3]
    if len(clean) == 1:
        return clean[0]
    return None


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    action = forms.alert(
        "What would you like to do?",
        title=TITLE,
        options=["Save", "Compare", "Cancel"],
    )
    if action == "Cancel":
        return

    if action == "Save":
        current_data, _child_map = _collect_current_data(doc)
        if not current_data:
            forms.alert("No Element_Linker payloads found in the current model.", title=TITLE)
            return
        output = {}
        for parent_id, locations in current_data.items():
            output[parent_id] = list(locations)
        save_path = forms.save_file(
            file_ext="txt",
            title="Save Post Audit file",
            default_name="post_audit_parent_locations.txt",
        )
        if not save_path:
            return
        try:
            with io.open(save_path, "w", encoding="utf-8") as handle:
                handle.write(repr(output))
                handle.write("\n")
        except Exception as exc:
            forms.alert("Failed to save file:\n\n{}".format(exc), title=TITLE)
            return
        forms.alert("Saved Post Audit file:\n{}".format(save_path), title=TITLE)
        return

    old_path = forms.pick_file(file_ext="txt", title="Select Post Audit file to compare")
    if not old_path:
        return

    try:
        old_data = _build_map(_load_dict(old_path))
    except Exception as exc:
        forms.alert("Failed to read file:\n\n{}".format(exc), title=TITLE)
        return

    new_data, child_map = _collect_current_data(doc)
    if not new_data:
        forms.alert("No Element_Linker payloads found in the current model.", title=TITLE)
        return

    old_ids = set(old_data.keys())
    new_ids = set(new_data.keys())

    removed_ids = sorted(old_ids - new_ids)
    added_ids = sorted(new_ids - old_ids)

    moved = []
    move_candidates = []
    common_ids = sorted(old_ids & new_ids)
    for pid in common_ids:
        old_locs = old_data.get(pid, set())
        new_locs = new_data.get(pid, set())
        if old_locs != new_locs:
            moved.append((pid, old_locs, new_locs))
            old_single = _single_location(old_locs)
            new_single = _single_location(new_locs)
            if old_single and new_single:
                delta = (
                    new_single[0] - old_single[0],
                    new_single[1] - old_single[1],
                    new_single[2] - old_single[2],
                )
                if abs(delta[0]) > 1e-9 or abs(delta[1]) > 1e-9 or abs(delta[2]) > 1e-9:
                    for child_id in child_map.get(pid, []):
                        move_candidates.append((child_id, delta, pid))

    output = script.get_output()
    output.print_md("# Post Audit Compare")
    output.print_md("## Summary")
    output.print_md("- Moved IDs: **{}**".format(len(moved)))
    output.print_md("- Removed IDs: **{}**".format(len(removed_ids)))
    output.print_md("- Added IDs: **{}**".format(len(added_ids)))

    if moved:
        output.print_md("\n## Moved IDs (old -> new)")
        for pid, old_locs, new_locs in moved:
            output.print_md("- {}".format(pid))
            output.print_md("  - old: {}".format(sorted(old_locs)))
            output.print_md("  - new: {}".format(sorted(new_locs)))
            old_single = _single_location(old_locs)
            new_single = _single_location(new_locs)
            if old_single and new_single:
                delta = (
                    new_single[0] - old_single[0],
                    new_single[1] - old_single[1],
                    new_single[2] - old_single[2],
                )
                output.print_md("  - delta: ({:.6f}, {:.6f}, {:.6f})".format(delta[0], delta[1], delta[2]))
            else:
                output.print_md("  - delta: ambiguous")
            children = child_map.get(pid, [])
            if children:
                output.print_md("  - child element ids: {}".format(sorted(children)))
            else:
                output.print_md("  - child element ids: []")

    if removed_ids:
        output.print_md("\n## Removed IDs")
        for pid in removed_ids:
            output.print_md("- {}".format(pid))

    if added_ids:
        output.print_md("\n## Added IDs")
        for pid in added_ids:
            output.print_md("- {}".format(pid))

    if move_candidates:
        parent_count = len(set([entry[2] for entry in move_candidates]))
        elem_count = len(set([entry[0] for entry in move_candidates]))
        choice = forms.alert(
            "Move child elements for moved parents?\n\n"
            "Parents with deltas: {}\n"
            "Child elements to move: {}\n\n"
            "This will move elements by their parent deltas.".format(parent_count, elem_count),
            title=TITLE,
            warn_icon=True,
            options=["Yes", "No"],
        )
        if choice == "Yes":
            moved_count = 0
            failed = []
            seen = set()
            t = Transaction(doc, "Post Audit Move Children")
            t.Start()
            for child_id, delta, _pid in move_candidates:
                if child_id in seen:
                    continue
                seen.add(child_id)
                try:
                    elem = doc.GetElement(ElementId(int(child_id)))
                except Exception:
                    elem = None
                if elem is None:
                    failed.append(child_id)
                    continue
                try:
                    ElementTransformUtils.MoveElement(doc, elem.Id, XYZ(delta[0], delta[1], delta[2]))
                    moved_count += 1
                except Exception:
                    failed.append(child_id)
            t.Commit()
            if failed:
                forms.alert(
                    "Moved {} element(s). Failed: {}.".format(moved_count, len(failed)),
                    title=TITLE,
                )
            else:
                forms.alert("Moved {} element(s).".format(moved_count), title=TITLE)

    forms.alert(
        "Compare complete.\n\nMoved: {}\nRemoved: {}\nAdded: {}\n\nSee the output panel for details.".format(
            len(moved), len(removed_ids), len(added_ids)
        ),
        title=TITLE,
    )


if __name__ == "__main__":
    main()
