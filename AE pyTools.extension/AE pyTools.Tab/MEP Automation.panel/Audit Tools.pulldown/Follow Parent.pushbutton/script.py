# -*- coding: utf-8 -*-
"""
Follow Parent
-------------
Move placed profile children to follow their current parent position/rotation
using Element_Linker metadata.
"""

import math
import re

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    Line,
    RevitLinkInstance,
    Transaction,
    XYZ,
)

output = script.get_output()
output.close_others()

TITLE = "Follow Parent"
LOG = script.get_logger()

LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
INLINE_LINKER_PATTERN = re.compile(
    r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|"
    r"Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
    r"Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation|"
    r"CKT_Circuit Number_CEDT|CKT_Panel_CEDT)\s*:\s*",
    re.IGNORECASE,
)

POSITION_TOL_FT = 1.0 / 256.0
ROTATION_TOL_DEG = 0.01


def _element_id_value(elem_id, default=None):
    if elem_id is None:
        return default
    for attr in ("Value", "IntegerValue"):
        try:
            value = getattr(elem_id, attr)
        except Exception:
            value = None
        if value is None:
            continue
        try:
            return int(value)
        except Exception:
            try:
                return value
            except Exception:
                continue
    return default


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _parse_int(value, default=None):
    if value in (None, ""):
        return default
    try:
        return int(value)
    except Exception:
        try:
            return int(float(value))
        except Exception:
            return default


def _parse_float(value, default=None):
    if value in (None, ""):
        return default
    try:
        return float(value)
    except Exception:
        return default


def _parse_xyz_string(value):
    if not value:
        return None
    parts = [part.strip() for part in str(value).split(",")]
    if len(parts) != 3:
        return None
    try:
        return XYZ(float(parts[0]), float(parts[1]), float(parts[2]))
    except Exception:
        return None


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)


def _get_point(elem):
    if elem is None:
        return None
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    point = getattr(loc, "Point", None)
    if point is not None:
        return point
    curve = getattr(loc, "Curve", None)
    if curve is not None:
        try:
            return curve.Evaluate(0.5, True)
        except Exception:
            try:
                return curve.GetEndPoint(0)
            except Exception:
                return None
    return None


def _vector_angle_deg(vec):
    if vec is None:
        return None
    try:
        x = float(vec.X)
        y = float(vec.Y)
    except Exception:
        return None
    if abs(x) < 1e-9 and abs(y) < 1e-9:
        return None
    try:
        return math.degrees(math.atan2(y, x))
    except Exception:
        return None


def _get_orientation_vector(elem):
    for attr in ("HandOrientation", "FacingOrientation"):
        try:
            vec = getattr(elem, attr, None)
        except Exception:
            vec = None
        if vec and (abs(vec.X) > 1e-9 or abs(vec.Y) > 1e-9):
            return XYZ(vec.X, vec.Y, vec.Z)
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            angle = float(loc.Rotation)
            return XYZ(math.cos(angle), math.sin(angle), 0.0)
        except Exception:
            pass
    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        for basis_name in ("BasisX", "BasisY"):
            basis = getattr(transform, basis_name, None)
            if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
                return XYZ(basis.X, basis.Y, basis.Z)
    return None


def _get_rotation_degrees(elem, link_transform=None):
    vec = _get_orientation_vector(elem)
    if vec is not None and link_transform is not None:
        try:
            vec = link_transform.OfVector(vec)
        except Exception:
            pass
    angle = _vector_angle_deg(vec)
    if angle is not None:
        return angle
    return 0.0


def _rotate_xy(vec, angle_degrees):
    if vec is None:
        return None
    if not angle_degrees:
        return XYZ(vec.X, vec.Y, vec.Z)
    try:
        ang = math.radians(float(angle_degrees))
    except Exception:
        return XYZ(vec.X, vec.Y, vec.Z)
    cos_a = math.cos(ang)
    sin_a = math.sin(ang)
    x = vec.X * cos_a - vec.Y * sin_a
    y = vec.X * sin_a + vec.Y * cos_a
    return XYZ(x, y, vec.Z)


def _normalize_angle(value):
    if value is None:
        return 0.0
    result = float(value)
    while result <= -180.0:
        result += 360.0
    while result > 180.0:
        result -= 360.0
    return result


def _name_variants(elem):
    names = set()
    if elem is None:
        return names
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            names.add(raw_name)
    except Exception:
        pass
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            names.add(u"{} : {}".format(fam_name, type_name))
            names.add(u"{} : {}".format(type_name, fam_name))
        if fam_name:
            names.add(fam_name)
        if type_name:
            names.add(type_name)
    elif isinstance(elem, Group):
        group_type = getattr(elem, "GroupType", None)
        gname = getattr(group_type, "Name", None) if group_type else None
        if gname:
            names.add(gname)
    return {_normalize_name(name) for name in names if _normalize_name(name)}


def _collect_family_and_group_instances(doc):
    elements = []
    seen = set()
    for cls in (FamilyInstance, Group):
        try:
            collector = FilteredElementCollector(doc).OfClass(cls).WhereElementIsNotElementType()
        except Exception:
            continue
        for elem in collector:
            elem_id = _element_id_value(getattr(elem, "Id", None), None)
            if elem_id is None:
                continue
            key = (cls.__name__, elem_id)
            if key in seen:
                continue
            seen.add(key)
            elements.append(elem)
    return elements


def _get_linker_text(elem):
    if elem is None:
        return ""
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
        value = None
        try:
            value = param.AsString()
        except Exception:
            value = None
        if not value:
            try:
                value = param.AsValueString()
            except Exception:
                value = None
        if value and str(value).strip():
            return str(value).strip()
    return ""


def _parse_linker_payload(payload_text):
    if not payload_text:
        return {}
    text = str(payload_text)
    entries = {}
    if "\n" in text:
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or ":" not in line:
                continue
            key, _, remainder = line.partition(":")
            entries[key.strip()] = remainder.strip()
    else:
        matches = list(INLINE_LINKER_PATTERN.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key.strip()] = value.strip(" ,")

    parent_element_id = _parse_int(entries.get("Parent ElementId"))
    if parent_element_id is None:
        parent_element_id = _parse_int(entries.get("Parent Element ID"))

    return {
        "led_id": (entries.get("Linked Element Definition ID") or "").strip(),
        "set_id": (entries.get("Set Definition ID") or "").strip(),
        "host_name": (entries.get("Host Name") or "").strip(),
        "parent_location": _parse_xyz_string(entries.get("Parent_location")),
        "location": _parse_xyz_string(entries.get("Location XYZ (ft)")),
        "rotation_deg": _parse_float(entries.get("Rotation (deg)"), 0.0),
        "parent_rotation_deg": _parse_float(entries.get("Parent Rotation (deg)"), None),
        "parent_element_id": parent_element_id,
        "level_id": _parse_int(entries.get("LevelId")),
        "element_id": _parse_int(entries.get("ElementId")),
        "facing": _parse_xyz_string(entries.get("FacingOrientation")),
        "entries": entries,
    }


def _build_linker_payload(parsed_payload, child_elem, target_point, target_rot, parent_point, parent_rot, parent_id):
    entries = dict((parsed_payload or {}).get("entries") or {})
    if "Parent Element ID" in entries and "Parent ElementId" not in entries:
        entries["Parent ElementId"] = entries.get("Parent Element ID")

    entries["Linked Element Definition ID"] = (parsed_payload.get("led_id") or entries.get("Linked Element Definition ID") or "").strip()
    entries["Set Definition ID"] = (parsed_payload.get("set_id") or entries.get("Set Definition ID") or "").strip()
    if "Host Name" in entries:
        entries["Host Name"] = (entries.get("Host Name") or "").strip()
    entries["Parent_location"] = _format_xyz(parent_point)
    entries["Location XYZ (ft)"] = _format_xyz(target_point)
    entries["Rotation (deg)"] = "{:.6f}".format(float(target_rot or 0.0))
    entries["Parent Rotation (deg)"] = "{:.6f}".format(float(parent_rot or 0.0))
    entries["Parent ElementId"] = str(parent_id if parent_id is not None else "")
    entries["LevelId"] = str(_element_id_value(getattr(child_elem, "LevelId", None), "") or "")
    entries["ElementId"] = str(_element_id_value(getattr(child_elem, "Id", None), "") or "")
    entries["FacingOrientation"] = _format_xyz(getattr(child_elem, "FacingOrientation", None))

    ordered_keys = [
        "Linked Element Definition ID",
        "Set Definition ID",
        "Host Name",
        "Parent_location",
        "Location XYZ (ft)",
        "Rotation (deg)",
        "Parent Rotation (deg)",
        "Parent ElementId",
        "LevelId",
        "ElementId",
        "FacingOrientation",
        "CKT_Circuit Number_CEDT",
        "CKT_Panel_CEDT",
    ]

    lines = []
    used = set()
    for key in ordered_keys:
        if key not in entries:
            continue
        value = entries.get(key)
        lines.append("{}: {}".format(key, value if value is not None else ""))
        used.add(key)
    for key, value in entries.items():
        if key in used:
            continue
        lines.append("{}: {}".format(key, value if value is not None else ""))
    return "\n".join(lines).strip()


def _set_linker_text(elem, payload_text):
    if elem is None or not payload_text:
        return False
    success = False
    for name in LINKER_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(payload_text)
            success = True
        except Exception:
            continue
    return success


def _doc_key(doc):
    if doc is None:
        return None
    try:
        return doc.PathName or doc.Title
    except Exception:
        return None


def _get_link_transform(link_inst):
    if link_inst is None:
        return None
    try:
        return link_inst.GetTotalTransform()
    except Exception:
        try:
            return link_inst.GetTransform()
        except Exception:
            return None


def _combine_transform(parent_transform, child_transform):
    if parent_transform is None:
        return child_transform
    if child_transform is None:
        return parent_transform
    try:
        return parent_transform.Multiply(child_transform)
    except Exception:
        return None


def _walk_link_documents(doc, parent_transform, doc_chain):
    if doc is None:
        return
    key = _doc_key(doc)
    if key and key in doc_chain:
        return
    next_chain = set(doc_chain or set())
    if key:
        next_chain.add(key)
    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        transform = _combine_transform(parent_transform, _get_link_transform(link_inst))
        yield link_doc, transform
        for nested in _walk_link_documents(link_doc, transform, next_chain):
            yield nested


def _iter_link_documents(doc):
    for link_doc, transform in _walk_link_documents(doc, None, set()):
        yield link_doc, transform


def _transform_point(transform, point):
    if transform is None or point is None:
        return point
    try:
        return transform.OfPoint(point)
    except Exception:
        return point


def _collect_parent_candidates_for_ids(doc, parent_ids):
    parent_ids = {int(pid) for pid in (parent_ids or set()) if pid not in (None, "")}
    by_parent_id = {}
    if not parent_ids:
        return by_parent_id

    def _add_candidate(parent_id, elem, point, is_linked, link_transform=None):
        if elem is None or point is None:
            return
        by_parent_id.setdefault(parent_id, []).append({
            "parent_id": parent_id,
            "element": elem if not is_linked else None,
            "point": point,
            "rotation_deg": _get_rotation_degrees(elem, link_transform=link_transform),
            "name_variants": _name_variants(elem),
            "is_linked": bool(is_linked),
        })

    for pid in parent_ids:
        try:
            elem = doc.GetElement(ElementId(pid))
        except Exception:
            elem = None
        if elem is not None:
            _add_candidate(pid, elem, _get_point(elem), is_linked=False)

    for link_doc, link_transform in _iter_link_documents(doc):
        for pid in parent_ids:
            try:
                elem = link_doc.GetElement(ElementId(pid))
            except Exception:
                elem = None
            if elem is None:
                continue
            link_point = _transform_point(link_transform, _get_point(elem))
            _add_candidate(pid, elem, link_point, is_linked=True, link_transform=link_transform)

    return by_parent_id


def _collect_selection_ids():
    ids = set()
    uidoc = getattr(revit, "uidoc", None)
    if uidoc is None:
        return ids
    try:
        selected = list(uidoc.Selection.GetElementIds())
    except Exception:
        selected = []
    for elem_id in selected:
        value = _element_id_value(elem_id, None)
        if value is not None:
            ids.add(int(value))
    return ids


def _collect_child_records(doc):
    records = []
    for elem in _collect_family_and_group_instances(doc):
        payload_text = _get_linker_text(elem)
        if not payload_text:
            continue
        payload = _parse_linker_payload(payload_text)
        if not payload:
            continue
        parent_id = payload.get("parent_element_id")
        if parent_id is None:
            continue
        child_id = _element_id_value(getattr(elem, "Id", None), None)
        if child_id is None:
            continue
        records.append({
            "element": elem,
            "child_id": child_id,
            "payload": payload,
        })
    return records


def _filter_child_records_for_selection(records, selection_ids):
    if not selection_ids:
        return list(records or [])
    selection_ids = {int(value) for value in selection_ids if value is not None}
    if not selection_ids:
        return list(records or [])

    target_parent_ids = set(selection_ids)
    for record in records or []:
        child_id = record.get("child_id")
        parent_id = (record.get("payload") or {}).get("parent_element_id")
        if child_id in selection_ids and parent_id is not None:
            target_parent_ids.add(parent_id)

    filtered = []
    for record in records or []:
        child_id = record.get("child_id")
        parent_id = (record.get("payload") or {}).get("parent_element_id")
        if child_id in selection_ids or parent_id in target_parent_ids:
            filtered.append(record)
    return filtered


def _choose_parent_candidate(payload, candidates):
    if not candidates:
        return None
    host_name_norm = _normalize_name(payload.get("host_name"))
    old_parent = payload.get("parent_location")

    best = None
    best_key = None
    for cand in candidates:
        name_match = 1 if host_name_norm and host_name_norm in (cand.get("name_variants") or set()) else 0
        if old_parent is not None and cand.get("point") is not None:
            try:
                distance = old_parent.DistanceTo(cand["point"])
            except Exception:
                distance = 1e9
        else:
            distance = 1e9
        host_pref = 1 if not cand.get("is_linked") else 0
        key = (name_match, -distance, host_pref)
        if best_key is None or key > best_key:
            best_key = key
            best = cand
    return best


def _move_and_rotate_child(doc, elem, target_point, target_rot_deg):
    current_point = _get_point(elem)
    if current_point is None:
        return False, False

    moved = False
    rotated = False

    move_vec = target_point - current_point
    move_len = 0.0
    try:
        move_len = move_vec.GetLength()
    except Exception:
        move_len = 0.0
    if move_len > POSITION_TOL_FT:
        try:
            ElementTransformUtils.MoveElement(doc, elem.Id, move_vec)
            moved = True
            current_point = target_point
        except Exception:
            pass

    current_rot = _get_rotation_degrees(elem)
    rot_delta = _normalize_angle(target_rot_deg - current_rot)
    if abs(rot_delta) > ROTATION_TOL_DEG:
        try:
            axis = Line.CreateBound(current_point, current_point + XYZ(0, 0, 1))
            ElementTransformUtils.RotateElement(doc, elem.Id, axis, math.radians(rot_delta))
            rotated = True
        except Exception:
            pass

    return moved, rotated


def _preserve_child_z(elem, target_point):
    if elem is None or target_point is None:
        return target_point
    current_point = _get_point(elem)
    if current_point is None:
        return target_point
    return XYZ(target_point.X, target_point.Y, current_point.Z)


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    if getattr(doc, "IsFamilyDocument", False):
        forms.alert("Open a project document before running Follow Parent.", title=TITLE)
        return

    selection_ids = _collect_selection_ids()
    all_child_records = _collect_child_records(doc)
    child_records = _filter_child_records_for_selection(all_child_records, selection_ids)
    if not child_records:
        if selection_ids:
            forms.alert(
                "No eligible placed profile children found for the current selection.\n"
                "Select placed children or host parents, then rerun.",
                title=TITLE,
            )
        else:
            forms.alert(
                "No placed profile children with Element_Linker parent metadata were found.",
                title=TITLE,
            )
        return

    parent_ids = {record["payload"].get("parent_element_id") for record in child_records}
    parent_candidates = _collect_parent_candidates_for_ids(doc, parent_ids)

    moved_count = 0
    rotated_count = 0
    payload_updates = 0
    unchanged_count = 0
    skipped_no_parent = 0
    skipped_no_point = 0
    processed = 0

    txn = Transaction(doc, "Follow Parent")
    txn.Start()
    try:
        for record in child_records:
            elem = record.get("element")
            payload = record.get("payload") or {}
            parent_id = payload.get("parent_element_id")
            candidates = parent_candidates.get(parent_id) or []
            parent_choice = _choose_parent_candidate(payload, candidates)
            if parent_choice is None:
                skipped_no_parent += 1
                continue

            child_point_old = payload.get("location") or _get_point(elem)
            parent_point_old = payload.get("parent_location") or parent_choice.get("point")
            parent_point_new = parent_choice.get("point")
            if child_point_old is None or parent_point_old is None or parent_point_new is None:
                skipped_no_point += 1
                continue

            parent_rot_old = payload.get("parent_rotation_deg")
            if parent_rot_old is None:
                parent_rot_old = parent_choice.get("rotation_deg") or 0.0

            child_rot_old = payload.get("rotation_deg")
            if child_rot_old is None:
                child_rot_old = _get_rotation_degrees(elem)

            local_offset_old = _rotate_xy(child_point_old - parent_point_old, -parent_rot_old)
            rotation_offset_old = _normalize_angle(child_rot_old - parent_rot_old)

            parent_rot_new = parent_choice.get("rotation_deg") or parent_rot_old
            target_point = parent_point_new + _rotate_xy(local_offset_old, parent_rot_new)
            target_point = _preserve_child_z(elem, target_point)
            target_rot = parent_rot_new + rotation_offset_old

            moved, rotated = _move_and_rotate_child(doc, elem, target_point, target_rot)
            if moved:
                moved_count += 1
            if rotated:
                rotated_count += 1
            if not moved and not rotated:
                unchanged_count += 1

            new_payload_text = _build_linker_payload(
                payload,
                elem,
                target_point,
                target_rot,
                parent_point_new,
                parent_rot_new,
                parent_id,
            )
            if _set_linker_text(elem, new_payload_text):
                payload_updates += 1

            processed += 1
        txn.Commit()
    except Exception as exc:
        try:
            txn.RollBack()
        except Exception:
            pass
        forms.alert("Follow Parent failed:\n\n{}".format(exc), title=TITLE)
        return

    lines = [
        "Children considered: {}".format(len(child_records)),
        "Children processed: {}".format(processed),
        "Moved: {}".format(moved_count),
        "Rotated: {}".format(rotated_count),
        "Unchanged: {}".format(unchanged_count),
        "Element_Linker updated: {}".format(payload_updates),
        "Skipped (parent not found): {}".format(skipped_no_parent),
        "Skipped (missing point data): {}".format(skipped_no_point),
    ]

    LOG.info("[Follow Parent] %s", " | ".join(lines))
    forms.alert("\n".join(lines), title=TITLE)


if __name__ == "__main__":
    main()
