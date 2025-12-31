# -*- coding: utf-8 -*-
"""
Update Vector
-------------
Reads the Element_Linker metadata stored on a selected element, compares the
current location/rotation to the original insertion point, and writes the
resulting offset + rotation back to the active YAML stored in Extensible Storage.
"""

import math
import os
import sys
import re

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementTransformUtils,
    FilteredElementCollector,
    Group,
    GroupType,
    IndependentTag,
    Line,
    TagOrientation,
    Transaction,
    XYZ,
    StorageType,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from LogicClasses.profile_schema import dump_data_to_string  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402
from ExtensibleStorage import ExtensibleStorage  # noqa: E402
ELEMENT_LINKER_PARAM_NAME = "Element_Linker Parameter"
ELEMENT_LINKER_SHARED_PARAM = "Element_Linker"
TITLE = "Update Vector"
LOG = script.get_logger()

try:
    basestring
except NameError:
    basestring = str


# --------------------------------------------------------------------------- #
# Geometry + element helpers
# --------------------------------------------------------------------------- #


def _get_point(elem):
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    if hasattr(loc, "Point") and loc.Point:
        return loc.Point
    if hasattr(loc, "Curve") and loc.Curve:
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            return None
    return None


def _get_tag_point(tag):
    try:
        return getattr(tag, "TagHeadPosition", None)
    except Exception:
        return None


def _vector_angle(vec):
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


def _get_rotation_degrees(elem):
    # Prefer basis vectors over Location.Rotation so mirrored/flipped families register 180° changes.
    for attr in ("HandOrientation", "FacingOrientation"):
        angle = _vector_angle(getattr(elem, attr, None))
        if angle is not None:
            return angle
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            pass
    try:
        transform = elem.GetTransform()
        angle = _vector_angle(getattr(transform, "BasisX", None))
        if angle is not None:
            return angle
    except Exception:
        pass
    return 0.0


def _feet_to_inches(value):
    try:
        return float(value) * 12.0
    except Exception:
        return 0.0


def _inches_to_feet(value):
    try:
        return float(value) / 12.0
    except Exception:
        return 0.0


def _level_relative_z_inches(elem, world_point):
    if elem is None:
        return 0.0
    doc = getattr(elem, "Document", None)
    level_elem = None
    level_id = getattr(elem, "LevelId", None)
    if level_id and doc:
        try:
            level_elem = doc.GetElement(level_id)
        except Exception:
            level_elem = None
    if not level_elem:
        level_param_names = (
            "INSTANCE_REFERENCE_LEVEL_PARAM",
            "FAMILY_LEVEL_PARAM",
            "INSTANCE_LEVEL_PARAM",
            "SCHEDULE_LEVEL_PARAM",
        )
        for name in level_param_names:
            bip = getattr(BuiltInParameter, name, None)
            if not bip:
                continue
            try:
                param = elem.get_Parameter(bip)
            except Exception:
                param = None
            if not param:
                continue
            try:
                eid = param.AsElementId()
            except Exception:
                eid = None
            if eid and doc:
                try:
                    level_elem = doc.GetElement(eid)
                except Exception:
                    level_elem = None
                if level_elem:
                    break
    level_elev = 0.0
    if level_elem:
        try:
            level_elev = getattr(level_elem, "Elevation", 0.0) or 0.0
        except Exception:
            level_elev = 0.0
    world_z = getattr(world_point, "Z", None)
    if world_z is None:
        world_z = 0.0
    try:
        relative_ft = float(world_z) - float(level_elev)
    except Exception:
        relative_ft = 0.0
    return _feet_to_inches(relative_ft)


def _rotate_xy(vec, angle_degrees):
    if vec is None:
        return None
    if not angle_degrees:
        return XYZ(vec.X, vec.Y, vec.Z)
    try:
        ang = math.radians(angle_degrees)
    except Exception:
        return XYZ(vec.X, vec.Y, vec.Z)
    cos_a = math.cos(ang)
    sin_a = math.sin(ang)
    x = vec.X * cos_a - vec.Y * sin_a
    y = vec.X * sin_a + vec.Y * cos_a
    return XYZ(x, y, vec.Z)


def _normalize_angle(delta_deg):
    if delta_deg is None:
        return 0.0
    value = float(delta_deg)
    while value <= -180.0:
        value += 360.0
    while value > 180.0:
        value -= 360.0
    return value


def _label_variants(value):
    base = (value or "").strip()
    variants = []
    if base:
        variants.append(base)
        if "#" in base:
            variants.append(base.split("#")[0].strip())
    return [v for v in variants if v]


WILDCARD_SIGNATURE = ("", "", "")


def _collect_hosted_tags(elem, host_point):
    doc = getattr(elem, "Document", None)
    if doc is None or host_point is None:
        return []
    try:
        deps = list(elem.GetDependentElements(None))
    except Exception:
        deps = []
    tags = []
    for dep_id in deps:
        try:
            tag = doc.GetElement(dep_id)
        except Exception:
            tag = None
        if not tag or not isinstance(tag, IndependentTag):
            continue
        try:
            tag_pt = tag.TagHeadPosition
        except Exception:
            tag_pt = None
        tag_symbol = None
        try:
            tag_symbol = doc.GetElement(tag.GetTypeId())
        except Exception:
            tag_symbol = None
        fam_name = None
        type_name = None
        category_name = None
        if tag_symbol:
            try:
                fam_name = getattr(tag_symbol, "FamilyName", None)
                if not fam_name:
                    fam = getattr(tag_symbol, "Family", None)
                    fam_name = getattr(fam, "Name", None) if fam else None
            except Exception:
                fam_name = None
            try:
                type_name = getattr(tag_symbol, "Name", None)
                if not type_name and hasattr(tag_symbol, "get_Parameter"):
                    try:
                        sparam = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if sparam:
                            type_name = sparam.AsString()
                    except Exception:
                        pass
            except Exception:
                type_name = None
            try:
                cat = getattr(tag_symbol, "Category", None)
                category_name = getattr(cat, "Name", None) if cat else None
            except Exception:
                category_name = None
        if not category_name:
            try:
                cat = getattr(tag, "Category", None)
                category_name = getattr(cat, "Name", None) if cat else None
            except Exception:
                category_name = None
        if not fam_name:
            try:
                sym = getattr(tag, "Symbol", None)
                fam = getattr(sym, "Family", None) if sym else None
                fam_name = getattr(fam, "Name", None) if fam else fam_name
            except Exception:
                pass
        if not type_name:
            try:
                tag_type = getattr(tag, "TagType", None)
                type_name = getattr(tag_type, "Name", None)
            except Exception:
                pass
        if not fam_name or not type_name:
            continue
        offsets = {
            "x_inches": 0.0,
            "y_inches": 0.0,
            "z_inches": 0.0,
            "rotation_deg": 0.0,
        }
        if tag_pt:
            delta = tag_pt - host_point
            offsets["x_inches"] = _feet_to_inches(delta.X)
            offsets["y_inches"] = _feet_to_inches(delta.Y)
            offsets["z_inches"] = _feet_to_inches(delta.Z)
        tags.append({
            "family_name": fam_name,
            "type_name": type_name,
            "category_name": category_name,
            "parameters": {},
            "offsets": offsets,
        })
    return tags


def _normalize_text(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _tag_entry_key(entry):
    if not isinstance(entry, dict):
        return None
    return _normalize_text(entry.get("type_name") or entry.get("type"))


def _tag_element_key(tag):
    if not isinstance(tag, IndependentTag):
        return None
    doc = getattr(tag, "Document", None)
    symbol = None
    try:
        if doc:
            symbol = doc.GetElement(tag.GetTypeId())
    except Exception:
        symbol = None
    family = None
    type_name = None
    if symbol:
        try:
            family = getattr(symbol, "FamilyName", None)
            if not family:
                fam = getattr(symbol, "Family", None)
                family = getattr(fam, "Name", None) if fam else None
        except Exception:
            family = None
        try:
            type_name = getattr(symbol, "Name", None)
            if not type_name:
                param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    type_name = param.AsString()
        except Exception:
            type_name = None
    if not family:
        try:
            sym = getattr(tag, "Symbol", None)
            fam = getattr(sym, "Family", None) if sym else None
            family = getattr(fam, "Name", None) if fam else family
        except Exception:
            pass
    if not type_name:
        try:
            tag_type = getattr(tag, "TagType", None)
            type_name = getattr(tag_type, "Name", None)
        except Exception:
            pass
    return _normalize_text(type_name)


def _collect_hosted_tag_elements(elem):
    doc = getattr(elem, "Document", None)
    if doc is None:
        return []
    try:
        dep_ids = list(elem.GetDependentElements(None))
    except Exception:
        dep_ids = []
    results = []
    host_id = getattr(getattr(elem, "Id", None), "IntegerValue", None)
    for dep_id in dep_ids:
        try:
            tag = doc.GetElement(dep_id)
        except Exception:
            tag = None
        if not tag or not isinstance(tag, IndependentTag):
            continue
        head = _get_tag_point(tag)
        if not head:
            continue
        signature = _tag_element_key(tag)
        orientation = None
        try:
            orientation = tag.TagOrientation
        except Exception:
            orientation = None
        results.append(
            {
                "element": tag,
                "head_point": head,
                "signature": signature,
                "orientation": orientation,
            }
        )
    if not results:
        try:
            all_tags = FilteredElementCollector(doc).OfClass(IndependentTag)
        except Exception:
            all_tags = []
        for tag in all_tags:
            host = _get_tag_host(tag)
            if host is None:
                continue
            try:
                tag_host_id = host.Id.IntegerValue
            except Exception:
                tag_host_id = None
            if host_id is not None and tag_host_id != host_id:
                continue
            head = _get_tag_point(tag)
            if not head:
                continue
            signature = _tag_element_key(tag)
            results.append(
                {
                    "element": tag,
                    "head_point": head,
                    "signature": signature,
                }
            )
    return results


def _group_tag_entries(tag_entries):
    grouped = {}
    for entry in tag_entries or []:
        key = _tag_entry_key(entry)
        if not key:
            continue
        grouped.setdefault(key, []).append(entry)
    return grouped


def _pop_matching_tag_entry(entry_map, target_key):
    if not entry_map:
        return None
    if target_key:
        bucket = entry_map.get(target_key)
        if bucket:
            entry = bucket.pop(0)
            if bucket:
                entry_map[target_key] = bucket
            else:
                entry_map.pop(target_key, None)
            return entry
    for key in list(entry_map.keys()):
        bucket = entry_map.get(key)
        if not bucket:
            continue
        entry = bucket.pop(0)
        if bucket:
            entry_map[key] = bucket
        else:
            entry_map.pop(key, None)
        return entry
    return None


def _ensure_tag_offset_entry(tag_entry):
    if not isinstance(tag_entry, dict):
        return {"x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0}
    offsets = tag_entry.get("offsets")
    entry = None
    if isinstance(offsets, list) and offsets:
        entry = offsets[0]
        if not isinstance(entry, dict):
            entry = {}
            offsets[0] = entry
    elif isinstance(offsets, dict):
        entry = offsets
    else:
        entry = {}
        tag_entry["offsets"] = [entry]
    entry.setdefault("x_inches", 0.0)
    entry.setdefault("y_inches", 0.0)
    entry.setdefault("z_inches", 0.0)
    entry.setdefault("rotation_deg", 0.0)
    return entry


def _update_tag_yaml_offsets(elem, led_entry, host_point, signature_filter=None):
    if not led_entry or host_point is None:
        return 0
    fresh_tags = _collect_hosted_tags(elem, host_point)
    if not fresh_tags:
        return 0
    tag_key = signature_filter
    if tag_key:
        existing = led_entry.get("tags") or []
        new_list = []
        replaced = False
        for entry in existing:
            key = _tag_entry_key(entry)
            if not replaced and key == tag_key:
                new_list.extend(fresh_tags)
                replaced = True
            else:
                new_list.append(entry)
        if not replaced:
            new_list.extend(fresh_tags)
        led_entry["tags"] = new_list
        return len(fresh_tags)
    led_entry["tags"] = fresh_tags
    return len(fresh_tags)


def _apply_tag_offsets_to_instance(doc, elem, tag_entries, host_point, rotation_delta_deg):
    if not doc or not elem or not tag_entries or host_point is None:
        return
    tag_infos = _collect_hosted_tag_elements(elem)
    if not tag_infos:
        return
    entry_map = _group_tag_entries(tag_entries)
    tol = 1e-6
    for info in tag_infos:
        signature = info.get("signature")
        if signature is None:
            continue
        target_sig = signature
        tag_entry = _pop_matching_tag_entry(entry_map, target_sig)
        if not tag_entry:
            continue
        offsets = _ensure_tag_offset_entry(tag_entry)
        delta = XYZ(
            _inches_to_feet(offsets.get("x_inches") or 0.0),
            _inches_to_feet(offsets.get("y_inches") or 0.0),
            _inches_to_feet(offsets.get("z_inches") or 0.0),
        )
        target_point = host_point + delta
        tag_elem = info.get("element")
        head_point = info.get("head_point")
        target_orientation = info.get("orientation")
        if not tag_elem or not head_point:
            continue

        move_vec = target_point - head_point
        try:
            move_len = move_vec.GetLength()
        except Exception:
            move_len = 0.0
        if move_len > tol:
            try:
                ElementTransformUtils.MoveElement(doc, tag_elem.Id, move_vec)
                head_point = target_point
            except Exception:
                head_point = target_point

        if rotation_delta_deg and abs(rotation_delta_deg) > tol:
            try:
                axis = Line.CreateBound(head_point, head_point + XYZ(0, 0, 1))
                ElementTransformUtils.RotateElement(doc, tag_elem.Id, axis, math.radians(rotation_delta_deg))
            except Exception:
                pass
        if target_orientation is not None:
            try:
                tag_elem.TagOrientation = target_orientation
            except Exception:
                pass


def _get_tag_host(tag):
    doc = getattr(tag, "Document", None) or revit.doc
    if doc is None:
        return None

    def _element_from_id(elem_id):
        if elem_id is None:
            return None
        target_id = getattr(elem_id, "ElementId", elem_id)
        try:
            return doc.GetElement(target_id)
        except Exception:
            return None

    try:
        local_ids = tag.GetTaggedLocalElementIds()
        if local_ids:
            for elem_id in local_ids:
                host = _element_from_id(elem_id)
                if host:
                    return host
    except Exception:
        pass
    try:
        refs = tag.GetTaggedReferences()
        if refs:
            for ref in refs:
                host = _element_from_id(getattr(ref, "ElementId", None))
                if host:
                    return host
    except Exception:
        pass
    try:
        tagged = getattr(tag, "TaggedElementId", None)
        if tagged:
            host = _element_from_id(tagged)
            if host:
                return host
    except Exception:
        pass
    return None


def _get_tag_metadata(tag):
    data = {
        "element": tag,
        "host_element": None,
        "host_point": None,
        "head_point": None,
        "family_name": None,
        "type_name": None,
        "category_name": None,
    }
    if not isinstance(tag, IndependentTag):
        return data
    host = _get_tag_host(tag)
    if host:
        data["host_element"] = host
        data["host_point"] = _get_point(host)
    data["head_point"] = _get_tag_point(tag)
    doc = getattr(tag, "Document", None) or revit.doc
    tag_symbol = None
    try:
        if doc:
            tag_symbol = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_symbol = None
    fam_name = None
    type_name = None
    category_name = None
    if tag_symbol:
        try:
            fam_name = getattr(tag_symbol, "FamilyName", None)
            if not fam_name:
                fam = getattr(tag_symbol, "Family", None)
                fam_name = getattr(fam, "Name", None) if fam else None
        except Exception:
            fam_name = None
        try:
            type_name = getattr(tag_symbol, "Name", None)
            if not type_name:
                param = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    type_name = param.AsString()
        except Exception:
            type_name = None
        try:
            cat = getattr(tag_symbol, "Category", None)
            category_name = getattr(cat, "Name", None) if cat else None
        except Exception:
            category_name = None
    if not category_name:
        try:
            cat = getattr(tag, "Category", None)
            category_name = getattr(cat, "Name", None) if cat else None
        except Exception:
            category_name = None
    data["family_name"] = fam_name
    data["type_name"] = type_name
    data["category_name"] = category_name
    return data


def _build_label(elem):
    fam_name = None
    type_name = None
    is_group = isinstance(elem, (Group, GroupType))

    if is_group:
        try:
            fam_name = getattr(elem, "Name", None)
            type_name = fam_name
        except Exception:
            fam_name = None
            type_name = None
    else:
        try:
            sym = getattr(elem, "Symbol", None) or getattr(elem, "GroupType", None)
            if sym:
                fam = getattr(sym, "Family", None)
                fam_name = getattr(fam, "Name", None) if fam else None
                type_name = getattr(sym, "Name", None)
                if not type_name:
                    try:
                        tparam = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if tparam:
                            type_name = tparam.AsString()
                    except Exception:
                        pass
        except Exception:
            fam_name = None
            type_name = None

    if fam_name and type_name:
        return u"{} : {}".format(fam_name, type_name)
    if type_name:
        return type_name
    if fam_name:
        return fam_name
    return ""


def _resolve_target_element():
    selection = None
    try:
        selection = revit.get_selection()
    except Exception:
        selection = None
    picked = None
    if selection:
        try:
            elems = list(selection.elements)
        except Exception:
            elems = []
        if elems:
            picked = elems[0]
    if picked:
        return picked
    try:
        return revit.pick_element(message="Select equipment to update vector")
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# Element_Linker helpers
# --------------------------------------------------------------------------- #


def _get_element_linker_payload(elem):
    param_names = (ELEMENT_LINKER_SHARED_PARAM, ELEMENT_LINKER_PARAM_NAME)
    for name in param_names:
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
        if value:
            stripped = value.strip()
            if stripped:
                return stripped
    return None


def _parse_xyz_string(text):
    if not text:
        return None
    parts = [p.strip() for p in text.split(",")]
    if len(parts) != 3:
        return None
    try:
        return XYZ(float(parts[0]), float(parts[1]), float(parts[2]))
    except Exception:
        return None


def _parse_float(value, default=None):
    if value is None:
        return default
    if isinstance(value, basestring):
        cleaned = value.strip()
        if not cleaned:
            return default
        try:
            return float(cleaned)
        except Exception:
            return default
    try:
        return float(value)
    except Exception:
        return default


def _parse_int(value, default=None):
    if value is None:
        return default
    if isinstance(value, basestring):
        cleaned = value.strip()
        if not cleaned:
            return default
        try:
            return int(cleaned)
        except Exception:
            return default
    try:
        return int(value)
    except Exception:
        return default


def _parse_element_linker_payload(payload_text):
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
        pattern = re.compile(
            r"(Linked Element Definition ID|Set Definition ID|Location XYZ \(ft\)|"
            r"Rotation \(deg\)|Parent Rotation \(deg\)|Parent ElementId|LevelId|"
            r"ElementId|FacingOrientation)\s*:\s*"
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")
    return {
        "led_id": entries.get("Linked Element Definition ID", "").strip(),
        "set_id": entries.get("Set Definition ID", "").strip(),
        "location": _parse_xyz_string(entries.get("Location XYZ (ft)")),
        "rotation_deg": _parse_float(entries.get("Rotation (deg)"), 0.0),
        "parent_rotation_deg": _parse_float(entries.get("Parent Rotation (deg)")),
        "parent_element_id": _parse_int(entries.get("Parent ElementId")),
        "facing": _parse_xyz_string(entries.get("FacingOrientation")),
        "raw": entries,
    }


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)


def _build_linker_payload(led_id, set_id, location, rotation_deg, level_id, element_id, facing, parent_rotation, parent_element_id=None):
    lines = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Location XYZ (ft): {}".format(_format_xyz(location)),
        "Rotation (deg): {:.6f}".format(rotation_deg or 0.0),
        "Parent Rotation (deg): {:.6f}".format(parent_rotation or 0.0),
        "Parent ElementId: {}".format(parent_element_id if parent_element_id is not None else ""),
        "LevelId: {}".format(level_id if level_id is not None else ""),
        "ElementId: {}".format(element_id if element_id is not None else ""),
        "FacingOrientation: {}".format(_format_xyz(facing)),
    ]
    return "\n".join(lines).strip()


def _set_element_linker_payload(elem, text_value):
    if not elem or not text_value:
        return False
    for name in (ELEMENT_LINKER_SHARED_PARAM, ELEMENT_LINKER_PARAM_NAME):
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        try:
            param.Set(text_value)
            return True
        except Exception:
            continue
    return False


def _write_element_linker_payload(elem, text_value, transaction_name=None):
    if not elem or not text_value:
        return False
    doc = getattr(elem, "Document", None) or revit.doc
    if doc is None:
        return False
    if getattr(doc, "IsModifiable", False):
        return _set_element_linker_payload(elem, text_value)
    txn = None
    try:
        txn = Transaction(doc, transaction_name or "Update Element_Linker payload")
        txn.Start()
        success = _set_element_linker_payload(elem, text_value)
        if success:
            txn.Commit()
        else:
            txn.RollBack()
        return success
    except Exception:
        if txn:
            try:
                txn.RollBack()
            except Exception:
                pass
        return False


def _collect_similar_elements(doc, led_id, original_elem):
    """
    Gather all non-type elements in the active document that share the same
    Linked Element Definition ID. Skips the originally selected element.
    """
    matches = []
    if doc is None:
        return matches
    target = (led_id or "").strip().lower()
    if not target:
        return matches
    try:
        collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    except Exception:
        return matches

    original_id = getattr(getattr(original_elem, "Id", None), "IntegerValue", None)
    for elem in collector:
        try:
            if original_id and elem.Id.IntegerValue == original_id:
                continue
        except Exception:
            pass
        payload_text = _get_element_linker_payload(elem)
        if not payload_text:
            continue
        payload = _parse_element_linker_payload(payload_text)
        cand_id = (payload.get("led_id") or "").strip().lower()
        if cand_id == target:
            matches.append(elem)
    return matches


def _apply_offsets_to_similar(doc, elements, original_local_offset, original_rotation_offset, new_local_offset, new_rotation_offset, rotation_delta_deg, tag_entries):
    """
    Move each similar element to match the updated offsets/rotation relative to
    their individual Element_Linker base points.
    """
    if doc is None or not elements:
        return 0, False, []
    local_offset = new_local_offset or XYZ(0, 0, 0)
    old_local_offset = original_local_offset or XYZ(0, 0, 0)
    tol = 1e-6
    rotation_offset = float(new_rotation_offset or 0.0)
    old_rotation_offset = float(original_rotation_offset or 0.0)
    rotation_change = abs(rotation_delta_deg or 0.0) > tol
    moved = []
    t = Transaction(doc, "Apply vector to similar equipment")
    try:
        t.Start()
        for elem in elements:
            elem_point = _get_point(elem)
            if not elem_point:
                continue
            payload_text = _get_element_linker_payload(elem)
            payload = _parse_element_linker_payload(payload_text)
            elem_rotation = _get_rotation_degrees(elem)
            parent_rot = payload.get("parent_rotation_deg")
            if parent_rot is not None:
                base_rotation = parent_rot
            else:
                base_rotation = elem_rotation - old_rotation_offset

            previous_world_offset = _rotate_xy(old_local_offset, base_rotation)
            base_point = elem_point - previous_world_offset

            target_rotation = base_rotation + rotation_offset
            world_offset = _rotate_xy(local_offset, base_rotation)
            target_point = base_point + world_offset

            move_vec = target_point - elem_point
            move_len = 0.0
            try:
                move_len = move_vec.GetLength()
            except Exception:
                move_len = 0.0
            if move_len > tol:
                ElementTransformUtils.MoveElement(doc, elem.Id, move_vec)
                elem_point = target_point

            rot_delta = _normalize_angle(target_rotation - elem_rotation)
            if rotation_change and abs(rot_delta) > tol:
                axis = Line.CreateBound(elem_point, elem_point + XYZ(0, 0, 1))
                ElementTransformUtils.RotateElement(doc, elem.Id, axis, math.radians(rot_delta))
                elem_rotation = target_rotation

            new_payload_text = _build_linker_payload(
                payload.get("led_id"),
                payload.get("set_id"),
                target_point,
                elem_rotation,
                getattr(getattr(elem, "LevelId", None), "IntegerValue", None),
                getattr(getattr(elem, "Id", None), "IntegerValue", None),
                getattr(elem, "FacingOrientation", None),
                base_rotation,
                payload.get("parent_element_id"),
            )
            _write_element_linker_payload(elem, new_payload_text)

            if tag_entries:
                _apply_tag_offsets_to_instance(
                    doc,
                    elem,
                    tag_entries,
                    target_point,
                    rotation_delta_deg if rotation_change else 0.0,
                )

            moved.append(elem)
        t.Commit()
        return len(moved), True, moved
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return len(moved), False, moved


# --------------------------------------------------------------------------- #
# Active YAML helpers
# --------------------------------------------------------------------------- #


def _get_element_param(elem, name):
    if elem is None or not name:
        return None
    try:
        param = elem.LookupParameter(name)
    except Exception:
        param = None
    if not param:
        return None
    try:
        if param.StorageType == StorageType.String:
            return param.AsString()
        if param.StorageType == StorageType.Double:
            return str(param.AsDouble())
        if param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
        return param.AsValueString()
    except Exception:
        return None


def _parameter_match_score(element, led_entry):
    if element is None or not isinstance(led_entry, dict):
        return 0
    params = led_entry.get("parameters") or {}
    score = 0
    for name, expected in params.items():
        expected_text = (str(expected or "").strip())
        if not expected_text:
            continue
        actual = _get_element_param(element, name)
        if actual is None:
            continue
        if actual.strip() == expected_text:
            score += 1
    return score


def _find_led_entry(data, led_id, set_hint=None, element=None):
    target = (led_id or "").strip().lower()
    if not target:
        return None
    set_target = (set_hint or "").strip().lower()
    fallback = None
    fallback_idx = None
    best_match = None
    best_idx = None
    best_score = -1
    for eq in data.get("equipment_definitions") or []:
        for set_entry in eq.get("linked_sets") or []:
            set_id = (set_entry.get("id") or "").strip()
            set_id_lower = set_id.lower()
            led_list = set_entry.get("linked_element_definitions") or []
            for idx, led_entry in enumerate(led_list):
                current_id = (led_entry.get("id") or led_entry.get("led_id") or "").strip()
                if not current_id:
                    continue
                if current_id.strip().lower() != target:
                    continue
                if set_target and set_id_lower == set_target:
                    return eq, set_entry, led_entry, idx
                score = _parameter_match_score(element, led_entry)
                if score > best_score:
                    best_match = (eq, set_entry, led_entry)
                    best_idx = idx
                    best_score = score
                if fallback is None:
                    fallback = (eq, set_entry, led_entry)
                    fallback_idx = idx
    if best_match is not None:
        return best_match[0], best_match[1], best_match[2], best_idx
    if fallback is None:
        return None
    return fallback[0], fallback[1], fallback[2], fallback_idx


def _ensure_offset_entry(led_entry):
    offsets = led_entry.setdefault("offsets", [])
    if not isinstance(offsets, list):
        offsets = [{}]
        led_entry["offsets"] = offsets
    if not offsets:
        offsets.append({})
    entry = offsets[0]
    if not isinstance(entry, dict):
        entry = {}
        offsets[0] = entry
    entry.setdefault("x_inches", 0.0)
    entry.setdefault("y_inches", 0.0)
    entry.setdefault("z_inches", 0.0)
    entry.setdefault("rotation_deg", 0.0)
    return entry


def _verify_saved_offsets(led_id, set_id, expected_entry):
    """Reload YAML after saving to confirm offsets persisted."""
    try:
        _, verify_data = load_active_yaml_data()
    except Exception as exc:
        try:
            LOG.warning("[Update Vector] verification reload failed for %s: %s", led_id, exc)
        except Exception:
            pass
        return False, "Reload failed: {}".format(exc)
    verify_lookup = _find_led_entry(verify_data, led_id, set_id)
    if not verify_lookup:
        try:
            LOG.warning("[Update Vector] verification could not find LED %s", led_id)
        except Exception:
            pass
        return False, "LED '{}' not found in refreshed YAML.".format(led_id)
    verify_entry = _ensure_offset_entry(verify_lookup[2])
    for key in ("x_inches", "y_inches", "z_inches", "rotation_deg"):
        exp = expected_entry.get(key, 0.0)
        act = verify_entry.get(key, 0.0)
        try:
            exp_val = float(exp)
        except Exception:
            exp_val = 0.0
        try:
            act_val = float(act)
        except Exception:
            act_val = 0.0
        if abs(exp_val - act_val) > 1e-6:
            try:
                LOG.warning("[Update Vector] verification mismatch for %s key=%s expected=%s actual=%s", led_id, key, exp_val, act_val)
            except Exception:
                pass
            return False, "Mismatch for {} (expected {}, found {})".format(key, exp_val, act_val)
    try:
        LOG.info(
            "[Update Vector] verification comparison matched for %s offsets=%s",
            led_id,
            {k: expected_entry.get(k) for k in ("x_inches", "y_inches", "z_inches", "rotation_deg")},
        )
    except Exception:
        pass
    return True, ""


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #


def main():
    selected = _resolve_target_element()
    if not selected:
        forms.alert("No element selected.", title=TITLE)
        return

    tag_only = False
    tag_signature_filter = None
    tag_meta = None
    elem = selected
    if isinstance(selected, IndependentTag):
        tag_meta = _get_tag_metadata(selected)
        host_elem = tag_meta.get("host_element")
        if not host_elem:
            forms.alert("Selected tag is not associated with a host element.", title=TITLE)
            return
        elem = host_elem
        tag_only = True
        tag_signature_filter = _tag_element_key(selected)

    elem_point = _get_point(elem)
    if tag_only:
        tag_host_point = tag_meta.get("host_point")
        if tag_host_point:
            elem_point = tag_host_point
    if not elem_point:
        forms.alert("Unable to read the element location.", title=TITLE)
        return

    elem_rotation = _get_rotation_degrees(elem)
    label = _build_label(elem) or (getattr(elem, "Name", None) or "")

    payload_text = _get_element_linker_payload(elem)
    if not payload_text:
        forms.alert("The selected element does not contain Element_Linker data. Re-run 'Add YAML Profiles' first.", title=TITLE)
        return
    payload = _parse_element_linker_payload(payload_text)
    led_id = payload.get("led_id")
    if not led_id:
        forms.alert("Element_Linker parameter is missing the 'Linked Element Definition ID' entry.", title=TITLE)
        return
    payload_location = payload.get("location")
    if not payload_location:
        forms.alert("Element_Linker parameter does not contain a valid base location.", title=TITLE)
        return
    payload_rotation = float(payload.get("rotation_deg") or 0.0)
    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)

    eq_entry = _find_led_entry(data, led_id, payload.get("set_id"), elem)
    if not eq_entry:
        forms.alert("Could not locate '{}' inside {}.".format(led_id, yaml_label), title=TITLE)
        return
    eq_def, set_entry, led_entry, led_index = eq_entry
    offset_entry = _ensure_offset_entry(led_entry)
    tag_entries = []

    original_local_offset = XYZ(
        _inches_to_feet(offset_entry.get("x_inches") or 0.0),
        _inches_to_feet(offset_entry.get("y_inches") or 0.0),
        _inches_to_feet(offset_entry.get("z_inches") or 0.0),
    )
    original_rotation_offset = float(offset_entry.get("rotation_deg") or 0.0)
    parent_rotation = payload.get("parent_rotation_deg")
    if parent_rotation is not None:
        base_rotation = parent_rotation
    else:
        base_rotation = payload_rotation - original_rotation_offset
    try:
        LOG.info(
            "[Update Vector] starting offsets led=%s label=%s inches=(%.3f, %.3f, %.3f) rot=%.3f",
            led_id,
            label,
            float(offset_entry.get("x_inches") or 0.0),
            float(offset_entry.get("y_inches") or 0.0),
            float(offset_entry.get("z_inches") or 0.0),
            original_rotation_offset,
        )
    except Exception:
        pass

    previous_world_offset = _rotate_xy(original_local_offset, base_rotation)
    base_point = payload_location - previous_world_offset
    delta_world = elem_point - base_point
    original_total_rotation = payload_rotation
    level_relative_z_inches = None
    if tag_only:
        rotation_delta = 0.0
        new_rotation_offset = original_rotation_offset
        total_rotation = base_rotation + new_rotation_offset
        local_offset = original_local_offset
        new_local_offset = XYZ(original_local_offset.X, original_local_offset.Y, original_local_offset.Z)
    else:
        new_rotation_offset = round(_normalize_angle(elem_rotation - base_rotation), 6)
        rotation_delta = _normalize_angle(new_rotation_offset - original_rotation_offset)
        local_offset = _rotate_xy(delta_world, -base_rotation)
        level_relative_z_inches = _level_relative_z_inches(elem, elem_point)
        local_offset = XYZ(local_offset.X, local_offset.Y, _inches_to_feet(level_relative_z_inches))
        total_rotation = base_rotation + new_rotation_offset
        new_local_offset = XYZ(local_offset.X, local_offset.Y, local_offset.Z)

    tags_changed = 0
    if led_entry:
        tags_changed = _update_tag_yaml_offsets(elem, led_entry, elem_point, tag_signature_filter)
        if tag_only and tags_changed == 0:
            forms.alert("No matching tag definition was found for the selected tag.", title=TITLE)
            return
        tag_entries = led_entry.get("tags") or []

    if not tag_only:
        offset_entry["x_inches"] = round(_feet_to_inches(local_offset.X if local_offset else 0.0), 6)
        offset_entry["y_inches"] = round(_feet_to_inches(local_offset.Y if local_offset else 0.0), 6)
        if level_relative_z_inches is None:
            level_relative_z_inches = _feet_to_inches(local_offset.Z if local_offset else 0.0)
        offset_entry["z_inches"] = round(level_relative_z_inches, 6)
        offset_entry["rotation_deg"] = new_rotation_offset
        try:
            LOG.info(
                "[Update Vector] updated offsets led=%s inches=(%.3f, %.3f, %.3f) rot=%.3f",
                led_id,
                offset_entry["x_inches"],
                offset_entry["y_inches"],
                offset_entry["z_inches"],
                offset_entry["rotation_deg"],
            )
        except Exception:
            pass
    if led_index is not None:
        les = set_entry.get("linked_element_definitions") or []
        if 0 <= led_index < len(les):
            les[led_index] = led_entry

    doc = getattr(revit, "doc", None)
    similar_elements = []
    apply_to_similar = False
    local_delta = XYZ(
        new_local_offset.X - original_local_offset.X,
        new_local_offset.Y - original_local_offset.Y,
        new_local_offset.Z - original_local_offset.Z,
    )
    delta_length = 0.0
    try:
        delta_length = local_delta.GetLength()
    except Exception:
        delta_length = 0.0
    should_propagate = (delta_length > 1e-6) or (abs(rotation_delta) > 1e-6) or (tags_changed > 0)
    if doc and led_id and should_propagate:
        similar_elements = _collect_similar_elements(doc, led_id, elem)
        if similar_elements:
            plural = "" if len(similar_elements) == 1 else "s"
            message = "Apply to all similar equipment? ({} additional instance{})".format(len(similar_elements), plural)
            apply_to_similar = bool(forms.alert(message, title=TITLE, yes=True, no=True))
        else:
            apply_to_similar = False
    propagate_requested = apply_to_similar
    if not apply_to_similar:
        similar_elements = []

    try:
        save_active_yaml_data(
            None,
            data,
            "Update Vector",
            "Updated offsets for {}".format(led_id),
        )
    except Exception as ex:
        forms.alert("Failed to update {}:\n\n{}".format(yaml_label, ex), title=TITLE)
        return

    if not tag_only:
        verify_ok, verify_msg = _verify_saved_offsets(led_id, payload.get("set_id"), offset_entry)
        if not verify_ok:
            forms.alert(
                "Warning: YAML verification after saving failed.\n{}\n"
                "Re-run Update Vector or review Extensible Storage history.".format(verify_msg),
                title=TITLE,
            )
        else:
            try:
                text = dump_data_to_string(data)
                ExtensibleStorage.update_active_text_only(revit.doc, data_path, text)
            except Exception as exc:
                LOG.warning("[Update Vector] failed to refresh active YAML snapshot: %s", exc)

    updated_payload_text = _build_linker_payload(
        led_id,
        payload.get("set_id"),
        elem_point,
        elem_rotation,
        getattr(getattr(elem, "LevelId", None), "IntegerValue", None),
        getattr(getattr(elem, "Id", None), "IntegerValue", None),
        getattr(elem, "FacingOrientation", None),
        base_rotation,
        payload.get("parent_element_id"),
    )
    _write_element_linker_payload(elem, updated_payload_text)

    moved_count = 0
    propagate_success = False
    moved_instances = []
    if apply_to_similar and similar_elements:
        moved_count, propagate_success, moved_instances = _apply_offsets_to_similar(
            doc,
            similar_elements,
            original_local_offset,
            original_rotation_offset,
            new_local_offset,
            new_rotation_offset,
            rotation_delta,
            tag_entries,
        )
        if propagate_success and moved_instances:
            try:
                selection = revit.get_selection()
                if selection:
                    selection.set_to([elem.Id for elem in moved_instances])
            except Exception:
                pass

    delta_inch_x = round(_feet_to_inches(local_delta.X), 6)
    delta_inch_y = round(_feet_to_inches(local_delta.Y), 6)
    delta_inch_z = round(_feet_to_inches(local_delta.Z), 6)

    def _format_float(val):
        try:
            return "{:.3f}".format(float(val))
        except Exception:
            return str(val)

    message = [
        "Updated vector for '{}' (LED: {})".format(label or led_id, led_id),
        "Rotation offset (deg): {}".format(_format_float(new_rotation_offset)),
        "Rotation change (deg): {}".format(_format_float(rotation_delta)),
        "Offsets (inches): X={}, Y={}, Z={}".format(
            _format_float(offset_entry["x_inches"]),
            _format_float(offset_entry["y_inches"]),
            _format_float(offset_entry["z_inches"]),
        ),
        "Applied move (inches): X={}, Y={}, Z={}".format(
            _format_float(delta_inch_x),
            _format_float(delta_inch_y),
            _format_float(delta_inch_z),
        ),
    ]
    if propagate_requested:
        message.append("")
        if propagate_success:
            message.append("Moved {} similar element{} by the same vector.".format(
                moved_count,
                "" if moved_count == 1 else "s",
            ))
            if moved_instances:
                message.append("The moved instances are now selected so you can review their locations.")
        elif not similar_elements:
            message.append("No additional similar elements were found.")
        else:
            message.append("Failed to move similar equipment; YAML offsets were still updated.")
    forms.alert("\n".join(message), title=TITLE)




if __name__ == '__main__':
    main()
