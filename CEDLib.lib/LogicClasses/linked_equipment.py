# -*- coding: utf-8 -*-
"""Helpers for managing parent/child equipment relationships and offsets."""

import math
from Autodesk.Revit.DB import XYZ

RELATIONS_KEY = "linked_relations"


def _ensure_dict(target, key, default):
    value = target.get(key)
    if isinstance(value, type(default)):
        return value
    target[key] = default
    return default


def ensure_relations(eq_def):
    if not isinstance(eq_def, dict):
        raise ValueError("Equipment definition must be a dictionary")
    relations = eq_def.get(RELATIONS_KEY)
    if not isinstance(relations, dict):
        relations = {}
        eq_def[RELATIONS_KEY] = relations
    children = _ensure_dict(relations, "children", [])
    _ensure_dict(relations, "parent", {})
    # Ensure new keys exist on existing child entries
    for entry in children:
        if isinstance(entry, dict):
            entry.setdefault("anchor_offsets", {})
            entry.setdefault("anchor_led_id", None)
    return relations


def get_parent(eq_def):
    relations = eq_def.get(RELATIONS_KEY)
    if isinstance(relations, dict):
        parent = relations.get("parent")
        if isinstance(parent, dict):
            return parent
    return {}


def get_parent_id(eq_def):
    parent = get_parent(eq_def)
    return (parent.get("equipment_id") or "").strip()


def set_parent(eq_def, equipment_id, offsets=None, led_id=None):
    relations = ensure_relations(eq_def)
    if equipment_id:
        parent_entry = {
            "equipment_id": equipment_id,
            "offsets": dict(offsets or {}),
        }
        if led_id:
            parent_entry["parent_led_id"] = led_id
        relations["parent"] = parent_entry
    else:
        relations["parent"] = {}
    return relations["parent"]


def upsert_child(eq_def, equipment_id, offsets=None, anchor_offsets=None, anchor_led_id=None):
    relations = ensure_relations(eq_def)
    children = relations.setdefault("children", [])
    for entry in children:
        if entry.get("equipment_id") == equipment_id:
            entry["offsets"] = dict(offsets or {})
            if anchor_offsets is not None:
                entry["anchor_offsets"] = dict(anchor_offsets or {})
            if anchor_led_id:
                entry["anchor_led_id"] = anchor_led_id
            return entry
    entry = {
        "equipment_id": equipment_id,
        "offsets": dict(offsets or {}),
    }
    if anchor_offsets:
        entry["anchor_offsets"] = dict(anchor_offsets or {})
    if anchor_led_id:
        entry["anchor_led_id"] = anchor_led_id
    children.append(entry)
    return entry


def remove_child(eq_def, equipment_id):
    relations = ensure_relations(eq_def)
    children = relations.setdefault("children", [])
    relations["children"] = [entry for entry in children if entry.get("equipment_id") != equipment_id]


def find_equipment_by_id(data, equipment_id):
    target = (equipment_id or "").strip().lower()
    if not target:
        return None
    for eq_def in data.get("equipment_definitions") or []:
        eq_id = (eq_def.get("id") or "").strip().lower()
        if eq_id == target:
            return eq_def
    return None


def find_equipment_by_name(data, name):
    target = (name or "").strip().lower()
    if not target:
        return None
    for eq_def in data.get("equipment_definitions") or []:
        eq_name = (eq_def.get("name") or eq_def.get("id") or "").strip().lower()
        if eq_name == target:
            return eq_def
    return None


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


def _rotate_xy(vec, angle_deg):
    if not isinstance(vec, XYZ):
        return XYZ(0, 0, 0)
    try:
        angle_rad = math.radians(float(angle_deg))
    except Exception:
        angle_rad = 0.0
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)
    x = vec.X * cos_a - vec.Y * sin_a
    y = vec.X * sin_a + vec.Y * cos_a
    return XYZ(x, y, vec.Z)


def _normalize_angle(angle_deg):
    try:
        value = float(angle_deg)
    except Exception:
        value = 0.0
    while value > 180.0:
        value -= 360.0
    while value <= -180.0:
        value += 360.0
    return value


def compute_offsets_from_points(parent_point, parent_rotation_deg, child_point, child_rotation_deg):
    if parent_point is None or child_point is None:
        return {
            "x_inches": 0.0,
            "y_inches": 0.0,
            "z_inches": 0.0,
            "rotation_deg": 0.0,
        }
    delta = child_point - parent_point
    local_vec = _rotate_xy(delta, -float(parent_rotation_deg or 0.0))
    offsets = {
        "x_inches": round(_feet_to_inches(local_vec.X), 6),
        "y_inches": round(_feet_to_inches(local_vec.Y), 6),
        "z_inches": round(_feet_to_inches(local_vec.Z), 6),
        "rotation_deg": round(_normalize_angle((child_rotation_deg or 0.0) - (parent_rotation_deg or 0.0)), 6),
    }
    return offsets


def local_vector_from_offsets(offsets):
    return XYZ(
        _inches_to_feet((offsets or {}).get("x_inches") or 0.0),
        _inches_to_feet((offsets or {}).get("y_inches") or 0.0),
        _inches_to_feet((offsets or {}).get("z_inches") or 0.0),
    )


def target_point_from_offsets(parent_point, parent_rotation_deg, offsets):
    if parent_point is None:
        parent_point = XYZ(0, 0, 0)
    local_vec = local_vector_from_offsets(offsets)
    world_vec = _rotate_xy(local_vec, parent_rotation_deg or 0.0)
    return parent_point + world_vec


def child_rotation_from_offsets(parent_rotation_deg, offsets):
    rot_delta = (offsets or {}).get("rotation_deg") or 0.0
    return _normalize_angle((parent_rotation_deg or 0.0) + float(rot_delta))


def build_child_requests(repo, data, parent_eq_def, parent_point, parent_rotation_deg, anchor_led_id=None):
    """Return placement requests for a given parent equipment definition and optional anchor LED."""
    requests = []
    if repo is None or parent_eq_def is None:
        return requests
    relations = parent_eq_def.get(RELATIONS_KEY) or {}
    children = relations.get("children") or []
    anchor_norm = (anchor_led_id or "").strip().lower()
    for entry in children:
        if not isinstance(entry, dict):
            continue
        eq_id = entry.get("equipment_id")
        child_eq = find_equipment_by_id(data, eq_id)
        if not child_eq:
            continue
        entry_anchor = (entry.get("anchor_led_id") or "").strip().lower()
        if anchor_norm and entry_anchor != anchor_norm:
            continue
        child_name = (child_eq.get("name") or child_eq.get("id") or "").strip()
        if not child_name:
            continue
        labels = repo.labels_for_cad(child_name)
        if not labels:
            continue
        offsets = entry.get("offsets") or {}
        anchor_offsets = entry.get("anchor_offsets") or {}
        anchor_point = target_point_from_offsets(parent_point, parent_rotation_deg, anchor_offsets)
        target_point = target_point_from_offsets(anchor_point, parent_rotation_deg, offsets)
        rotation = child_rotation_from_offsets(parent_rotation_deg, offsets)
        requests.append({
            "equipment": child_eq,
            "equipment_id": child_eq.get("id"),
            "name": child_name,
            "labels": labels,
            "target_point": target_point,
            "rotation": rotation,
            "offsets": offsets,
        })
    return requests
