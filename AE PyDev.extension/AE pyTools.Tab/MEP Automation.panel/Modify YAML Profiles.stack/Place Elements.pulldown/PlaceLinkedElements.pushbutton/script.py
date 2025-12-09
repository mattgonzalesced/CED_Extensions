# -*- coding: utf-8 -*-
"""Place every equipment definition on matching linked elements automatically using the active YAML store."""

import math
import os
import sys

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import BuiltInParameter, FamilyInstance, FilteredElementCollector, Group, RevitLinkInstance, XYZ

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import PlaceElementsEngine, ProfileRepository  # noqa: E402
from profile_schema import equipment_defs_to_legacy  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

TITLE = "Place Linked Elements"
LOG = script.get_logger()

try:
    basestring
except NameError:
    basestring = str


def _build_repository(data):
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def _place_requests(doc, repo, selection_map, rows, default_level=None):
    if not selection_map or not rows:
        return {"placed": 0}
    engine = PlaceElementsEngine(
        doc,
        repo,
        default_level=default_level,
        allow_tags=False,
        transaction_name="Place Linked Elements",
    )
    return engine.place_from_csv(rows, selection_map)


def _build_row(name, point, rotation_deg):
    return {
        "Name": name,
        "Count": "1",
        "Position X": str(point.X * 12.0),
        "Position Y": str(point.Y * 12.0),
        "Position Z": str(point.Z * 12.0),
        "Rotation": str(rotation_deg or 0.0)
    }


def _normalize_name(value):
    if not value:
        return ""
    normalized = " ".join(str(value).strip().lower().split())
    return normalized


def _get_symbol(elem):
    symbol = getattr(elem, "Symbol", None)
    if symbol is not None:
        return symbol
    try:
        type_id = elem.GetTypeId()
    except Exception:
        type_id = None
    if type_id:
        doc = getattr(elem, "Document", None)
        if doc:
            try:
                return doc.GetElement(type_id)
            except Exception:
                return None
    return None


def _name_variants(elem):
    names = set()
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            names.add(raw_name)
    except Exception:
        pass
    if isinstance(elem, FamilyInstance):
        symbol = _get_symbol(elem)
        family = getattr(symbol, "Family", None) if symbol else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        family_name = getattr(family, "Name", None) if family else None
        if not family_name or not type_name:
            try:
                fam_param = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
            except Exception:
                fam_param = None
            if fam_param and not family_name:
                try:
                    fam_id = fam_param.AsElementId()
                except Exception:
                    fam_id = None
                if fam_id:
                    doc = getattr(elem, "Document", None)
                    if doc:
                        try:
                            fam_elem = doc.GetElement(fam_id)
                        except Exception:
                            fam_elem = None
                        if fam_elem is not None:
                            family_name = getattr(fam_elem, "Name", None)
                if not family_name:
                    try:
                        family_name = fam_param.AsValueString()
                    except Exception:
                        family_name = None
            try:
                type_param = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
            except Exception:
                type_param = None
            if type_param and not type_name:
                try:
                    type_id = type_param.AsElementId()
                except Exception:
                    type_id = None
                if type_id:
                    doc = getattr(elem, "Document", None)
                    if doc:
                        try:
                            type_elem = doc.GetElement(type_id)
                        except Exception:
                            type_elem = None
                        if type_elem is not None:
                            type_name = getattr(type_elem, "Name", None)
                if not type_name:
                    try:
                        type_name = type_param.AsValueString()
                    except Exception:
                        type_name = None
        if family_name and type_name:
            names.add(u"{} : {}".format(family_name, type_name))
            names.add(u"{} : {}".format(type_name, family_name))
        if type_name:
            names.add(type_name)
        if family_name:
            names.add(family_name)
    elif isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        group_name = getattr(gtype, "Name", None) if gtype else None
        if group_name:
            names.add(group_name)
    return {_normalize_name(name) for name in names if _normalize_name(name)}


def _parse_payload_pose(payload_text):
    if not payload_text:
        return None
    location = None
    rotation = None
    parent_rotation = None
    for raw_line in payload_text.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, _, remainder = line.partition(":")
        key = key.strip().lower()
        value = remainder.strip()
        if key.startswith("location xyz"):
            parts = [p.strip() for p in value.split(",")]
            if len(parts) == 3:
                try:
                    location = tuple(float(p) for p in parts)
                except Exception:
                    location = None
        elif key.startswith("parent rotation"):
            try:
                parent_rotation = float(value)
            except Exception:
                parent_rotation = None
        elif key.startswith("rotation"):
            try:
                rotation = float(value)
            except Exception:
                rotation = None
    if not location:
        return None
    point = XYZ(location[0], location[1], location[2])
    final_rotation = parent_rotation if parent_rotation is not None else (rotation or 0.0)
    return {"point": point, "rotation": final_rotation}


def _get_element_point(elem):
    location = getattr(elem, "Location", None)
    if not location:
        return None
    point = getattr(location, "Point", None)
    if point:
        return point
    curve = getattr(location, "Curve", None)
    if curve:
        try:
            return curve.Evaluate(0.5, True)
        except Exception:
            try:
                return curve.GetEndPoint(0)
            except Exception:
                return None
    return None


def _get_orientation_vector(elem):
    location = getattr(elem, "Location", None)
    if location and hasattr(location, "Rotation"):
        try:
            angle = float(location.Rotation)
            return XYZ(math.cos(angle), math.sin(angle), 0.0)
        except Exception:
            pass
    try:
        facing = getattr(elem, "FacingOrientation", None)
        if facing and (abs(facing.X) > 1e-9 or abs(facing.Y) > 1e-9):
            return XYZ(facing.X, facing.Y, 0.0)
    except Exception:
        pass
    try:
        hand = getattr(elem, "HandOrientation", None)
        if hand and (abs(hand.X) > 1e-9 or abs(hand.Y) > 1e-9):
            return XYZ(hand.X, hand.Y, 0.0)
    except Exception:
        pass
    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        basis = getattr(transform, "BasisX", None)
        if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
            return XYZ(basis.X, basis.Y, 0.0)
        basis = getattr(transform, "BasisY", None)
        if basis and (abs(basis.X) > 1e-9 or abs(basis.Y) > 1e-9):
            return XYZ(basis.X, basis.Y, 0.0)
    return None


def _angle_from_vector(vec):
    if not vec:
        return 0.0
    try:
        return math.degrees(math.atan2(vec.Y, vec.X))
    except Exception:
        return 0.0


def _transform_point(transform, point):
    if point is None or transform is None:
        return point
    try:
        return transform.OfPoint(point)
    except Exception:
        return point


def _collect_placeholders(doc, normalized_targets):
    placements = {}
    if not normalized_targets:
        return placements

    def _store(name, point, rotation):
        if not name or point is None:
            return
        placements.setdefault(name, []).append({
            "point": point,
            "rotation": rotation,
        })

    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for elem in collector:
        if not isinstance(elem, (FamilyInstance, Group)):
            continue
        variants = _name_variants(elem)
        if not variants:
            continue
        if not any(name in normalized_targets for name in variants):
            continue
        point = _get_element_point(elem)
        if point is None:
            continue
        rotation = _angle_from_vector(_get_orientation_vector(elem))
        for name in variants:
            if name in normalized_targets:
                _store(name, point, rotation)

    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_inst.GetLinkDocument()
        if link_doc is None:
            continue
        try:
            transform = link_inst.GetTotalTransform()
        except Exception:
            try:
                transform = link_inst.GetTransform()
            except Exception:
                transform = None
        linked_instances = FilteredElementCollector(link_doc).OfClass(FamilyInstance)
        for inst in linked_instances:
            variants = _name_variants(inst)
            if not variants:
                continue
            if not any(name in normalized_targets for name in variants):
                continue
            point = _transform_point(transform, _get_element_point(inst))
            if point is None:
                continue
            vec = _get_orientation_vector(inst)
            if vec and transform is not None:
                try:
                    vec = transform.OfVector(vec)
                except Exception:
                    pass
            rotation = _angle_from_vector(vec)
            for name in variants:
                if name in normalized_targets:
                    _store(name, point, rotation)

    return placements


def _anchor_rows_for_cad(repo, cad_name):
    anchors = []
    if hasattr(repo, "anchor_definitions_for_cad"):
        try:
            anchors = repo.anchor_definitions_for_cad(cad_name)
        except Exception:
            anchors = []
    rows = []
    for anchor in anchors or []:
        params = anchor.get_static_params() or {}
        payload = None
        for key in ("Element_Linker Parameter", "Element_Linker"):
            value = params.get(key)
            if value:
                payload = value
                break
        pose = _parse_payload_pose(payload)
        if pose:
            rows.append({"point": pose["point"], "rotation": pose["rotation"]})
    return rows


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)
    repo = _build_repository(data)

    equipment_names = repo.cad_names()
    if not equipment_names:
        forms.alert("No equipment definitions found in {}.".format(yaml_label), title=TITLE)
        return

    target_names = {_normalize_name(name) for name in equipment_names if _normalize_name(name)}
    placeholders = _collect_placeholders(doc, target_names)
    if not placeholders:
        forms.alert(
            "No linked elements were found whose names match equipment definitions. "
            "Verify the Revit model contains linked elements with matching names.",
            title=TITLE,
        )
        return

    rows = []
    selection_map = {}
    placed_defs = set()
    missing_labels = []
    for cad_name in equipment_names:
        normalized = _normalize_name(cad_name)
        matches = placeholders.get(normalized)
        if not matches:
            matches = _anchor_rows_for_cad(repo, cad_name)
        if not matches:
            continue
        labels = repo.labels_for_cad(cad_name)
        if not labels:
            missing_labels.append(cad_name)
            continue
        selection_map[cad_name] = labels
        placed_defs.add(cad_name)
        for match in matches:
            point = match.get("point")
            rotation = match.get("rotation")
            if point is None:
                continue
            rows.append(_build_row(cad_name, point, rotation))

    if not rows:
        forms.alert(
            "No placements were generated. Ensure equipment definitions include linked types "
            "and that matching linked elements exist in the model.",
            title=TITLE,
        )
        return

    level = None
    level_sel = forms.select_levels(multiple=False)
    if isinstance(level_sel, list) and level_sel:
        level = level_sel[0]
    elif level_sel:
        level = level_sel

    results = _place_requests(doc, repo, selection_map, rows, default_level=level)
    placed = results.get("placed", 0)
    summary = [
        "Processed {} linked element(s).".format(len(rows)),
        "Placed {} element(s).".format(placed),
    ]

    unmatched_defs = sorted(name for name in equipment_names if name not in placed_defs)
    if unmatched_defs:
        summary.append("")
        summary.append("No matching linked elements for:")
        sample = unmatched_defs[:5]
        summary.extend(" - {}".format(name) for name in sample)
        if len(unmatched_defs) > len(sample):
            summary.append("   (+{} more)".format(len(unmatched_defs) - len(sample)))

    if missing_labels:
        summary.append("")
        summary.append("Definitions missing linked types:")
        sample = missing_labels[:5]
        summary.extend(" - {}".format(name) for name in sample)
        if len(missing_labels) > len(sample):
            summary.append("   (+{} more)".format(len(missing_labels) - len(sample)))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
