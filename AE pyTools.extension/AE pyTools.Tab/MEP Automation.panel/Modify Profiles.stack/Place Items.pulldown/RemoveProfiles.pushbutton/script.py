# -*- coding: utf-8 -*-
"""
Remove Profiles
---------------
Delete placed profile elements from the active view based on Element_Linker
Parent ElementId grouping.
"""

import math
import os
import re
import sys

from pyrevit import forms, revit, script
output = script.get_output()
output.close_others()
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    IndependentTag,
    TextNote,
    Transaction,
    ViewType,
    XYZ,
)
from System.Collections.Generic import List

LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if not os.path.isdir(LIB_ROOT):
    alt_root = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "..", "CEDLib.lib")
    )
    if os.path.isdir(alt_root):
        LIB_ROOT = alt_root
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402
from UIClasses.ProfileSelectionWindow import show_profile_selection_window  # noqa: E402

TITLE = "Remove Profiles"
LOG = script.get_logger()
ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
PAYLOAD_PATTERN = re.compile(
    r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|Location XYZ \(ft\)|"
    r"Rotation \(deg\)|Parent Rotation \(deg\)|Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation)\s*:\s*",
    re.IGNORECASE,
)


def _get_element_linker_payload(elem):
    if elem is None:
        return None
    for name in ELEMENT_LINKER_PARAM_NAMES:
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
    return None


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
            entries[key.strip().lower()] = remainder.strip()
    else:
        matches = list(PAYLOAD_PATTERN.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1).strip().lower()
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")

    def _as_int(value):
        try:
            return int(value)
        except Exception:
            try:
                return int(float(value))
            except Exception:
                return None

    def _parse_xyz(value):
        if not value:
            return None
        text = str(value).strip()
        if not text or "not found" in text.lower():
            return None
        parts = [p.strip() for p in text.split(",")]
        if len(parts) != 3:
            return None
        try:
            return (float(parts[0]), float(parts[1]), float(parts[2]))
        except Exception:
            return None

    return {
        "led_id": (entries.get("linked element definition id") or "").strip(),
        "set_id": (entries.get("set definition id") or "").strip(),
        "host_name": (entries.get("host name") or "").strip(),
        "parent_element_id": _as_int(entries.get("parent elementid") or entries.get("parent element id")),
        "parent_location": _parse_xyz(entries.get("parent_location")),
        "location": _parse_xyz(entries.get("location xyz (ft)")),
    }


def _collect_profile_placements(doc, view):
    placements = {}
    if not doc or view is None:
        return placements
    try:
        elems = list(FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType())
    except Exception:
        elems = []
    for elem in elems:
        payload = _get_element_linker_payload(elem)
        if not payload:
            continue
        parsed = _parse_linker_payload(payload)
        parent_id = parsed.get("parent_element_id")
        parent_loc = parsed.get("parent_location")
        host_name = parsed.get("host_name") or ""
        try:
            elem_id_val = elem.Id.IntegerValue
        except Exception:
            elem_id_val = None
        host_key = host_name.lower()
        if parent_id is not None:
            key = ("id", host_key, parent_id)
        else:
            # Independent profiles: group by profile name so all placements
            # of that profile can be removed together.
            if host_key:
                key = ("ind", host_key)
            elif parent_loc:
                parent_loc_key = (
                    round(parent_loc[0], 6),
                    round(parent_loc[1], 6),
                    round(parent_loc[2], 6),
                )
                key = ("loc", parent_loc_key)
            elif elem_id_val is not None:
                key = ("elem", elem_id_val)
            else:
                continue
        entry = placements.get(key)
        if entry is None:
            entry = {
                "host_name": host_name,
                "parent_id": parent_id,
                "parent_location": parent_loc,
                "element_ids": [],
            }
            placements[key] = entry
        try:
            if elem_id_val is not None:
                entry["element_ids"].append(elem_id_val)
        except Exception:
            continue
    return placements


def _build_option_label(entry):
    host_name = entry.get("host_name") or "<Unknown Profile>"
    parent_id = entry.get("parent_id")
    parent_loc = entry.get("parent_location")
    count = len(entry.get("element_ids") or [])
    if parent_id is not None:
        parent_label = "Parent {}".format(parent_id)
    elif parent_loc:
        parent_label = "Independent @ {:.3f},{:.3f},{:.3f}".format(
            parent_loc[0], parent_loc[1], parent_loc[2]
        )
    else:
        parent_label = "Independent"
    return "{} ({}) - {} item(s)".format(host_name, parent_label, count)


def _select_profiles(options, xaml_path):
    ok, selected = show_profile_selection_window(xaml_path, options)
    if not ok:
        return []
    return selected or []


def _normalize_keynote_family(value):
    if not value:
        return ""
    text = str(value)
    if ":" in text:
        text = text.split(":", 1)[0]
    return "".join([ch for ch in text.lower() if ch.isalnum()])


def _is_ga_keynote_symbol(family_name):
    return _normalize_keynote_family(family_name) == "gakeynotesymbolced"


def _get_point(elem):
    if elem is None:
        return None
    loc = getattr(elem, "Location", None)
    if loc is None:
        try:
            coord = getattr(elem, "Coord", None)
        except Exception:
            coord = None
        if coord is not None:
            return coord
        return None
    try:
        if hasattr(loc, "Point") and loc.Point is not None:
            return loc.Point
    except Exception:
        pass
    try:
        if hasattr(loc, "Curve") and loc.Curve is not None:
            return loc.Curve.Evaluate(0.5, True)
    except Exception:
        pass
    return None


def _get_rotation_degrees(elem):
    loc = getattr(elem, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            pass
    try:
        transform = elem.GetTransform()
    except Exception:
        transform = None
    if transform is not None:
        basis = getattr(transform, "BasisX", None)
        if basis:
            try:
                return math.degrees(math.atan2(basis.Y, basis.X))
            except Exception:
                pass
    return 0.0


def _offsets_to_ft(offsets):
    if not offsets:
        return (0.0, 0.0, 0.0)
    if isinstance(offsets, dict):
        try:
            return (
                float(offsets.get("x_inches", 0.0) or 0.0) / 12.0,
                float(offsets.get("y_inches", 0.0) or 0.0) / 12.0,
                float(offsets.get("z_inches", 0.0) or 0.0) / 12.0,
            )
        except Exception:
            return (0.0, 0.0, 0.0)
    if isinstance(offsets, (list, tuple)) and len(offsets) >= 3:
        try:
            return (float(offsets[0]), float(offsets[1]), float(offsets[2]))
        except Exception:
            return (0.0, 0.0, 0.0)
    return (0.0, 0.0, 0.0)


def _rotate_offset(offsets, rotation_deg):
    try:
        ang = math.radians(rotation_deg or 0.0)
    except Exception:
        ang = 0.0
    cos_a = math.cos(ang)
    sin_a = math.sin(ang)
    ox, oy, oz = offsets
    return (ox * cos_a - oy * sin_a, ox * sin_a + oy * cos_a, oz)


def _distance_xy(a, b):
    if a is None or b is None:
        return None
    try:
        dx = a.X - b.X
        dy = a.Y - b.Y
        return math.sqrt((dx * dx) + (dy * dy))
    except Exception:
        return None


def _tag_symbol_label(tag):
    if tag is None:
        return ("", "")
    doc = getattr(tag, "Document", None)
    tag_type = None
    if doc is not None:
        try:
            tag_type = doc.GetElement(tag.GetTypeId())
        except Exception:
            tag_type = None
    fam_name = None
    type_name = None
    if tag_type:
        try:
            fam = getattr(tag_type, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else getattr(tag_type, "FamilyName", None)
        except Exception:
            fam_name = None
        try:
            type_name = getattr(tag_type, "Name", None)
        except Exception:
            type_name = None
    return (fam_name or "", type_name or "")


def _family_instance_label(elem):
    if elem is None:
        return ("", "")
    symbol = getattr(elem, "Symbol", None)
    fam_name = ""
    type_name = ""
    if symbol is not None:
        try:
            fam = getattr(symbol, "Family", None)
            fam_name = getattr(fam, "Name", None) if fam else ""
        except Exception:
            fam_name = ""
        try:
            type_name = getattr(symbol, "Name", None) or ""
        except Exception:
            type_name = ""
    return (fam_name or "", type_name or "")


def _note_type_label(note):
    if note is None:
        return ("", "")
    doc = getattr(note, "Document", None)
    type_elem = None
    try:
        type_id = note.GetTypeId()
    except Exception:
        type_id = None
    if doc is not None and type_id:
        try:
            type_elem = doc.GetElement(type_id)
        except Exception:
            type_elem = None
    if not type_elem:
        return ("", "")
    try:
        type_name = getattr(type_elem, "Name", None) or ""
    except Exception:
        type_name = ""
    family_name = ""
    try:
        fam = getattr(type_elem, "Family", None)
        family_name = getattr(fam, "Name", None) if fam else ""
    except Exception:
        family_name = ""
    return (family_name or "", type_name or "")


def _normalize_label(value):
    return " ".join(str(value or "").strip().lower().split())


def _normalize_text(value):
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return " ".join(text.strip().lower().split())


def _matches_label(def_family, def_type, cand_family, cand_type):
    def_family = _normalize_label(def_family)
    def_type = _normalize_label(def_type)
    cand_family = _normalize_label(cand_family)
    cand_type = _normalize_label(cand_type)
    if def_family and cand_family and def_family != cand_family:
        return False
    if def_type and cand_type and def_type != cand_type:
        return False
    return True


def _build_led_annotation_map(data):
    led_map = {}
    if not data:
        return led_map
    for eq in data.get("equipment_definitions") or []:
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_id = (led.get("id") or "").strip()
                if not led_id:
                    continue
                led_map[led_id] = {
                    "tags": list(led.get("tags") or []),
                    "keynotes": list(led.get("keynotes") or []),
                    "text_notes": list(led.get("text_notes") or []),
                }
    return led_map


def _build_profile_text_signatures(data):
    signatures = {}
    if not data:
        return signatures
    for eq in data.get("equipment_definitions") or []:
        profile_name = (eq.get("name") or eq.get("id") or "").strip()
        if not profile_name:
            continue
        key = _normalize_label(profile_name)
        if not key:
            continue
        entries = signatures.setdefault(key, set())
        for linked_set in eq.get("linked_sets") or []:
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                for note_def in led.get("text_notes") or []:
                    if not isinstance(note_def, dict):
                        continue
                    note_text = _normalize_text(note_def.get("text") or "")
                    if not note_text:
                        continue
                    type_name = _normalize_label(note_def.get("type_name") or "")
                    entries.add((note_text, type_name))
    return signatures


def _collect_annotation_candidates(doc, view, search_all_views=False):
    tags = []
    notes = []
    annos = []
    try:
        if search_all_views:
            tags = list(FilteredElementCollector(doc).OfClass(IndependentTag))
        else:
            tags = list(FilteredElementCollector(doc, view.Id).OfClass(IndependentTag))
    except Exception:
        tags = []
    try:
        if search_all_views:
            notes = list(FilteredElementCollector(doc).OfClass(TextNote))
        else:
            notes = list(FilteredElementCollector(doc, view.Id).OfClass(TextNote))
    except Exception:
        notes = []
    try:
        if search_all_views:
            annos = list(FilteredElementCollector(doc).OfClass(FamilyInstance))
        else:
            annos = list(FilteredElementCollector(doc, view.Id).OfClass(FamilyInstance))
    except Exception:
        annos = []
    generic_annos = []
    for inst in annos:
        try:
            cat = getattr(inst, "Category", None)
            cat_name = getattr(cat, "Name", "") if cat else ""
        except Exception:
            cat_name = ""
        if str(cat_name).strip().lower() == "generic annotations":
            generic_annos.append(inst)
    return tags, notes, generic_annos


def _collect_annotation_ids(
    doc,
    view,
    selected_entries,
    led_map,
    max_distance_ft=3.0,
    loose_text_notes=False,
    ultra_loose_text_notes=False,
    search_all_views=False,
    profile_text_signatures=None,
):
    if not doc or view is None or not selected_entries:
        return set()
    tag_candidates, note_candidates, anno_candidates = _collect_annotation_candidates(doc, view, search_all_views=search_all_views)
    if not tag_candidates and not note_candidates and not anno_candidates:
        return set()
    max_dist = float(max_distance_ft or 0.0)
    if max_dist <= 0.0:
        return set()

    collected = set()
    note_signatures = set()
    if ultra_loose_text_notes:
        for entry in selected_entries:
            for elem_id in entry.get("element_ids") or []:
                try:
                    elem = doc.GetElement(ElementId(int(elem_id)))
                except Exception:
                    elem = None
                if elem is None:
                    continue
                payload = _get_element_linker_payload(elem)
                parsed = _parse_linker_payload(payload) if payload else {}
                led_id = (parsed.get("led_id") or "").strip()
                if led_id and led_map:
                    ann_def = led_map.get(led_id)
                    if ann_def:
                        for note_def in ann_def.get("text_notes") or []:
                            if not isinstance(note_def, dict):
                                continue
                            note_text = _normalize_text(note_def.get("text") or "")
                            if not note_text:
                                continue
                            type_name = _normalize_label(note_def.get("type_name") or "")
                            note_signatures.add((note_text, type_name))
                if profile_text_signatures:
                    host_name = entry.get("display_name") or entry.get("host_name") or parsed.get("host_name") or ""
                    host_key = _normalize_label(host_name)
                    for sig in profile_text_signatures.get(host_key, set()):
                        note_signatures.add(sig)
    for entry in selected_entries:
        for elem_id in entry.get("element_ids") or []:
            try:
                elem = doc.GetElement(ElementId(int(elem_id)))
            except Exception:
                elem = None
            if elem is None:
                continue
            payload = _get_element_linker_payload(elem)
            parsed = _parse_linker_payload(payload) if payload else {}
            host_point = _get_point(elem)
            if host_point is None and parsed.get("location"):
                try:
                    loc = parsed.get("location")
                    host_point = XYZ(loc[0], loc[1], loc[2])
                except Exception:
                    host_point = None
            if host_point is None:
                continue
            host_rot = _get_rotation_degrees(elem)
            led_id = (parsed.get("led_id") or "").strip()
            if not led_id or not led_map:
                continue
            ann_def = led_map.get(led_id)
            if not ann_def:
                continue

            for tag_def in list(ann_def.get("tags") or []) + list(ann_def.get("keynotes") or []):
                if not isinstance(tag_def, dict):
                    continue
                family = tag_def.get("family_name") or tag_def.get("family") or ""
                type_name = tag_def.get("type_name") or tag_def.get("type") or ""
                offsets = tag_def.get("offsets") or tag_def.get("offset") or {}
                offset_ft = _offsets_to_ft(offsets)
                if _is_ga_keynote_symbol(family):
                    offset_ft = _rotate_offset(offset_ft, host_rot)
                    candidates = anno_candidates
                else:
                    candidates = tag_candidates
                expected = XYZ(
                    host_point.X + offset_ft[0],
                    host_point.Y + offset_ft[1],
                    host_point.Z + offset_ft[2],
                )
                for cand in candidates:
                    if isinstance(cand, IndependentTag):
                        cand_family, cand_type = _tag_symbol_label(cand)
                        cand_point = getattr(cand, "TagHeadPosition", None)
                    else:
                        cand_family, cand_type = _family_instance_label(cand)
                        cand_point = _get_point(cand)
                    if cand_point is None:
                        continue
                    if not _matches_label(family, type_name, cand_family, cand_type):
                        continue
                    dist = _distance_xy(expected, cand_point)
                    if dist is None or dist > max_dist:
                        continue
                    try:
                        collected.add(cand.Id.IntegerValue)
                    except Exception:
                        continue

            for note_def in ann_def.get("text_notes") or []:
                if not isinstance(note_def, dict):
                    continue
                offsets = note_def.get("offsets") or note_def.get("offset") or {}
                offset_ft = _offsets_to_ft(offsets)
                expected = XYZ(
                    host_point.X + offset_ft[0],
                    host_point.Y + offset_ft[1],
                    host_point.Z + offset_ft[2],
                )
                note_text = (note_def.get("text") or "").strip()
                note_text_norm = _normalize_text(note_text)
                def_type = (note_def.get("type_name") or "").strip()
                strict_hit = False
                for note in note_candidates:
                    cand_point = _get_point(note)
                    if cand_point is None:
                        continue
                    dist = _distance_xy(expected, cand_point)
                    if dist is None or dist > max_dist:
                        continue
                    if note_text:
                        try:
                            cand_text_norm = _normalize_text(note.Text or "")
                            if note_text_norm and cand_text_norm:
                                if (
                                    note_text_norm != cand_text_norm
                                    and note_text_norm not in cand_text_norm
                                    and cand_text_norm not in note_text_norm
                                ):
                                    continue
                            elif note_text_norm != cand_text_norm:
                                continue
                        except Exception:
                            continue
                    if def_type:
                        fam_name, type_name = _note_type_label(note)
                        candidate_labels = {_normalize_label(type_name)}
                        if fam_name and type_name:
                            candidate_labels.add(_normalize_label("{} : {}".format(fam_name, type_name)))
                        if _normalize_label(def_type) not in candidate_labels:
                            continue
                    try:
                        collected.add(note.Id.IntegerValue)
                        strict_hit = True
                    except Exception:
                        continue
                if loose_text_notes and not strict_hit and note_text_norm:
                    for note in note_candidates:
                        cand_point = _get_point(note)
                        if cand_point is None:
                            continue
                        dist = _distance_xy(expected, cand_point)
                        if dist is None or dist > max_dist:
                            continue
                        try:
                            cand_text_norm = _normalize_text(note.Text or "")
                        except Exception:
                            continue
                        if not cand_text_norm:
                            continue
                        if (
                            note_text_norm != cand_text_norm
                            and note_text_norm not in cand_text_norm
                            and cand_text_norm not in note_text_norm
                        ):
                            continue
                        try:
                            collected.add(note.Id.IntegerValue)
                        except Exception:
                            continue
    if ultra_loose_text_notes and note_signatures:
        for note in note_candidates:
            try:
                cand_text_norm = _normalize_text(note.Text or "")
            except Exception:
                cand_text_norm = ""
            if not cand_text_norm:
                continue
            for sig_text, sig_type in note_signatures:
                if sig_text and sig_text not in cand_text_norm and cand_text_norm not in sig_text:
                    continue
                try:
                    collected.add(note.Id.IntegerValue)
                except Exception:
                    pass
                break
    return collected


def _delete_elements(doc, element_ids):
    if not element_ids:
        return 0, []
    unique_ids = []
    seen = set()
    for raw_id in element_ids:
        try:
            int_id = int(raw_id)
        except Exception:
            continue
        if int_id in seen:
            continue
        seen.add(int_id)
        unique_ids.append(ElementId(int_id))
    if not unique_ids:
        return 0, []

    id_list = List[ElementId]()
    for eid in unique_ids:
        id_list.Add(eid)

    txn = Transaction(doc, "Remove Profiles")
    deleted_ids = []
    try:
        txn.Start()
        deleted_ids = list(doc.Delete(id_list) or [])
        txn.Commit()
    except Exception:
        try:
            txn.RollBack()
        except Exception:
            pass
        raise
    return len(unique_ids), deleted_ids


def main():
    doc = getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return
    view = getattr(doc, "ActiveView", None)
    if view is None:
        forms.alert("No active view detected.", title=TITLE)
        return
    if hasattr(view, "IsTemplate") and view.IsTemplate:
        forms.alert("Active view is a template. Open a plan view to remove profiles.", title=TITLE)
        return
    if hasattr(view, "ViewType") and view.ViewType == ViewType.ThreeD:
        forms.alert("Active view is 3D. Open a plan view to remove profiles.", title=TITLE)
        return

    placements = _collect_profile_placements(doc, view)
    if not placements:
        forms.alert(
            "No profile placements were found in the active view using Element_Linker.",
            title=TITLE,
        )
        return

    options = []
    for entry in placements.values():
        label = _build_option_label(entry)
        options.append({
            "label": label,
            "display_name": entry.get("host_name") or "",
            "host_name": entry.get("host_name") or "",
            "parent_id": entry.get("parent_id"),
            "element_ids": list(entry.get("element_ids") or []),
        })
    options.sort(key=lambda opt: (opt.get("host_name") or "", opt.get("parent_id") or 0))

    xaml_path = os.path.join(os.path.dirname(__file__), "RemoveProfilesWindow.xaml")
    selected = _select_profiles(options, xaml_path)
    if not selected:
        forms.alert("No profiles selected. Nothing to remove.", title=TITLE)
        return

    all_ids = []
    selected_labels = []
    for entry in selected:
        selected_labels.append(entry.get("label") or "")
        all_ids.extend(entry.get("element_ids") or [])
    if not all_ids:
        forms.alert("Selected profiles contain no elements in the active view.", title=TITLE)
        return

    led_map = {}
    profile_text_signatures = {}
    try:
        _, data = load_active_yaml_data()
        led_map = _build_led_annotation_map(data)
        profile_text_signatures = _build_profile_text_signatures(data)
    except Exception:
        led_map = {}
        profile_text_signatures = {}

    search_all_views = False
    try:
        search_all_views = bool(window.FindName("SearchAllViewsCheckBox").IsChecked)
    except Exception:
        search_all_views = False

    annotation_ids = _collect_annotation_ids(
        doc,
        view,
        selected,
        led_map,
        max_distance_ft=20.0,
        loose_text_notes=True,
        ultra_loose_text_notes=True,
        search_all_views=search_all_views,
        profile_text_signatures=profile_text_signatures,
    )
    if annotation_ids:
        all_ids.extend(list(annotation_ids))

    confirm = forms.alert(
        "Remove {} element(s) for {} profile placement(s) from the active view?".format(
            len(set(all_ids)), len(selected)
        ),
        title=TITLE,
        yes=True,
        no=True,
    )
    if not confirm:
        return

    try:
        requested_count, deleted_ids = _delete_elements(doc, all_ids)
    except Exception as exc:
        LOG.error("Remove Profiles failed: %s", exc)
        forms.alert("Removal failed: {}".format(exc), title=TITLE)
        return

    removed_count = len(deleted_ids or [])
    forms.alert(
        "Removed {} element(s) from the active view.".format(removed_count),
        title=TITLE,
    )


if __name__ == "__main__":
    main()
