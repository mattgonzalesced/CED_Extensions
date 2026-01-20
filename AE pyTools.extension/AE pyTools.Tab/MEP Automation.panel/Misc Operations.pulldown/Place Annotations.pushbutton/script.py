# -*- coding: utf-8 -*-
"""
Place Annotations (filtered)
----------------------------
Loads tag, keynote, and text note definitions from the active YAML stored in
Extensible Storage and allows the user to place them on matching family
instances in the active view. Annotations are filtered so that only instances
that belong to the same dev-Group (via parameter value) are annotated.
"""

import math
import os
import sys
import re

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInParameter,
    FamilyInstance,
    FilteredElementCollector,
    Transaction,
    ViewType,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.PlaceElementsLogic import (  # noqa: E402
    PlaceElementsEngine,
    ProfileRepository,
    tag_key_from_dict,
)
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
TITLE = "Place Annotations"

try:
    basestring
except NameError:
    basestring = str


def _build_repository(data):
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def _parse_linker_payload(text):
    if not text:
        return {}
    payload = str(text)
    entries = {}
    if "\n" in payload:
        for raw_line in payload.splitlines():
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
        matches = list(pattern.finditer(payload))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(payload)
            value = payload[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")
    return {
        "led_id": entries.get("Linked Element Definition ID", "").strip(),
        "set_id": entries.get("Set Definition ID", "").strip(),
    }


def _extract_led_id_from_linked_def(linked_def):
    if not linked_def:
        return None
    raw = (
        linked_def.get_static_param("Element_Linker Parameter")
        or linked_def.get_static_param("Element_Linker")
    )
    payload = _parse_linker_payload(raw)
    return payload.get("led_id") or None


def _collect_tag_entries(repo, tag_filter=None):
    grouped = {}
    for equipment_name in repo.cad_names():
        labels = repo.labels_for_cad(equipment_name)
        for label in labels:
            linked_def = repo.definition_for_label(equipment_name, label)
            if not linked_def:
                continue
            placement = linked_def.get_placement()
            if not placement:
                continue
            group_id = (
                linked_def.get_static_param("dev-Group ID")
                or linked_def.get_static_param("dev_Group ID")
            )
            family_label = None
            try:
                fam = linked_def.get_family()
                typ = linked_def.get_type()
                if fam and typ:
                    family_label = u"{} : {}".format(fam, typ)
            except Exception:
                family_label = None

            for tag in placement.get_tags() or []:
                if tag_filter and not tag_filter(tag):
                    continue
                if not _has_tag_definition(tag):
                    continue
                key = tag_key_from_dict(tag)
                fam = tag.get("family") or tag.get("family_name") or "<Family?>"
                typ = tag.get("type") or tag.get("type_name") or "<Type?>"
                if not key:
                    key = (
                        (tag.get("category") or tag.get("category_name") or "").lower(),
                        fam,
                        typ,
                    )
                entry = grouped.get(key)
                if not entry:
                    entry = {
                        "key": key,
                        "tag_family": fam,
                        "tag_type": typ,
                        "contexts": [],
                    }
                    grouped[key] = entry
                if isinstance(tag, dict):
                    tag_copy = dict(tag)
                else:
                    tag_copy = {
                        "family": fam,
                        "type": typ,
                        "category": tag.get("category") if hasattr(tag, "get") else None,
                    }
                entry["contexts"].append({
                    "equipment_name": equipment_name,
                    "label": label,
                    "canonical": family_label,
                    "group_id": group_id,
                    "led_id": _extract_led_id_from_linked_def(linked_def),
                    "tag": tag_copy,
                })

    entries = []
    for entry in grouped.values():
        contexts = entry.get("contexts") or []
        fam = entry.get("tag_family") or "<Family?>"
        typ = entry.get("tag_type") or "<Type?>"
        if len(contexts) == 1:
            ctx = contexts[0]
            entry["display"] = u"{family} : {type}  ({equip} :: {label})".format(
                family=fam,
                type=typ,
                equip=ctx.get("equipment_name") or "<Equipment?>",
                label=ctx.get("label") or "<Label?>",
            )
        else:
            equip_count = len({ctx.get("equipment_name") for ctx in contexts if ctx.get("equipment_name")})
            label_count = len(contexts)
            entry["display"] = u"{family} : {type}  ({labels} labels / {defs} equipment definitions)".format(
                family=fam,
                type=typ,
                labels=label_count,
                defs=equip_count or 1,
            )
        entries.append(entry)
    entries.sort(key=lambda e: (e.get("display") or "").lower())
    return entries


def _get_label(inst):
    try:
        symbol = inst.Symbol
    except Exception:
        symbol = None
    fam_name = None
    type_name = None
    if symbol:
        try:
            fam = symbol.Family
            fam_name = fam.Name if fam else None
        except Exception:
            fam_name = None
        try:
            type_name = symbol.Name
            if not type_name:
                param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    type_name = param.AsString()
        except Exception:
            type_name = None
    if fam_name and type_name:
        return u"{} : {}".format(fam_name, type_name)
    return (type_name or fam_name or "").strip()


def _base_label(label):
    if not label:
        return ""
    if "#" in label:
        return label.split("#", 1)[0].strip()
    return label.strip()


def _canonical_from_label(label):
    if not label:
        return ""
    if "::" in label:
        return label.split("::")[-1].strip()
    return label.strip()


def _normalize_key(value):
    if not value:
        return ""
    return " ".join(str(value).lower().replace("_", " ").split())


def _has_tag_definition(tag_dict):
    if not isinstance(tag_dict, dict):
        return False
    keys = (
        "family",
        "family_name",
        "type",
        "type_name",
        "category",
        "category_name",
    )
    return any(tag_dict.get(key) for key in keys)


def _normalize_keynote_family(value):
    if not value:
        return ""
    text = str(value)
    if ":" in text:
        text = text.split(":", 1)[0]
    return "".join([ch for ch in text.lower() if ch.isalnum()])


def _is_ga_keynote_symbol(family_name):
    return _normalize_keynote_family(family_name) == "gakeynotesymbolced"


def _is_keynote_entry(tag_entry):
    if not tag_entry:
        return False
    if isinstance(tag_entry, dict):
        family = tag_entry.get("family_name") or tag_entry.get("family") or ""
    else:
        family = getattr(tag_entry, "family_name", None) or getattr(tag_entry, "family", None) or ""
    return _is_ga_keynote_symbol(family)


def _has_text_note_definition(note_dict):
    if not isinstance(note_dict, dict):
        return False
    text_value = (note_dict.get("text") or "").strip()
    type_name = (note_dict.get("type_name") or "").strip()
    return bool(text_value or type_name)


def _text_note_preview(text_value, limit=60):
    text_value = (text_value or "").replace("\r", " ").replace("\n", " ").strip()
    if not text_value:
        text_value = "<text note>"
    if limit and len(text_value) > limit:
        text_value = text_value[: limit - 3] + "..."
    return text_value


def _text_note_group_key(note_dict):
    if not isinstance(note_dict, dict):
        return None
    text_value = (note_dict.get("text") or "").strip()
    type_name = (note_dict.get("type_name") or "").strip()
    width_val = note_dict.get("width")
    try:
        width_val = round(float(width_val), 6) if width_val not in (None, "") else None
    except Exception:
        width_val = None
    return (text_value, type_name, width_val)


def _text_note_instance_key(note_dict):
    if not isinstance(note_dict, dict):
        return None
    text_value = (note_dict.get("text") or "").strip()
    type_name = (note_dict.get("type_name") or "").strip()
    width_val = note_dict.get("width")
    try:
        width_val = float(width_val or 0.0)
    except Exception:
        width_val = 0.0
    offsets = note_dict.get("offset") or (0.0, 0.0, 0.0)
    try:
        offset_key = (float(offsets[0]), float(offsets[1]), float(offsets[2]))
    except Exception:
        offset_key = tuple(offsets) if isinstance(offsets, (list, tuple)) else ()
    try:
        rotation = float(note_dict.get("rotation_deg") or 0.0)
    except Exception:
        rotation = 0.0
    return (text_value, type_name, width_val, offset_key, rotation)


def _collect_text_note_entries(repo):
    grouped = {}
    for equipment_name in repo.cad_names():
        labels = repo.labels_for_cad(equipment_name)
        for label in labels:
            linked_def = repo.definition_for_label(equipment_name, label)
            if not linked_def:
                continue
            placement = linked_def.get_placement()
            if not placement:
                continue
            group_id = (
                linked_def.get_static_param("dev-Group ID")
                or linked_def.get_static_param("dev_Group ID")
            )
            family_label = None
            try:
                fam = linked_def.get_family()
                typ = linked_def.get_type()
                if fam and typ:
                    family_label = u"{} : {}".format(fam, typ)
            except Exception:
                family_label = None

            for note in placement.get_text_notes() or []:
                if not _has_text_note_definition(note):
                    continue
                key = _text_note_group_key(note)
                text_value = note.get("text") if isinstance(note, dict) else ""
                type_name = note.get("type_name") if isinstance(note, dict) else None
                entry = grouped.get(key)
                if not entry:
                    entry = {
                        "key": key,
                        "note_text": text_value,
                        "note_type": type_name,
                        "contexts": [],
                    }
                    grouped[key] = entry
                note_copy = dict(note) if isinstance(note, dict) else {
                    "text": text_value or "",
                    "type_name": type_name,
                }
                entry["contexts"].append({
                    "equipment_name": equipment_name,
                    "label": label,
                    "canonical": family_label,
                    "group_id": group_id,
                    "led_id": _extract_led_id_from_linked_def(linked_def),
                    "note": note_copy,
                })

    entries = []
    for entry in grouped.values():
        contexts = entry.get("contexts") or []
        preview = _text_note_preview(entry.get("note_text"))
        type_name = entry.get("note_type") or "<Type?>"
        if len(contexts) == 1:
            ctx = contexts[0]
            entry["display"] = u"Text Note: \"{text}\" ({type})  ({equip} :: {label})".format(
                text=preview,
                type=type_name,
                equip=ctx.get("equipment_name") or "<Equipment?>",
                label=ctx.get("label") or "<Label?>",
            )
        else:
            equip_count = len({ctx.get("equipment_name") for ctx in contexts if ctx.get("equipment_name")})
            label_count = len(contexts)
            entry["display"] = u"Text Note: \"{text}\" ({type})  ({labels} labels / {defs} equipment definitions)".format(
                text=preview,
                type=type_name,
                labels=label_count,
                defs=equip_count or 1,
            )
        entries.append(entry)
    entries.sort(key=lambda e: (e.get("display") or "").lower())
    return entries


def _get_point(inst):
    loc = getattr(inst, "Location", None)
    if not loc:
        return None
    if hasattr(loc, "Point") and loc.Point:
        return loc.Point
    if hasattr(loc, "Curve") and loc.Curve:
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            return None
    return None


def _get_rotation_degrees(inst):
    loc = getattr(inst, "Location", None)
    if loc is not None and hasattr(loc, "Rotation"):
        try:
            return math.degrees(loc.Rotation)
        except Exception:
            pass
    try:
        transform = inst.GetTransform()
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


def _collect_instance_lookup(doc, view, elements=None):
    lookup = {}
    symbol_lookup = {}

    def _add(key, inst):
        norm = _normalize_key(key)
        if not norm:
            return
        lookup.setdefault(norm, []).append(inst)

    def _add_symbol(fam, typ, inst):
        if not fam or not typ:
            return
        f_norm = _normalize_key(fam)
        t_norm = _normalize_key(typ)
        if not f_norm or not t_norm:
            return
        symbol_lookup.setdefault((f_norm, t_norm), []).append(inst)
        lookup.setdefault(f_norm, []).append(inst)

    if elements is None:
        instances = FilteredElementCollector(doc, view.Id).OfClass(FamilyInstance)
    else:
        instances = [elem for elem in elements if isinstance(elem, FamilyInstance)]
    for inst in instances:
        label = _get_label(inst)
        if not label:
            continue
        _add(label, inst)
        _add(_base_label(label), inst)
        _add(_canonical_from_label(label), inst)
        sym = getattr(inst, "Symbol", None)
        if sym:
            fam_name = None
            type_name = None
            try:
                fam = sym.Family
                fam_name = fam.Name if fam else None
            except Exception:
                fam_name = None
            try:
                type_name = sym.Name
                if not type_name:
                    param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if param:
                        type_name = param.AsString()
            except Exception:
                type_name = None
            if fam_name and type_name:
                _add_symbol(fam_name, type_name, inst)
    return lookup, symbol_lookup


def _filter_by_group_id(instances, group_id):
    if not group_id:
        return instances
    target = _normalize_key(group_id)
    filtered = []
    for inst in instances:
        param = inst.LookupParameter("dev-Group ID") or inst.LookupParameter("dev_Group ID")
        if not param:
            continue
        try:
            value = param.AsString()
        except Exception:
            try:
                value = param.AsValueString()
            except Exception:
                value = None
        if value and _normalize_key(value) == target:
            filtered.append(inst)
    return filtered


def _get_linker_payload_from_instance(inst):
    for name in LINKER_PARAM_NAMES:
        try:
            param = inst.LookupParameter(name)
        except Exception:
            param = None
        if not param:
            continue
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
        if text:
            payload = _parse_linker_payload(text)
            if payload.get("led_id"):
                return payload
    return {}


def _filter_by_led_id(instances, led_id):
    if not led_id:
        return instances
    target = led_id.strip().lower()
    if not target:
        return instances
    filtered = []
    missing_payload = []
    for inst in instances:
        payload = _get_linker_payload_from_instance(inst)
        cand = (payload.get("led_id") or "").strip().lower()
        if not cand:
            missing_payload.append(inst)
        elif cand == target:
            filtered.append(inst)
    if filtered:
        return filtered
    return missing_payload


def _resolve_hosts(host_lookup, symbol_lookup, label, canonical):
    """
    Attempt to find host instances using multiple label variants and a final
    fallback that matches only by the family name.
    """
    tried = set()
    candidates = [label, _base_label(label), canonical]
    for candidate in candidates:
        norm = _normalize_key(candidate)
        if not norm or norm in tried:
            continue
        tried.add(norm)
        hosts = host_lookup.get(norm)
        if hosts:
            return hosts
    family = None
    if canonical and ":" in canonical:
        family = canonical.split(":", 1)[0]
    elif label and ":" in label:
        family = label.split(":", 1)[0]
    if family:
        fam_norm = _normalize_key(family)
        if fam_norm and fam_norm not in tried:
            hosts = host_lookup.get(fam_norm)
            if hosts:
                return hosts
    return []


def _select_entries(entries, title, view_name, button_name):
    if not entries:
        return []
    choices = [entry["display"] for entry in entries]
    selection = forms.SelectFromList.show(
        choices,
        title="Select {0} to place in '{1}'".format(title, view_name),
        multiselect=True,
        button_name=button_name,
    )
    if not selection:
        return None
    if isinstance(selection, basestring):
        selected_entries = [entry for entry in entries if entry["display"] == selection]
    else:
        selected_entries = [entry for entry in entries if entry["display"] in selection]
    if not selected_entries:
        return None
    return selected_entries


def _place_tag_entries(entries, engine, host_lookup, symbol_lookup, placed_tag_pairs, missing):
    total = 0
    for entry in entries:
        contexts = entry.get("contexts") or []
        for ctx in contexts:
            target_label = ctx.get("label")
            if not target_label:
                continue
            tag_def = ctx.get("tag")
            if not _has_tag_definition(tag_def):
                continue
            canonical = ctx.get("canonical")
            host_list = _resolve_hosts(host_lookup, symbol_lookup, target_label, canonical)
            if not host_list:
                missing.append("{0} :: {1}".format(ctx.get("equipment_name") or "<Equipment?>", target_label))
                continue
            hosts = _filter_by_group_id(host_list, ctx.get("group_id"))
            if not hosts:
                continue
            hosts = _filter_by_led_id(hosts, ctx.get("led_id"))
            if not hosts:
                continue

            tag_key = tag_key_from_dict(tag_def)
            for inst in hosts:
                host_id = getattr(getattr(inst, "Id", None), "IntegerValue", None)
                host_key = (host_id, tag_key)
                if host_key in placed_tag_pairs:
                    continue
                base_pt = _get_point(inst)
                if not base_pt:
                    continue
                rot_deg = _get_rotation_degrees(inst)
                engine._place_tags([tag_def], inst, base_pt, rot_deg)
                total += 1
                placed_tag_pairs.add(host_key)
    return total


def _place_text_note_entries(entries, engine, host_lookup, symbol_lookup, placed_note_pairs, missing):
    total = 0
    for entry in entries:
        contexts = entry.get("contexts") or []
        for ctx in contexts:
            target_label = ctx.get("label")
            if not target_label:
                continue
            note_def = ctx.get("note")
            if not _has_text_note_definition(note_def):
                continue
            canonical = ctx.get("canonical")
            host_list = _resolve_hosts(host_lookup, symbol_lookup, target_label, canonical)
            if not host_list:
                missing.append("{0} :: {1}".format(ctx.get("equipment_name") or "<Equipment?>", target_label))
                continue
            hosts = _filter_by_group_id(host_list, ctx.get("group_id"))
            if not hosts:
                continue
            hosts = _filter_by_led_id(hosts, ctx.get("led_id"))
            if not hosts:
                continue

            note_key = _text_note_instance_key(note_def)
            for inst in hosts:
                host_id = getattr(getattr(inst, "Id", None), "IntegerValue", None)
                dedupe_key = (host_id, note_key) if note_key is not None else None
                if dedupe_key and dedupe_key in placed_note_pairs:
                    continue
                base_pt = _get_point(inst)
                if not base_pt:
                    continue
                rot_deg = _get_rotation_degrees(inst)
                engine._place_text_notes([note_def], base_pt, rot_deg, host_instance=inst, host_location=base_pt)
                total += 1
                if dedupe_key:
                    placed_note_pairs.add(dedupe_key)
    return total


def main():
    doc = revit.doc
    active_view = getattr(doc, "ActiveView", None)
    if not active_view:
        forms.alert("No active view detected.", title=TITLE)
        return
    if active_view.ViewType == ViewType.ThreeD:
        forms.alert("Place Annotations only works in plan/elevation views.", title=TITLE)
        return

    try:
        data_path, data = load_active_yaml_data()
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    yaml_label = get_yaml_display_name(data_path)
    repo = _build_repository(data)
    tag_entries = _collect_tag_entries(repo, tag_filter=lambda tag: not _is_keynote_entry(tag))
    keynote_entries = _collect_tag_entries(repo, tag_filter=_is_keynote_entry)
    text_note_entries = _collect_text_note_entries(repo)
    if not tag_entries and not keynote_entries and not text_note_entries:
        forms.alert("No annotations found in {}.".format(yaml_label), title=TITLE)
        return

    annotation_choices = []
    if tag_entries:
        annotation_choices.append("Tags")
    if keynote_entries:
        annotation_choices.append("Keynotes")
    if text_note_entries:
        annotation_choices.append("Text Notes")
    if len(annotation_choices) == 1:
        selection = annotation_choices
    else:
        selection = forms.SelectFromList.show(
            annotation_choices,
            title="Select annotations to place in '{}'".format(active_view.Name),
            multiselect=True,
            button_name="Next",
        )
    if not selection:
        return
    if isinstance(selection, basestring):
        selection = [selection]

    selected_tags = []
    selected_keynotes = []
    selected_text_notes = []
    if "Tags" in selection:
        selected_tags = _select_entries(tag_entries, "tags", active_view.Name, "Place")
        if selected_tags is None:
            return
    if "Keynotes" in selection:
        selected_keynotes = _select_entries(keynote_entries, "keynotes", active_view.Name, "Place")
        if selected_keynotes is None:
            return
    if "Text Notes" in selection:
        selected_text_notes = _select_entries(text_note_entries, "text notes", active_view.Name, "Place")
        if selected_text_notes is None:
            return
    if not selected_tags and not selected_keynotes and not selected_text_notes:
        return

    selection = revit.get_selection()
    selected_elements = list(getattr(selection, "elements", []) or []) if selection else []
    if selected_elements:
        host_lookup, symbol_lookup = _collect_instance_lookup(doc, active_view, selected_elements)
        if not host_lookup and not symbol_lookup:
            forms.alert("No family instances found in the current selection.", title=TITLE)
            return
    else:
        host_lookup, symbol_lookup = _collect_instance_lookup(doc, active_view)
    if not host_lookup and not symbol_lookup:
        forms.alert("No family instances found in the active view.", title=TITLE)
        return

    tag_view_map = {}
    for entry in selected_tags + selected_keynotes:
        if entry.get("key"):
            tag_view_map.setdefault(entry["key"], []).append(active_view.Id.IntegerValue)

    engine = PlaceElementsEngine(doc, repo, tag_view_map=tag_view_map)
    total_tags = 0
    total_keynotes = 0
    total_text_notes = 0
    missing = []

    placed_tag_pairs = set()
    placed_note_pairs = set()
    t = Transaction(doc, "Place Annotations (YAML)")
    t.Start()
    try:
        if selected_tags:
            total_tags = _place_tag_entries(
                selected_tags,
                engine,
                host_lookup,
                symbol_lookup,
                placed_tag_pairs,
                missing,
            )
        if selected_keynotes:
            total_keynotes = _place_tag_entries(
                selected_keynotes,
                engine,
                host_lookup,
                symbol_lookup,
                placed_tag_pairs,
                missing,
            )
        if selected_text_notes:
            total_text_notes = _place_text_note_entries(
                selected_text_notes,
                engine,
                host_lookup,
                symbol_lookup,
                placed_note_pairs,
                missing,
            )
        t.Commit()
    except Exception as exc:
        t.RollBack()
        forms.alert("Failed while placing annotations:\n\n{0}".format(exc), title=TITLE)
        return

    summary = [
        "Placed annotations in view '{0}':".format(active_view.Name),
        " - Tags: {0}".format(total_tags),
        " - Keynotes: {0}".format(total_keynotes),
        " - Text Notes: {0}".format(total_text_notes),
    ]
    if missing:
        summary.append("")
        summary.append("No host elements found for:")
        sample = missing[:8]
        summary.extend(" - {0}".format(label) for label in sample)
        if len(missing) > len(sample):
            summary.append("   (+{0} more)".format(len(missing) - len(sample)))

    forms.alert("\n".join(summary), title=TITLE)


if __name__ == "__main__":
    main()
