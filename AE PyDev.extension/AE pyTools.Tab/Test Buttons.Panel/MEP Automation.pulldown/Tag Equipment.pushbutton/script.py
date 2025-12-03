# -*- coding: utf-8 -*-
"""
Tag Equipment (filtered)
------------------------
Loads tag definitions from profileData.yaml and allows the user to place tags on
matching family instances in the active view. Tags are filtered so that only
instances that belong to the same dev-Group (via parameter value) are tagged.
"""

import io
import math
import os
import sys

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
from profile_schema import equipment_defs_to_legacy, load_data as load_profile_data  # noqa: E402
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")
LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")

try:
    basestring
except NameError:
    basestring = str


def _simple_yaml_parse(text):
    lines = text.splitlines()

    def parse_block(start_idx, base_indent):
        idx = start_idx
        result = None
        while idx < len(lines):
            raw_line = lines[idx]
            stripped = raw_line.strip()
            if not stripped or stripped.startswith("#"):
                idx += 1
                continue
            indent = len(raw_line) - len(raw_line.lstrip(" "))
            if indent < base_indent:
                break
            if stripped.startswith("-"):
                if result is None:
                    result = []
                elif not isinstance(result, list):
                    break
                remainder = stripped[1:].strip()
                if remainder:
                    result.append(remainder)
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result.append(value)
            else:
                if result is None:
                    result = {}
                elif isinstance(result, list):
                    break
                key, _, remainder = stripped.partition(":")
                key = key.strip().strip('"')
                remainder = remainder.strip()
                if remainder:
                    result[key] = remainder
                    idx += 1
                else:
                    value, idx = parse_block(idx + 1, indent + 2)
                    result[key] = value
        if result is None:
            result = {}
        return result, idx

    parsed, _ = parse_block(0, 0)
    return parsed if isinstance(parsed, dict) else {}


def _pick_profile_data_path():
    cached = get_cached_yaml_path()
    if cached and os.path.exists(cached):
        return cached
    path = forms.pick_file(
        file_ext="yaml",
        title="Select profileData YAML file",
        init_dir=os.path.dirname(DEFAULT_DATA_PATH),
    )
    if path:
        set_cached_yaml_path(path)
    return path


def _load_profile_store(data_path):
    data = load_profile_data(data_path)
    if data.get("equipment_definitions"):
        return data
    try:
        with io.open(data_path, "r", encoding="utf-8") as handle:
            fallback = _simple_yaml_parse(handle.read())
        if fallback.get("equipment_definitions"):
            return fallback
    except Exception:
        pass
    return data


def _build_repository(data_path):
    data = _load_profile_store(data_path)
    legacy_profiles = equipment_defs_to_legacy(data.get("equipment_definitions") or [])
    eq_defs = ProfileRepository._parse_profiles(legacy_profiles)
    return ProfileRepository(eq_defs)


def _parse_linker_payload(text):
    if not text:
        return {}
    entries = {}
    for raw_line in str(text).splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, _, remainder = line.partition(":")
        entries[key.strip()] = remainder.strip()
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


def _collect_tag_entries(repo):
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
    if loc is None:
        return 0.0
    rot = getattr(loc, "Rotation", None)
    if rot:
        try:
            return math.degrees(rot)
        except Exception:
            return 0.0
    basis = getattr(inst, "FacingOrientation", None)
    if basis:
        try:
            return math.degrees(math.atan2(basis.Y, basis.X))
        except Exception:
            return 0.0
    return 0.0


def _collect_instance_lookup(doc, view):
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

    collector = FilteredElementCollector(doc, view.Id).OfClass(FamilyInstance)
    for inst in collector:
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
    for inst in instances:
        payload = _get_linker_payload_from_instance(inst)
        cand = (payload.get("led_id") or "").strip().lower()
        if cand == target:
            filtered.append(inst)
    return filtered


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


def main():
    doc = revit.doc
    active_view = getattr(doc, "ActiveView", None)
    if not active_view:
        forms.alert("No active view detected.", title="Tag Equipment")
        return
    if active_view.ViewType == ViewType.ThreeD:
        forms.alert("Tag Equipment only works in plan/elevation views.", title="Tag Equipment")
        return

    data_path = _pick_profile_data_path()
    if not data_path:
        return

    repo = _build_repository(data_path)
    tag_entries = _collect_tag_entries(repo)
    if not tag_entries:
        forms.alert("No tags found in the selected YAML.", title="Tag Equipment")
        return

    choices = [entry["display"] for entry in tag_entries]
    selection = forms.SelectFromList.show(
        choices,
        title="Select tags to place in '{}'".format(active_view.Name),
        multiselect=True,
        button_name="Tag",
    )
    if not selection:
        return

    if isinstance(selection, basestring):
        selected_entries = [entry for entry in tag_entries if entry["display"] == selection]
    else:
        selected_entries = [entry for entry in tag_entries if entry["display"] in selection]
    if not selected_entries:
        return

    host_lookup, symbol_lookup = _collect_instance_lookup(doc, active_view)
    if not host_lookup and not symbol_lookup:
        forms.alert("No family instances found in the active view.", title="Tag Equipment")
        return

    tag_view_map = {}
    for entry in selected_entries:
        if entry.get("key"):
            tag_view_map.setdefault(entry["key"], []).append(active_view.Id.IntegerValue)

    engine = PlaceElementsEngine(doc, repo, tag_view_map=tag_view_map)
    total_tags = 0
    missing = []

    placed_tag_pairs = set()
    t = Transaction(doc, "Tag Equipment (YAML)")
    t.Start()
    try:
        for entry in selected_entries:
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
                    total_tags += 1
                    placed_tag_pairs.add(host_key)
        t.Commit()
    except Exception as exc:
        t.RollBack()
        forms.alert("Failed while placing tags:\n\n{0}".format(exc), title="Tag Equipment")
        return

    summary = [
        "Placed {0} tag(s) in view '{1}'.".format(total_tags, active_view.Name),
    ]
    if missing:
        summary.append("")
        summary.append("No host elements found for:")
        sample = missing[:8]
        summary.extend(" - {0}".format(label) for label in sample)
        if len(missing) > len(sample):
            summary.append("   (+{0} more)".format(len(missing) - len(sample)))

    forms.alert("\n".join(summary), title="Tag Equipment")


if __name__ == "__main__":
    main()
