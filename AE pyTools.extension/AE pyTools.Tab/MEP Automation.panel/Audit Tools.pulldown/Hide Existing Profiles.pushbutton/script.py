# -*- coding: utf-8 -*-
"""
Toggle Existing Profiles
------------------------
Hide/unhide linked model elements in the active view when their names already
exist as equipment profiles in the active YAML (stored in Extensible Storage).
"""

import os
import sys
import System

from pyrevit import forms, revit, script
from pyrevit.revit import ui
import pyrevit.extensions as exts
output = script.get_output()
output.close_others()
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Reference,
    RevitLinkInstance,
    TemporaryViewMode,
    View,
    ViewDuplicateOption,
)
from System.Collections.Generic import List

try:
    from Autodesk.Revit.DB import LinkElementId
except Exception:
    LinkElementId = None
try:
    from Autodesk.Revit.UI import PostableCommand, RevitCommandId
except Exception:
    PostableCommand = None
    RevitCommandId = None


LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402
from ExtensibleStorage import ExtensibleStorage  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402

TITLE = "Hide Existing Profiles"
BUILD_TAG = "2026-04-01-hide-only3"
WORK_VIEW_NAME = "TemporaryHideExistingProfiles"
SETTING_KEY = "mep_automation.toggle_existing_profiles_hidden"
SETTING_IDS_KEY = "mep_automation.toggle_existing_profiles_pairs"
SETTING_STRATEGY_KEY = "mep_automation.toggle_existing_profiles_strategy"
SETTING_HIDDEN_HOST_IDS_KEY = "mep_automation.toggle_existing_profiles_hidden_host_ids"
USE_TEMPORARY_HIDE_ISOLATE = False


try:
    basestring
except NameError:
    basestring = str


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


def _with_build(message):
    return "{}\n\nBuild: {}".format(message, BUILD_TAG)


def _normalize_family_type_key(value):
    text = _normalize_name(value)
    if not text:
        return ""
    # Normalize full-width colon and spacing around separator.
    text = text.replace(u"\uff1a", ":")
    if ":" in text:
        left, right = text.split(":", 1)
        left = left.strip()
        right = right.strip()
        if left and right:
            return "{}:{}".format(left, right)
    return text


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


def _family_type_label(elem):
    if not isinstance(elem, FamilyInstance):
        return ""
    symbol = getattr(elem, "Symbol", None)
    if symbol is None:
        try:
            type_id = elem.GetTypeId()
        except Exception:
            type_id = None
        if type_id is not None:
            try:
                symbol = elem.Document.GetElement(type_id)
            except Exception:
                symbol = None
    family = getattr(symbol, "Family", None) if symbol else None
    fam_name = getattr(family, "Name", None) if family else None
    type_name = getattr(symbol, "Name", None) if symbol else None
    if not type_name and symbol is not None:
        try:
            type_name = getattr(elem, "Name", None)
        except Exception:
            type_name = None
    if fam_name and type_name:
        return u"{} : {}".format(fam_name, type_name)
    return ""


def _linked_name_candidates(elem):
    candidates = []
    family_type = _family_type_label(elem)
    if family_type:
        key = _normalize_family_type_key(family_type)
        if key:
            candidates.append(("family_type", family_type, key))
    return candidates


def _match_profile_for_element(elem, profile_lookup):
    for basis, raw_name, key in _linked_name_candidates(elem):
        profiles = profile_lookup.get(key) or []
        if profiles:
            return profiles[0], raw_name, basis, key
    return None, None, None, None


def _collect_profile_name_lookup(data):
    lookup = {}

    def _add_profile_name(raw_value):
        raw_text = (raw_value or "").strip()
        if not raw_text:
            return
        normalized = _normalize_family_type_key(raw_text)
        if not normalized:
            return
        existing = lookup.setdefault(normalized, [])
        if raw_text not in existing:
            existing.append(raw_text)

    for eq in data.get("equipment_definitions") or []:
        if not isinstance(eq, dict):
            continue
        _add_profile_name(eq.get("name"))
        linked_sets = eq.get("linked_sets") or []
        if (not linked_sets) and isinstance(eq.get("linked_element_definitions"), list):
            linked_sets = [{"linked_element_definitions": eq.get("linked_element_definitions")}]
        for linked_set in linked_sets or []:
            if not isinstance(linked_set, dict):
                continue
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                _add_profile_name(led.get("label"))
                _add_profile_name(led.get("name"))
    return lookup


def _iter_target_link_elements(link_doc):
    try:
        collector = FilteredElementCollector(link_doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    except Exception:
        collector = []
    target_cat_id = None
    try:
        target_cat_id = int(BuiltInCategory.OST_SpecialityEquipment)
    except Exception:
        target_cat_id = None
    for elem in collector:
        if target_cat_id is not None:
            try:
                cat = getattr(elem, "Category", None)
                cat_id = _element_id_value(getattr(cat, "Id", None), default=None) if cat is not None else None
            except Exception:
                cat_id = None
            if cat_id != target_cat_id:
                continue
        yield elem


def _doc_key(doc):
    if doc is None:
        return None
    path = None
    title = None
    hash_code = None
    try:
        path = doc.PathName
    except Exception:
        path = None
    try:
        title = doc.Title
    except Exception:
        title = None
    try:
        hash_code = doc.GetHashCode()
    except Exception:
        hash_code = None
    return "{}||{}||{}".format(path or "", title or "", hash_code if hash_code is not None else "")


def _iter_link_doc_chains(doc):
    if doc is None:
        return
    root_key = _doc_key(doc)
    start_seen = set([root_key]) if root_key else set()

    def _walk(source_doc, chain, seen_keys):
        try:
            link_instances = list(FilteredElementCollector(source_doc).OfClass(RevitLinkInstance))
        except Exception:
            link_instances = []
        for link_inst in link_instances:
            try:
                link_doc = link_inst.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                continue
            key = _doc_key(link_doc)
            if key and key in seen_keys:
                continue
            next_seen = set(seen_keys)
            if key:
                next_seen.add(key)
            next_chain = list(chain)
            next_chain.append(link_inst)
            yield link_doc, next_chain
            for nested in _walk(link_doc, next_chain, next_seen):
                yield nested

    for item in _walk(doc, [], start_seen):
        yield item


def _chain_ids_from_instances(link_chain):
    chain_ids = []
    for link_inst in link_chain or []:
        link_id_int = _element_id_value(getattr(link_inst, "Id", None), default=None)
        if link_id_int is None:
            return None
        chain_ids.append(int(link_id_int))
    if not chain_ids:
        return None
    return tuple(chain_ids)


def _normalize_target_entry(entry):
    if not isinstance(entry, (list, tuple)) or len(entry) != 2:
        return None
    chain_part, elem_part = entry
    try:
        elem_id_int = int(elem_part)
    except Exception:
        return None
    if isinstance(chain_part, (list, tuple)):
        chain_ids = []
        for raw in chain_part:
            try:
                chain_ids.append(int(raw))
            except Exception:
                return None
    else:
        try:
            chain_ids = [int(chain_part)]
        except Exception:
            return None
    if not chain_ids:
        return None
    return (tuple(chain_ids), elem_id_int)


def _has_nested_targets(targets):
    for target in targets or []:
        normalized = _normalize_target_entry(target)
        if not normalized:
            continue
        chain_ids, _elem_id_int = normalized
        if len(chain_ids) > 1:
            return True
    return False


def _resolve_chain_instances(doc, chain_ids):
    if doc is None or not chain_ids:
        return None, None
    current_doc = doc
    chain_instances = []
    for raw_link_id in chain_ids:
        try:
            link_id = ElementId(int(raw_link_id))
        except Exception:
            return None, None
        try:
            link_inst = current_doc.GetElement(link_id)
        except Exception:
            link_inst = None
        if link_inst is None:
            return None, None
        chain_instances.append(link_inst)
        try:
            current_doc = link_inst.GetLinkDocument()
        except Exception:
            current_doc = None
        if current_doc is None:
            return None, None
    return chain_instances, current_doc


def _collect_matching_link_element_ids(doc, profile_lookup):
    link_element_ids = []
    matched_pairs = []
    match_records = []
    linked_name_samples = {}
    seen = set()
    scanned = 0
    matched = 0
    link_count = 0

    for link_doc, link_chain in _iter_link_doc_chains(doc):
        link_count += 1
        chain_ids = _chain_ids_from_instances(link_chain)
        if not chain_ids:
            continue
        chain_names = []
        for inst in link_chain or []:
            try:
                chain_names.append(getattr(inst, "Name", None) or "<link>")
            except Exception:
                chain_names.append("<link>")
        chain_text = " > ".join(chain_names) if chain_names else "<link>"
        link_doc_name = getattr(link_doc, "Title", None) or "<linked doc>"
        is_direct = len(chain_ids) == 1
        root_link_inst = link_chain[0] if link_chain else None
        for elem in _iter_target_link_elements(link_doc):
            scanned += 1
            elem_id_int = _element_id_value(getattr(elem, "Id", None), default=None)
            if elem_id_int is None:
                continue
            for _basis, raw_name, key in _linked_name_candidates(elem):
                if key and key not in linked_name_samples:
                    linked_name_samples[key] = raw_name
            profile_name, matched_name, match_basis, matched_key = _match_profile_for_element(elem, profile_lookup)
            if not profile_name:
                continue
            key = (chain_ids, int(elem_id_int))
            if key in seen:
                continue
            seen.add(key)
            matched_pairs.append(key)
            match_records.append({
                "chain_ids": chain_ids,
                "chain_text": chain_text,
                "linked_doc_name": link_doc_name,
                "linked_element_id": int(elem_id_int),
                "linked_name": matched_name,
                "linked_family_type": _family_type_label(elem),
                "match_basis": match_basis,
                "match_key": matched_key,
                "profile_name": profile_name,
            })
            if is_direct and LinkElementId is not None and root_link_inst is not None:
                try:
                    link_element_ids.append(LinkElementId(root_link_inst.Id, elem.Id))
                except Exception:
                    pass
            matched += 1

    return link_element_ids, matched_pairs, match_records, linked_name_samples, scanned, matched, link_count


def _get_toggle_state(doc, default=False):
    try:
        value = ExtensibleStorage.get_user_setting(doc, SETTING_KEY, default=None)
    except Exception:
        return bool(default)
    if value is None:
        return bool(default)
    if isinstance(value, basestring):
        token = value.strip().lower()
        if token in ("1", "true", "yes", "on"):
            return True
        if token in ("0", "false", "no", "off"):
            return False
    return bool(value)


def _set_toggle_state(doc, value):
    try:
        return ExtensibleStorage.set_user_setting(
            doc,
            SETTING_KEY,
            bool(value),
            transaction_name="TOGGLE_EXISTING_PROFILES_STATE",
        )
    except Exception:
        return False


def _get_last_hide_strategy(doc, default_value="targeted"):
    try:
        value = ExtensibleStorage.get_user_setting(doc, SETTING_STRATEGY_KEY, default=None)
    except Exception:
        value = None
    if not value:
        return default_value
    try:
        text = str(value).strip().lower()
    except Exception:
        return default_value
    return text or default_value


def _set_last_hide_strategy(doc, strategy):
    try:
        return ExtensibleStorage.set_user_setting(
            doc,
            SETTING_STRATEGY_KEY,
            str(strategy or "targeted"),
            transaction_name="TOGGLE_EXISTING_PROFILES_STRATEGY",
        )
    except Exception:
        return False


def _set_button_icon(is_on, script_cmp=None, ui_button_cmp=None):
    if script_cmp is not None and ui_button_cmp is not None:
        try:
            off_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_OFF_ICON_FILE)
            on_icon = ui.resolve_icon_file(script_cmp.directory, "on.png")
            icon_path = on_icon if is_on else off_icon
            ui_button_cmp.set_icon(icon_path)
            return
        except Exception:
            pass
    try:
        script.toggle_icon(bool(is_on))
    except Exception:
        pass


def _serialize_pairs(pairs):
    if not pairs:
        return ""
    encoded = []
    for entry in pairs:
        normalized = _normalize_target_entry(entry)
        if not normalized:
            continue
        chain_ids, elem_id_int = normalized
        chain_token = ">".join([str(int(link_id)) for link_id in chain_ids])
        encoded.append("{}|{}".format(chain_token, int(elem_id_int)))
    return ";".join(encoded)


def _deserialize_pairs(raw):
    results = []
    if not raw:
        return results
    for chunk in str(raw).split(";"):
        token = chunk.strip()
        if not token:
            continue
        if "|" in token:
            chain_raw, elem_raw = token.split("|", 1)
        elif "," in token:
            # Backward-compatible legacy format: "link_id,elem_id"
            chain_raw, elem_raw = token.split(",", 1)
        else:
            continue
        chain_ids = []
        for part in str(chain_raw).split(">"):
            part = part.strip()
            if not part:
                continue
            try:
                chain_ids.append(int(part))
            except Exception:
                chain_ids = []
                break
        if not chain_ids:
            continue
        try:
            elem_id_int = int(elem_raw)
        except Exception:
            continue
        results.append((tuple(chain_ids), elem_id_int))
    return results


def _get_hidden_pairs(doc):
    try:
        raw = ExtensibleStorage.get_user_setting(doc, SETTING_IDS_KEY, default=None)
    except Exception:
        raw = None
    return _deserialize_pairs(raw)


def _set_hidden_pairs(doc, pairs):
    payload = _serialize_pairs(pairs)
    try:
        return ExtensibleStorage.set_user_setting(
            doc,
            SETTING_IDS_KEY,
            payload,
            transaction_name="TOGGLE_EXISTING_PROFILES_IDS",
        )
    except Exception:
        return False


def _serialize_int_ids(values):
    if not values:
        return ""
    parts = []
    seen = set()
    for raw in values:
        try:
            val = int(raw)
        except Exception:
            continue
        if val <= 0 or val in seen:
            continue
        seen.add(val)
        parts.append(str(val))
    return ",".join(parts)


def _deserialize_int_ids(raw):
    result = []
    if not raw:
        return result
    seen = set()
    for part in str(raw).split(","):
        token = part.strip()
        if not token:
            continue
        try:
            val = int(token)
        except Exception:
            continue
        if val <= 0 or val in seen:
            continue
        seen.add(val)
        result.append(val)
    return result


def _get_hidden_host_ids(doc):
    try:
        raw = ExtensibleStorage.get_user_setting(doc, SETTING_HIDDEN_HOST_IDS_KEY, default=None)
    except Exception:
        raw = None
    return _deserialize_int_ids(raw)


def _set_hidden_host_ids(doc, host_ids):
    payload = _serialize_int_ids(host_ids)
    try:
        return ExtensibleStorage.set_user_setting(
            doc,
            SETTING_HIDDEN_HOST_IDS_KEY,
            payload,
            transaction_name="TOGGLE_EXISTING_PROFILES_HOST_IDS",
        )
    except Exception:
        return False


def _get_view_hidden_id_values(view):
    values = set()
    if view is None:
        return values
    try:
        hidden_ids = view.GetHiddenElementIds()
    except Exception:
        hidden_ids = None
    if hidden_ids is None:
        return values
    for elem_id in hidden_ids:
        val = _element_id_value(elem_id, default=None)
        if val is None:
            continue
        try:
            val = int(val)
        except Exception:
            continue
        if val > 0:
            values.add(val)
    return values


def _unhide_via_stored_host_ids(doc, view, host_id_values):
    if doc is None or view is None or not host_id_values:
        return 0
    currently_hidden = _get_view_hidden_id_values(view)
    if currently_hidden:
        candidate_vals = [val for val in host_id_values if int(val) in currently_hidden]
    else:
        candidate_vals = list(host_id_values)
    if not candidate_vals:
        return 0

    success = 0
    with revit.Transaction("Unhide Existing Profiles (Stored Host Ids)"):
        for raw_val in candidate_vals:
            try:
                elem_id = ElementId(int(raw_val))
            except Exception:
                continue
            # Keep this path targeted; do not treat root link instances as element surrogates.
            try:
                host_elem = doc.GetElement(elem_id)
            except Exception:
                host_elem = None
            if isinstance(host_elem, RevitLinkInstance):
                continue
            ids = List[ElementId]()
            ids.Add(elem_id)
            try:
                view.UnhideElements(ids)
                success += 1
            except Exception:
                continue
    return success


def _build_link_element_ids_from_pairs(doc, pairs):
    if LinkElementId is None:
        raise RuntimeError("Current Revit API does not expose LinkElementId for linked-element visibility control.")
    if not pairs:
        return []

    results = []
    seen = set()
    for entry in pairs:
        normalized = _normalize_target_entry(entry)
        if not normalized:
            continue
        chain_ids, elem_id_int = normalized
        if len(chain_ids) != 1:
            continue
        chain_instances, link_doc = _resolve_chain_instances(doc, chain_ids)
        if not chain_instances or link_doc is None:
            continue
        link_inst = chain_instances[0]
        if link_doc is None:
            continue
        try:
            linked_elem = link_doc.GetElement(ElementId(int(elem_id_int)))
        except Exception:
            linked_elem = None
        if linked_elem is None:
            continue
        key = (tuple(chain_ids), int(elem_id_int))
        if key in seen:
            continue
        seen.add(key)
        try:
            results.append(LinkElementId(link_inst.Id, linked_elem.Id))
        except Exception:
            continue
    return results


def _build_link_references_from_pairs(doc, pairs):
    if not pairs:
        return []

    refs = []
    seen = set()
    for entry in pairs:
        normalized = _normalize_target_entry(entry)
        if not normalized:
            continue
        chain_ids, elem_id_int = normalized
        chain_instances, link_doc = _resolve_chain_instances(doc, chain_ids)
        if not chain_instances or link_doc is None:
            continue
        if link_doc is None:
            continue
        try:
            linked_elem = link_doc.GetElement(ElementId(int(elem_id_int)))
        except Exception:
            linked_elem = None
        if linked_elem is None:
            continue
        key = (tuple(chain_ids), int(elem_id_int))
        if key in seen:
            continue
        seen.add(key)
        try:
            ref = Reference(linked_elem)
            for link_inst in reversed(chain_instances):
                ref = ref.CreateLinkReference(link_inst)
            refs.append(ref)
        except Exception:
            continue
    return refs


def _collect_root_link_ids_from_pairs(doc, pairs):
    link_ids = []
    seen = set()
    if doc is None:
        return link_ids
    for entry in pairs or []:
        normalized = _normalize_target_entry(entry)
        if not normalized:
            continue
        chain_ids, _elem_id = normalized
        if not chain_ids:
            continue
        root_id_int = int(chain_ids[0])
        if root_id_int in seen:
            continue
        seen.add(root_id_int)
        try:
            link_id = ElementId(root_id_int)
            if doc.GetElement(link_id) is None:
                continue
            link_ids.append(link_id)
        except Exception:
            continue
    return link_ids


def _unhide_via_root_links(doc, view, pairs):
    root_link_ids = _collect_root_link_ids_from_pairs(doc, pairs)
    if not root_link_ids:
        return 0

    success = 0
    with revit.Transaction("Unhide Existing Profile Link Instances"):
        for link_id in root_link_ids:
            link_elem = None
            try:
                link_elem = doc.GetElement(link_id)
            except Exception:
                link_elem = None
            if link_elem is None:
                continue
            # Count success only when the link instance was actually hidden.
            try:
                was_hidden = bool(link_elem.IsHidden(view))
            except Exception:
                was_hidden = False
            if not was_hidden:
                continue
            ids = List[ElementId]()
            ids.Add(link_id)
            try:
                view.UnhideElements(ids)
                success += 1
            except Exception:
                continue
    return success


def _hide_via_root_links(doc, view, pairs):
    root_link_ids = _collect_root_link_ids_from_pairs(doc, pairs)
    if not root_link_ids:
        return 0

    success = 0
    with revit.Transaction("Hide Existing Profile Link Instances"):
        for link_id in root_link_ids:
            ids = List[ElementId]()
            ids.Add(link_id)
            try:
                view.HideElements(ids)
                success += 1
            except Exception:
                continue
    return success


def _build_report_records_from_pairs(doc, pairs, profile_lookup):
    records = []
    seen = set()
    for entry in pairs or []:
        normalized = _normalize_target_entry(entry)
        if not normalized:
            continue
        chain_ids, elem_id_int = normalized
        key = (tuple(chain_ids), int(elem_id_int))
        if key in seen:
            continue
        seen.add(key)

        chain_instances, link_doc = _resolve_chain_instances(doc, chain_ids)
        chain_names = []
        for inst in chain_instances or []:
            try:
                chain_names.append(getattr(inst, "Name", None) or "<link>")
            except Exception:
                chain_names.append("<link>")
        chain_text = " > ".join(chain_names) if chain_names else "<missing link path>"
        link_doc_name = getattr(link_doc, "Title", None) if link_doc is not None else "<missing linked doc>"

        linked_elem = None
        if link_doc is not None:
            try:
                linked_elem = link_doc.GetElement(ElementId(int(elem_id_int)))
            except Exception:
                linked_elem = None

        profile_name, matched_name, match_basis, matched_key = _match_profile_for_element(linked_elem, profile_lookup)
        family_type = _family_type_label(linked_elem) if linked_elem is not None else ""

        records.append({
            "chain_ids": chain_ids,
            "chain_text": chain_text,
            "linked_doc_name": link_doc_name,
            "linked_element_id": int(elem_id_int),
            "linked_name": matched_name or family_type or "<missing family:type>",
            "linked_family_type": family_type or "<missing family:type>",
            "match_basis": match_basis or "<not resolved>",
            "match_key": matched_key or "",
            "profile_name": profile_name or "<profile name not resolved>",
        })
    return records


def _print_match_report(action_text, apply_mode, yaml_label, records, scanned, link_count, matched):
    output.print_md("### {} Report".format(TITLE))
    output.print_md(
        "- Action: **{}** | Apply mode: **{}** | YAML: **{}**".format(
            action_text,
            str(apply_mode or "").upper(),
            yaml_label or "",
        )
    )
    output.print_md(
        "- Scope: YAML profile **name** compared against linked **Family : Type** only."
    )
    output.print_md("- Note: YAML lookup includes equipment profile names and linked definition labels.")
    output.print_md(
        "- Links scanned: **{}** | Elements scanned: **{}** | Matches: **{}**".format(
            link_count,
            scanned,
            matched,
        )
    )
    if not records:
        output.print_md("_No matched linked elements to report._")
        return

    rows = []
    for rec in records:
        rows.append([
            rec.get("profile_name") or "",
            rec.get("linked_name") or "",
            rec.get("linked_family_type") or "",
            rec.get("match_basis") or "",
            rec.get("linked_element_id"),
            rec.get("linked_doc_name") or "",
            rec.get("chain_text") or "",
        ])
    rows.sort(key=lambda item: (str(item[0]).lower(), str(item[1]).lower(), str(item[6]).lower(), int(item[4] or 0)))
    output.print_table(
        table_data=rows,
        columns=["Profile Name", "Matched Linked Family : Type", "Linked Family : Type", "Match Basis", "Linked Element Id", "Linked Doc", "Link Path"],
    )


def _print_no_match_diagnostics(yaml_lookup, linked_name_samples):
    yaml_keys = set(yaml_lookup.keys())
    linked_keys = set(linked_name_samples.keys())
    only_yaml = sorted(list(yaml_keys - linked_keys))
    only_linked = sorted(list(linked_keys - yaml_keys))

    output.print_md("### {} Diagnostics".format(TITLE))
    output.print_md("- No matches found for strict linked `Family : Type` vs profile `name` matching.")
    output.print_md("- YAML profile names loaded: **{}**".format(len(yaml_keys)))
    output.print_md("- Unique linked Family : Type names scanned: **{}**".format(len(linked_keys)))

    if only_yaml:
        output.print_md("#### YAML Names Not Found In Linked Scan (sample)")
        sample_rows = []
        for key in only_yaml[:30]:
            raw = (yaml_lookup.get(key) or [key])[0]
            sample_rows.append([raw, key])
        output.print_table(table_data=sample_rows, columns=["YAML Profile Name", "Normalized Key"])

    if only_linked:
        output.print_md("#### Linked Names Not Found In YAML (sample)")
        sample_rows = []
        for key in only_linked[:30]:
            raw = linked_name_samples.get(key) or key
            sample_rows.append([raw, key])
        output.print_table(table_data=sample_rows, columns=["Linked Family : Type", "Normalized Key"])


def _build_host_surrogate_ids_from_link_ids(doc, link_element_ids):
    elem_ids = List[ElementId]()
    seen = set()
    if not link_element_ids:
        return elem_ids
    for lid in link_element_ids:
        host_id = None
        try:
            host_id = getattr(lid, "HostElementId", None)
        except Exception:
            host_id = None
        host_val = _element_id_value(host_id, default=None)
        if host_val is None:
            continue
        try:
            host_val = int(host_val)
        except Exception:
            continue
        if host_val <= 0 or host_val in seen:
            continue
        # Guard: never treat a root link instance id as a per-element surrogate.
        if doc is not None:
            try:
                host_elem = doc.GetElement(host_id)
            except Exception:
                host_elem = None
            if isinstance(host_elem, RevitLinkInstance):
                continue
        seen.add(host_val)
        try:
            elem_ids.Add(host_id)
        except Exception:
            continue
    return elem_ids


def _host_surrogate_values_from_link_ids(doc, link_element_ids):
    values = []
    seen = set()
    elem_ids = _build_host_surrogate_ids_from_link_ids(doc, link_element_ids)
    for elem_id in elem_ids:
        val = _element_id_value(elem_id, default=None)
        if val is None:
            continue
        try:
            val = int(val)
        except Exception:
            continue
        if val <= 0 or val in seen:
            continue
        seen.add(val)
        values.append(val)
    return values


def _hide_via_host_ids(doc, view, host_id_values):
    if doc is None or view is None or not host_id_values:
        return 0
    ids = List[ElementId]()
    seen = set()
    for raw_val in host_id_values:
        try:
            val = int(raw_val)
        except Exception:
            continue
        if val <= 0 or val in seen:
            continue
        seen.add(val)
        try:
            elem_id = ElementId(val)
        except Exception:
            continue
        # Never hide whole link instances from surrogate fallback.
        try:
            host_elem = doc.GetElement(elem_id)
        except Exception:
            host_elem = None
        if isinstance(host_elem, RevitLinkInstance):
            continue
        try:
            ids.Add(elem_id)
        except Exception:
            continue

    if ids.Count <= 0:
        return 0

    try:
        view.HideElements(ids)
        return int(ids.Count)
    except Exception:
        # Keep transaction alive and salvage with per-id calls.
        success = 0
        for elem_id in ids:
            one = List[ElementId]()
            one.Add(elem_id)
            try:
                view.HideElements(one)
                success += 1
            except Exception:
                continue
        return success


def _unhide_all_hidden_non_link_elements(doc, view):
    if doc is None or view is None:
        return 0
    try:
        hidden_ids = list(view.GetHiddenElementIds() or [])
    except Exception:
        hidden_ids = []
    if not hidden_ids:
        return 0

    ids = List[ElementId]()
    seen = set()
    for elem_id in hidden_ids:
        val = _element_id_value(elem_id, default=None)
        if val is None:
            continue
        try:
            val = int(val)
        except Exception:
            continue
        if val <= 0 or val in seen:
            continue
        seen.add(val)
        try:
            elem = doc.GetElement(elem_id)
        except Exception:
            elem = None
        if isinstance(elem, RevitLinkInstance):
            continue
        try:
            ids.Add(elem_id)
        except Exception:
            continue

    if ids.Count <= 0:
        return 0

    with revit.Transaction("Unhide Existing Profiles (Emergency Non-Link Recovery)"):
        view.UnhideElements(ids)
    return int(ids.Count)


def _apply_visibility(view, link_element_ids, hide_mode):
    typed_ids = List[LinkElementId]()
    for item in link_element_ids:
        typed_ids.Add(item)

    def _invoke_element_ids_from_link_ids():
        doc = getattr(view, "Document", None)
        elem_ids = _build_host_surrogate_ids_from_link_ids(doc, link_element_ids)
        if elem_ids.Count <= 0:
            return False
        if hide_mode:
            view.HideElements(elem_ids)
        else:
            view.UnhideElements(elem_ids)
        return True

    def _invoke_link_overload(method_name):
        try:
            view_type = view.GetType()
            methods = [m for m in view_type.GetMethods() if m.Name == method_name]
        except Exception:
            methods = []
        for method in methods:
            try:
                params = method.GetParameters()
                if len(params) <= 0:
                    continue
                ptype = params[0].ParameterType
                if not getattr(ptype, "IsGenericType", False):
                    continue
                gargs = ptype.GetGenericArguments()
                if not gargs or len(gargs) != 1:
                    continue
                target_arg = getattr(gargs[0], "FullName", None) or ""
                if target_arg != "Autodesk.Revit.DB.LinkElementId":
                    continue
                invoke_args = [typed_ids]
                if len(params) == 2:
                    second_type = params[1].ParameterType
                    second_name = getattr(second_type, "FullName", None) or ""
                    if second_name in ("System.Boolean", "bool"):
                        invoke_args.append(False)
                    elif second_name in ("System.Int32", "int"):
                        invoke_args.append(0)
                    elif getattr(second_type, "IsEnum", False):
                        invoke_args.append(System.Enum.ToObject(second_type, 0))
                    else:
                        continue
                elif len(params) > 2:
                    continue
                args = System.Array[System.Object](invoke_args)
                method.Invoke(view, args)
                return True
            except Exception:
                continue
        return False

    method_name = "HideElements" if hide_mode else "UnhideElements"
    if hide_mode:
        # Hide: prefer link-id overload first, then host surrogate ids.
        if _invoke_link_overload(method_name):
            return
        if _invoke_element_ids_from_link_ids():
            return
    else:
        # Unhide: prefer host surrogate ids first to avoid brittle link-id overloads.
        if _invoke_element_ids_from_link_ids():
            return
        if _invoke_link_overload(method_name):
            return

    raise RuntimeError(
        "No compatible {} overload accepted the linked element collection.".format(method_name)
    )


def _has_link_overload(view, method_name):
    try:
        view_type = view.GetType()
        methods = [m for m in view_type.GetMethods() if m.Name == method_name]
    except Exception:
        methods = []
    for method in methods:
        try:
            params = method.GetParameters()
            if len(params) <= 0:
                continue
            ptype = params[0].ParameterType
            if not getattr(ptype, "IsGenericType", False):
                continue
            gargs = ptype.GetGenericArguments()
            if not gargs or len(gargs) != 1:
                continue
            target_arg = getattr(gargs[0], "FullName", None) or ""
            if target_arg == "Autodesk.Revit.DB.LinkElementId":
                return True
        except Exception:
            continue
    return False


def _post_visibility_command(uidoc, refs, hide_mode):
    if PostableCommand is None or RevitCommandId is None:
        raise RuntimeError("Revit UI command API is unavailable in this environment.")
    if uidoc is None:
        raise RuntimeError("No active UI document available.")
    if not refs:
        raise RuntimeError("No linked references available for visibility command.")

    typed_refs = List[Reference]()
    for ref in refs:
        typed_refs.Add(ref)
    uidoc.Selection.SetReferences(typed_refs)

    if hide_mode:
        cmd_enum = getattr(PostableCommand, "HideElements", None)
    else:
        cmd_enum = getattr(PostableCommand, "UnhideElements", None) or getattr(PostableCommand, "UnhideElement", None)
    if cmd_enum is not None:
        cmd_id = RevitCommandId.LookupPostableCommandId(cmd_enum)
        if _try_post_command_id(cmd_id, require_can_post=hide_mode):
            return

    if hide_mode:
        raise RuntimeError("No postable hide command is available.")

    # Revit 2025 can omit Unhide from PostableCommand even when an internal
    # command id exists; try known command id names directly.
    candidate_names = [
        "ID_VIEW_UNHIDE_ELEMENTS",
        "ID_VIEW_UNHIDE_ELEMENT",
        "ID_VIEW_UNHIDE_ELEMS",
        "ID_UNHIDE_ELEMENTS",
        "ID_UNHIDE_ELEMENT",
        "ID_EDIT_UNHIDE_ELEMENTS",
        "ID_VIEW_UNHIDE",
        "ID_VIEW_UNHIDE_IN_VIEW",
        "ID_UNHIDE_IN_VIEW_ELEMENTS",
        "ID_VIEW_UNHIDE_IN_VIEW_ELEMENTS",
    ]
    for name in candidate_names:
        try:
            candidate_cmd = RevitCommandId.LookupCommandId(name)
        except Exception:
            candidate_cmd = None
        if _try_post_command_id(candidate_cmd, require_can_post=False):
            return

    raise RuntimeError("No postable unhide command is available.")


def _can_post_unhide_command():
    if PostableCommand is None or RevitCommandId is None:
        return False
    can_post = getattr(__revit__, "CanPostCommand", None)
    if not callable(can_post):
        return False

    enum_candidates = [
        getattr(PostableCommand, "UnhideElements", None),
        getattr(PostableCommand, "UnhideElement", None),
    ]
    for cmd_enum in enum_candidates:
        if cmd_enum is None:
            continue
        try:
            cmd_id = RevitCommandId.LookupPostableCommandId(cmd_enum)
        except Exception:
            cmd_id = None
        if cmd_id is None:
            continue
        try:
            if can_post(cmd_id):
                return True
        except Exception:
            continue

    candidate_names = [
        "ID_VIEW_UNHIDE_ELEMENTS",
        "ID_VIEW_UNHIDE_ELEMS",
        "ID_UNHIDE_ELEMENTS",
        "ID_UNHIDE_IN_VIEW_ELEMENTS",
        "ID_VIEW_UNHIDE_IN_VIEW_ELEMENTS",
    ]
    for name in candidate_names:
        try:
            cmd_id = RevitCommandId.LookupCommandId(name)
        except Exception:
            cmd_id = None
        if cmd_id is None:
            continue
        try:
            if can_post(cmd_id):
                return True
        except Exception:
            continue
    return False


def _try_post_command_id(cmd_id, require_can_post=True):
    if cmd_id is None:
        return False
    can_post = getattr(__revit__, "CanPostCommand", None)
    if callable(can_post):
        try:
            can_post_now = bool(can_post(cmd_id))
        except Exception:
            can_post_now = False
        if require_can_post and (not can_post_now):
            return False
    try:
        __revit__.PostCommand(cmd_id)
        return True
    except Exception:
        return False


def _post_temp_hide_command(uidoc, refs):
    if PostableCommand is None or RevitCommandId is None:
        raise RuntimeError("Revit UI command API is unavailable in this environment.")
    if uidoc is None:
        raise RuntimeError("No active UI document available.")
    if not refs:
        raise RuntimeError("No linked references available for temporary hide command.")

    typed_refs = List[Reference]()
    for ref in refs:
        typed_refs.Add(ref)
    uidoc.Selection.SetReferences(typed_refs)

    enum_candidates = [
        getattr(PostableCommand, "TemporaryHideIsolateElement", None),
        getattr(PostableCommand, "TemporaryHideIsolateElements", None),
    ]
    for cmd_enum in enum_candidates:
        if cmd_enum is None:
            continue
        try:
            cmd_id = RevitCommandId.LookupPostableCommandId(cmd_enum)
        except Exception:
            cmd_id = None
        if _try_post_command_id(cmd_id):
            return

    candidate_names = [
        "ID_VIEW_TEMP_HIDE_ELEMENTS",
        "ID_VIEW_TEMP_HIDE_ELEMS",
        "ID_VIEW_TEMPORARY_HIDE_ELEMENTS",
        "ID_VIEW_TEMP_HIDE_ISOLATE_ELEMENTS",
    ]
    for name in candidate_names:
        try:
            cmd_id = RevitCommandId.LookupCommandId(name)
        except Exception:
            cmd_id = None
        if _try_post_command_id(cmd_id):
            return

    raise RuntimeError("No postable temporary hide command is available.")


def _hide_elements_temporary(view, link_element_ids):
    if LinkElementId is None:
        raise RuntimeError("LinkElementId is unavailable for temporary linked hide.")
    if view is None:
        raise RuntimeError("No active view is available.")
    if not link_element_ids:
        raise RuntimeError("No linked element ids were supplied for temporary hide.")

    typed_ids = List[LinkElementId]()
    for item in link_element_ids:
        typed_ids.Add(item)

    try:
        view_type = view.GetType()
        methods = [m for m in view_type.GetMethods() if m.Name == "HideElementsTemporary"]
    except Exception:
        methods = []

    for method in methods:
        try:
            params = method.GetParameters()
            if len(params) <= 0:
                continue
            first_type = params[0].ParameterType
            if not getattr(first_type, "IsGenericType", False):
                continue
            gargs = first_type.GetGenericArguments()
            if not gargs or len(gargs) != 1:
                continue
            target_arg = getattr(gargs[0], "FullName", None) or ""
            if target_arg != "Autodesk.Revit.DB.LinkElementId":
                continue

            invoke_args = [typed_ids]
            if len(params) == 2:
                second_type = params[1].ParameterType
                second_name = getattr(second_type, "FullName", None) or ""
                if second_name in ("System.Boolean", "bool"):
                    invoke_args.append(False)
                elif second_name in ("System.Int32", "int"):
                    invoke_args.append(0)
                elif getattr(second_type, "IsEnum", False):
                    invoke_args.append(System.Enum.ToObject(second_type, 0))
                else:
                    continue
            elif len(params) > 2:
                continue

            args = System.Array[System.Object](invoke_args)
            method.Invoke(view, args)
            return
        except Exception:
            continue

    # Fallback: some Revit builds only expose ICollection<ElementId> here.
    # Use host surrogate ids, but skip root link instance ids to avoid hiding walls/doors/etc.
    doc = getattr(view, "Document", None)
    elem_ids = _build_host_surrogate_ids_from_link_ids(doc, link_element_ids)
    if elem_ids.Count > 0:
        try:
            view.HideElementsTemporary(elem_ids)
            return
        except Exception:
            pass

    raise RuntimeError("No compatible HideElementsTemporary overload accepted linked ids or host-surrogate ids.")


def _clear_temporary_hide_isolate(view):
    if view is None:
        raise RuntimeError("No active view is available.")
    with revit.Transaction("Clear Temporary Hide/Isolate Existing Profiles"):
        view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)


def _set_reference_selection(uidoc, refs):
    if uidoc is None or not refs:
        return 0
    typed_refs = List[Reference]()
    for ref in refs:
        try:
            typed_refs.Add(ref)
        except Exception:
            continue
    if typed_refs.Count <= 0:
        return 0
    uidoc.Selection.SetReferences(typed_refs)
    return int(typed_refs.Count)


def _enable_reveal_hidden_mode(view):
    if view is None:
        raise RuntimeError("No active view is available.")
    try:
        if view.IsInTemporaryViewMode(TemporaryViewMode.RevealHiddenElements):
            return
    except Exception:
        pass
    try:
        view.EnableRevealHiddenMode()
        return
    except Exception:
        pass
    with revit.Transaction("Enable Reveal Hidden Elements"):
        view.EnableRevealHiddenMode()


def _disable_reveal_hidden_mode(view):
    if view is None:
        return
    try:
        if not view.IsInTemporaryViewMode(TemporaryViewMode.RevealHiddenElements):
            return
    except Exception:
        return
    try:
        view.DisableTemporaryViewMode(TemporaryViewMode.RevealHiddenElements)
    except Exception:
        with revit.Transaction("Disable Reveal Hidden Elements"):
            view.DisableTemporaryViewMode(TemporaryViewMode.RevealHiddenElements)


def _post_reveal_hidden_command():
    if PostableCommand is None or RevitCommandId is None:
        raise RuntimeError("Revit UI command API is unavailable in this environment.")
    cmd_enum = getattr(PostableCommand, "RevealHiddenElements", None)
    if cmd_enum is not None:
        try:
            cmd_id = RevitCommandId.LookupPostableCommandId(cmd_enum)
        except Exception:
            cmd_id = None
        if _try_post_command_id(cmd_id):
            return
    candidate_names = [
        "ID_VIEW_REVEAL_HIDDEN_ELEMENTS",
        "ID_REVEAL_HIDDEN_ELEMENTS",
    ]
    for name in candidate_names:
        try:
            cmd_id = RevitCommandId.LookupCommandId(name)
        except Exception:
            cmd_id = None
        if _try_post_command_id(cmd_id):
            return
    raise RuntimeError("No postable reveal hidden command is available.")


def _is_revit_2025_or_newer(doc):
    if doc is None:
        return False
    try:
        version_raw = getattr(doc.Application, "VersionNumber", "") or ""
        return int(str(version_raw).strip()) >= 2025
    except Exception:
        return False


def _prepare_working_view(doc, source_view, view_name):
    if doc is None or source_view is None:
        raise RuntimeError("No source view is available.")
    if getattr(source_view, "IsTemplate", False):
        raise RuntimeError("Source view cannot be a template.")
    try:
        source_name = (source_view.Name or "").strip()
    except Exception:
        source_name = ""
    if source_name.lower() == str(view_name or "").strip().lower():
        raise RuntimeError(
            "Run this tool from a production/source view, not from '{}'.".format(view_name)
        )

    can_dup = False
    try:
        can_dup = bool(source_view.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
    except Exception:
        can_dup = False
    if not can_dup:
        raise RuntimeError("Current view type cannot be duplicated.")

    existing_ids = []
    for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType():
        if v is None:
            continue
        try:
            if getattr(v, "IsTemplate", False):
                continue
            if (getattr(v, "Name", None) or "").strip().lower() != str(view_name).strip().lower():
                continue
            if _element_id_value(v.Id, default=None) == _element_id_value(source_view.Id, default=None):
                continue
            existing_ids.append(v.Id)
        except Exception:
            continue

    new_view = None
    with revit.Transaction("Prepare Temporary Hide Existing Profiles View"):
        for old_id in existing_ids:
            try:
                doc.Delete(old_id)
            except Exception:
                continue
        new_id = source_view.Duplicate(ViewDuplicateOption.Duplicate)
        new_view = doc.GetElement(new_id)
        if new_view is None:
            raise RuntimeError("Failed to duplicate active view.")
        new_view.Name = view_name

    return new_view


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    source_view = getattr(doc, "ActiveView", None)
    if source_view is None:
        forms.alert("No active view detected.", title=TITLE)
        return
    if getattr(source_view, "IsTemplate", False):
        forms.alert("Active view is a template. Open a model view and try again.", title=TITLE)
        return

    try:
        view = _prepare_working_view(doc, source_view, WORK_VIEW_NAME)
    except Exception as exc:
        forms.alert(
            _with_build("Unable to prepare working view '{}':\n\n{}".format(WORK_VIEW_NAME, exc)),
            title=TITLE,
        )
        return

    try:
        if revit.uidoc is not None:
            revit.uidoc.ActiveView = view
    except Exception:
        pass

    try:
        data_path, data = load_active_yaml_data(doc)
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    except Exception as exc:
        forms.alert("Failed to load active YAML data:\n\n{}".format(exc), title=TITLE)
        return

    profile_lookup = _collect_profile_name_lookup(data)
    if not profile_lookup:
        forms.alert("No equipment profile names were found in the active YAML.", title=TITLE)
        return

    scanned = 0
    matched = 0
    link_count = 0
    matched_pairs = []
    match_records = []
    linked_name_samples = {}
    link_element_ids = []
    try:
        link_element_ids, matched_pairs, match_records, linked_name_samples, scanned, matched, link_count = _collect_matching_link_element_ids(doc, profile_lookup)
    except Exception as exc:
        forms.alert(
            "Unable to collect linked elements for hiding:\n\n{}".format(exc),
            title=TITLE,
        )
        return

    if not matched_pairs:
        _print_no_match_diagnostics(profile_lookup, linked_name_samples)
        forms.alert(
            _with_build(
                "No linked Family : Type names matched YAML profile names.\n\n"
                "Links scanned: {}\n"
                "Elements scanned: {}\n"
                "Matches: {}\n\nSee pyRevit output for mismatch diagnostics.".format(link_count, scanned, matched)
            ),
            title=TITLE,
        )
        return

    action_text = "Hide"
    apply_mode = "api-host-ids"
    hidden_count = 0
    target_host_ids = _host_surrogate_values_from_link_ids(doc, link_element_ids) if link_element_ids else []
    hide_errors = []

    if target_host_ids or link_element_ids:
        try:
            with revit.Transaction("Hide Existing Profiles"):
                if target_host_ids and hidden_count <= 0:
                    try:
                        hidden_count = _hide_via_host_ids(doc, view, target_host_ids)
                        if hidden_count > 0:
                            apply_mode = "api-host-ids"
                    except Exception as exc:
                        hide_errors.append("Host-id hide error: {}".format(exc))

                if hidden_count <= 0 and link_element_ids:
                    try:
                        _apply_visibility(view, link_element_ids, True)
                        hidden_count = len(link_element_ids)
                        apply_mode = "api-link"
                    except Exception as exc:
                        hide_errors.append("Linked API hide error: {}".format(exc))
        except Exception as exc:
            hide_errors.append("API transaction error: {}".format(exc))

    if hidden_count <= 0:
        try:
            refs = _build_link_references_from_pairs(doc, matched_pairs)
            _post_visibility_command(revit.uidoc, refs, True)
            hidden_count = len(refs)
            apply_mode = "ui"
        except Exception as exc:
            hide_errors.append("UI hide error: {}".format(exc))

    if hidden_count <= 0:
        detail = "\n\n".join(hide_errors) if hide_errors else "No hide method succeeded."
        forms.alert(_with_build("Hide failed.\n\n{}".format(detail)), title=TITLE)
        return

    # Keep this command single-transaction: no ExtensibleStorage writes.

    yaml_label = get_yaml_display_name(data_path)
    _print_match_report(action_text, apply_mode, yaml_label, match_records, scanned, link_count, matched)
    forms.show_balloon(
        TITLE,
        "Working view: {}\nHidden linked elements: {}\nMatched linked elements: {}\nElements scanned: {}\nApply mode: {}\nYAML: {}\nBuild: {}\nTo reveal: use Undo in Revit.".format(
            WORK_VIEW_NAME,
            hidden_count,
            matched,
            scanned,
            apply_mode.upper(),
            yaml_label,
            BUILD_TAG,
        ),
    )


if __name__ == "__main__":
    main()
