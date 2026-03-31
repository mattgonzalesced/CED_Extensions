# -*- coding: utf-8 -*-
"""
Toggle Existing Profiles
------------------------
Hide/unhide linked model elements in the active view when their names already
exist as equipment profiles in the active YAML (stored in Extensible Storage).
"""

import os
import sys

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

TITLE = "Toggle Existing Profiles"
SETTING_KEY = "mep_automation.toggle_existing_profiles_hidden"
SETTING_IDS_KEY = "mep_automation.toggle_existing_profiles_pairs"


try:
    basestring
except NameError:
    basestring = str


def _normalize_name(value):
    if not value:
        return ""
    return " ".join(str(value).strip().lower().split())


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


def _apply_visibility(view, link_element_ids, hide_mode):
    typed_ids = List[LinkElementId]()
    for item in link_element_ids:
        typed_ids.Add(item)
    if hide_mode:
        view.HideElements(typed_ids)
    else:
        view.UnhideElements(typed_ids)


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
    if cmd_enum is None:
        raise RuntimeError("No postable {} command is available.".format("hide" if hide_mode else "unhide"))
    cmd_id = RevitCommandId.LookupPostableCommandId(cmd_enum)
    if cmd_id is None:
        raise RuntimeError("Unable to resolve '{}' command id.".format(cmd_enum))
    can_post = getattr(__revit__, "CanPostCommand", None)
    if callable(can_post):
        if not can_post(cmd_id):
            raise RuntimeError(
                "The '{}' command is currently unavailable in this view/context.".format(cmd_enum)
            )
    __revit__.PostCommand(cmd_id)


def _is_revit_2025_or_newer(doc):
    if doc is None:
        return False
    try:
        version_raw = getattr(doc.Application, "VersionNumber", "") or ""
        return int(str(version_raw).strip()) >= 2025
    except Exception:
        return False


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    view = getattr(doc, "ActiveView", None)
    if view is None:
        forms.alert("No active view detected.", title=TITLE)
        return
    if getattr(view, "IsTemplate", False):
        forms.alert("Active view is a template. Open a model view and try again.", title=TITLE)
        return

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

    hidden_now = _get_toggle_state(doc, default=False)
    hide_mode = not hidden_now
    prefer_ui_mode = _is_revit_2025_or_newer(doc) or (LinkElementId is None)

    scanned = 0
    matched = 0
    link_count = 0
    matched_pairs = []
    match_records = []
    linked_name_samples = {}
    link_element_ids = []

    if hide_mode:
        try:
            link_element_ids, matched_pairs, match_records, linked_name_samples, scanned, matched, link_count = _collect_matching_link_element_ids(doc, profile_lookup)
            if _has_nested_targets(matched_pairs):
                prefer_ui_mode = True
        except Exception as exc:
            forms.alert(
                "Unable to collect linked elements for toggling:\n\n{}".format(exc),
                title=TITLE,
            )
            return
    else:
        stored_pairs = _get_hidden_pairs(doc)
        if stored_pairs:
            matched_pairs = list(stored_pairs)
            matched = len(matched_pairs)
            if _has_nested_targets(matched_pairs):
                prefer_ui_mode = True
            if not prefer_ui_mode:
                try:
                    link_element_ids = _build_link_element_ids_from_pairs(doc, stored_pairs)
                    matched = len(link_element_ids) or matched
                except Exception:
                    link_element_ids = []
        else:
            matched_pairs = []
            link_element_ids = []
        needs_rescan = (not link_element_ids) and (not (prefer_ui_mode and matched_pairs))
        if needs_rescan:
            try:
                coll_link_ids, coll_pairs, coll_records, coll_linked_samples, coll_scanned, coll_matched, coll_link_count = _collect_matching_link_element_ids(doc, profile_lookup)
                if not matched_pairs:
                    matched_pairs = coll_pairs
                    matched = coll_matched
                    match_records = coll_records
                    linked_name_samples = coll_linked_samples
                    if _has_nested_targets(matched_pairs):
                        prefer_ui_mode = True
                scanned = coll_scanned
                link_count = coll_link_count
                if not prefer_ui_mode:
                    link_element_ids = coll_link_ids
            except Exception as exc:
                forms.alert(
                    "Unable to collect linked elements for toggling:\n\n{}".format(exc),
                    title=TITLE,
                )
                return

    if not matched_pairs:
        _set_toggle_state(doc, False)
        _set_hidden_pairs(doc, [])
        _set_button_icon(False)
        _print_no_match_diagnostics(profile_lookup, linked_name_samples)
        forms.alert(
            "No linked Family : Type names matched YAML profile names.\n\n"
            "Links scanned: {}\n"
            "Elements scanned: {}\n"
            "Matches: {}\n\nSee pyRevit output for mismatch diagnostics.".format(link_count, scanned, matched),
            title=TITLE,
        )
        return

    action_text = "Hide" if hide_mode else "Unhide"
    apply_mode = "ui" if prefer_ui_mode else "api"

    if prefer_ui_mode:
        try:
            refs = _build_link_references_from_pairs(doc, matched_pairs)
            _post_visibility_command(revit.uidoc, refs, hide_mode)
        except Exception as ui_exc:
            if LinkElementId is None:
                forms.alert(
                    "{} failed.\n\n{}".format(action_text, ui_exc),
                    title=TITLE,
                )
                return
            try:
                if not link_element_ids:
                    link_element_ids = _build_link_element_ids_from_pairs(doc, matched_pairs)
                if not link_element_ids:
                    raise RuntimeError("No API-compatible linked element ids were available for fallback.")
                with revit.Transaction("{} Existing Profiles".format(action_text)):
                    _apply_visibility(view, link_element_ids, hide_mode)
                apply_mode = "api"
            except Exception as api_fallback_exc:
                forms.alert(
                    "{} failed.\n\nUI command error:\n{}\n\nAPI fallback error:\n{}".format(
                        action_text,
                        ui_exc,
                        api_fallback_exc,
                    ),
                    title=TITLE,
                )
                return
    else:
        if not link_element_ids and LinkElementId is not None:
            try:
                link_element_ids = _build_link_element_ids_from_pairs(doc, matched_pairs)
            except Exception:
                link_element_ids = []
        txn_name = "{} Existing Profiles".format(action_text)
        try:
            if not link_element_ids:
                raise RuntimeError("No API-compatible linked element ids were available for visibility update.")
            with revit.Transaction(txn_name):
                _apply_visibility(view, link_element_ids, hide_mode)
        except Exception as api_exc:
            try:
                refs = _build_link_references_from_pairs(doc, matched_pairs)
                _post_visibility_command(revit.uidoc, refs, hide_mode)
                apply_mode = "ui"
            except Exception as ui_fallback_exc:
                forms.alert(
                    "{} failed.\n\nAPI error:\n{}\n\nUI fallback error:\n{}".format(
                        action_text,
                        api_exc,
                        ui_fallback_exc,
                    ),
                    title=TITLE,
                )
                return

    _set_toggle_state(doc, hide_mode)
    if hide_mode:
        _set_hidden_pairs(doc, matched_pairs)
    else:
        _set_hidden_pairs(doc, [])
    _set_button_icon(hide_mode)
    state_text = "ON (hidden)" if hide_mode else "OFF (visible)"
    yaml_label = get_yaml_display_name(data_path)
    if not match_records:
        match_records = _build_report_records_from_pairs(doc, matched_pairs, profile_lookup)
    _print_match_report(action_text, apply_mode, yaml_label, match_records, scanned, link_count, matched)
    forms.show_balloon(
        TITLE,
        "State: {}\nMatched linked elements: {}\nElements scanned: {}\nApply mode: {}\nYAML: {}\nReport: pyRevit output panel".format(
            state_text,
            matched,
            scanned,
            apply_mode.upper(),
            yaml_label,
        ),
    )


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    doc = None
    try:
        uidoc = getattr(__rvt__, "ActiveUIDocument", None)
        doc = uidoc.Document if uidoc else None
    except Exception:
        doc = None
    state = _get_toggle_state(doc, default=False)
    _set_button_icon(state, script_cmp=script_cmp, ui_button_cmp=ui_button_cmp)


if __name__ == "__main__":
    main()
