# -*- coding: utf-8 -*-
"""
Toggle Existing Profiles
------------------------
Hide/unhide linked model elements in the active view when their names already
exist as equipment profiles in the active YAML (stored in Extensible Storage).
"""

import os
import re
import sys

from pyrevit import forms, revit, script
from pyrevit.revit import ui
import pyrevit.extensions as exts
output = script.get_output()
output.close_others()
from Autodesk.Revit.DB import (
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
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


def _normalize_name_ignoring_default_suffix(value):
    normalized = _normalize_name(value)
    if not normalized:
        return ""
    cleaned = re.sub(
        r"\s*:\s*(?:default(?:\s*\d+)?|defaulttype|default\s*type)$",
        "",
        normalized,
        flags=re.IGNORECASE,
    ).strip()
    if ":" in cleaned:
        left, right = cleaned.split(":", 1)
        family = (left or "").strip()
        type_name = (right or "").strip()
        if family and (not type_name or type_name == family):
            return family
    return cleaned


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


def _build_label(elem):
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            return u"{} : {}".format(fam_name, type_name)
        if type_name:
            return type_name
        if fam_name:
            return fam_name
    try:
        name = getattr(elem, "Name", None)
        if name:
            return name
    except Exception:
        pass
    return ""


def _name_variants(elem):
    variants = set()
    label = _build_label(elem)
    if label:
        variants.add(_normalize_name(label))
        variants.add(_normalize_name_ignoring_default_suffix(label))
    try:
        raw_name = getattr(elem, "Name", None)
        if raw_name:
            variants.add(_normalize_name(raw_name))
            variants.add(_normalize_name_ignoring_default_suffix(raw_name))
    except Exception:
        pass
    return {value for value in variants if value}


def _collect_profile_name_norms(data):
    norms = set()
    for eq in data.get("equipment_definitions") or []:
        if not isinstance(eq, dict):
            continue
        for raw in (eq.get("name"), eq.get("id")):
            if not raw:
                continue
            normalized = _normalize_name(raw)
            if normalized:
                norms.add(normalized)
            normalized_loose = _normalize_name_ignoring_default_suffix(raw)
            if normalized_loose:
                norms.add(normalized_loose)
    return norms


def _iter_target_link_elements(link_doc):
    for cls in (FamilyInstance, Group):
        try:
            collector = FilteredElementCollector(link_doc).OfClass(cls).WhereElementIsNotElementType()
        except Exception:
            collector = []
        for elem in collector:
            yield elem


def _collect_matching_link_element_ids(doc, profile_norms):
    if LinkElementId is None:
        raise RuntimeError("Current Revit API does not expose LinkElementId for linked-element visibility control.")

    link_element_ids = []
    matched_pairs = []
    seen = set()
    scanned = 0
    matched = 0
    link_count = 0

    try:
        link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
    except Exception:
        link_instances = []

    for link_inst in link_instances:
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        link_count += 1
        link_id_int = _element_id_value(getattr(link_inst, "Id", None), default=None)
        if link_id_int is None:
            continue
        for elem in _iter_target_link_elements(link_doc):
            scanned += 1
            elem_id_int = _element_id_value(getattr(elem, "Id", None), default=None)
            if elem_id_int is None:
                continue
            if not (_name_variants(elem) & profile_norms):
                continue
            key = (link_id_int, elem_id_int)
            if key in seen:
                continue
            seen.add(key)
            try:
                link_element_ids.append(LinkElementId(link_inst.Id, elem.Id))
            except Exception:
                continue
            matched_pairs.append(key)
            matched += 1

    return link_element_ids, matched_pairs, scanned, matched, link_count


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
    return ";".join(["{},{}".format(int(link_id), int(elem_id)) for link_id, elem_id in pairs])


def _deserialize_pairs(raw):
    results = []
    if not raw:
        return results
    for chunk in str(raw).split(";"):
        token = chunk.strip()
        if not token or "," not in token:
            continue
        link_raw, elem_raw = token.split(",", 1)
        try:
            results.append((int(link_raw), int(elem_raw)))
        except Exception:
            continue
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

    link_map = {}
    try:
        link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
    except Exception:
        link_instances = []
    for link_inst in link_instances:
        link_id_int = _element_id_value(getattr(link_inst, "Id", None), default=None)
        if link_id_int is None or link_id_int in link_map:
            continue
        link_map[link_id_int] = link_inst

    results = []
    seen = set()
    for link_id_int, elem_id_int in pairs:
        link_inst = link_map.get(link_id_int)
        if link_inst is None:
            continue
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        try:
            linked_elem = link_doc.GetElement(ElementId(int(elem_id_int)))
        except Exception:
            linked_elem = None
        if linked_elem is None:
            continue
        key = (int(link_id_int), int(elem_id_int))
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
    link_map = {}
    try:
        link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
    except Exception:
        link_instances = []
    for link_inst in link_instances:
        link_id_int = _element_id_value(getattr(link_inst, "Id", None), default=None)
        if link_id_int is None or link_id_int in link_map:
            continue
        link_map[link_id_int] = link_inst

    refs = []
    seen = set()
    for link_id_int, elem_id_int in pairs:
        link_inst = link_map.get(link_id_int)
        if link_inst is None:
            continue
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        try:
            linked_elem = link_doc.GetElement(ElementId(int(elem_id_int)))
        except Exception:
            linked_elem = None
        if linked_elem is None:
            continue
        key = (int(link_id_int), int(elem_id_int))
        if key in seen:
            continue
        seen.add(key)
        try:
            refs.append(Reference(linked_elem).CreateLinkReference(link_inst))
        except Exception:
            continue
    return refs


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
    __revit__.PostCommand(cmd_id)


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

    profile_norms = _collect_profile_name_norms(data)
    if not profile_norms:
        forms.alert("No equipment profile names were found in the active YAML.", title=TITLE)
        return

    hidden_now = _get_toggle_state(doc, default=False)
    hide_mode = not hidden_now

    scanned = 0
    matched = 0
    link_count = 0
    matched_pairs = []

    if hide_mode:
        try:
            link_element_ids, matched_pairs, scanned, matched, link_count = _collect_matching_link_element_ids(doc, profile_norms)
        except Exception as exc:
            forms.alert(
                "Unable to collect linked elements for toggling:\n\n{}".format(exc),
                title=TITLE,
            )
            return
    else:
        stored_pairs = _get_hidden_pairs(doc)
        if stored_pairs:
            try:
                link_element_ids = _build_link_element_ids_from_pairs(doc, stored_pairs)
                matched_pairs = list(stored_pairs)
                matched = len(link_element_ids)
            except Exception:
                link_element_ids = []
        else:
            link_element_ids = []
        if not link_element_ids:
            try:
                link_element_ids, matched_pairs, scanned, matched, link_count = _collect_matching_link_element_ids(doc, profile_norms)
            except Exception as exc:
                forms.alert(
                    "Unable to collect linked elements for toggling:\n\n{}".format(exc),
                    title=TITLE,
                )
                return

    if not link_element_ids:
        _set_toggle_state(doc, False)
        _set_hidden_pairs(doc, [])
        _set_button_icon(False)
        forms.alert(
            "No linked elements matched active YAML profile names.\n\n"
            "Links scanned: {}\n"
            "Elements scanned: {}\n"
            "Matches: {}".format(link_count, scanned, matched),
            title=TITLE,
        )
        return

    action_text = "Hide" if hide_mode else "Unhide"
    apply_mode = "api"
    txn_name = "{} Existing Profiles".format(action_text)
    try:
        with revit.Transaction(txn_name):
            _apply_visibility(view, link_element_ids, hide_mode)
    except Exception as exc:
        message = str(exc or "")
        mismatch = ("ICollection[ElementId]" in message) or ("LinkElementId" in message)
        if not mismatch:
            forms.alert(
                "{} failed.\n\n{}".format(action_text, exc),
                title=TITLE,
            )
            return
        try:
            refs = _build_link_references_from_pairs(doc, matched_pairs)
            _post_visibility_command(revit.uidoc, refs, hide_mode)
            apply_mode = "ui"
        except Exception as fallback_exc:
            forms.alert(
                "{} failed.\n\n{}\n\nFallback command path also failed:\n\n{}".format(
                    action_text,
                    exc,
                    fallback_exc,
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
    forms.show_balloon(
        TITLE,
        "State: {}\nMatched linked elements: {}\nElements scanned: {}\nApply mode: {}\nYAML: {}".format(
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
