# -*- coding: utf-8 -*-
"""
Parent parameter conflict detection and resolution.
Runs after sync to warn when child parameters that depend on parent_parameter
mappings no longer match the parent element values.
"""

import io
import json
import os
import re
import sys
import time

from pyrevit import forms, script
from Autodesk.Revit.DB import (
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    StorageType,
    Transaction,
)

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402

try:
    basestring
except NameError:  # pragma: no cover
    basestring = str

SETTING_KEY = "parent_param_conflict_check"
CONFIG_FILE = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "CEDLib.lib", "LetThereBeYAML.settings.json")
)

LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
LED_ID_PARAM_NAMES = (
    "Linked Element Definition ID",
    "Linked Element Definition Id",
    "Linked Element DefinitionID",
)
PARENT_ID_PARAM_NAMES = ("Parent ElementId", "Parent Element ID")
PARENT_PARAMETER_PATTERN = re.compile(
    r'^\s*parent_parameter\s*:\s*(?:"([^"]+)"|\'([^\']+)\'|(.+))\s*$',
    re.IGNORECASE,
)
NUMERIC_TOKEN_PATTERN = re.compile(r"[-+]?\d+(?:[.,]\d+)?")

ACTION_UPDATE_CHILD = "update_child"
ACTION_UPDATE_PARENT = "update_parent"
ACTION_SKIP = "skip"

ENV_RUNNING_KEY = "ced_parent_param_conflict_running"
ENV_LAST_RUN_KEY = "ced_parent_param_conflict_last_run"
LOCK_WINDOW_SECONDS = 60.0
LOCK_FILE = os.path.abspath(
    os.path.join(os.path.dirname(CONFIG_FILE), "parent_param_conflicts.lock.json")
)


def _read_config():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with io.open(CONFIG_FILE, "r", encoding="utf-8") as handle:  # type: ignore
            return json.load(handle)
    except Exception:
        try:
            with open(CONFIG_FILE, "r") as handle:
                return json.load(handle)
        except Exception:
            return {}


def _write_config(data):
    directory = os.path.dirname(CONFIG_FILE)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)
    with open(CONFIG_FILE, "w") as handle:
        json.dump(data, handle, indent=2)


def _get_env(name, default=None):
    try:
        value = script.get_envvar(name)
    except Exception:
        return default
    if value in (None, ""):
        return default
    return value


def _set_env(name, value):
    try:
        script.set_envvar(name, value)
    except Exception:
        return False
    return True


def _read_lock_file():
    if not os.path.exists(LOCK_FILE):
        return {}
    try:
        with io.open(LOCK_FILE, "r", encoding="utf-8") as handle:  # type: ignore
            return json.load(handle)
    except Exception:
        try:
            with open(LOCK_FILE, "r") as handle:
                return json.load(handle)
        except Exception:
            return {}


def _write_lock_file(payload):
    directory = os.path.dirname(LOCK_FILE)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)
    try:
        with io.open(LOCK_FILE, "w", encoding="utf-8") as handle:  # type: ignore
            json.dump(payload, handle, indent=2)
    except Exception:
        try:
            with open(LOCK_FILE, "w") as handle:
                json.dump(payload, handle, indent=2)
        except Exception:
            pass


def _doc_key(doc):
    if doc is None:
        return None
    try:
        return doc.PathName or doc.Title
    except Exception:
        return None


def _should_open_ui(doc):
    doc_key = _doc_key(doc)
    now = time.time()
    lock_payload = _read_lock_file()
    last_key = lock_payload.get("doc_key")
    last_ts = lock_payload.get("timestamp") or 0.0
    is_running = bool(lock_payload.get("running"))
    if is_running and last_key == doc_key and now - last_ts < LOCK_WINDOW_SECONDS:
        return False
    if last_key == doc_key and now - last_ts < LOCK_WINDOW_SECONDS:
        return False
    if str(_get_env(ENV_RUNNING_KEY, "0")).strip() == "1":
        return False
    if not doc_key:
        _set_env(ENV_RUNNING_KEY, "1")
        _write_lock_file({"doc_key": None, "timestamp": now, "running": True})
        return True
    raw = _get_env(ENV_LAST_RUN_KEY, "{}")
    try:
        payload = json.loads(raw)
    except Exception:
        payload = {}
    last_key = payload.get("doc_key")
    last_ts = payload.get("timestamp") or 0.0
    if last_key == doc_key and now - last_ts < 20.0:
        return False
    payload = {"doc_key": doc_key, "timestamp": now}
    _set_env(ENV_LAST_RUN_KEY, json.dumps(payload))
    _set_env(ENV_RUNNING_KEY, "1")
    _write_lock_file({"doc_key": doc_key, "timestamp": now, "running": True})
    return True


def _release_ui_lock():
    _set_env(ENV_RUNNING_KEY, "0")
    now = time.time()
    payload = _read_lock_file()
    payload["timestamp"] = now
    payload["running"] = False
    _write_lock_file(payload)


def get_setting(default=True):
    data = _read_config()
    if SETTING_KEY not in data:
        return bool(default)
    return bool(data.get(SETTING_KEY))


def set_setting(value):
    data = _read_config()
    data[SETTING_KEY] = bool(value)
    _write_config(data)


def _extract_parent_parameter_name(value):
    if isinstance(value, dict):
        for key in ("parent_parameter", "Parent Parameter", "parent parameter"):
            if key in value:
                raw = value.get(key)
                if raw is None:
                    return None
                name = str(raw).strip()
                return name or None
    if not isinstance(value, basestring):
        return None
    match = PARENT_PARAMETER_PATTERN.match(value)
    if not match:
        return None
    name = (match.group(1) or match.group(2) or match.group(3) or "").strip()
    return name or None


def _collect_parent_param_mappings(data):
    led_map = {}
    for eq_def in data.get("equipment_definitions") or []:
        if not isinstance(eq_def, dict):
            continue
        for linked_set in eq_def.get("linked_sets") or []:
            if not isinstance(linked_set, dict):
                continue
            for led in linked_set.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_id = led.get("id")
                if not led_id:
                    continue
                params = led.get("parameters") or {}
                if not isinstance(params, dict):
                    continue
                mappings = []
                for child_param, raw_value in params.items():
                    parent_param = _extract_parent_parameter_name(raw_value)
                    if parent_param:
                        mappings.append((child_param, parent_param))
                if mappings:
                    led_map.setdefault(led_id, []).extend(mappings)
    return led_map


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
        if text and str(text).strip():
            return str(text)
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
        pattern = re.compile(
            r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|"
            r"Location XYZ \(ft\)|Rotation \(deg\)|Parent Rotation \(deg\)|"
            r"Parent ElementId|Parent Element ID|LevelId|ElementId|FacingOrientation)\s*:\s*"
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")
    led_id = (entries.get("Linked Element Definition ID") or "").strip()
    parent_element_id = None
    raw_parent = entries.get("Parent ElementId")
    if raw_parent not in (None, ""):
        try:
            parent_element_id = int(raw_parent)
        except Exception:
            try:
                parent_element_id = int(float(raw_parent))
            except Exception:
                parent_element_id = None
    return {"led_id": led_id, "parent_element_id": parent_element_id}


def _read_param_text(param):
    if not param:
        return ""
    storage = getattr(param, "StorageType", None)
    if storage == StorageType.Integer:
        try:
            return str(param.AsInteger())
        except Exception:
            return ""
    if storage == StorageType.ElementId:
        try:
            elem_id = param.AsElementId()
        except Exception:
            elem_id = None
        if elem_id is None:
            return ""
        try:
            return str(elem_id.IntegerValue)
        except Exception:
            return str(elem_id)
    if storage == StorageType.Double:
        try:
            return str(param.AsDouble())
        except Exception:
            pass
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
    return str(text or "")


def _read_param_int(param):
    if not param:
        return None
    storage = getattr(param, "StorageType", None)
    if storage == StorageType.Integer:
        try:
            return int(param.AsInteger())
        except Exception:
            return None
    if storage == StorageType.ElementId:
        try:
            elem_id = param.AsElementId()
        except Exception:
            elem_id = None
        if elem_id is None:
            return None
        try:
            return int(elem_id.IntegerValue)
        except Exception:
            return None
    if storage == StorageType.Double:
        try:
            return int(round(param.AsDouble()))
        except Exception:
            return None
    text = _read_param_text(param)
    if text not in (None, ""):
        try:
            return int(text)
        except Exception:
            try:
                return int(float(text))
            except Exception:
                return None
    return None


def _payload_from_element_params(elem):
    if elem is None:
        return {}
    led_id = ""
    for name in LED_ID_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        text = _read_param_text(param)
        if text:
            led_id = text.strip()
            break
    parent_id = None
    for name in PARENT_ID_PARAM_NAMES:
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        parent_id = _read_param_int(param)
        if parent_id is not None:
            break
    payload = {}
    if led_id:
        payload["led_id"] = led_id
    if parent_id is not None:
        payload["parent_element_id"] = parent_id
    return payload


def _collect_candidate_elements(doc):
    elements = []
    seen = set()
    for cls in (FamilyInstance, Group):
        try:
            collector = FilteredElementCollector(doc).OfClass(cls).WhereElementIsNotElementType()
        except Exception:
            continue
        for elem in collector:
            try:
                elem_id = elem.Id.IntegerValue
            except Exception:
                elem_id = None
            if elem_id is None or elem_id in seen:
                continue
            seen.add(elem_id)
            elements.append(elem)
    return elements


def _element_label(elem):
    if elem is None:
        return "<missing>"
    label = None
    if isinstance(elem, FamilyInstance):
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            label = "{} : {}".format(fam_name, type_name)
    if not label:
        try:
            label = getattr(elem, "Name", None)
        except Exception:
            label = None
    if not label:
        label = "<element>"
    try:
        elem_id = elem.Id.IntegerValue
    except Exception:
        elem_id = None
    if elem_id is not None:
        return "{} (Id:{})".format(label, elem_id)
    return label


def _numeric_token(text):
    if text in (None, ""):
        return None
    match = NUMERIC_TOKEN_PATTERN.search(str(text))
    if not match:
        return None
    return match.group(0).replace(",", ".")


def _try_float(text):
    token = _numeric_token(text)
    if not token:
        return None
    try:
        return float(token)
    except Exception:
        return None


def _param_info(param):
    info = {
        "exists": bool(param),
        "has_value": False,
        "display": "<missing>",
        "numeric": None,
        "storage": None,
    }
    if not param:
        return info
    info["storage"] = getattr(param, "StorageType", None)
    has_value = True
    try:
        has_value = param.HasValue
    except Exception:
        has_value = True
    storage = info["storage"]
    if storage == StorageType.String:
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
        if value not in (None, "") and str(value).strip():
            info["display"] = str(value)
            info["has_value"] = True
        else:
            info["display"] = "<unset>"
            info["has_value"] = False
        return info
    if not has_value:
        info["display"] = "<unset>"
        info["has_value"] = False
        return info
    if storage == StorageType.Integer:
        try:
            value = param.AsInteger()
        except Exception:
            value = None
        if value is None:
            info["display"] = "<unset>"
            info["has_value"] = False
        else:
            info["display"] = str(value)
            info["numeric"] = float(value)
            info["has_value"] = True
        return info
    if storage == StorageType.Double:
        try:
            raw = param.AsDouble()
        except Exception:
            raw = None
        if raw is None:
            info["display"] = "<unset>"
            info["has_value"] = False
        else:
            display = None
            try:
                display = param.AsValueString()
            except Exception:
                display = None
            if display in (None, ""):
                display = raw
            info["display"] = str(display)
            info["numeric"] = _try_float(info["display"])
            info["has_value"] = True
        return info
    if storage == StorageType.ElementId:
        try:
            elem_id = param.AsElementId()
        except Exception:
            elem_id = None
        if elem_id is None:
            info["display"] = "<unset>"
            info["has_value"] = False
        else:
            try:
                int_val = elem_id.IntegerValue
            except Exception:
                int_val = None
            if int_val is None:
                info["display"] = str(elem_id)
            else:
                info["display"] = str(int_val)
                info["numeric"] = float(int_val)
            info["has_value"] = True
        return info
    info["display"] = "<unset>" if not has_value else "<value>"
    info["has_value"] = bool(has_value)
    return info


def _values_match(parent_info, child_info):
    if parent_info is None or child_info is None:
        return False
    parent_num = parent_info.get("numeric")
    child_num = child_info.get("numeric")
    if parent_num is None:
        parent_num = _try_float(parent_info.get("display"))
    if child_num is None:
        child_num = _try_float(child_info.get("display"))
    if parent_num is not None and child_num is not None:
        return abs(parent_num - child_num) <= 1e-6
    parent_text = (parent_info.get("display") or "").strip().lower()
    child_text = (child_info.get("display") or "").strip().lower()
    return parent_text == child_text


def _get_type_param(elem, param_name):
    if elem is None or not param_name or not hasattr(elem, "GetTypeId"):
        return None
    try:
        type_id = elem.GetTypeId()
    except Exception:
        type_id = None
    if not type_id:
        return None
    try:
        type_elem = elem.Document.GetElement(type_id)
    except Exception:
        type_elem = None
    if type_elem is None:
        return None
    try:
        return type_elem.LookupParameter(param_name)
    except Exception:
        return None


def _is_type_param(elem, param_name, instance_param, type_param):
    if type_param is None:
        return False
    if instance_param is None:
        return True
    try:
        return instance_param.Id == type_param.Id
    except Exception:
        return False


def _copy_param_value(source_param, target_param):
    if source_param is None or target_param is None:
        return False
    if getattr(target_param, "IsReadOnly", False):
        return False
    source_storage = getattr(source_param, "StorageType", None)
    target_storage = getattr(target_param, "StorageType", None)
    source_text = None
    try:
        source_text = source_param.AsString()
    except Exception:
        source_text = None
    if source_text in (None, ""):
        try:
            source_text = source_param.AsValueString()
        except Exception:
            source_text = None
    source_token = _numeric_token(source_text)
    source_numeric = _try_float(source_text)
    if target_storage == StorageType.String:
        if source_numeric is not None:
            value = str(int(round(source_numeric))) if abs(source_numeric - round(source_numeric)) < 1e-6 else str(source_numeric)
        elif source_token is not None:
            value = source_token
        else:
            value = source_text if source_text is not None else ""
        try:
            target_param.Set(str(value))
            return True
        except Exception:
            return False
    if target_storage == StorageType.Integer:
        value = None
        if source_storage == StorageType.Integer:
            try:
                value = source_param.AsInteger()
            except Exception:
                value = None
        elif source_storage == StorageType.Double:
            try:
                value = int(round(source_param.AsDouble()))
            except Exception:
                value = None
        else:
            value = source_numeric
            if value is not None:
                value = int(round(value))
        if value is None:
            return False
        try:
            target_param.Set(int(value))
            return True
        except Exception:
            return False
    if target_storage == StorageType.Double:
        value = None
        if source_storage == StorageType.Double:
            try:
                value = float(source_param.AsDouble())
            except Exception:
                value = None
        elif source_storage == StorageType.Integer:
            try:
                value = float(source_param.AsInteger())
            except Exception:
                value = None
        else:
            value = source_numeric
        if value is None:
            return False
        try:
            target_param.Set(float(value))
            return True
        except Exception:
            return False
    if target_storage == StorageType.ElementId:
        if source_storage == StorageType.ElementId:
            try:
                target_param.Set(source_param.AsElementId())
                return True
            except Exception:
                return False
        value = source_numeric
        if value is None:
            return False
        try:
            target_param.Set(ElementId(int(round(value))))
            return True
        except Exception:
            return False
    return False


def collect_conflicts(doc, data):
    led_map = _collect_parent_param_mappings(data)
    if not led_map:
        return []
    elements = _collect_candidate_elements(doc)
    if not elements:
        return []
    conflicts = []
    seen = set()
    for elem in elements:
        linker_text = _get_linker_text(elem)
        payload = _parse_linker_payload(linker_text) if linker_text else {}
        if not payload.get("led_id") or payload.get("parent_element_id") is None:
            param_payload = _payload_from_element_params(elem)
            for key, value in param_payload.items():
                if payload.get(key) in (None, ""):
                    payload[key] = value
        led_id = payload.get("led_id")
        if not led_id:
            continue
        mappings = led_map.get(led_id)
        if not mappings:
            continue
        parent_id = payload.get("parent_element_id")
        if parent_id is None:
            continue
        try:
            parent_elem = doc.GetElement(ElementId(int(parent_id)))
        except Exception:
            parent_elem = None
        if parent_elem is None:
            continue
        parent_label = _element_label(parent_elem)
        child_label = _element_label(elem)
        for child_param_name, parent_param_name in mappings:
            key = (elem.Id.IntegerValue, child_param_name, parent_param_name)
            if key in seen:
                continue
            seen.add(key)
            try:
                child_param = elem.LookupParameter(child_param_name)
            except Exception:
                child_param = None
            try:
                parent_param = parent_elem.LookupParameter(parent_param_name)
            except Exception:
                parent_param = None
            type_param = _get_type_param(parent_elem, parent_param_name)
            parent_is_type = _is_type_param(parent_elem, parent_param_name, parent_param, type_param)
            if parent_param is None and type_param is not None:
                parent_param = type_param
            child_info = _param_info(child_param)
            parent_info = _param_info(parent_param)
            if not parent_info["exists"] and not child_info["exists"]:
                continue
            if parent_info["exists"] and not parent_info["has_value"]:
                is_conflict = True
            elif not parent_info["exists"]:
                is_conflict = True
            elif not child_info["exists"]:
                is_conflict = True
            elif not child_info["has_value"]:
                is_conflict = True
            else:
                is_conflict = not _values_match(parent_info, child_info)
            if not is_conflict:
                continue
            allow_update_child = bool(child_param and not child_param.IsReadOnly)
            allow_update_parent = bool(
                parent_param and not parent_param.IsReadOnly and not parent_is_type
            )
            conflicts.append({
                "id": "{}:{}:{}".format(elem.Id.IntegerValue, child_param_name, parent_param_name),
                "led_id": led_id,
                "parent_id": parent_id,
                "child_id": elem.Id.IntegerValue,
                "parent_label": parent_label,
                "child_label": child_label,
                "child_param": child_param_name,
                "parent_param": parent_param_name,
                "child_display": child_info["display"],
                "parent_display": parent_info["display"],
                "parent_is_type": parent_is_type,
                "allow_update_child": allow_update_child,
                "allow_update_parent": allow_update_parent,
                "param_key": "{} -> {}".format(child_param_name, parent_param_name),
            })
    return conflicts


def resolve_conflicts(doc, conflicts, decisions):
    if not conflicts or not decisions:
        return {"updated_child": 0, "updated_parent": 0, "skipped": 0}
    updated_child = 0
    updated_parent = 0
    skipped = 0
    t = Transaction(doc, "Resolve Parent Parameter Conflicts")
    t.Start()
    try:
        for conflict in conflicts:
            action = decisions.get(conflict["id"], ACTION_SKIP)
            if action not in (ACTION_UPDATE_CHILD, ACTION_UPDATE_PARENT):
                skipped += 1
                continue
            try:
                child_elem = doc.GetElement(ElementId(int(conflict["child_id"])))
            except Exception:
                child_elem = None
            try:
                parent_elem = doc.GetElement(ElementId(int(conflict["parent_id"])))
            except Exception:
                parent_elem = None
            if child_elem is None or parent_elem is None:
                skipped += 1
                continue
            child_param = None
            parent_param = None
            try:
                child_param = child_elem.LookupParameter(conflict["child_param"])
            except Exception:
                child_param = None
            try:
                parent_param = parent_elem.LookupParameter(conflict["parent_param"])
            except Exception:
                parent_param = None
            if action == ACTION_UPDATE_CHILD:
                if not child_param or child_param.IsReadOnly or not parent_param:
                    skipped += 1
                    continue
                if _copy_param_value(parent_param, child_param):
                    updated_child += 1
                else:
                    skipped += 1
            elif action == ACTION_UPDATE_PARENT:
                if conflict.get("parent_is_type"):
                    skipped += 1
                    continue
                if not parent_param or parent_param.IsReadOnly or not child_param:
                    skipped += 1
                    continue
                if _copy_param_value(child_param, parent_param):
                    updated_parent += 1
                else:
                    skipped += 1
    except Exception:
        t.RollBack()
        raise
    else:
        t.Commit()
    return {"updated_child": updated_child, "updated_parent": updated_parent, "skipped": skipped}


def _load_ui_module():
    module_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "ParentParamConflictsWindow.py"))
    if not os.path.exists(module_path):
        return None
    try:
        import imp
        return imp.load_source("ced_parent_param_conflicts_ui", module_path)
    except Exception:
        return None


def run_sync_check(doc):
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return
    if not get_setting(default=True):
        return
    try:
        _, data = load_active_yaml_data(doc)
    except Exception:
        return
    conflicts = collect_conflicts(doc, data)
    if not conflicts:
        return
    if not _should_open_ui(doc):
        return
    try:
        param_keys = sorted({item.get("param_key") for item in conflicts if item.get("param_key")})
        ui_module = _load_ui_module()
        if ui_module is None:
            forms.alert("Parent parameter conflicts found, but UI failed to load.", title="Parent Parameter Conflicts")
            return
        xaml_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "ParentParamConflictsWindow.xaml"))
        window = ui_module.ParentParamConflictsWindow(xaml_path, conflicts, param_keys)
        result = window.show_dialog()
        if not result:
            return
        decisions = getattr(window, "decisions", {}) or {}
        if not decisions:
            return
        counts = resolve_conflicts(doc, conflicts, decisions)
        summary = [
            "Updated child parameters: {}".format(counts.get("updated_child", 0)),
            "Updated parent parameters: {}".format(counts.get("updated_parent", 0)),
            "Skipped: {}".format(counts.get("skipped", 0)),
            "",
            "Note: changes were applied after sync; sync again to publish updates.",
        ]
        forms.alert("\n".join(summary), title="Parent Parameter Conflicts")
    finally:
        _release_ui_lock()


__all__ = [
    "get_setting",
    "set_setting",
    "collect_conflicts",
    "resolve_conflicts",
    "run_sync_check",
    "ACTION_UPDATE_CHILD",
    "ACTION_UPDATE_PARENT",
    "ACTION_SKIP",
]
