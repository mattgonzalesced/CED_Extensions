# -*- coding: utf-8 -*-
"""
Edit YAML Profiles
------------------
UI to edit CEDLib.lib/profileData.yaml (used by Place Elements - YAML).
Allows editing offsets, parameters, tags, category, and is_group for each CAD profile/type.
"""

import io
import json
import os
import hashlib
import datetime

from pyrevit import script, forms

# Add CEDLib.lib to sys.path for shared UI/logic classes
import sys
LIB_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib"))
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from UIClasses.ProfileEditorWindow import ProfileEditorWindow  # noqa: E402
from profile_schema import (  # noqa: E402
    equipment_defs_to_legacy,
    legacy_to_equipment_defs,
    load_data as load_profile_data,
    save_data as save_profile_data,
)
from LogicClasses.yaml_path_cache import get_cached_yaml_path, set_cached_yaml_path  # noqa: E402

DEFAULT_DATA_PATH = os.path.join(LIB_ROOT, "profileData.yaml")


# --------------------------------------------------------------------------- #
# YAML helpers
# --------------------------------------------------------------------------- #


def _parse_scalar(token):
    token = (token or "").strip()
    if not token:
        return ""
    if token in ("{}",):
        return {}
    if token in ("[]",):
        return []
    if token.startswith('"') and token.endswith('"'):
        return token[1:-1]
    lowered = token.lower()
    if lowered in ("true", "false"):
        return lowered == "true"
    if lowered == "null":
        return None
    try:
        if "." in token:
            return float(token)
        return int(token)
    except Exception:
        return token


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
                    result.append(_parse_scalar(remainder))
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
                    result[key] = _parse_scalar(remainder)
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


# --------------------------------------------------------------------------- #
# Simple shims so the existing ProfileEditorWindow can work on profileData.yaml
# --------------------------------------------------------------------------- #

class OffsetShim(object):
    def __init__(self, dct=None):
        dct = dct or {}
        self.x_inches = float(dct.get("x_inches", 0.0) or 0.0)
        self.y_inches = float(dct.get("y_inches", 0.0) or 0.0)
        self.z_inches = float(dct.get("z_inches", 0.0) or 0.0)
        self.rotation_deg = float(dct.get("rotation_deg", 0.0) or 0.0)


class InstanceConfigShim(object):
    def __init__(self, dct=None):
        dct = dct or {}
        offs = dct.get("offsets") or [{}]
        self.offsets = [OffsetShim(o) for o in offs]
        self.parameters = dict(dct.get("parameters") or {})
        raw_tags = dct.get("tags") or []
        shim_tags = []
        for tg in raw_tags:
            if isinstance(tg, dict):
                shim_tags.append({
                    "category_name": tg.get("category_name"),
                    "family_name": tg.get("family_name"),
                    "type_name": tg.get("type_name"),
                    "parameters": tg.get("parameters") or {},
                    "offsets": tg.get("offsets") or {},
                })
            else:
                shim_tags.append(tg)
        self.tags = shim_tags

    def get_offset(self, idx):
        if not self.offsets:
            self.offsets = [OffsetShim()]
        try:
            return self.offsets[idx]
        except Exception:
            return self.offsets[0]


class TypeConfigShim(object):
    def __init__(self, dct=None):
        dct = dct or {}
        self.label = dct.get("label")
        self.category_name = dct.get("category_name")
        self.is_group = bool(dct.get("is_group", False))
        self.instance_config = InstanceConfigShim(dct.get("instance_config") or {})


class ProfileShim(object):
    def __init__(self, dct=None):
        dct = dct or {}
        self.cad_name = dct.get("cad_name")
        self._types = [TypeConfigShim(t) for t in (dct.get("types") or [])]

    def get_types(self):
        return list(self._types)

    def find_type_by_label(self, label):
        for t in self._types:
            if getattr(t, "label", None) == label:
                return t
        return None


def _file_hash(path):
    if not os.path.exists(path):
        return ""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def _append_log(action, cad_names, before_hash, after_hash, log_path):
    entry = {
        "timestamp": datetime.datetime.utcnow().isoformat() + "Z",
        "user": os.getenv("USERNAME") or os.getenv("USER") or "unknown",
        "action": action,
        "cad_names": list(cad_names or []),
        "before_hash": before_hash,
        "after_hash": after_hash,
    }
    with io.open(log_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=True) + "\n")


def _dict_from_shims(profiles):
    out = {"profiles": []}
    for p in profiles.values():
        types = []
        for t in p.get_types():
            inst = t.instance_config
            offsets = []
            for off in getattr(inst, "offsets", []) or []:
                offsets.append({
                    "x_inches": off.x_inches,
                    "y_inches": off.y_inches,
                    "z_inches": off.z_inches,
                    "rotation_deg": off.rotation_deg,
                })
            # Filter parameters to only electrical-like categories
            params = getattr(inst, "parameters", {}) or {}
            cat_l = (t.category_name or "").lower() if hasattr(t, "category_name") else ""
            is_electrical = ("electrical" in cat_l) or ("lighting" in cat_l) or ("data" in cat_l)
            if not is_electrical:
                params = {}
            types.append({
                "label": t.label,
                "category_name": t.category_name,
                "is_group": t.is_group,
                "instance_config": {
                    "offsets": offsets or [{}],
                    "parameters": params,
                    "tags": _serialize_tags(getattr(inst, "tags", []) or []),
                },
            })
        out["profiles"].append({
            "cad_name": p.cad_name,
            "types": types,
        })
    return out


def _shims_from_dict(data):
    profiles = {}
    for p in data.get("profiles") or []:
        cad = p.get("cad_name")
        if not cad:
            continue
        profiles[cad] = ProfileShim(p)
    return profiles


def main():
    data_path = _pick_profile_data_path()
    if not data_path:
        return
    log_path = os.path.join(os.path.dirname(data_path), "profileData.log")

    # XAML lives alongside the UI class in CEDLib.lib/UIClasses
    xaml_path = os.path.join(LIB_ROOT, "UIClasses", "ProfileEditorWindow.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("ProfileEditorWindow.xaml not found under CEDLib.lib/UIClasses.", title="Edit YAML Profiles")
        return

    raw_data = _load_profile_store(data_path)
    raw_defs = raw_data.get("equipment_definitions") or []
    legacy_dict = {"profiles": equipment_defs_to_legacy(raw_defs)}
    shim_profiles = _shims_from_dict(legacy_dict)

    window = ProfileEditorWindow(xaml_path, shim_profiles)
    result = window.show_dialog()
    if not result:
        return

    try:
        before_hash = _file_hash(data_path)
        updated_dict = _dict_from_shims(shim_profiles)
        updated_defs = legacy_to_equipment_defs(
            updated_dict.get("profiles") or [],
            raw_defs,
        )
        save_profile_data(data_path, {"equipment_definitions": updated_defs})
        after_hash = _file_hash(data_path)
        _append_log("edit", sorted(shim_profiles.keys()), before_hash, after_hash, log_path)
        forms.alert("Saved profile changes to profileData.yaml.\nReload Place Elements (YAML) to use the updates.", title="Edit YAML Profiles")
    except Exception as ex:
        forms.alert("Failed to save profileData.yaml:\n\n{}".format(ex), title="Edit YAML Profiles")


def _serialize_tags(tags):
    serialized = []
    for tg in tags:
        if isinstance(tg, dict):
            serialized.append(tg)
            continue
        family = getattr(tg, "family_name", None) or getattr(tg, "family", None)
        type_name = getattr(tg, "type_name", None) or getattr(tg, "type", None)
        cat = getattr(tg, "category_name", None) or getattr(tg, "category", None)
        offsets = getattr(tg, "offsets", None)
        offsets_dict = {}
        if isinstance(offsets, dict):
            offsets_dict = {
                "x_inches": float(offsets.get("x_inches", 0.0) or 0.0),
                "y_inches": float(offsets.get("y_inches", 0.0) or 0.0),
                "z_inches": float(offsets.get("z_inches", 0.0) or 0.0),
                "rotation_deg": float(offsets.get("rotation_deg", 0.0) or 0.0),
            }
        elif hasattr(offsets, "x_inches"):
            offsets_dict = {
                "x_inches": float(getattr(offsets, "x_inches", 0.0) or 0.0),
                "y_inches": float(getattr(offsets, "y_inches", 0.0) or 0.0),
                "z_inches": float(getattr(offsets, "z_inches", 0.0) or 0.0),
                "rotation_deg": float(getattr(offsets, "rotation_deg", 0.0) or 0.0),
            }
        serialized.append({
            "category_name": cat,
            "family_name": family,
            "type_name": type_name,
            "parameters": getattr(tg, "parameters", {}) or {},
            "offsets": offsets_dict,
        })
    return serialized


if __name__ == "__main__":
    main()
