# -*- coding: utf-8 -*-
# Edit Element Linker profiles (CadBlockProfile / TypeConfig) at runtime and save to element_data.yaml

import io
import json
import os
import hashlib
import datetime

from pyrevit import script, forms

from Element_Linker import CAD_BLOCK_PROFILES
from ProfileEditorWindow import ProfileEditorWindow

ELEMENT_DATA_PATH = script.get_bundle_file(os.path.join("..", "..", "..", "..", "lib", "element_data.yaml"))
if not ELEMENT_DATA_PATH or not os.path.exists(ELEMENT_DATA_PATH):
    ELEMENT_DATA_PATH = os.path.abspath(os.path.join(script.get_script_path(), "..", "..", "..", "..", "lib", "element_data.yaml"))


def _offset_to_dict(offset_obj):
    return {
        "x_inches": getattr(offset_obj, "x_inches", 0.0) or 0.0,
        "y_inches": getattr(offset_obj, "y_inches", 0.0) or 0.0,
        "z_inches": getattr(offset_obj, "z_inches", 0.0) or 0.0,
        "rotation_deg": getattr(offset_obj, "rotation_deg", 0.0) or 0.0,
    }


def _serialize_instance(instance_config):
    params = getattr(instance_config, "parameters", {}) or {}

    offsets = []
    # InstanceConfig stores offsets in _offsets; also accept .offsets if set by UI
    offsets_src = getattr(instance_config, "_offsets", None) or getattr(instance_config, "offsets", None) or []
    for off in offsets_src:
        offsets.append(_offset_to_dict(off))

    tags = []
    tag_list = getattr(instance_config, "tags", None)
    if tag_list is None and hasattr(instance_config, "get_tags"):
        try:
            tag_list = instance_config.get_tags()
        except Exception:
            tag_list = []
    tag_list = tag_list or []
    for tg in tag_list:
        tg_offsets = getattr(tg, "offsets", None)
        tags.append({
            "category_name": getattr(tg, "category_name", None),
            "family_name": getattr(tg, "family_name", None),
            "type_name": getattr(tg, "type_name", None),
            "parameters": getattr(tg, "parameters", {}) or {},
            "offsets": _offset_to_dict(tg_offsets) if tg_offsets else _offset_to_dict(None),
        })

    return {
        "parameters": params,
        "offsets": offsets,
        "tags": tags,
    }


def _serialize_profiles(registry):
    profiles = []
    for cad_name, profile in registry.items():
        types = []
        type_list = []
        if hasattr(profile, "get_types"):
            type_list = profile.get_types()
        elif hasattr(profile, "_types"):
            type_list = profile._types or []

        for tc in type_list:
            inst_cfg = getattr(tc, "instance_config", None)
            types.append({
                "label": getattr(tc, "label", None),
                "category_name": getattr(tc, "category_name", None),
                "is_group": getattr(tc, "is_group", False),
                "instance_config": _serialize_instance(inst_cfg) if inst_cfg else {},
            })

        profiles.append({
            "cad_name": cad_name,
            "types": types,
        })

    return {"profiles": profiles}


def _profile_priority(profile_dict):
    cats = set()
    for t in profile_dict.get("types", []):
        cat = t.get("category_name")
        if cat:
            cats.add(cat)

    if "Electrical Fixtures" in cats:
        order = 0
    elif cats and cats.issubset({"Data Devices"}):
        order = 1
    elif "Lighting Fixtures" in cats:
        order = 2
    elif "Plumbing Fixtures" in cats:
        order = 3
    else:
        order = 4

    return (order, profile_dict.get("cad_name", ""))


def _sort_profiles_in_place(data):
    profiles = data.get("profiles") or []
    profiles.sort(key=_profile_priority)
    data["profiles"] = profiles


def _persist_to_yaml(path, data):
    if not path:
        raise IOError("element_data.yaml path not resolved")
    _sort_profiles_in_place(data)
    with io.open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def _file_hash(path):
    if not os.path.exists(path):
        return ""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def _append_log(action, cad_names, before_hash, after_hash):
    log_path = os.path.join(os.path.dirname(ELEMENT_DATA_PATH), "element_data.log")
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


def main():
    xaml_path = script.get_bundle_file('ProfileEditorWindow.xaml')
    if not xaml_path:
        forms.alert(
            "ProfileEditorWindow.xaml not found in the bundle.",
            title="Edit Element Linker Profiles"
        )
        return

    window = ProfileEditorWindow(xaml_path, CAD_BLOCK_PROFILES)
    result = window.show_dialog()
    if result:
        try:
            before_hash = _file_hash(ELEMENT_DATA_PATH)
            data = _serialize_profiles(CAD_BLOCK_PROFILES)
            _persist_to_yaml(ELEMENT_DATA_PATH, data)
            after_hash = _file_hash(ELEMENT_DATA_PATH)
            _append_log("edit", sorted(CAD_BLOCK_PROFILES.keys()), before_hash, after_hash)
            forms.alert("Saved profile changes to element_data.yaml.\nReload Populate Elements to see your updates.", title="Edit Element Linker Profiles")
        except Exception as ex:
            forms.alert("Failed to save to element_data.yaml:\n\n{}".format(ex), title="Edit Element Linker Profiles")
    else:
        # Nothing to persist if dialog was cancelled
        pass


if __name__ == "__main__":
    main()
