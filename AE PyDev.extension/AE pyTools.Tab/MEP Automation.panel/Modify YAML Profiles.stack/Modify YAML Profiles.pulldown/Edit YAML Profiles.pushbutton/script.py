# -*- coding: utf-8 -*-
"""
Edit YAML Profiles
------------------
Edits the active YAML payload stored in Extensible Storage so Place Elements
and other tools always reference the in-model definition.
Allows editing offsets, parameters, tags, category, and is_group for each equipment definition/type.
"""

import os

from pyrevit import script, forms, revit
from Autodesk.Revit.DB import TransactionGroup

# Add CEDLib.lib to sys.path for shared UI/logic classes
import sys


def _find_cedlib_root():
    current = os.path.abspath(os.path.dirname(__file__))
    while True:
        candidate = os.path.join(current, "CEDLib.lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    raise RuntimeError("Unable to locate CEDLib.lib relative to {}".format(__file__))


try:
    LIB_ROOT = _find_cedlib_root()
except RuntimeError as exc:
    forms.alert(str(exc), title="Edit YAML Profiles")
    raise
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from UIClasses.ProfileEditorWindow import ProfileEditorWindow  # noqa: E402
from profile_schema import (  # noqa: E402
    equipment_defs_to_legacy,
    legacy_to_equipment_defs,
)
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data, save_active_yaml_data  # noqa: E402


# --------------------------------------------------------------------------- #
# YAML helpers
# --------------------------------------------------------------------------- #


# --------------------------------------------------------------------------- #
# Simple shims so the existing ProfileEditorWindow can work on stored YAML
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
        self.led_id = dct.get("led_id")
        self.element_def_id = dct.get("id") or dct.get("led_id")
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
            params = getattr(inst, "parameters", {}) or {}
            types.append({
                "label": t.label,
                "id": getattr(t, "element_def_id", None) or getattr(t, "led_id", None),
                "led_id": getattr(t, "led_id", None),
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


def _has_negative_z(value):
    if value is None:
        return False
    try:
        return float(value) < 0.0
    except Exception:
        return False


def _find_negative_z_offsets(profile_dict):
    negatives = []
    for profile in profile_dict.get("profiles") or []:
        cad_name = profile.get("cad_name") or "<Unnamed CAD>"
        for type_entry in profile.get("types") or []:
            label = type_entry.get("label") or type_entry.get("id") or "<Unnamed Type>"
            inst = type_entry.get("instance_config") or {}
            for idx, offset in enumerate(inst.get("offsets") or []):
                if _has_negative_z(offset.get("z_inches")):
                    negatives.append({
                        "cad": cad_name,
                        "label": label,
                        "index": idx + 1,
                        "value": float(offset.get("z_inches") or 0.0),
                        "source": "offset",
                    })
            for tag in inst.get("tags") or []:
                offsets = tag.get("offsets") or {}
                if _has_negative_z(offsets.get("z_inches")):
                    negatives.append({
                        "cad": cad_name,
                        "label": label,
                        "index": None,
                        "value": float(offsets.get("z_inches") or 0.0),
                        "source": "tag",
                    })
    return negatives


def _shims_from_dict(data):
    profiles = {}
    for p in data.get("profiles") or []:
        cad = p.get("cad_name")
        if not cad:
            continue
        profiles[cad] = ProfileShim(p)
    return profiles


def _build_relations_index(equipment_defs):
    id_to_name = {}
    for entry in equipment_defs or []:
        eq_id = (entry.get("id") or "").strip()
        eq_name = (entry.get("name") or eq_id or "").strip()
        if eq_id:
            id_to_name[eq_id] = eq_name
    relations = {}
    for entry in equipment_defs or []:
        profile_name = (entry.get("name") or entry.get("id") or "").strip()
        entry_id = (entry.get("id") or "").strip()
        rel = entry.get("linked_relations") or {}
        parent_block = rel.get("parent") or {}
        parent_id = (parent_block.get("equipment_id") or "").strip()
        parent_led = (parent_block.get("parent_led_id") or "").strip()
        child_entries = []
        for child in rel.get("children") or []:
            cid = (child.get("equipment_id") or "").strip()
            if not cid:
                continue
            anchor_led = (child.get("anchor_led_id") or "").strip()
            child_entries.append({
                "id": cid,
                "name": id_to_name.get(cid, ""),
                "anchor_led_id": anchor_led,
            })
        data = {
            "parent_id": parent_id,
            "parent_name": id_to_name.get(parent_id, ""),
            "parent_led_id": parent_led,
            "children": child_entries,
        }
        if profile_name:
            relations[profile_name] = data
        if entry_id and entry_id not in relations:
            relations[entry_id] = data
    return relations


def main():
    doc = getattr(revit, "doc", None)
    trans_group = TransactionGroup(doc, "Edit YAML Profiles") if doc else None
    if trans_group:
        trans_group.Start()
    success = False
    try:
        try:
            data_path, raw_data = load_active_yaml_data()
        except RuntimeError as exc:
            forms.alert(str(exc), title="Edit YAML Profiles")
            return
        yaml_label = get_yaml_display_name(data_path)
        # XAML lives alongside the UI class in CEDLib.lib/UIClasses
        xaml_path = os.path.join(LIB_ROOT, "UIClasses", "ProfileEditorWindow.xaml")
        if not os.path.exists(xaml_path):
            forms.alert("ProfileEditorWindow.xaml not found under CEDLib.lib/UIClasses.", title="Edit YAML Profiles")
            return

        raw_defs = raw_data.get("equipment_definitions") or []
        relations_index = _build_relations_index(raw_defs)
        legacy_dict = {"profiles": equipment_defs_to_legacy(raw_defs)}
        shim_profiles = _shims_from_dict(legacy_dict)

        window = ProfileEditorWindow(xaml_path, shim_profiles, relations_index)
        result = window.show_dialog()
        if not result:
            return

        try:
            updated_dict = _dict_from_shims(shim_profiles)
            negatives = _find_negative_z_offsets(updated_dict)
            if negatives:
                lines = ["Negative Z-offsets detected:"]
                for entry in negatives[:5]:
                    if entry["source"] == "tag":
                        lines.append(" - {} / {} tag offsets = {:.2f}\"".format(entry["cad"], entry["label"], entry["value"]))
                    else:
                        lines.append(" - {} / {} offset #{} = {:.2f}\"".format(entry["cad"], entry["label"], entry["index"], entry["value"]))
                if len(negatives) > len(lines) - 1:
                    lines.append(" - (+{} more)".format(len(negatives) - (len(lines) - 1)))
                lines.append("")
                lines.append("Continue saving anyway?")
                proceed = forms.alert(
                    "\n".join(lines),
                    title="Edit YAML Profiles",
                    ok=False,
                    yes=True,
                    no=True,
                )
                if not proceed:
                    forms.alert("Save canceled. No changes were written.", title="Edit YAML Profiles")
                    return
            updated_defs = legacy_to_equipment_defs(
                updated_dict.get("profiles") or [],
                raw_defs,
            )
            raw_data["equipment_definitions"] = updated_defs
            save_active_yaml_data(
                None,
                raw_data,
                "Edit YAML Profiles",
                "Updated YAML profiles via editor window",
            )
            forms.alert(
                "Saved profile changes to {}.\nReload Place Elements (YAML) to use the updates.".format(yaml_label),
                title="Edit YAML Profiles",
            )
            success = True
        except Exception as ex:
            forms.alert("Failed to save {}:\n\n{}".format(yaml_label, ex), title="Edit YAML Profiles")
    finally:
        if trans_group:
            try:
                if success:
                    trans_group.Assimilate()
                else:
                    trans_group.RollBack()
            except Exception:
                pass


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
