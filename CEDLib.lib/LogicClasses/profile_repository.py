# -*- coding: utf-8 -*-
"""
ProfileRepository
-----------------
Loads profileData.yaml into EquipmentDefinition objects and exposes lookup helpers
keyed by CAD block name.
"""

import os

from LogicClasses.EquipmentDefinition import EquipmentDefinition
from LogicClasses.LinkedElementSet import LinkedElementSet
from LogicClasses.LinkedElementDefinition import LinkedElementDefinition
from LogicClasses.PlacementRule import PlacementRule

try:
    from .profile_schema import equipment_defs_to_legacy, load_data
except Exception:  # pragma: no cover
    from LogicClasses.profile_schema import equipment_defs_to_legacy, load_data  # type: ignore


class ProfileRepository(object):
    """
    Converts profileData.yaml (JSON format) into EquipmentDefinition objects and
    exposes lookup helpers keyed by CAD name.
    """

    def __init__(self, equipment_definitions):
        self._by_cad = {}
        self._label_map = {}
        self._anchors_by_cad = {}

        for eq_def in equipment_definitions:
            cad_name = eq_def.get_equipment_def_id()
            if not cad_name:
                continue
            self._by_cad[cad_name] = eq_def
            labels = {}
            for linked_set in eq_def.get_linked_sets() or []:
                for linked_def in linked_set.get_elements() or []:
                    lbl = linked_def.get_element_def_id()
                    if not lbl:
                        continue
                    if getattr(linked_def, "is_parent_anchor", lambda: False)():
                        self._anchors_by_cad.setdefault(cad_name, []).append(linked_def)
                        continue
                    base_lbl = lbl
                    unique_lbl = base_lbl
                    idx = 2
                    while unique_lbl in labels:
                        unique_lbl = u"{} #{}".format(base_lbl, idx)
                        idx += 1
                    labels[unique_lbl] = linked_def
            self._label_map[cad_name] = labels

    @classmethod
    def from_file(cls, data_path=None):
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        path = data_path or os.path.join(base_dir, "profileData.yaml")
        if not os.path.exists(path):
            return cls([])
        data = load_data(path)
        equipment_defs = [d for d in (data.get("equipment_definitions") or []) if isinstance(d, dict)]
        profiles = equipment_defs_to_legacy(equipment_defs)
        eq_defs = cls._parse_profiles(profiles)
        return cls(eq_defs)

    @staticmethod
    def _parse_profiles(profiles):
        """
        Build a list of EquipmentDefinition objects from the raw dict profiles.
        The YAML is shaped to match the class structure instead of changing
        constructors.
        """
        eq_defs = []
        for prof in profiles:
            if not isinstance(prof, dict):
                continue
            cad_name = (prof.get("cad_name") or prof.get("equipment_def_id") or "").strip()
            if not cad_name:
                continue

            linked_defs = []
            for type_entry in prof.get("types") or []:
                if not isinstance(type_entry, dict):
                    continue
                led_id = type_entry.get("led_id")
                set_id = type_entry.get("set_id")
                label = (type_entry.get("label") or "").strip()
                if not label:
                    continue
                # Split "Family : Type" gracefully
                if ":" in label:
                    fam_part, type_part = label.split(":", 1)
                    family_name = fam_part.strip()
                    type_name = type_part.strip()
                else:
                    family_name = None
                    type_name = label

                inst_cfg = type_entry.get("instance_config")
                if not isinstance(inst_cfg, dict):
                    inst_cfg = {}
                params = inst_cfg.get("parameters")
                if not isinstance(params, dict):
                    params = {}
                offsets = inst_cfg.get("offsets")
                if not isinstance(offsets, list) or not offsets:
                    offsets = [{}]
                offsets = [off if isinstance(off, dict) else {} for off in offsets]

                def _inch_to_ft(val):
                    try:
                        return float(val) / 12.0
                    except Exception:
                        return 0.0

                def _offset_dict_to_tuple(data):
                    if not isinstance(data, dict):
                        return None
                    return (
                        _inch_to_ft(data.get("x_inches", 0.0) or 0.0),
                        _inch_to_ft(data.get("y_inches", 0.0) or 0.0),
                        _inch_to_ft(data.get("z_inches", 0.0) or 0.0),
                    )

                tag_defs = []
                tag_sources = []
                tag_sources.extend(inst_cfg.get("tags") or [])
                tag_sources.extend(inst_cfg.get("keynotes") or [])
                def _normalize_keynote_family(value):
                    if not value:
                        return ""
                    text = str(value)
                    if ":" in text:
                        text = text.split(":", 1)[0]
                    return "".join([ch for ch in text.lower() if ch.isalnum()])

                def _is_ga_keynote_symbol(family_name):
                    return _normalize_keynote_family(family_name) == "gakeynotesymbolced"

                def _is_builtin_keynote_tag(tag_data):
                    family = tag_data.get("family_name") or tag_data.get("family") or ""
                    category = tag_data.get("category_name") or tag_data.get("category") or ""
                    if _is_ga_keynote_symbol(family):
                        return False
                    fam_text = (family or "").lower()
                    cat_text = (category or "").lower()
                    if "keynote tags" in cat_text:
                        return True
                    if "keynote tag" in fam_text:
                        return True
                    return False

                def _extract_key_value_from_params(param_dict):
                    if not isinstance(param_dict, dict):
                        return None
                    for key, value in param_dict.items():
                        if (key or "").strip().lower() in ("keynote value", "key value", "keynote"):
                            return value
                    return None

                def _normalize_tag_parameters(tag_data):
                    if not isinstance(tag_data, dict):
                        return {}
                    params = tag_data.get("parameters")
                    if isinstance(params, dict):
                        params = dict(params)
                    else:
                        params = {}
                    # Some converted YAMLs flatten keynote parameters at the top level.
                    flattened = False
                    for key in tag_data.keys():
                        key_lower = (key or "").strip().lower()
                        if "keynote" in key_lower or "key value" in key_lower:
                            flattened = True
                            break
                    if flattened:
                        for key, value in tag_data.items():
                            if key in (
                                "parameters",
                                "offsets",
                                "offset",
                                "rotation_deg",
                                "leader_elbow",
                                "leader_end",
                                "family_name",
                                "family",
                                "type_name",
                                "type",
                                "category_name",
                                "category",
                                "key_value",
                            ):
                                continue
                            key_lower = (key or "").strip().lower()
                            if "keynote" in key_lower or "key value" in key_lower:
                                if key not in params:
                                    params[key] = value
                    return params
                for tag_data in tag_sources:
                    if not isinstance(tag_data, dict):
                        continue
                    if _is_builtin_keynote_tag(tag_data):
                        continue
                    if _is_builtin_keynote_tag(tag_data):
                        continue
                    offsets_dict = tag_data.get("offsets") or {}
                    if not isinstance(offsets_dict, dict):
                        offsets_dict = {}

                    leader_elbow = _offset_dict_to_tuple(tag_data.get("leader_elbow"))
                    leader_end = _offset_dict_to_tuple(tag_data.get("leader_end"))
                    parameters = _normalize_tag_parameters(tag_data)
                    key_value = tag_data.get("key_value")
                    if key_value in (None, ""):
                        for key in ("Keynote Value", "Key Value", "Keynote"):
                            if key in tag_data and tag_data.get(key) not in (None, ""):
                                key_value = tag_data.get(key)
                                break
                    if key_value in (None, ""):
                        key_value = _extract_key_value_from_params(parameters)
                    if key_value not in (None, ""):
                        has_key = False
                        for key in parameters.keys():
                            if (key or "").strip().lower() in ("keynote value", "key value", "keynote"):
                                has_key = True
                                break
                        if not has_key:
                            parameters["Keynote Value"] = key_value
                    tag_defs.append({
                        "family": tag_data.get("family_name") or tag_data.get("family"),
                        "type": tag_data.get("type_name") or tag_data.get("type"),
                        "category": tag_data.get("category_name") or tag_data.get("category"),
                        "parameters": parameters,
                        "key_value": key_value,
                        "offset": (
                            _inch_to_ft(offsets_dict.get("x_inches", 0.0) or 0.0),
                            _inch_to_ft(offsets_dict.get("y_inches", 0.0) or 0.0),
                            _inch_to_ft(offsets_dict.get("z_inches", 0.0) or 0.0),
                        ),
                        "rotation_deg": float(offsets_dict.get("rotation_deg", 0.0) or 0.0),
                        "leader_elbow": leader_elbow,
                        "leader_end": leader_end,
                    })

                text_note_defs = []
                for note_data in inst_cfg.get("text_notes") or []:
                    if not isinstance(note_data, dict):
                        continue
                    offsets_dict = note_data.get("offsets") or {}
                    if not isinstance(offsets_dict, dict):
                        offsets_dict = {}
                    leader_defs = []
                    for leader in note_data.get("leaders") or []:
                        if not isinstance(leader, dict):
                            continue
                        leader_defs.append(dict(leader))
                    text_note_defs.append({
                        "text": note_data.get("text") or "",
                        "type_name": note_data.get("type_name"),
                        "offset": (
                            _inch_to_ft(offsets_dict.get("x_inches", 0.0) or 0.0),
                            _inch_to_ft(offsets_dict.get("y_inches", 0.0) or 0.0),
                            _inch_to_ft(offsets_dict.get("z_inches", 0.0) or 0.0),
                        ),
                        "rotation_deg": float(offsets_dict.get("rotation_deg", 0.0) or 0.0),
                        "width": _inch_to_ft(note_data.get("width_inches", 0.0) or 0.0),
                        "leaders": leader_defs,
                    })

                is_group_flag = bool(type_entry.get("is_group")) or (type_entry.get("category_name") == "Model Groups")

                for idx, off in enumerate(offsets):
                    placement = PlacementRule(
                        offset_xyz=(
                            float(off.get("x_inches", 0.0)) / 12.0,
                            float(off.get("y_inches", 0.0)) / 12.0,
                            float(off.get("z_inches", 0.0)) / 12.0,
                        ),
                        rotation_degrees=float(off.get("rotation_deg", 0.0)),
                        placement_mode="group" if is_group_flag else None,
                        tags=tag_defs,
                        text_notes=text_note_defs,
                    )
                    element_def_id = label if idx == 0 else u"{} #{}".format(label, idx + 1)
                    led = LinkedElementDefinition(
                        element_def_id=element_def_id,
                        category=type_entry.get("category_name"),
                        family=family_name,
                        type_name=type_name,
                        placement=placement,
                        static_params=params,
                        dynamic_params=None,
                        allow_recreate=False,
                        is_optional=False,
                        is_parent_anchor=bool(type_entry.get("is_parent_anchor")),
                    )
                    if led_id:
                        setattr(led, "_ced_led_id", led_id)
                    if set_id:
                        setattr(led, "_ced_set_id", set_id)
                    linked_defs.append(led)

            linked_set = LinkedElementSet(
                set_def_id=cad_name,
                name=cad_name,
                elements=linked_defs,
            )
            eq_def = EquipmentDefinition(
                equipment_def_id=cad_name,
                name=cad_name,
                linked_sets=[linked_set],
            )
            eq_defs.append(eq_def)
        return eq_defs

    def cad_names(self):
        return sorted(self._by_cad.keys())

    def labels_for_cad(self, cad_name):
        return list((self._label_map.get(cad_name) or {}).keys())

    def definition_for_label(self, cad_name, label):
        return (self._label_map.get(cad_name) or {}).get(label)

    def anchor_definitions_for_cad(self, cad_name):
        return list(self._anchors_by_cad.get(cad_name, []))


__all__ = ["ProfileRepository"]
