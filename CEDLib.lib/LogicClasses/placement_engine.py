# -*- coding: utf-8 -*-
"""
Placement engine that consumes profile definitions and places Revit elements.
"""

import os
import math
import re

from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    FamilySymbol,
    GroupType,
    Group,
    Element,
    BuiltInCategory,
    BuiltInParameter,
    XYZ,
    Line,
    Structure,
    ElementTransformUtils,
    Level,
    ViewType,
    IndependentTag,
    TextNote,
    TextNoteType,
    TextNoteLeaderTypes,
    Reference,
    TagMode,
    TagOrientation,
    ElementId,
)
from System import Enum

from LogicClasses.csv_helpers import feet_inch_to_inches
from LogicClasses.tag_utils import tag_key_from_dict

try:
    basestring
except NameError:  # Python 3 fallback
    basestring = str

ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")
SAFE_HASH = u"\uff03"
PARENT_PARAMETER_PATTERN = re.compile(
    r'^\s*parent_parameter\s*:\s*(?:"([^"]+)"|\'([^\']+)\')\s*$',
    re.IGNORECASE,
)


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
            r"(Linked Element Definition ID|Set Definition ID|Host Name|Parent_location|Location XYZ \(ft\)|"
            r"Rotation \(deg\)|Parent Rotation \(deg\)|Parent ElementId|LevelId|"
            r"ElementId|FacingOrientation)\s*:\s*"
        )
        matches = list(pattern.finditer(text))
        for idx, match in enumerate(matches):
            key = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[start:end].strip().rstrip(",")
            entries[key] = value.strip(" ,")

    def _as_int(value):
        try:
            return int(value)
        except Exception:
            return None

    def _as_float(value):
        try:
            return float(value)
        except Exception:
            return None

    def _as_xyz(value):
        if not value:
            return None
        parts = [p.strip() for p in value.split(",")]
        if len(parts) != 3:
            return None
        try:
            return tuple(float(p) for p in parts)
        except Exception:
            return None

    return {
        "led_id": (entries.get("Linked Element Definition ID", "") or "").strip(),
        "set_id": (entries.get("Set Definition ID", "") or "").strip(),
        "host_name": (entries.get("Host Name", "") or "").strip(),
        "parent_location": _as_xyz(entries.get("Parent_location", "")),
        "level_id": _as_int(entries.get("LevelId", "")),
        "element_id": _as_int(entries.get("ElementId", "")),
        "parent_element_id": _as_int(entries.get("Parent ElementId", "")),
        "location": _as_xyz(entries.get("Location XYZ (ft)", "")),
        "rotation": _as_float(entries.get("Rotation (deg)", "")),
        "parent_rotation": _as_float(entries.get("Parent Rotation (deg)", "")),
    }


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)

def _format_xyz_value(value):
    if not value:
        return ""
    try:
        return "{:.6f},{:.6f},{:.6f}".format(value.X, value.Y, value.Z)
    except Exception:
        pass
    try:
        return "{:.6f},{:.6f},{:.6f}".format(value[0], value[1], value[2])
    except Exception:
        return ""


def _build_linker_payload(
    led_id,
    set_id,
    location,
    rotation_deg,
    level_id,
    element_id,
    facing,
    parent_rotation_deg=None,
    host_name=None,
    parent_element_id=None,
    parent_location=None,
):
    rotation = float(rotation_deg or 0.0)
    parent_rotation = float(parent_rotation_deg or 0.0)
    parent_location_text = _format_xyz_value(parent_location)
    if not parent_location_text:
        parent_location_text = "Not found"
    parts = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Host Name: {}".format(host_name or ""),
        "Parent_location: {}".format(parent_location_text),
        "Location XYZ (ft): {}".format(_format_xyz(location)),
        "Rotation (deg): {:.6f}".format(rotation),
        "Parent Rotation (deg): {:.6f}".format(parent_rotation),
        "Parent ElementId: {}".format(parent_element_id if parent_element_id is not None else ""),
        "LevelId: {}".format(level_id if level_id is not None else ""),
        "ElementId: {}".format(element_id if element_id is not None else ""),
        "FacingOrientation: {}".format(_format_xyz(facing)),
    ]
    return ", ".join(parts).strip()


class PlaceElementsEngine(object):
    def __init__(
        self,
        doc,
        repo,
        default_level=None,
        tag_view_map=None,
        allow_tags=True,
        allow_text_notes=False,
        max_tag_distance_feet=None,
        transaction_name="Place Elements (YAML)",
        apply_recorded_level=True,
    ):
        self.doc = doc
        self.repo = repo
        self.default_level = default_level
        self.tag_view_map = tag_view_map or {}
        self.allow_tags = bool(allow_tags)
        self.allow_text_notes = bool(allow_text_notes)
        self.max_tag_distance_feet = max_tag_distance_feet
        self.transaction_name = transaction_name or "Place Elements (YAML)"
        self.apply_recorded_level = bool(apply_recorded_level)
        self._init_symbol_map()
        self._init_group_map()
        self._init_text_note_types()
        self._build_repo_name_lookup()

    def _init_symbol_map(self):
        """Map 'Family : Type' to FamilySymbol."""
        self.symbol_label_map = {}
        self.symbol_label_map_by_category = {}
        self._activated_symbols = set()
        symbols = list(FilteredElementCollector(self.doc).OfClass(FamilySymbol).ToElements())
        for sym in symbols:
            try:
                family = getattr(sym, "Family", None)
                fam_name = getattr(family, "Name", None) if family else None
                if not fam_name:
                    continue
                type_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                type_name = type_param.AsString() if type_param else None
                if not type_name and hasattr(sym, "Name"):
                    type_name = sym.Name
                if not type_name:
                    continue
                label = u"{} : {}".format(fam_name, type_name)
                self.symbol_label_map[label] = sym
                cat = getattr(sym, "Category", None)
                cat_name = getattr(cat, "Name", None) if cat else None
                if cat_name:
                    key = (label.strip().lower(), cat_name.strip().lower())
                    if key not in self.symbol_label_map_by_category:
                        self.symbol_label_map_by_category[key] = sym
            except Exception:
                continue

    def _init_group_map(self):
        """
        Map model group names (and attached detail variants) to group type tuples.

        Values are stored as (model_group_type, detail_group_type_or_None) so that
        YAML labels like "Detail : Model" from the old JSON tool can be resolved.
        """
        self.group_label_map = {}
        self._all_group_types = []
        self._detail_group_map = {}

        def _store_label(label, entry):
            """
            Store label in map (exact + lowercase alias). Lowercase alias allows case-insensitive lookup
            without scanning all group types.
            """
            cleaned = (label or "").strip()
            if not cleaned:
                return
            self.group_label_map[cleaned] = entry
            lower = cleaned.lower()
            if lower not in self.group_label_map:
                self.group_label_map[lower] = entry

        # Match original JSON tool: collect model group TYPES (be permissive)
        groups = []
        try:
            groups += list(
                FilteredElementCollector(self.doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .ToElements()
            )
        except Exception:
            pass

        # Extra pass without WhereElementIsElementType (some hosts return types as "instances")
        try:
            groups += list(
                FilteredElementCollector(self.doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .ToElements()
            )
        except Exception:
            pass

        # Fallback: any GroupType found
        try:
            groups += list(
                FilteredElementCollector(self.doc)
                .OfClass(GroupType)
                .ToElements()
            )
        except Exception:
            pass

        # Deduplicate by ElementId
        dedup = {}
        for g in groups:
            try:
                dedup[g.Id.IntegerValue] = g
            except Exception:
                continue
        groups = list(dedup.values())

        # Keep only GroupType elements (skip instances that slipped through)
        try:
            groups = [g for g in groups if isinstance(g, GroupType)]
        except Exception:
            filtered = []
            for g in groups:
                try:
                    if g and g.GetType() == GroupType:
                        filtered.append(g)
                except Exception:
                    continue
            groups = filtered

        collected_names = []
        for gtype in groups:
            try:
                # Do not filter on Category; some GroupTypes return Category=None even when valid.
                try:
                    from Autodesk.Revit.DB import Element  # late import for safer Name access

                    name = (Element.Name.__get__(gtype) or "").strip()
                except Exception:
                    name = (getattr(gtype, "Name", "") or "").strip()
                if not name:
                    continue
                collected_names.append(name)
                self._all_group_types.append(gtype)
                # Base entries (no attached detail)
                base_entry = (gtype, None)
                _store_label(name, base_entry)
                _store_label(u"None : {}".format(name), base_entry)
                _store_label(u"{} : None".format(name), base_entry)

                # Add attached detail group variants so labels from the JSON tool work
                try:
                    detail_ids = gtype.GetAvailableAttachedDetailGroupTypeIds()
                except Exception:
                    detail_ids = []

                for did in detail_ids or []:
                    try:
                        detail_type = self.doc.GetElement(did)
                        detail_name = (detail_type.Name or "").strip()
                    except Exception:
                        detail_type = None
                        detail_name = ""
                    if not detail_name:
                        continue
                    label = u"{} : {}".format(detail_name, name)  # "<Detail> : <Model>"
                    entry = (gtype, detail_type)
                    _store_label(label, entry)

            except Exception:
                continue
        # Optional debug logging removed for performance/clean output

    def _build_repo_name_lookup(self):
        self._repo_name_lookup = {}
        try:
            names = self.repo.cad_names()
        except Exception:
            names = []
        for name in names:
            stripped = (name or "").strip()
            variants = set()
            if stripped:
                variants.add(stripped)
                variants.add(stripped.lower())
                if ":" in stripped:
                    root = stripped.split(":", 1)[0].strip()
                    if root:
                        variants.add(root)
                        variants.add(root.lower())
            for variant in variants:
                if variant and variant not in self._repo_name_lookup:
                    self._repo_name_lookup[variant] = name

    def _init_text_note_types(self):
        self._text_note_types = {}
        self._text_note_types_lower = {}
        self._default_text_note_type = None
        note_types = self._collect_text_note_types()
        if not note_types:
            try:
                default_id = TextNoteType.GetDefault(self.doc)
            except Exception:
                default_id = None
            if default_id:
                try:
                    default_type = self.doc.GetElement(default_id)
                except Exception:
                    default_type = None
                if default_type:
                    note_types = [default_type]
        for note_type in note_types:
            name = self._text_note_type_name(note_type)
            if not name:
                continue
            self._register_text_note_label(name, note_type)
            family = self._text_note_family_label(note_type)
            if family:
                combo = u"{} : {}".format(family, name).strip()
                self._register_text_note_label(combo, note_type)
            if self._default_text_note_type is None:
                self._default_text_note_type = note_type

    def _register_text_note_label(self, label, note_type):
        cleaned = (label or "").strip()
        if not cleaned:
            return
        self._text_note_types[cleaned] = note_type
        lower = cleaned.lower()
        if lower not in self._text_note_types_lower:
            self._text_note_types_lower[lower] = note_type

    def _text_note_family_label(self, note_type):
        if note_type is None:
            return ""
        try:
            family = getattr(note_type, "Family", None)
            fam_name = getattr(family, "Name", None) if family else None
            if fam_name:
                return fam_name
        except Exception:
            fam_name = None
        try:
            fam_name = getattr(note_type, "FamilyName", None)
            if fam_name:
                return fam_name
        except Exception:
            fam_name = None
        if hasattr(note_type, "get_Parameter"):
            for bip in (BuiltInParameter.ALL_MODEL_FAMILY_NAME, BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM):
                if not bip:
                    continue
                try:
                    param = note_type.get_Parameter(bip)
                except Exception:
                    param = None
                if param:
                    try:
                        value = (param.AsString() or "").strip()
                    except Exception:
                        value = ""
                    if value:
                        return value
        return ""

    def _ensure_text_note_type(self, label):
        cleaned = (label or "").strip()
        if not cleaned:
            return None
        base = self._default_text_note_type
        if base is None and self._text_note_types:
            try:
                base = list(self._text_note_types.values())[0]
            except Exception:
                base = None
        if base is None:
            return None
        try:
            new_type = base.Duplicate(cleaned)
        except Exception:
            return None
        self._register_text_note_label(cleaned, new_type)
        family = self._text_note_family_label(new_type)
        if family:
            combo = u"{} : {}".format(family, cleaned).strip()
            self._register_text_note_label(combo, new_type)
        return new_type

    def _text_note_label_variants(self, label):
        variants = []
        cleaned = (label or "").strip()
        if cleaned:
            variants.append(cleaned)
        if ":" in cleaned:
            parts = [part.strip() for part in cleaned.split(":")]
            for part in parts:
                if part:
                    variants.append(part)
        if not variants:
            variants.append("")
        return variants

    def _scan_text_note_types_in_doc(self, variants):
        normalized = [(val or "").strip() for val in variants if val is not None]
        normalized = [val for val in normalized if val]
        lowered = [val.lower() for val in normalized]
        if not normalized:
            return None
        note_types = self._collect_text_note_types()
        for note_type in note_types:
            name = self._text_note_type_name(note_type)
            if not name:
                continue
            if name in normalized or name.lower() in lowered:
                self._register_text_note_label(name, note_type)
                family = self._text_note_family_label(note_type)
                if family:
                    combo = u"{} : {}".format(family, name).strip()
                    self._register_text_note_label(combo, note_type)
                return note_type
        return None

    def _log_text_note_types(self, requested):
        logger = self._get_logger()
        if not logger or not requested:
            return
        available = set(self._text_note_types.keys())
        for note_type in self._collect_text_note_types():
            name = self._text_note_type_name(note_type)
            if name:
                available.add(name)
        if not available:
            logger.info("[Place Elements] No text note styles available while looking for '%s'.", requested)
            return
        sample = sorted(available)
        preview = sample[:10]
        suffix = ""
        if len(sample) > len(preview):
            suffix = " (+{} more)".format(len(sample) - len(preview))
        logger.info("[Place Elements] Available text note styles: %s%s", ", ".join(preview), suffix)

    def _collect_text_note_types(self):
        collected = []
        seen = set()
        logger = self._get_logger()
        source_counts = []

        def _add(elem):
            if elem is None:
                return
            try:
                key = elem.Id.IntegerValue
            except Exception:
                key = None
            if key in seen or key is None:
                return
            seen.add(key)
            collected.append(elem)

        collectors = []
        try:
            collectors.append(FilteredElementCollector(self.doc).OfClass(TextNoteType))
        except Exception:
            pass
        try:
            collectors.append(
                FilteredElementCollector(self.doc)
                .OfCategory(BuiltInCategory.OST_TextNotes)
                .WhereElementIsElementType()
            )
        except Exception:
            pass
        for collector in collectors:
            try:
                elements = list(collector)
            except Exception:
                elements = []
            source_counts.append(len(elements))
            for elem in elements:
                _add(elem)

        try:
            note_instances = list(FilteredElementCollector(self.doc).OfClass(TextNote))
        except Exception:
            note_instances = []
        source_counts.append(len(note_instances))
        for inst in note_instances:
            try:
                type_id = inst.GetTypeId()
            except Exception:
                type_id = None
            if not type_id:
                continue
            try:
                type_elem = self.doc.GetElement(type_id)
            except Exception:
                type_elem = None
            _add(type_elem)

        if logger:
            try:
                logger.info(
                    "[Place Elements] Text note type sources: OfClass=%s, OfCategory=%s, InstanceTypes=%s, total unique=%s",
                    source_counts[0] if len(source_counts) > 0 else 0,
                    source_counts[1] if len(source_counts) > 1 else 0,
                    source_counts[2] if len(source_counts) > 2 else 0,
                    len(collected),
                )
            except Exception:
                pass

        return collected

    def _element_location_point(self, elem):
        if elem is None:
            return None
        loc = getattr(elem, "Location", None)
        if loc is None:
            return None
        try:
            if hasattr(loc, "Point") and loc.Point is not None:
                return loc.Point
        except Exception:
            pass
        try:
            if hasattr(loc, "Curve") and loc.Curve is not None:
                return loc.Curve.Evaluate(0.5, True)
        except Exception:
            pass
        return None

    def _apply_text_note_leaders(self, text_note, leaders, host_loc):
        if not leaders or text_note is None or host_loc is None:
            return
        try:
            existing = list(getattr(text_note, "GetLeaders", lambda: [])() or [])
        except Exception:
            existing = []
        for leader in existing or []:
            try:
                text_note.RemoveLeader(leader)
            except Exception:
                continue
        logger = self._get_logger()
        note_point = self._element_location_point(text_note)
        host_end_offsets = {"x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0}
        host_elbow_offsets = None
        if note_point is not None and host_loc is not None:
            host_elbow_offsets = {
                "x_inches": self._feet_to_inches((note_point.X - host_loc.X) * 0.5),
                "y_inches": self._feet_to_inches((note_point.Y - host_loc.Y) * 0.5),
                "z_inches": self._feet_to_inches((note_point.Z - host_loc.Z) * 0.5),
            }
        for leader_data in leaders:
            leader_type = self._leader_type_from_string(leader_data.get("type"))
            try:
                new_leader = text_note.AddLeader(leader_type)
            except Exception:
                continue
            aimed = self._aim_leader_at_host(new_leader, text_note, host_loc)
            if aimed:
                continue
            elif logger:
                try:
                    logger.info(
                        "[Place Elements] Text note %s leader using stored offsets; host aim unavailable.",
                        getattr(text_note, "Id", "<unknown>"),
                    )
                except Exception:
                    pass
            fallback_data = dict(leader_data)
            if host_loc is not None:
                fallback_data["end"] = dict(host_end_offsets)
                if host_elbow_offsets:
                    fallback_data["elbow"] = dict(host_elbow_offsets)
            end_point = self._offset_dict_to_point(fallback_data.get("end"), host_loc)
            if end_point is not None:
                try:
                    new_leader.SetEndPosition(end_point)
                except Exception:
                    pass
            elbow_point = self._offset_dict_to_point(fallback_data.get("elbow"), host_loc)
            if elbow_point is not None:
                try:
                    new_leader.SetElbowPosition(elbow_point)
                except Exception:
                    pass

    def _aim_leader_at_host(self, leader, text_note, host_loc):
        if leader is None or host_loc is None:
            return False
        logger = self._get_logger()
        note_point = getattr(text_note, "Coord", None)
        if note_point is None:
            note_point = self._element_location_point(text_note)
        target = XYZ(host_loc.X, host_loc.Y, note_point.Z) if note_point is not None else host_loc
        if not self._set_leader_point(leader, target, primary="EndPosition"):
            if logger:
                try:
                    logger.info(
                        "[Place Elements] Failed to aim leader to host for text note %s (no setter for end position).",
                        getattr(text_note, "Id", "<unknown>"),
                    )
                except Exception:
                    pass
            return False
        if note_point is not None:
            elbow = XYZ(
                (target.X + note_point.X) * 0.5,
                (target.Y + note_point.Y) * 0.5,
                note_point.Z,
            )
            if not self._set_leader_point(leader, elbow, primary="ElbowPosition"):
                if logger:
                    try:
                        logger.info(
                            "[Place Elements] Failed to set leader elbow for text note %s (no setter for elbow).",
                            getattr(text_note, "Id", "<unknown>"),
                        )
                    except Exception:
                        pass
        return True

    def _set_leader_point(self, leader, point, primary):
        if leader is None or point is None:
            return False
        candidates = []
        if primary:
            candidates.append(primary)
        # Known alternate names
        mapping = {
            "EndPosition": ["End", "LeaderEnd"],
            "ElbowPosition": ["Elbow", "LeaderElbow"],
        }
        candidates.extend(mapping.get(primary, []))
        for name in list(candidates):
            method = getattr(leader, "Set{}".format(name), None)
            if callable(method):
                try:
                    method(point)
                    return True
                except Exception:
                    continue
            attr = getattr(leader, name, None)
            if attr is None:
                continue
            try:
                setattr(leader, name, point)
                return True
            except Exception:
                continue
        return False

    def _leader_type_from_string(self, value):
        target = "StraightLeader"
        if value:
            lookup = str(value).strip().lower()
            if "arc" in lookup:
                target = "ArcLeader"
            elif "free" in lookup:
                target = "FreeLeader"
            elif "text" in lookup:
                target = "TextNoteLeader"
        return self._resolve_leader_enum(target)

    def _resolve_leader_enum(self, name):
        candidates = []
        if name:
            candidates.append(name)
        candidates.append("StraightLeader")
        tried = set()
        for candidate in candidates:
            if not candidate or candidate in tried:
                continue
            tried.add(candidate)
            # direct attribute lookup
            try:
                enum_value = getattr(TextNoteLeaderTypes, candidate)
                if enum_value is not None:
                    return enum_value
            except Exception:
                pass
            # Enum.Parse fallback for builds where attributes are not exposed
            try:
                parsed = Enum.Parse(TextNoteLeaderTypes, candidate, True)
                if parsed is not None:
                    return parsed
            except Exception:
                pass
        try:
            values = list(Enum.GetValues(TextNoteLeaderTypes))
            if values:
                return values[0]
        except Exception:
            pass
        return None

    def _offset_dict_to_point(self, data, origin):
        if not data or origin is None:
            return None
        try:
            x = self._inch_to_ft(data.get("x_inches", 0.0) or 0.0)
            y = self._inch_to_ft(data.get("y_inches", 0.0) or 0.0)
            z = self._inch_to_ft(data.get("z_inches", 0.0) or 0.0)
        except Exception:
            return None
        return XYZ(origin.X + x, origin.Y + y, origin.Z + z)

    def _text_note_type_name(self, note_type):
        if note_type is None:
            return ""
        try:
            name = (note_type.Name or "").strip()
        except Exception:
            name = ""
        if name:
            return name
        if hasattr(note_type, "get_Parameter"):
            for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM):
                if not bip:
                    continue
                try:
                    param = note_type.get_Parameter(bip)
                except Exception:
                    param = None
                if param:
                    try:
                        value = (param.AsString() or "").strip()
                    except Exception:
                        value = ""
                    if value:
                        return value
        return ""

    def _resolve_row_level(self, row):
        if not isinstance(row, dict):
            return None
        level_value = row.get("LevelId")
        if level_value in (None, ""):
            return None
        if isinstance(level_value, Level):
            return level_value
        if hasattr(level_value, "IntegerValue"):
            try:
                element = self.doc.GetElement(level_value)
            except Exception:
                element = None
            if isinstance(element, Level):
                return element
        try:
            level_id = int(level_value)
        except Exception:
            return None
        try:
            element = self.doc.GetElement(ElementId(level_id))
        except Exception:
            element = None
        if isinstance(element, Level):
            return element
        return None

    def place_from_csv(self, csv_rows, cad_selection_map):
        """
        csv_rows: list of CAD CSV rows
        cad_selection_map: { cad_name: [element_def_id, ...] }
        """
        if not csv_rows:
            return {"placed": 0, "total_rows": 0, "rows_with_coords": 0, "rows_with_mapping": 0}

        level = self.default_level
        if level is None:
            level = FilteredElementCollector(self.doc).OfClass(Level).FirstElement()
            if level is None:
                raise Exception("No Level found in this document; cannot place elements.")
        self.default_level = level
        fallback_level = level

        occurrence_counter = {}
        total_rows = 0
        rows_with_mapping = 0
        rows_with_coords = 0
        placed_count = 0
        self._group_fail_notfound = []
        self._group_fail_error = []
        self._group_fail_error_msgs = []

        t = Transaction(self.doc, self.transaction_name)
        t.Start()

        try:
            for row in csv_rows:
                total_rows += 1
                cad_name = (row.get("Name") or "").strip()
                if not cad_name:
                    continue
                labels, repo_key = self._resolve_selection_map(cad_selection_map, cad_name)
                if not labels:
                    continue
                rows_with_mapping += 1
                if isinstance(labels, basestring):
                    labels = [labels]

                x_raw = (row.get("Position X") or "").strip()
                y_raw = (row.get("Position Y") or "").strip()
                z_raw = (row.get("Position Z") or "").strip()
                if not x_raw or not y_raw or not z_raw:
                    continue

                x_inches = feet_inch_to_inches(x_raw)
                y_inches = feet_inch_to_inches(y_raw)
                z_inches = feet_inch_to_inches(z_raw)

                def _to_inches(raw, parsed):
                    if parsed is not None:
                        return parsed
                    try:
                        return float(raw) * 12.0
                    except Exception:
                        return None

                x_inches = _to_inches(x_raw, x_inches)
                y_inches = _to_inches(y_raw, y_inches)
                z_inches = _to_inches(z_raw, z_inches)
                if x_inches is None or y_inches is None or z_inches is None:
                    continue
                rows_with_coords += 1

                base_loc = XYZ(x_inches / 12.0, y_inches / 12.0, z_inches / 12.0)
                try:
                    base_rot_deg = float(row.get("Rotation", 0.0))
                except Exception:
                    base_rot_deg = 0.0
                parent_element_id = None
                parent_raw = row.get("Parent ElementId")
                if parent_raw not in (None, ""):
                    try:
                        parent_element_id = int(parent_raw)
                    except Exception:
                        try:
                            parent_element_id = int(float(parent_raw))
                        except Exception:
                            parent_element_id = None

                row_level = self._resolve_row_level(row)
                if row_level is not None:
                    self.default_level = row_level
                else:
                    self.default_level = fallback_level

                canonical_name = repo_key or cad_name
                for label in labels:
                    key = (canonical_name, label)
                    occ_index = occurrence_counter.get(key, 0)
                    occurrence_counter[key] = occ_index + 1
                    linked_def = self.repo.definition_for_label(canonical_name, label)
                    if not linked_def:
                        continue
                    placed = self._place_one(
                        linked_def,
                        base_loc,
                        base_rot_deg,
                        occ_index,
                        parent_element_id,
                        canonical_name,
                    )
                    if placed:
                        placed_count += 1
                        placement = linked_def.get_placement()
                        offsets_ft = (0.0, 0.0, 0.0)
                        rot_off = 0.0
                        placement_mode = None
                        tags = []
                        if placement:
                            off_xyz = placement.get_offset_xyz()
                            if off_xyz:
                                offsets_ft = off_xyz
                            rot_off = placement.get_rotation_degrees() or 0.0
                            placement_mode = placement.get_placement_mode()
                            if hasattr(placement, "get_tags"):
                                try:
                                    tags = placement.get_tags()
                                except Exception:
                                    tags = []
                        # identification logging removed
                    else:
                        try:
                            from pyrevit import script

                            logger = script.get_logger()
                        except Exception:
                            logger = None
                        if logger:
                            logger.warning(
                                "[Place Linked Elements] Skipped '%s' for '%s' because the matching family/type is not loaded in this model.",
                                label,
                                canonical_name,
                            )
            self.default_level = fallback_level
            t.Commit()
        except Exception:
            self.default_level = fallback_level
            t.RollBack()
            raise

        # Persist identification log sorted by equipment_id
        return {
            "placed": placed_count,
            "total_rows": total_rows,
            "rows_with_coords": rows_with_coords,
            "rows_with_mapping": rows_with_mapping,
            "group_notfound": len(self._group_fail_notfound),
            "group_errors": len(self._group_fail_error),
            "group_missing_labels": list(self._group_fail_notfound),
            "group_error_labels": list(self._group_fail_error),
        }

    def _resolve_selection_map(self, selection_map, lookup_name):
        labels = selection_map.get(lookup_name)
        if labels:
            canonical = self._canonical_repo_name(lookup_name)
            return labels, canonical
        normalized = lookup_name.strip()
        normalized_lower = normalized.lower()
        normalized_root = normalized_lower.split(":", 1)[0].strip()
        for key, value in selection_map.items():
            if not isinstance(key, basestring):
                continue
            key_stripped = key.strip()
            if key_stripped == normalized or key_stripped.lower() == normalized_lower:
                canonical = self._canonical_repo_name(key)
                return value, canonical
            key_root = key_stripped.lower().split(":", 1)[0].strip()
            if key_root and key_root == normalized_root:
                canonical = self._canonical_repo_name(key)
                return value, canonical
        canonical = self._canonical_repo_name(lookup_name)
        if canonical and canonical in selection_map:
            return selection_map[canonical], canonical
        return None, None

    def _canonical_repo_name(self, value):
        if not isinstance(value, basestring):
            return value
        stripped = value.strip()
        if not stripped:
            return value
        lowered = stripped.lower()
        root = lowered.split(":", 1)[0].strip()
        variants = [stripped, lowered]
        if ":" in lowered:
            variants.append(lowered.split(":", 1)[0].strip())
        if root and ":" in root:
            variants.append(root.split(":", 1)[0].strip())
        for key in variants:
            if not key:
                continue
            canonical = self._repo_name_lookup.get(key)
            if canonical:
                return canonical
        return stripped

    def _place_one(self, linked_def, base_loc, base_rot_deg, occurrence_index, parent_element_id=None, host_name=None):
        placement = linked_def.get_placement()
        offset_xyz = placement.get_offset_xyz() if placement else None
        offset = offset_xyz or (0.0, 0.0, 0.0)
        rot_offset = placement.get_rotation_degrees() if placement else 0.0
        placement_mode = placement.get_placement_mode() if placement else None
        is_group = bool(placement_mode and str(placement_mode).lower() == "group")

        offset_rotation = base_rot_deg
        if offset_rotation:
            ang = math.radians(offset_rotation)
            cos_a = math.cos(ang)
            sin_a = math.sin(ang)
            ox, oy = offset[0], offset[1]
            offset = (
                ox * cos_a - oy * sin_a,
                ox * sin_a + oy * cos_a,
                offset[2],
            )

        loc = XYZ(
            base_loc.X + offset[0],
            base_loc.Y + offset[1],
            base_loc.Z + offset[2],
        )
        if loc.Z < 0.0:
            loc = XYZ(loc.X, loc.Y, 1.0)

        parent_element = None
        if parent_element_id not in (None, ""):
            try:
                parent_element = self.doc.GetElement(ElementId(int(parent_element_id)))
            except Exception:
                parent_element = None
        if parent_element_id not in (None, "") and parent_element is None:
            return False

        label = linked_def.get_element_def_id()
        family = linked_def.get_family()
        type_name = linked_def.get_type()
        final_rot_deg = base_rot_deg + (rot_offset or 0.0)

        gtype, detail_type = (None, None)
        if not is_group:
            gtype, detail_type = self._find_group_type(label, family, type_name)
            if gtype:
                is_group = True

        tags = placement.get_tags() if placement else []
        text_notes = placement.get_text_notes() if placement else []

        instance = None
        if is_group:
            instance = self._place_group(label, family, type_name, linked_def, loc, gtype, detail_type)
        if not instance:
            instance = self._place_symbol(label, family, type_name, linked_def, loc, offset[2], parent_element)
        if instance:
            if self.apply_recorded_level:
                self._apply_recorded_level(instance, linked_def)
            if abs(final_rot_deg) > 1e-6:
                self._rotate_instance(instance, loc, final_rot_deg)
            self._update_element_linker_parameter(
                instance,
                linked_def,
                loc,
                final_rot_deg,
                base_rot_deg,
                parent_element_id,
                host_name,
                base_loc,
            )
            if self.allow_tags:
                self._place_tags(tags, instance, loc, final_rot_deg)
            if self.allow_text_notes and text_notes:
                self._place_text_notes(text_notes, loc, final_rot_deg, host_instance=instance, host_location=loc)
            return True
        return False

    def _find_group_type(self, label, family_name, type_name):
        """Case-insensitive matching of possible keys to a model group type (and attached detail)."""
        candidates = []
        if label:
            candidates.append(label)
            parts = label.split(":")
            left = parts[0].strip() if parts else ""
            right = parts[-1].strip() if parts else ""
            for seg in (left, right):
                base = seg.split("#")[0].strip()
                if base:
                    candidates.append(base)
            no_colon = label.replace(":", " ").strip()
            if no_colon:
                candidates.append(no_colon)
        if type_name:
            candidates.append(type_name)
            clean_type = self._clean_type_name(type_name)
            if clean_type:
                candidates.append(clean_type)
        if family_name:
            candidates.append(family_name)

        extended = []
        for c in candidates:
            if not c:
                continue
            extended.append(c)
            extended.append(u"None : {0}".format(c))
        candidates = extended

        for cand in candidates:
            key = (cand or "").strip()
            if not key:
                continue
            key_lower = key.lower()
            entry = self.group_label_map.get(key)
            if entry:
                return entry
            if key_lower != key:
                entry = self.group_label_map.get(key_lower)
                if entry:
                    return entry
            none_prefix = u"None : {}".format(key)
            entry = self.group_label_map.get(none_prefix)
            if entry:
                return entry
            none_prefix_lower = none_prefix.lower()
            entry = self.group_label_map.get(none_prefix_lower)
            if entry:
                return entry
        return (None, None)

    def _clean_type_name(self, value):
        if value is None:
            return ""
        v = value.lower().strip()
        prefixes = ["floor plan:", "plan:", "fp:"]
        for p in prefixes:
            if v.startswith(p):
                v = v[len(p):].strip()
                break
        v = v.split("#")[0].strip()
        return v

    def _place_group(self, label, family_name, type_name, linked_def, location, gtype=None, detail_type=None):
        if gtype is None:
            gtype, detail_type = self._find_group_type(label, family_name, type_name)
        if not gtype:
            self._group_fail_notfound.append(label)
            return None

        try:
            instance = self.doc.Create.PlaceGroup(location, gtype)

            try:
                view = getattr(self.doc, "ActiveView", None)
            except Exception:
                view = None
            if view:
                detail_ids = []
                try:
                    detail_ids = list(instance.GetAvailableAttachedDetailGroupTypeIds() or [])
                except Exception:
                    detail_ids = []
                if detail_type is not None:
                    try:
                        detail_ids.append(detail_type.Id)
                    except Exception:
                        pass
                seen = set()
                for did in detail_ids or []:
                    try:
                        if did is None or did.IntegerValue in seen:
                            continue
                        instance.ShowAttachedDetailGroups(view, did)
                        seen.add(did.IntegerValue)
                    except Exception:
                        continue

            return instance
        except Exception as exc:
            self._group_fail_error.append(label)
            try:
                self._group_fail_error_msgs.append(u"{}: {}".format(label, exc))
            except Exception:
                pass
            return None

    def _activate_symbol(self, symbol):
        if symbol is None:
            return False
        if symbol.Id in self._activated_symbols:
            return True
        try:
            if not symbol.IsActive:
                symbol.Activate()
                self.doc.Regenerate()
            self._activated_symbols.add(symbol.Id)
            return True
        except Exception:
            return False

    def _place_symbol(self, label, family_name, type_name, linked_def, location, z_offset_feet, parent_element=None):
        symbol = self.symbol_label_map.get(label)
        if not symbol and family_name and type_name:
            key = u"{} : {}".format(family_name, type_name)
            symbol = self.symbol_label_map.get(key)
        if not symbol:
            return None

        if not self._activate_symbol(symbol):
            return None

        cat = getattr(symbol, "Category", None)
        is_generic_annotation = cat and getattr(cat, "Name", None) == "Generic Annotations"
        level = self.default_level
        if level is None:
            level = FilteredElementCollector(self.doc).OfClass(Level).FirstElement()
            self.default_level = level
        if level is None:
            return None

        try:
            if is_generic_annotation:
                instance = self.doc.Create.NewFamilyInstance(location, symbol, self.doc.ActiveView)
            else:
                instance = self.doc.Create.NewFamilyInstance(
                    location,
                    symbol,
                    level,
                    Structure.StructuralType.NonStructural,
                )
                if abs(z_offset_feet) > 1e-6:
                    elev_param = instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                    if elev_param and not elev_param.IsReadOnly:
                        elev_param.Set(z_offset_feet)

            self._apply_parameters(instance, linked_def.get_static_params(), parent_element)
            return instance
        except Exception:
            return None

    def _rotate_instance(self, instance, location, angle_deg):
        try:
            angle_rad = math.radians(angle_deg)
            axis = Line.CreateBound(location, location + XYZ(0, 0, 1))
            ElementTransformUtils.RotateElement(self.doc, instance.Id, axis, angle_rad)
        except Exception:
            pass

    def _tagged_element_ids(self, tag):
        if tag is None:
            return []
        try:
            getter = getattr(tag, "GetTaggedLocalElementIds", None)
            if callable(getter):
                return list(getter() or [])
        except Exception:
            pass
        ids = []
        for attr in ("TaggedLocalElementId", "TaggedElementId"):
            try:
                value = getattr(tag, attr, None)
            except Exception:
                value = None
            if value:
                ids.append(value)
        return ids

    def _existing_tag_for_host(self, host_instance, symbol_id, view_obj):
        if not host_instance or not symbol_id or not view_obj:
            return False
        try:
            tags = FilteredElementCollector(self.doc, view_obj.Id).OfClass(IndependentTag)
        except Exception:
            tags = []
        if not tags:
            return False
        try:
            host_id_val = host_instance.Id.IntegerValue
        except Exception:
            host_id_val = None
        for tag in tags:
            try:
                if tag.GetTypeId() != symbol_id:
                    continue
            except Exception:
                continue
            if host_id_val is None:
                continue
            for tagged_id in self._tagged_element_ids(tag):
                try:
                    tagged_val = tagged_id.IntegerValue
                except Exception:
                    try:
                        tagged_val = int(tagged_id)
                    except Exception:
                        tagged_val = None
                if tagged_val is not None and tagged_val == host_id_val:
                    return True
        return False

    def _place_tags(self, tag_defs, host_instance, base_loc, final_rot_deg):
        if not tag_defs:
            return
        logger = self._get_logger()
        active_view = getattr(self.doc, "ActiveView", None)
        for tag in tag_defs:
            family = tag.get("family")
            type_name = tag.get("type")
            category_name = tag.get("category") or ""
            def _normalize_keynote_family(value):
                if not value:
                    return ""
                text = str(value)
                if ":" in text:
                    text = text.split(":", 1)[0]
                return "".join([ch for ch in text.lower() if ch.isalnum()])
            def _is_ga_keynote_symbol(name):
                return _normalize_keynote_family(name) == "gakeynotesymbolced"
            family_text = (family or "").lower()
            type_text = (type_name or "").lower()
            is_ga_keynote = _is_ga_keynote_symbol(family)
            is_keynote_def = "keynote" in (category_name or "").lower() or "keynote" in family_text or "keynote" in type_text
            if is_ga_keynote:
                is_keynote_def = False
            label = None
            if family and type_name:
                label = u"{} : {}".format(family, type_name)
            elif type_name:
                label = type_name
            elif family:
                label = family
            else:
                continue
            symbol = None
            if category_name:
                key = (label.strip().lower(), category_name.strip().lower())
                symbol = self.symbol_label_map_by_category.get(key)
            if not symbol:
                symbol = self.symbol_label_map.get(label)
            if not symbol and family and type_name:
                key = u"{} : {}".format(family, type_name)
                symbol = self.symbol_label_map.get(key)
            if not symbol:
                if logger and is_keynote_def:
                    logger.info(
                        "[Place Elements] Keynote symbol not found: label='%s' family='%s' type='%s' category='%s'.",
                        label,
                        family or "",
                        type_name or "",
                        category_name or "",
                    )
                continue
            if not self._activate_symbol(symbol):
                if logger and is_keynote_def:
                    logger.info(
                        "[Place Elements] Keynote symbol activation failed: label='%s'.",
                        label,
                    )
                continue

            offsets = tag.get("offset") or (0.0, 0.0, 0.0)
            if is_ga_keynote:
                try:
                    ang = math.radians(final_rot_deg or 0.0)
                    cos_a = math.cos(ang)
                    sin_a = math.sin(ang)
                    ox, oy, oz = offsets[0] or 0.0, offsets[1] or 0.0, offsets[2] or 0.0
                    offsets = (ox * cos_a - oy * sin_a, ox * sin_a + oy * cos_a, oz)
                except Exception:
                    pass
            if self.max_tag_distance_feet not in (None, "") and not is_ga_keynote:
                try:
                    limit = float(self.max_tag_distance_feet)
                except Exception:
                    limit = None
                if limit:
                    try:
                        dist = math.sqrt(
                            float(offsets[0] or 0.0) ** 2
                            + float(offsets[1] or 0.0) ** 2
                            + float(offsets[2] or 0.0) ** 2
                        )
                    except Exception:
                        dist = None
                    if dist is not None and dist > limit:
                        if logger:
                            logger.info(
                                "[Place Elements] Skipping tag '%s' offset distance %.2f ft > %.2f ft.",
                                label,
                                dist,
                                limit,
                            )
                        continue
            leader_elbow = self._convert_offset_to_tuple(tag.get("leader_elbow"))
            leader_end = self._convert_offset_to_tuple(tag.get("leader_end"))
            category_name = (category_name or "").lower()
            parameters = dict(tag.get("parameters") or {})
            keynote_value = None
            keynote_source = ""
            keynote_text = ""
            if is_keynote_def:
                key_value = tag.get("key_value")
                if key_value not in (None, ""):
                    has_key = False
                    for key in parameters.keys():
                        if (key or "").strip().lower() in ("keynote value", "key value", "keynote"):
                            has_key = True
                            break
                    if not has_key:
                        parameters["Keynote Value"] = key_value
                keynote_value = self._extract_keynote_value(parameters)
                keynote_source = self._extract_keynote_source(parameters)
                keynote_text = self._extract_keynote_text(parameters)

            sym_cat = getattr(symbol, "Category", None)
            fam_cat = None
            try:
                fam = getattr(symbol, "Family", None)
                fam_cat = getattr(fam, "FamilyCategory", None)
            except Exception:
                fam_cat = None
            sym_cat_name = ((sym_cat.Name or "") if sym_cat else "").lower()
            fam_cat_name = ((fam_cat.Name or "") if fam_cat else "").lower()
            combined_cat = " ".join([category_name, sym_cat_name, fam_cat_name])
            is_tag_family = "tag" in combined_cat
            is_annotation_family = ("annotation" in combined_cat) and not is_tag_family
            if is_ga_keynote:
                is_tag_family = False
                is_annotation_family = True
            keynote_by_element = bool(keynote_source and "element" in keynote_source)
            keynote_host_applied = False
            if is_keynote_def and is_tag_family and host_instance and keynote_value not in (None, "") and keynote_by_element:
                keynote_host_applied = self._apply_keynote_value_to_host(host_instance, keynote_value, logger)
                if not keynote_host_applied and logger:
                    logger.info(
                        "[Place Elements] Host has no writable keynote param; placing keynote as by-keynote tag for '%s'.",
                        label,
                    )

            key = tag_key_from_dict(tag)
            target_views = []
            if key and self.tag_view_map:
                view_ids = self.tag_view_map.get(key) or []
                for vid in view_ids:
                    try:
                        view_obj = self.doc.GetElement(ElementId(int(vid)))
                    except Exception:
                        view_obj = None
                    if view_obj:
                        target_views.append(view_obj)
            if not target_views and (is_tag_family or is_annotation_family):
                if active_view:
                    target_views.append(active_view)

            if not (is_tag_family or is_annotation_family):
                target_views = [None]
            elif not target_views:
                continue

            for view_obj in target_views:
                tag_loc = XYZ(
                    base_loc.X + (offsets[0] or 0.0),
                    base_loc.Y + (offsets[1] or 0.0),
                    base_loc.Z + (offsets[2] or 0.0),
                )
                tag_rotation = final_rot_deg + float(tag.get("rotation_deg", 0.0) or 0.0)
                instance = None
                try:
                    if is_tag_family:
                        if not view_obj or not host_instance:
                            continue
                        if self._existing_tag_for_host(host_instance, symbol.Id, view_obj):
                            if logger and is_keynote_def:
                                logger.info(
                                    "[Place Elements] Keynote skipped: tag already exists for host (%s).",
                                    label,
                                )
                            continue
                        try:
                            reference = Reference(host_instance)
                        except Exception:
                            continue
                        tag_mode = TagMode.TM_ADDBY_CATEGORY
                        if is_keynote_def:
                            tag_mode = getattr(TagMode, "TM_ADDBY_KEYNOTE", TagMode.TM_ADDBY_CATEGORY)
                        created_with_type_id = False
                        create_with_default = False
                        if symbol is not None:
                            try:
                                overload = IndependentTag.Create.Overloads[
                                    Document, ElementId, Reference, bool, TagMode, TagOrientation, XYZ, ElementId
                                ]
                                independent = overload(
                                    self.doc,
                                    view_obj.Id,
                                    reference,
                                    True,
                                    tag_mode,
                                    TagOrientation.Horizontal,
                                    tag_loc,
                                    symbol.Id,
                                )
                                created_with_type_id = True
                            except Exception:
                                independent = None
                        if independent is None and is_keynote_def and symbol is not None:
                            try:
                                keynote_cat_id = ElementId(BuiltInCategory.OST_KeynoteTags)
                                get_default = getattr(self.doc, "GetDefaultFamilyTypeId", None)
                                set_default = getattr(self.doc, "SetDefaultFamilyTypeId", None)
                                default_type_id = get_default(keynote_cat_id) if callable(get_default) else None
                                if callable(set_default):
                                    set_default(keynote_cat_id, symbol.Id)
                                    create_with_default = True
                                independent = IndependentTag.Create(
                                    self.doc,
                                    view_obj.Id,
                                    reference,
                                    True,
                                    tag_mode,
                                    TagOrientation.Horizontal,
                                    tag_loc,
                                )
                            except Exception as exc:
                                if logger and is_keynote_def:
                                    logger.info(
                                        "[Place Elements] Keynote default-type create failed for '%s': %s",
                                        label,
                                        exc,
                                    )
                                independent = None
                            finally:
                                if create_with_default and callable(set_default):
                                    try:
                                        if default_type_id:
                                            set_default(keynote_cat_id, default_type_id)
                                    except Exception:
                                        pass
                        if independent is None:
                            independent = IndependentTag.Create(
                                self.doc,
                                view_obj.Id,
                                reference,
                                True,
                                tag_mode,
                                TagOrientation.Horizontal,
                                tag_loc,
                            )
                        if not independent:
                            continue
                        if not created_with_type_id and not create_with_default:
                            try:
                                independent.ChangeTypeId(symbol.Id)
                            except Exception as exc:
                                if logger and is_keynote_def:
                                    sym_cat = getattr(symbol, "Category", None)
                                    sym_cat_name = getattr(sym_cat, "Name", None) if sym_cat else ""
                                    logger.info(
                                        "[Place Elements] Keynote ChangeTypeId failed for '%s': %s (symbol category=%s)",
                                        label,
                                        exc,
                                        sym_cat_name or "",
                                    )
                                    try:
                                        self.doc.Delete(independent.Id)
                                        continue
                                    except Exception:
                                        pass
                        instance = independent
                    elif is_annotation_family:
                        if not view_obj or (hasattr(view_obj, "ViewType") and view_obj.ViewType == ViewType.ThreeD):
                            continue
                        instance = self.doc.Create.NewFamilyInstance(tag_loc, symbol, view_obj)
                    else:
                        level = self.default_level
                        if level is None:
                            level = FilteredElementCollector(self.doc).OfClass(Level).FirstElement()
                            self.default_level = level
                        if level is None:
                            continue
                        instance = self.doc.Create.NewFamilyInstance(
                            tag_loc,
                            symbol,
                            level,
                            Structure.StructuralType.NonStructural,
                        )
                except Exception as exc:
                    if logger and is_keynote_def:
                        logger.info(
                            "[Place Elements] Keynote placement failed for '%s': %s",
                            label,
                            exc,
                        )
                    instance = None
                if not instance:
                    continue
                if isinstance(instance, IndependentTag):
                    try:
                        instance.TagHeadPosition = tag_loc
                    except Exception:
                        pass
                    if leader_end:
                        try:
                            instance.LeaderEnd = XYZ(
                                base_loc.X + (leader_end[0] or 0.0),
                                base_loc.Y + (leader_end[1] or 0.0),
                                base_loc.Z + (leader_end[2] or 0.0),
                            )
                        except Exception:
                            pass
                    if leader_elbow:
                        try:
                            instance.LeaderElbow = XYZ(
                                base_loc.X + (leader_elbow[0] or 0.0),
                                base_loc.Y + (leader_elbow[1] or 0.0),
                                base_loc.Z + (leader_elbow[2] or 0.0),
                            )
                        except Exception:
                            pass
                    if is_keynote_def and keynote_value not in (None, ""):
                        self._apply_keynote_value_to_tag(instance, keynote_value, keynote_text, logger)
                else:
                    if abs(tag_rotation) > 1e-6:
                        try:
                            axis = Line.CreateBound(tag_loc, tag_loc + XYZ(0, 0, 1))
                            ElementTransformUtils.RotateElement(self.doc, instance.Id, axis, math.radians(tag_rotation))
                        except Exception:
                            pass
                self._apply_parameters(instance, parameters)

    def _resolve_text_note_type(self, type_name):
        if not self._text_note_types:
            self._init_text_note_types()
            if not self._text_note_types:
                return None
        variants = self._text_note_label_variants(type_name)
        for variant in variants:
            exact = self._text_note_types.get(variant)
            if exact:
                return exact
            lower = variant.lower()
            scoped = self._text_note_types_lower.get(lower)
            if scoped:
                return scoped
        match = self._scan_text_note_types_in_doc(variants)
        if match:
            return match
        primary = variants[0] if variants else type_name
        self._log_text_note_types(primary)
        created = self._ensure_text_note_type(primary)
        if created:
            return created
        return self._default_text_note_type

    def _convert_offset_to_tuple(self, offsets):
        if offsets is None:
            return None
        if isinstance(offsets, dict):
            return (
                self._coerce_length(offsets.get("x_inches"), offsets.get("x")),
                self._coerce_length(offsets.get("y_inches"), offsets.get("y")),
                self._coerce_length(offsets.get("z_inches"), offsets.get("z")),
            )
        if isinstance(offsets, (list, tuple)):
            values = list(offsets) + [0.0, 0.0, 0.0]
            return (
                self._coerce_float(values[0]),
                self._coerce_float(values[1]),
                self._coerce_float(values[2]),
            )
        try:
            if hasattr(offsets, "x_inches"):
                x_val = self._inch_to_ft(getattr(offsets, "x_inches", 0.0))
            else:
                x_val = self._coerce_float(getattr(offsets, "X", getattr(offsets, "x", 0.0)))
            if hasattr(offsets, "y_inches"):
                y_val = self._inch_to_ft(getattr(offsets, "y_inches", 0.0))
            else:
                y_val = self._coerce_float(getattr(offsets, "Y", getattr(offsets, "y", 0.0)))
            if hasattr(offsets, "z_inches"):
                z_val = self._inch_to_ft(getattr(offsets, "z_inches", 0.0))
            else:
                z_val = self._coerce_float(getattr(offsets, "Z", getattr(offsets, "z", 0.0)))
            return (x_val, y_val, z_val)
        except Exception:
            return None

    def _inch_to_ft(self, value):
        try:
            return float(value) / 12.0
        except Exception:
            return 0.0

    def _coerce_length(self, inches_value, feet_value):
        if inches_value not in (None, ""):
            return self._inch_to_ft(inches_value)
        if feet_value not in (None, ""):
            try:
                return float(feet_value)
            except Exception:
                return 0.0
        return 0.0

    def _coerce_float(self, value):
        try:
            return float(value)
        except Exception:
            return 0.0

    def _place_text_notes(self, text_defs, base_loc, final_rot_deg, host_instance=None, host_location=None):
        if not text_defs:
            return
        active_view = getattr(self.doc, "ActiveView", None)
        if not active_view or active_view.ViewType == ViewType.ThreeD:
            return
        logger = self._get_logger()
        host_point = host_location or (self._element_location_point(host_instance) if host_instance is not None else None)
        if host_point is None:
            host_point = self._element_location_point(host_instance)
        origin = host_point or base_loc
        if origin is None:
            return
        if logger:
            try:
                logger.info(
                    "[Place Elements] Text note origin base=(%0.3f,%0.3f,%0.3f) host=(%s)",
                    base_loc.X if base_loc else 0.0,
                    base_loc.Y if base_loc else 0.0,
                    base_loc.Z if base_loc else 0.0,
                    _format_xyz(host_point) if host_point else "<none>",
                )
            except Exception:
                pass
        for note in text_defs:
            if isinstance(note, dict):
                text_value = (note.get("text") or "").strip()
                offsets, rotation_delta = self._resolve_note_offset_rotation(note)
                note_type_name = note.get("type_name")
                width_val = note.get("width")
                if width_val is None and note.get("width_inches") is not None:
                    width_val = self._inch_to_ft(note.get("width_inches"))
                leader_data = note.get("leaders") or []
            else:
                text_value = getattr(note, "text", "") or ""
                offsets, rotation_delta = self._resolve_note_offset_rotation(note)
                note_type_name = getattr(note, "type_name", None)
                width_val = getattr(note, "width", None)
                leader_data = getattr(note, "leaders", []) or []
            if not text_value:
                continue
            note_type = self._resolve_text_note_type(note_type_name)
            if note_type is None:
                self._log_text_note_types(note_type_name)
                if logger:
                    logger.warning(
                        "[Place Elements] Skipping text note '%s' because type '%s' is not loaded.",
                        text_value,
                        note_type_name or "<unspecified>",
                    )
                continue
            try:
                dx, dy, dz = offsets
            except Exception:
                dx = dy = dz = 0.0
            loc = XYZ(
                origin.X + (dx or 0.0),
                origin.Y + (dy or 0.0),
                origin.Z + (dz or 0.0),
            )
            if logger:
                try:
                    logger.info(
                        "[Place Elements] Text note '%s' offsets=(%0.3f,%0.3f,%0.3f) origin=(%0.3f,%0.3f,%0.3f)",
                        text_value,
                        dx or 0.0,
                        dy or 0.0,
                        dz or 0.0,
                        origin.X,
                        origin.Y,
                        origin.Z,
                    )
                except Exception:
                    pass
            total_rotation = final_rot_deg + rotation_delta
            try:
                created = TextNote.Create(self.doc, active_view.Id, loc, text_value, note_type.Id)
            except Exception as exc:
                if logger:
                    logger.warning(
                        "[Place Elements] Failed to place text note '%s' using type '%s': %s",
                        text_value,
                        note_type_name or getattr(note_type, "Name", None),
                        exc,
                    )
                continue
            if width_val:
                try:
                    created.Width = float(width_val)
                except Exception:
                    pass
            if abs(total_rotation) > 1e-6:
                try:
                    axis = Line.CreateBound(loc, loc + XYZ(0, 0, 1))
                    ElementTransformUtils.RotateElement(self.doc, created.Id, axis, math.radians(total_rotation))
                except Exception:
                    pass
            if logger:
                try:
                    logger.info(
                        "[Place Elements] Placed text note '%s' using type '%s' at (%0.3f,%0.3f,%0.3f).",
                        text_value,
                        note_type_name or getattr(note_type, "Name", None),
                        loc.X,
                        loc.Y,
                        loc.Z,
                    )
                except Exception:
                    pass
            self._apply_text_note_leaders(created, leader_data, host_point or origin)
            self._apply_parameters(created, {})

    def _resolve_note_offset_rotation(self, note):
        raw_offsets = None
        if isinstance(note, dict):
            raw_offsets = note.get("offsets") or note.get("offset")
        else:
            raw_offsets = getattr(note, "offsets", None) or getattr(note, "offset", None)
        offsets = self._convert_offset_to_tuple(raw_offsets)
        if offsets is None:
            offsets = (0.0, 0.0, 0.0)
        rotation = self._extract_note_rotation(note, raw_offsets)
        return offsets, rotation

    def _extract_note_rotation(self, note, offsets_source):
        candidates = []
        if isinstance(note, dict):
            candidates.append(note.get("rotation_deg"))
        else:
            candidates.append(getattr(note, "rotation_deg", None))
        if isinstance(offsets_source, dict):
            candidates.append(offsets_source.get("rotation_deg"))
            candidates.append(offsets_source.get("rotation"))
        elif offsets_source is not None:
            try:
                candidates.append(getattr(offsets_source, "rotation_deg", None))
            except Exception:
                pass
        for value in candidates:
            if value not in (None, ""):
                try:
                    return float(value)
                except Exception:
                    continue
        return 0.0

    def _get_logger(self):
        try:
            from pyrevit import script

            return script.get_logger()
        except Exception:
            return None

    def _extract_parent_parameter_name(self, value):
        if not isinstance(value, basestring):
            return None
        match = PARENT_PARAMETER_PATTERN.match(value)
        if not match:
            return None
        name = (match.group(1) or match.group(2) or "").strip()
        if not name:
            return None
        return name.replace(SAFE_HASH, "#")

    def _get_param_unit_type_id(self, param):
        get_unit = getattr(param, "GetUnitTypeId", None)
        if callable(get_unit):
            try:
                return get_unit()
            except Exception:
                return None
        return None

    def _apply_parent_parameter(self, child_param, parent_element, parent_param_name):
        if not child_param or child_param.IsReadOnly or not parent_param_name:
            return False
        if not parent_element:
            return False
        try:
            parent_param = parent_element.LookupParameter(parent_param_name)
        except Exception:
            parent_param = None
        if not parent_param:
            return False

        from Autodesk.Revit.DB import StorageType, UnitUtils

        try:
            child_storage = child_param.StorageType
        except Exception:
            return False
        try:
            parent_storage = parent_param.StorageType
        except Exception:
            parent_storage = None

        def _parent_string_value():
            try:
                value = parent_param.AsString()
            except Exception:
                value = None
            if value not in (None, ""):
                return value
            try:
                return parent_param.AsValueString()
            except Exception:
                return None

        if child_storage == StorageType.String:
            value = _parent_string_value()
            if value is None:
                value = ""
            try:
                child_param.Set(str(value))
                return True
            except Exception:
                return False

        def _parent_numeric_value():
            if parent_storage == StorageType.Double:
                try:
                    internal = parent_param.AsDouble()
                except Exception:
                    internal = None
                parent_unit_id = self._get_param_unit_type_id(parent_param)
                if parent_unit_id is not None and internal is not None:
                    try:
                        return UnitUtils.ConvertFromInternalUnits(float(internal), parent_unit_id), True
                    except Exception:
                        pass
                raw = _parent_string_value()
                if raw not in (None, ""):
                    try:
                        return float(raw), True
                    except Exception:
                        pass
                if internal is not None:
                    return float(internal), False
                return None, False
            if parent_storage == StorageType.Integer:
                try:
                    return float(parent_param.AsInteger()), True
                except Exception:
                    return None, False
            raw = _parent_string_value()
            if raw not in (None, ""):
                try:
                    return float(raw), True
                except Exception:
                    return None, False
            return None, False

        if child_storage == StorageType.Integer:
            value, is_display = _parent_numeric_value()
            if value is None:
                return False
            child_unit_id = self._get_param_unit_type_id(child_param)
            if child_unit_id is not None and is_display:
                try:
                    value = UnitUtils.ConvertToInternalUnits(float(value), child_unit_id)
                except Exception:
                    pass
            try:
                child_param.Set(int(round(value)))
                return True
            except Exception:
                return False

        if child_storage == StorageType.Double:
            value, is_display = _parent_numeric_value()
            if value is None:
                return False
            child_unit_id = self._get_param_unit_type_id(child_param)
            if child_unit_id is not None and is_display:
                try:
                    value = UnitUtils.ConvertToInternalUnits(float(value), child_unit_id)
                except Exception:
                    pass
            try:
                child_param.Set(float(value))
                return True
            except Exception:
                return False

        if child_storage == StorageType.ElementId:
            if parent_storage == StorageType.ElementId:
                try:
                    child_param.Set(parent_param.AsElementId())
                    return True
                except Exception:
                    return False
            return False

        raw = _parent_string_value()
        if raw is None:
            return False
        try:
            child_param.Set(str(raw))
            return True
        except Exception:
            return False

    def _is_load_classification_param(self, name):
        if not name:
            return False
        return "load classification" in str(name).lower()

    def _load_classification_map(self):
        cache = getattr(self, "_load_classification_cache", None)
        if cache is not None:
            return cache
        mapping = {}
        try:
            from Autodesk.Revit.DB.Electrical import ElectricalLoadClassification
        except Exception:
            ElectricalLoadClassification = None
        if ElectricalLoadClassification is not None:
            try:
                classifications = FilteredElementCollector(self.doc).OfClass(ElectricalLoadClassification)
            except Exception:
                classifications = []
            for classification in classifications:
                name = getattr(classification, "Name", None)
                if not name:
                    continue
                key = " ".join(str(name).strip().lower().split())
                mapping[key] = classification.Id
        self._load_classification_cache = mapping
        return mapping

    def _resolve_load_classification_id(self, value):
        if value is None:
            return None
        if isinstance(value, ElementId):
            return value
        key = " ".join(str(value).strip().lower().split())
        if not key:
            return None
        return self._load_classification_map().get(key)

    def _keynote_entry_map(self):
        cache = getattr(self, "_keynote_entry_cache", None)
        if cache is not None:
            return cache
        mapping = {}
        normalized = {}
        text_map = {}
        entries_list = []
        try:
            from Autodesk.Revit.DB import KeynoteTable
        except Exception:
            KeynoteTable = None
        if KeynoteTable is not None:
            try:
                table = KeynoteTable.GetKeynoteTable(self.doc)
            except Exception:
                table = None
            if table:
                entries = None
                for method_name in ("GetKeynoteEntries", "GetKeynoteEntriesByCategory", "GetKeynoteEntriesByGroup"):
                    if hasattr(table, method_name):
                        try:
                            entries = getattr(table, method_name)()
                        except Exception:
                            entries = None
                        if entries:
                            break
                if entries is None:
                    try:
                        entries = table.KeynoteEntries
                    except Exception:
                        entries = None
                if entries:
                    try:
                        entries_list = list(entries)
                    except Exception:
                        entries_list = []
                    for entry in entries_list:
                        try:
                            key = getattr(entry, "Key", None)
                        except Exception:
                            key = None
                        if not key:
                            continue
                        key_text = str(key).strip()
                        mapping[key_text] = entry.Id
                        norm_key = self._normalize_keynote_key(key_text)
                        if norm_key and norm_key not in normalized:
                            normalized[norm_key] = entry.Id
                        text_value = None
                        for attr in ("Text", "KeynoteText"):
                            try:
                                text_value = getattr(entry, attr, None)
                            except Exception:
                                text_value = None
                            if text_value:
                                break
                        if text_value:
                            norm_text = self._normalize_keynote_text(text_value)
                            if norm_text and norm_text not in text_map:
                                text_map[norm_text] = entry.Id
        self._keynote_entry_cache = mapping
        self._keynote_entry_norm_cache = normalized
        self._keynote_entry_text_cache = text_map
        self._keynote_entry_list = entries_list
        return mapping

    def _normalize_keynote_key(self, value):
        if value is None:
            return ""
        text = re.sub(r"[^0-9A-Za-z]+", "", str(value))
        return text.upper().strip()

    def _normalize_keynote_text(self, value):
        if value is None:
            return ""
        return " ".join(str(value).strip().lower().split())

    def _resolve_keynote_entry_id(self, value, fallback_text=None):
        if value is None:
            return None
        if isinstance(value, ElementId):
            return value
        key = str(value).strip()
        if not key:
            return None
        mapping = self._keynote_entry_map()
        entry_id = mapping.get(key)
        if entry_id:
            return entry_id
        normalized = getattr(self, "_keynote_entry_norm_cache", {}) or {}
        norm_key = self._normalize_keynote_key(key)
        entry_id = normalized.get(norm_key)
        if entry_id:
            return entry_id
        if fallback_text:
            text_map = getattr(self, "_keynote_entry_text_cache", {}) or {}
            norm_text = self._normalize_keynote_text(fallback_text)
            entry_id = text_map.get(norm_text)
            if entry_id:
                return entry_id
        return None

    def _extract_keynote_source(self, params_dict):
        if not params_dict:
            return ""
        for key, value in params_dict.items():
            key_name = (key or "").strip().lower()
            if "key source" in key_name or "keynote source" in key_name:
                return (str(value).strip().lower() if value is not None else "")
        return ""

    def _extract_keynote_text(self, params_dict):
        if not params_dict:
            return ""
        for key in ("Keynote Text", "Keynote Description"):
            value = params_dict.get(key)
            if value not in (None, ""):
                return str(value)
        for key, value in params_dict.items():
            key_name = (key or "").strip().lower()
            if "keynote text" in key_name or "keynote description" in key_name:
                if value not in (None, ""):
                    return str(value)
        return ""

    def _is_keynote_param_name(self, name):
        if not name:
            return False
        lowered = str(name).strip().lower()
        return "keynote" in lowered or lowered == "key value"

    def _extract_keynote_value(self, params_dict):
        if not params_dict:
            return None
        for key in ("Key Value", "Keynote Value", "Keynote"):
            value = params_dict.get(key)
            if value not in (None, ""):
                return value
        for key, value in params_dict.items():
            if (key or "").strip().lower() in ("key value", "keynote value", "keynote"):
                if value not in (None, ""):
                    return value
        return None

    def _apply_keynote_value_to_host(self, host_element, key_value, logger=None):
        if host_element is None or key_value in (None, ""):
            return False
        def _resolve_param(target):
            try:
                param = target.get_Parameter(BuiltInParameter.KEYNOTE_PARAM)
            except Exception:
                param = None
            if not param:
                for name in ("Keynote", "Keynote Value"):
                    try:
                        param = target.LookupParameter(name)
                    except Exception:
                        param = None
                    if param:
                        break
            return param

        param = _resolve_param(host_element)
        if not param or param.IsReadOnly:
            type_elem = None
            try:
                type_id = host_element.GetTypeId()
                type_elem = host_element.Document.GetElement(type_id)
            except Exception:
                type_elem = None
            if type_elem:
                param = _resolve_param(type_elem)
        if not param or param.IsReadOnly:
            return False
        try:
            from Autodesk.Revit.DB import StorageType
        except Exception:
            StorageType = None
        try:
            storage_type = param.StorageType if StorageType else None
        except Exception:
            storage_type = None
        try:
            if storage_type and storage_type == StorageType.ElementId:
                resolved = None
                try:
                    resolved = ElementId(int(key_value))
                except Exception:
                    resolved = self._resolve_keynote_entry_id(key_value)
                if resolved:
                    param.Set(resolved)
                else:
                    param.SetValueString(str(key_value))
            elif storage_type and storage_type == StorageType.Integer:
                param.Set(int(key_value))
            elif storage_type and storage_type == StorageType.Double:
                param.Set(float(key_value))
            else:
                param.Set(str(key_value))
            if logger:
                try:
                    pname = getattr(getattr(param, "Definition", None), "Name", None) or ""
                except Exception:
                    pname = ""
                logger.info(
                    "[Place Elements] Applied keynote value '%s' to host (%s).",
                    key_value,
                    pname or "Keynote",
                )
            return True
        except Exception:
            try:
                param.SetValueString(str(key_value))
                if logger:
                    logger.info(
                        "[Place Elements] Applied keynote value '%s' to host (value string).",
                        key_value,
                    )
                return True
            except Exception:
                return False

    def _apply_keynote_value_to_tag(self, tag_element, key_value, keynote_text=None, logger=None):
        if tag_element is None or key_value in (None, ""):
            return False
        diagnostics = []
        resolved_entry = self._resolve_keynote_entry_id(key_value, fallback_text=keynote_text)
        if resolved_entry is None and logger:
            if keynote_text:
                diagnostics.append(
                    "no keynote entry for '{}' (text='{}')".format(key_value, keynote_text)
                )
            else:
                diagnostics.append("no keynote entry for '{}'".format(key_value))
        param = None
        try:
            param = tag_element.get_Parameter(BuiltInParameter.KEYNOTE_PARAM)
        except Exception:
            param = None
        if param:
            try:
                storage = param.StorageType
                storage_name = storage.ToString() if storage else ""
            except Exception:
                storage_name = ""
            diagnostics.append("KEYNOTE_PARAM storage={}".format(storage_name or "?"))
            if param.IsReadOnly:
                diagnostics.append("KEYNOTE_PARAM readonly")
            else:
                if resolved_entry is not None:
                    try:
                        param.Set(resolved_entry)
                        if logger:
                            logger.info(
                                "[Place Elements] Applied keynote value '%s' to tag (entry id).",
                                key_value,
                            )
                        return True
                    except Exception as exc:
                        diagnostics.append("KEYNOTE_PARAM set id failed: {}".format(exc))
                try:
                    param.SetValueString(str(key_value))
                    if logger:
                        logger.info(
                            "[Place Elements] Applied keynote value '%s' to tag (value string).",
                            key_value,
                        )
                    return True
                except Exception as exc:
                    diagnostics.append("KEYNOTE_PARAM set value string failed: {}".format(exc))
        for name in ("Key Value", "Keynote Value", "Keynote"):
            try:
                param = tag_element.LookupParameter(name)
            except Exception:
                param = None
            if not param:
                continue
            if param.IsReadOnly:
                diagnostics.append("{} readonly".format(name))
                continue
            try:
                param.Set(str(key_value))
                if logger:
                    logger.info(
                        "[Place Elements] Applied keynote value '%s' to tag (%s).",
                        key_value,
                        name,
                    )
                return True
            except Exception as exc:
                diagnostics.append("{} set failed: {}".format(name, exc))
                continue

        type_elem = None
        try:
            type_id = tag_element.GetTypeId()
            type_elem = tag_element.Document.GetElement(type_id)
        except Exception:
            type_elem = None
        if type_elem:
            try:
                param = type_elem.get_Parameter(BuiltInParameter.KEYNOTE_PARAM)
            except Exception:
                param = None
            if param:
                try:
                    if param.IsReadOnly:
                        diagnostics.append("type KEYNOTE_PARAM readonly")
                    else:
                        if resolved_entry is not None:
                            param.Set(resolved_entry)
                        else:
                            param.SetValueString(str(key_value))
                        if logger:
                            logger.info(
                                "[Place Elements] Applied keynote value '%s' to tag type.",
                                key_value,
                            )
                        return True
                except Exception as exc:
                    diagnostics.append("type KEYNOTE_PARAM set failed: {}".format(exc))
            for name in ("Key Value", "Keynote Value", "Keynote"):
                try:
                    param = type_elem.LookupParameter(name)
                except Exception:
                    param = None
                if not param:
                    continue
                if param.IsReadOnly:
                    diagnostics.append("type {} readonly".format(name))
                    continue
                try:
                    param.Set(str(key_value))
                    if logger:
                        logger.info(
                            "[Place Elements] Applied keynote value '%s' to tag type (%s).",
                            key_value,
                            name,
                        )
                    return True
                except Exception as exc:
                    diagnostics.append("type {} set failed: {}".format(name, exc))

        if logger:
            logger.info(
                "[Place Elements] Failed to apply keynote value '%s' to tag.",
                key_value,
            )
            if diagnostics:
                logger.info(
                    "[Place Elements] Keynote tag parameter diagnostics: %s",
                    "; ".join(diagnostics),
                )
        return False

    def _apply_parameters(self, element, params_dict, parent_element=None):
        from Autodesk.Revit.DB import StorageType, UnitUtils

        if not params_dict:
            return
        for raw_name, value in params_dict.items():
            name = (raw_name or "").replace(SAFE_HASH, "#")
            try:
                param = element.LookupParameter(name)
            except Exception:
                param = None
            if (not param or param.IsReadOnly) and isinstance(element, IndependentTag):
                lowered = name.strip().lower()
                if lowered in ("key value", "keynote value", "keynote"):
                    if not param or param.IsReadOnly:
                        for alias in ("Key Value", "Keynote Value", "Keynote"):
                            try:
                                param = element.LookupParameter(alias)
                            except Exception:
                                param = None
                            if param and not param.IsReadOnly:
                                break
                            if param and param.IsReadOnly:
                                param = None
                    if not param:
                        try:
                            param = element.get_Parameter(BuiltInParameter.KEYNOTE_PARAM)
                        except Exception:
                            param = None
            if not param or param.IsReadOnly:
                continue
            parent_param_name = self._extract_parent_parameter_name(value)
            if parent_param_name:
                self._apply_parent_parameter(param, parent_element, parent_param_name)
                continue
            try:
                storage_type = param.StorageType
                if storage_type == StorageType.ElementId:
                    param_name = None
                    try:
                        definition = getattr(param, "Definition", None)
                        param_name = getattr(definition, "Name", None) if definition else None
                    except Exception:
                        param_name = None
                    name_hint = param_name or name
                    if self._is_load_classification_param(name_hint):
                        classification_id = self._resolve_load_classification_id(value)
                        if classification_id is not None:
                            param.Set(classification_id)
                            continue
                    try:
                        param.Set(ElementId(int(value)))
                    except Exception:
                        if self._is_keynote_param_name(name_hint):
                            resolved = self._resolve_keynote_entry_id(value)
                            if resolved is not None:
                                try:
                                    param.Set(resolved)
                                    continue
                                except Exception:
                                    pass
                            try:
                                param.SetValueString(str(value))
                                continue
                            except Exception:
                                pass
                        continue
                if storage_type == StorageType.Integer:
                    param.Set(int(value))
                elif storage_type == StorageType.Double:
                    needs_va = bool(name and "Apparent Load" in name)
                    needs_voltage = bool(name and "Voltage" in name)
                    if needs_va or needs_voltage:
                        get_unit = getattr(param, "GetUnitTypeId", None)
                        unit_id = None
                        if callable(get_unit):
                            try:
                                unit_id = get_unit()
                            except Exception:
                                unit_id = None
                        if unit_id:
                            try:
                                converted = UnitUtils.ConvertToInternalUnits(float(value), unit_id)
                                param.Set(converted)
                                continue
                            except Exception:
                                pass
                    param.Set(float(value))
                elif storage_type == StorageType.String:
                    param.Set(str(value))
                else:
                    param.Set(str(value))
            except Exception:
                continue

    def _get_linker_template(self, linked_def):
        cache = getattr(linked_def, "_ced_linker_template", None)
        if cache is not None:
            return cache
        params = linked_def.get_static_params() or {}
        for name in ELEMENT_LINKER_PARAM_NAMES:
            value = params.get(name)
            if isinstance(value, basestring) and value.strip():
                parsed = _parse_linker_payload(value)
                cache = {
                    "param_name": name,
                    "led_id": parsed.get("led_id"),
                    "set_id": parsed.get("set_id"),
                    "level_id": parsed.get("level_id"),
                }
                setattr(linked_def, "_ced_linker_template", cache)
                return cache
        led_id = getattr(linked_def, "_ced_led_id", None)
        set_id = getattr(linked_def, "_ced_set_id", None)
        if led_id or set_id:
            cache = {
                "param_name": None,
                "led_id": led_id,
                "set_id": set_id,
                "level_id": None,
            }
            setattr(linked_def, "_ced_linker_template", cache)
            return cache
        cache = {}
        setattr(linked_def, "_ced_linker_template", cache)
        return cache

    def _is_pipe_category(self, element):
        if element is None:
            return False
        cat = getattr(element, "Category", None)
        if not cat:
            return False
        try:
            cat_id = cat.Id.IntegerValue
        except Exception:
            cat_id = None
        for bip in (BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeAccessory):
            try:
                if cat_id == int(bip):
                    return True
            except Exception:
                continue
        name = (getattr(cat, "Name", "") or "").lower()
        return ("pipe fitting" in name) or ("pipe accessory" in name)

    def _set_element_linker_param(self, element, payload_value):
        if not element or not payload_value:
            return False
        success = False
        for name in ELEMENT_LINKER_PARAM_NAMES:
            try:
                param = element.LookupParameter(name)
            except Exception:
                param = None
            if not param or param.IsReadOnly:
                continue
            try:
                param.Set(payload_value)
                success = True
            except Exception:
                continue
        return success

    def _update_element_linker_parameter(
        self,
        instance,
        linked_def,
        location,
        rotation_deg,
        parent_rotation_deg=None,
        parent_element_id=None,
        host_name=None,
        parent_location=None,
    ):
        if not instance or not linked_def:
            return
        template = self._get_linker_template(linked_def)
        if not template:
            template = {}
        led_id = template.get("led_id")
        set_id = template.get("set_id")
        if not led_id and not set_id and not self._is_pipe_category(instance):
            return
        level_id = None
        level_ref = getattr(instance, "LevelId", None)
        if level_ref is not None:
            try:
                level_id = level_ref.IntegerValue
            except Exception:
                level_id = None
        element_id = None
        try:
            element_id = instance.Id.IntegerValue
        except Exception:
            element_id = None
        facing = getattr(instance, "FacingOrientation", None)
        payload = _build_linker_payload(
            led_id=led_id,
            set_id=set_id,
            location=location,
            rotation_deg=rotation_deg,
            level_id=level_id,
            element_id=element_id,
            facing=facing,
            parent_rotation_deg=parent_rotation_deg,
            host_name=host_name,
            parent_element_id=parent_element_id,
            parent_location=parent_location,
        )
        self._set_element_linker_param(instance, payload)

    def _apply_recorded_level(self, instance, linked_def):
        if not instance or not linked_def:
            return
        template = self._get_linker_template(linked_def)
        if not template:
            return
        level_id_val = template.get("level_id")
        if not level_id_val:
            return
        try:
            level_element = self.doc.GetElement(ElementId(int(level_id_val)))
        except Exception:
            level_element = None
        if not level_element:
            return
        level_id = level_element.Id
        level_param_names = (
            "INSTANCE_LEVEL_PARAM",
            "FAMILY_LEVEL_PARAM",
            "SCHEDULE_LEVEL_PARAM",
            "INSTANCE_REFERENCE_LEVEL_PARAM",
        )
        for name in level_param_names:
            bip = getattr(BuiltInParameter, name, None)
            if bip is None:
                continue
            try:
                param = instance.get_Parameter(bip)
            except Exception:
                param = None
            if not param or param.IsReadOnly:
                continue
            try:
                param.Set(level_id)
                return
            except Exception:
                continue


__all__ = ["PlaceElementsEngine"]
