# -*- coding: utf-8 -*-
"""
Placement engine that consumes profile definitions and places Revit elements.
"""

import os
import math

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
    Reference,
    TagMode,
    TagOrientation,
    ElementId,
)

from LogicClasses.csv_helpers import feet_inch_to_inches
from LogicClasses.tag_utils import tag_key_from_dict

try:
    basestring
except NameError:  # Python 3 fallback
    basestring = str

ELEMENT_LINKER_PARAM_NAMES = ("Element_Linker", "Element_Linker Parameter")


def _parse_linker_payload(payload_text):
    if not payload_text:
        return {}
    entries = {}
    for raw_line in payload_text.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, _, remainder = line.partition(":")
        entries[key.strip()] = remainder.strip()
    return {
        "led_id": (entries.get("Linked Element Definition ID", "") or "").strip(),
        "set_id": (entries.get("Set Definition ID", "") or "").strip(),
    }


def _format_xyz(vec):
    if not vec:
        return ""
    return "{:.6f},{:.6f},{:.6f}".format(vec.X, vec.Y, vec.Z)


def _build_linker_payload(led_id, set_id, location, rotation_deg, level_id, element_id, facing):
    rotation = float(rotation_deg or 0.0)
    lines = [
        "Linked Element Definition ID: {}".format(led_id or ""),
        "Set Definition ID: {}".format(set_id or ""),
        "Location XYZ (ft): {}".format(_format_xyz(location)),
        "Rotation (deg): {:.6f}".format(rotation),
        "LevelId: {}".format(level_id if level_id is not None else ""),
        "ElementId: {}".format(element_id if element_id is not None else ""),
        "FacingOrientation: {}".format(_format_xyz(facing)),
    ]
    return "\n".join(lines).strip()


class PlaceElementsEngine(object):
    def __init__(self, doc, repo, default_level=None, tag_view_map=None, allow_tags=True, transaction_name="Place Elements (YAML)"):
        self.doc = doc
        self.repo = repo
        self.default_level = default_level
        self.tag_view_map = tag_view_map or {}
        self.allow_tags = bool(allow_tags)
        self.transaction_name = transaction_name or "Place Elements (YAML)"
        self._init_symbol_map()
        self._init_group_map()

    def _init_symbol_map(self):
        """Map 'Family : Type' to FamilySymbol."""
        self.symbol_label_map = {}
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
                labels = cad_selection_map.get(cad_name)
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

                for label in labels:
                    key = (cad_name, label)
                    occ_index = occurrence_counter.get(key, 0)
                    occurrence_counter[key] = occ_index + 1
                    linked_def = self.repo.definition_for_label(cad_name, label)
                    if linked_def and self._place_one(linked_def, base_loc, base_rot_deg, occ_index):
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

            t.Commit()
        except Exception:
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

    def _place_one(self, linked_def, base_loc, base_rot_deg, occurrence_index):
        placement = linked_def.get_placement()
        offset_xyz = placement.get_offset_xyz() if placement else None
        offset = offset_xyz or (0.0, 0.0, 0.0)
        rot_offset = placement.get_rotation_degrees() if placement else 0.0
        placement_mode = placement.get_placement_mode() if placement else None
        is_group = bool(placement_mode and str(placement_mode).lower() == "group")

        total_offset_rotation = base_rot_deg + (rot_offset or 0.0)
        if total_offset_rotation:
            ang = math.radians(total_offset_rotation)
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

        instance = None
        if is_group:
            instance = self._place_group(label, family, type_name, linked_def, loc, gtype, detail_type)
        if not instance:
            instance = self._place_symbol(label, family, type_name, linked_def, loc, offset[2])
        if instance:
            if abs(final_rot_deg) > 1e-6:
                self._rotate_instance(instance, loc, final_rot_deg)
            self._update_element_linker_parameter(instance, linked_def, loc, final_rot_deg)
            if self.allow_tags:
                self._place_tags(tags, instance, loc, final_rot_deg)
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

            if detail_type is not None:
                try:
                    view = getattr(self.doc, "ActiveView", None)
                    if view:
                        instance.ShowAttachedDetailGroups(view, detail_type.Id)
                except Exception:
                    pass

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

    def _place_symbol(self, label, family_name, type_name, linked_def, location, z_offset_feet):
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

            self._apply_parameters(instance, linked_def.get_static_params())
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

    def _place_tags(self, tag_defs, host_instance, base_loc, final_rot_deg):
        if not tag_defs:
            return
        active_view = getattr(self.doc, "ActiveView", None)
        for tag in tag_defs:
            family = tag.get("family")
            type_name = tag.get("type")
            label = None
            if family and type_name:
                label = u"{} : {}".format(family, type_name)
            elif type_name:
                label = type_name
            elif family:
                label = family
            else:
                continue
            symbol = self.symbol_label_map.get(label)
            if not symbol and family and type_name:
                key = u"{} : {}".format(family, type_name)
                symbol = self.symbol_label_map.get(key)
            if not symbol:
                continue
            if not self._activate_symbol(symbol):
                continue

            offsets = tag.get("offset") or (0.0, 0.0, 0.0)
            category_name = (tag.get("category") or "").lower()
            parameters = tag.get("parameters") or {}

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
                        try:
                            reference = Reference(host_instance)
                        except Exception:
                            continue
                        independent = IndependentTag.Create(
                            self.doc,
                            view_obj.Id,
                            reference,
                            True,
                            TagMode.TM_ADDBY_CATEGORY,
                            TagOrientation.Horizontal,
                            tag_loc,
                        )
                        if not independent:
                            continue
                        independent.ChangeTypeId(symbol.Id)
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
                except Exception:
                    instance = None
                if not instance:
                    continue
                if isinstance(instance, IndependentTag):
                    try:
                        instance.TagHeadPosition = tag_loc
                    except Exception:
                        pass
                else:
                    if abs(tag_rotation) > 1e-6:
                        try:
                            axis = Line.CreateBound(tag_loc, tag_loc + XYZ(0, 0, 1))
                            ElementTransformUtils.RotateElement(self.doc, instance.Id, axis, math.radians(tag_rotation))
                        except Exception:
                            pass
                self._apply_parameters(instance, parameters)

    def _apply_parameters(self, element, params_dict):
        from Autodesk.Revit.DB import StorageType

        if not params_dict:
            return
        for name, value in params_dict.items():
            param = element.LookupParameter(name)
            if not param or param.IsReadOnly:
                continue
            try:
                if param.StorageType == StorageType.Integer:
                    param.Set(int(value))
                elif param.StorageType == StorageType.Double:
                    param.Set(float(value))
                elif param.StorageType == StorageType.String:
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
                }
                setattr(linked_def, "_ced_linker_template", cache)
                return cache
        cache = {}
        setattr(linked_def, "_ced_linker_template", cache)
        return cache

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

    def _update_element_linker_parameter(self, instance, linked_def, location, rotation_deg):
        if not instance or not linked_def:
            return
        template = self._get_linker_template(linked_def)
        if not template or not template.get("led_id"):
            return
        led_id = template.get("led_id")
        set_id = template.get("set_id")
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
        )
        self._set_element_linker_param(instance, payload)


__all__ = ["PlaceElementsEngine"]
