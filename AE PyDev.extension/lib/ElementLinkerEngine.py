# -*- coding: utf-8 -*-
"""
ElementLinkerEngine
-------------------
Real placement engine for the Element Linker workflow.

- Uses labels in format "FamilyName : TypeName"
- Places family instances for each mapped CAD row
- Applies offsets, rotation, and simple instance parameters
- Single transaction for all placements
- No logging, no tags, no model groups
"""

import math

from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    FamilySymbol,
    BuiltInParameter,
    XYZ,
    Line,
    Structure,
    ElementTransformUtils,
    Level
)

from pyrevit import forms

from Element_Linker import CAD_BLOCK_PROFILES
from ElementLinkerUtils import feet_inch_to_inches


try:
    basestring
except NameError:
    basestring = str


class ElementPlacementEngine(object):
    def __init__(self, doc, default_level=None):
        """
        doc: Revit Document
        default_level: Level element to use for placing model families
        """
        self.doc = doc
        self.default_level = default_level
        self._init_symbol_map()

    # ------------------------------------------------------------------ #
    #  Symbol map (Family : Type) with safe Name access
    # ------------------------------------------------------------------ #
    def _init_symbol_map(self):
        """Build a map: 'FamilyName : TypeName' -> FamilySymbol."""
        doc = self.doc
        symbols = list(
            FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
        )

        self.symbol_label_map = {}
        self._activated_symbols = set()

        for sym in symbols:
            try:
                # Family name
                family = getattr(sym, "Family", None)
                if family is None:
                    continue
                fam_name = getattr(family, "Name", None)
                if not fam_name:
                    continue

                # Type name: prefer SYMBOL_NAME_PARAM, fall back to sym.Name
                type_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                type_name = None
                if type_param:
                    type_name = type_param.AsString()
                if not type_name and hasattr(sym, "Name"):
                    type_name = sym.Name
                if not type_name:
                    continue

                label = u"{} : {}".format(fam_name, type_name)  # FAMILY : TYPE
                self.symbol_label_map[label] = sym

            except Exception:
                # Skip anything weird
                continue

    # ------------------------------------------------------------------ #
    #  Public API
    # ------------------------------------------------------------------ #
    def place_from_csv(self, csv_rows, cad_selection_map):
        """
        csv_rows: list[dict] with at least 'Name', 'Position X/Y/Z' (strings)
        cad_selection_map: { cad_name: label or [labels...] }

        Each label must be in the form "FamilyName : TypeName" and match both:
        - a TypeConfig.label in the profile for that CAD name
        - a real FamilySymbol in the current project
        """
        if not csv_rows:
            forms.alert("No rows in CSV for placement.", title="Element Linker")
            return

        doc = self.doc

        # Get a level if none was provided
        level = self.default_level
        if level is None:
            level = FilteredElementCollector(doc).OfClass(Level).FirstElement()
            if level is None:
                forms.alert(
                    "No Level found in this document; cannot place model families.",
                    title="Element Linker",
                )
                return
        self.default_level = level

        t = Transaction(doc, "Element Linker Placement")
        t.Start()

        total_rows = 0
        rows_with_mapping = 0
        rows_with_coords = 0
        placed_count = 0

        # (cad_name, label) -> occurrence index, used for offset patterns
        occurrence_counter = {}

        try:
            for row in csv_rows:
                total_rows += 1

                cad_name = (row.get("Name") or "").strip()
                if not cad_name:
                    continue

                labels = cad_selection_map.get(cad_name)
                if not labels:
                    # no mapping chosen for this CAD name
                    continue

                rows_with_mapping += 1

                if isinstance(labels, basestring):
                    labels = [labels]

                # Coordinates: X/Y/Z as feet+inches or numeric feet
                x_raw = (row.get("Position X") or "").strip()
                y_raw = (row.get("Position Y") or "").strip()
                z_raw = (row.get("Position Z") or "").strip()

                if not x_raw or not y_raw or not z_raw:
                    continue

                x_inches = feet_inch_to_inches(x_raw)
                y_inches = feet_inch_to_inches(y_raw)
                z_inches = feet_inch_to_inches(z_raw)

                # Fallback: treat as feet if feet_inch_to_inches returns None
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

                base_loc = XYZ(
                    x_inches / 12.0,
                    y_inches / 12.0,
                    z_inches / 12.0,
                )

                try:
                    rot_deg = float(row.get("Rotation", 0.0))
                except Exception:
                    rot_deg = 0.0

                for label in labels:
                    key = (cad_name, label)
                    occ_index = occurrence_counter.get(key, 0)
                    occurrence_counter[key] = occ_index + 1

                    if self._place_one_instance(
                        cad_name=cad_name,
                        label=label,
                        base_loc=base_loc,
                        base_rot_deg=rot_deg,
                        occurrence_index=occ_index,
                    ):
                        placed_count += 1

            t.Commit()

        except Exception as e:
            t.RollBack()
            forms.alert(
                "Error during placement:\n\n{0}".format(e),
                title="Element Linker",
            )
            return

        # Simple summary
        msg_lines = [
            "Element Linker placement complete.",
            "",
            "Total CSV rows: {0}".format(total_rows),
            "Rows with a selected CAD mapping: {0}".format(rows_with_mapping),
            "Rows with valid coordinates: {0}".format(rows_with_coords),
            "Family instances placed: {0}".format(placed_count),
        ]

        if placed_count == 0:
            msg_lines.append("")
            msg_lines.append("No elements were placed.")
            msg_lines.append("Check that:")
            msg_lines.append(" • Families/types are loaded in this project.")
            msg_lines.append(" • Labels in profiles match 'Family : Type' exactly.")

        forms.alert("\n".join(msg_lines), title="Element Linker")

    # ------------------------------------------------------------------ #
    #  Internal placement helper
    # ------------------------------------------------------------------ #
    def _place_one_instance(self, cad_name, label, base_loc, base_rot_deg, occurrence_index):
        """
        Place a single family instance for the given CAD row and label.

        - Uses InstanceConfig.get_offset(occurrence_index) for patterned offsets
        - Applies rotation around Z
        - Applies simple instance parameters
        - Handles both model families and Generic Annotations
        """
        profile = CAD_BLOCK_PROFILES.get(cad_name)
        if not profile:
            return False

        type_cfg = profile.find_type_by_label(label)
        if not type_cfg:
            return False

        # This version only handles family instances; group types are ignored
        if getattr(type_cfg, "is_group", False):
            return False

        inst_cfg = type_cfg.instance_config
        offset_cfg = inst_cfg.get_offset(occurrence_index)

        # Look up the symbol by "Family : Type" label (works for ALL categories)
        symbol = self.symbol_label_map.get(label)
        if not symbol:
            return False

        # Activate symbol once
        if symbol.Id not in self._activated_symbols:
            try:
                if not symbol.IsActive:
                    symbol.Activate()
                    self.doc.Regenerate()
                self._activated_symbols.add(symbol.Id)
            except Exception:
                return False

        # Determine if this is a Generic Annotation (view-based family)
        cat = getattr(symbol, "Category", None)
        is_generic_annotation = False
        if cat and getattr(cat, "Name", None) == "Generic Annotations":
            is_generic_annotation = True
        elif getattr(type_cfg, "category_name", "") == "Generic Annotations":
            # Fallback to profile metadata if needed
            is_generic_annotation = True

        # Offsets in feet
        ox = offset_cfg.x_inches / 12.0
        oy = offset_cfg.y_inches / 12.0
        oz = offset_cfg.z_inches / 12.0

        loc = XYZ(
            base_loc.X + ox,
            base_loc.Y + oy,
            base_loc.Z + oz,
        )
        z_offset_feet = oz

        try:
            if is_generic_annotation:
                # --- Generic Annotation placement: view-based ---
                view = self.doc.ActiveView
                if view is None:
                    return False

                instance = self.doc.Create.NewFamilyInstance(
                    loc,
                    symbol,
                    view
                )

                # No INSTANCE_ELEVATION_PARAM logic here; GA are view elements

            else:
                # --- Normal model family placement (what you already had) ---
                level = self.default_level
                if level is None:
                    return False

                instance = self.doc.Create.NewFamilyInstance(
                    loc,
                    symbol,
                    level,
                    Structure.StructuralType.NonStructural,
                )

                # Override elevation if we have a vertical offset
                if abs(z_offset_feet) > 1e-6:
                    elev_param = instance.get_Parameter(
                        BuiltInParameter.INSTANCE_ELEVATION_PARAM
                    )
                    if elev_param and not elev_param.IsReadOnly:
                        elev_param.Set(z_offset_feet)

            # Apply rotation (works for both model and GA in plan views)
            total_rot_deg = base_rot_deg + offset_cfg.rotation_deg
            if abs(total_rot_deg) > 1e-6:
                angle_rad = math.radians(total_rot_deg)
                axis = Line.CreateBound(loc, loc + XYZ(0, 0, 1))
                ElementTransformUtils.RotateElement(
                    self.doc,
                    instance.Id,
                    axis,
                    angle_rad,
                )

            # Apply simple name->value parameters
            self._apply_parameters(instance, inst_cfg.get_parameters())

        except Exception:
            return False

        return True


    # ------------------------------------------------------------------ #
    #  Parameter helper
    # ------------------------------------------------------------------ #
    def _apply_parameters(self, element, params_dict):
        """Apply simple string/int/float parameters by name."""
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
                # StorageType.ElementId not handled in this simple version
            except Exception:
                # Silently skip parameters that fail to set
                continue
