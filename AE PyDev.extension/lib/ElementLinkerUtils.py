# -*- coding: utf-8 -*-
"""
ElementLinkerUtils
------------------
Utility helpers for the Element Linker workflow.

- feet_inch_to_inches: parse strings like 139'-10 3/16"
- read_xyz_csv: read CAD CSV (Name, Position X/Y/Z, Rotation, Count = 1)
- organize_symbols_by_category: build label maps for family symbols
- organize_model_groups: basic support for model groups (no attached details yet)
"""

import csv
import codecs

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilySymbol,
    GroupType,
    BuiltInCategory,
)


# ---------------------------------------------------------------------------
# UNIT PARSER
# ---------------------------------------------------------------------------

def feet_inch_to_inches(value):
    """
    Converts strings like "139'-10 3/16\"" into total inches (float).
    Handles:
        5'-6"
        -5'-6"
        10'
        10 1/2"
        6 3/4
    Returns None if conversion fails.
    """
    try:
        if value is None:
            return None
        s = value.strip()
        if not s:
            return None

        # Remove trailing quote symbols
        s = s.replace('"', '').replace('”', '').replace('“', '')

        sign = 1.0
        if s.startswith('-'):
            sign = -1.0
            s = s[1:].strip()

        # Split feet and the rest on '
        feet = 0.0
        inches = 0.0

        if "'" in s:
            ft_part, rest = s.split("'", 1)
            ft_part = ft_part.strip()
            if ft_part:
                feet = float(ft_part)
            s = rest.strip()
        else:
            # no explicit feet portion; treat as inches-only
            s = s.strip()

        # Remaining s may be something like: "10 3/16" or "10" or "3/16"
        if s:
            parts = s.split()
            if len(parts) == 1:
                # "10" or "3/16"
                if '/' in parts[0]:
                    num, den = parts[0].split('/')
                    inches = float(num) / float(den)
                else:
                    inches = float(parts[0])
            elif len(parts) == 2:
                # "10 3/16"
                whole = float(parts[0])
                num, den = parts[1].split('/')
                frac = float(num) / float(den)
                inches = whole + frac

        total_inches = sign * (feet * 12.0 + inches)
        return total_inches
    except Exception:
        return None


# ---------------------------------------------------------------------------
# CSV READER
# ---------------------------------------------------------------------------

def read_xyz_csv(csv_path):
    """
    Reads the CAD CSV and returns:
        rows: list[dict]
        unique_names: list of CAD block names (strings)
    Skips rows where Count != "1" or Position X is blank.
    """
    xyz_rows = []
    unique_names = set()

    with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=',')
        for row in reader:
            if row.get("Count", "").strip() != "1":
                continue
            if not row.get("Position X", "").strip():
                continue

            xyz_rows.append(row)
            cad_name = row.get("Name", "").strip()
            if cad_name:
                unique_names.add(cad_name)

    return xyz_rows, list(unique_names)


# ---------------------------------------------------------------------------
# SYMBOL ORGANIZATION
# ---------------------------------------------------------------------------

def organize_symbols_by_category(family_symbols):
    """
    Build lookup tables for FamilySymbols.

    Returns:
        symbol_label_map:  { "Type : Family": FamilySymbol }
        symbols_by_category: { category_name: ["Type : Family", ...] }
        families_by_category: { category_name: [family_name, ...] }
        types_by_family: { family_name: [type_name, ...] }
        family_symbols: [FamilySymbol, ...]
    """
    symbol_label_map = {}
    symbols_by_category = {}
    families_by_category = {}
    types_by_family = {}
    family_symbols = []

    for sym in family_symbols:
        try:
            cat = sym.Category
            if cat is None:
                continue

            category_name = cat.Name
            family = getattr(sym, "Family", None)
            if family is None:
                continue

            family_name = family.Name
            type_name = sym.Name

            label = u"{} : {}".format(type_name, family_name)

            symbol_label_map[label] = sym
            family_symbols.append(sym)

            # symbols_by_category
            if category_name not in symbols_by_category:
                symbols_by_category[category_name] = []
            if label not in symbols_by_category[category_name]:
                symbols_by_category[category_name].append(label)

            # families_by_category
            if category_name not in families_by_category:
                families_by_category[category_name] = []
            if family_name not in families_by_category[category_name]:
                families_by_category[category_name].append(family_name)

            # types_by_family
            if family_name not in types_by_family:
                types_by_family[family_name] = []
            if type_name not in types_by_family[family_name]:
                types_by_family[family_name].append(type_name)

        except Exception:
            # Best effort; skip bad symbols
            continue

    # Sort lists
    for cat_name in symbols_by_category:
        symbols_by_category[cat_name].sort()
    for cat_name in families_by_category:
        families_by_category[cat_name].sort()
    for fam_name in types_by_family:
        types_by_family[fam_name].sort()

    return symbol_label_map, symbols_by_category, families_by_category, types_by_family, family_symbols


# ---------------------------------------------------------------------------
# MODEL GROUP ORGANIZATION (simplified)
# ---------------------------------------------------------------------------

def organize_model_groups(doc):
    """
    Basic model group organization.

    Returns:
        group_label_map: { "None : ModelGroupName": (GroupType, None) }
        details_by_model_group: { GroupType: [None] }  # placeholder
        model_group_names: [ModelGroupName, ...]
    """
    group_label_map = {}
    details_by_model_group = {}
    model_group_names = []

    # Collect all GroupTypes that are model groups
    groups = FilteredElementCollector(doc).OfClass(GroupType)
    for gtype in groups:
        try:
            cat = gtype.Category
            if not cat:
                continue
            if cat.Id.IntegerValue != int(BuiltInCategory.OST_IOSModelGroups):
                continue

            model_group_name = gtype.Name
            label = u"None : {}".format(model_group_name)

            group_label_map[label] = (gtype, None)
            details_by_model_group[gtype] = [None]
            model_group_names.append(model_group_name)
        except Exception:
            continue

    model_group_names.sort()
    return group_label_map, details_by_model_group, model_group_names
