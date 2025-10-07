# -*- coding: utf-8 -*-
"""
Utility functions for the Family Placement script.
"""

def feet_inch_to_inches(value):
    """
    Converts a string like "139'-10 3/16"" into total inches (float).
    If the feet part is negative, the inches are subtracted.
    For example, "-5'6"" returns -66.
    Returns None if conversion fails.
    """
    try:
        value = value.strip()
        if not value:
            return None
        parts = value.split("'")
        if len(parts) < 2:
            return float(value)
        # Convert the feet part.
        feet = float(parts[0])
        # Process the inches part.
        inch_part = parts[1].replace('"', '').strip()
        # Remove any negative sign from inches (we rely on feet sign)
        if inch_part.startswith("-"):
            inch_part = inch_part[1:]
        # Convert the inches part, handling fractions if present.
        if " " in inch_part:
            inch_parts = inch_part.split(" ")
            inches = float(inch_parts[0])
            if len(inch_parts) > 1:
                fraction = inch_parts[1]
                num, denom = fraction.split("/")
                inches += float(num) / float(denom)
        else:
            if inch_part == "":
                inches = 0.0
            else:
                inches = float(inch_part)
        # If feet is negative, subtract the inches instead of adding.
        if feet < 0:
            return feet * 12 - inches
        else:
            return feet * 12 + inches
    except Exception as ex:
        print("Error converting '{0}' to inches: {1}".format(value, ex))
        return None


def create_safe_control_name(cad_name, control_counter):
    """
    Create a valid WPF control name by removing invalid characters.
    """
    safe_name = "".join(c for c in cad_name if c.isalnum() or c == "_")[:20]  # Limit length
    return safe_name


def read_xyz_csv(csv_path):
    """Read XYZ CSV and return filtered rows and unique names."""
    import csv
    import codecs

    xyz_rows = []
    unique_names = set()

    with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=',')
        reader.fieldnames = [h.strip() for h in reader.fieldnames or []]

        for row in reader:
            # Skip invalid rows
            if row.get("Count", "").strip() != "1" or not row.get("Position X", "").strip():
                continue

            xyz_rows.append(row)
            cad_name = row.get("Name", "").strip()
            if cad_name:
                unique_names.add(cad_name)

    return xyz_rows, unique_names


def read_matchings_csv(csv_path):
    """Read Structured Matchings CSV and return mapping and parameter dictionaries.

    Returns:
        matchings_dict: CAD_Block_Name -> list of family labels ("Type : Family")
        groups_dict: CAD_Block_Name -> list of group labels ("DetailGroup : ModelGroup" or "None : ModelGroup")
        parameters_dict: CAD_Block_Name -> {label -> {param_name -> param_value}}
    """
    import csv
    import codecs

    matchings_dict = {}
    groups_dict = {}
    parameters_dict = {}

    with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=',')
        reader.fieldnames = [h.strip() for h in reader.fieldnames or []]

        # Separate columns once (not per row)
        excluded = {"CAD_Block_Name", "Notes", "Model_Groups"}
        family_columns = [c for c in reader.fieldnames if c not in excluded and not c.endswith("_Parameters")]
        parameter_columns = [c for c in reader.fieldnames if c.endswith("_Parameters")]
        has_model_groups_col = "Model_Groups" in reader.fieldnames

        for row in reader:
            cad_name = row.get("CAD_Block_Name", "").strip()
            if not cad_name:
                continue

            families = []
            groups = []

            # Process family columns (existing logic)
            for col in family_columns:
                data = row.get(col, "").strip()
                if data:
                    for entry in data.split(","):
                        parts = entry.strip().split(":", 1)
                        if len(parts) == 2:
                            family_name, type_name = parts[0].strip(), parts[1].strip()
                            if family_name and type_name:
                                families.append("{} : {}".format(type_name, family_name))

            # Process Model_Groups column (NEW)
            if has_model_groups_col:
                data = row.get("Model_Groups", "").strip()
                if data:
                    for entry in data.split(","):
                        parts = entry.strip().split(":", 1)
                        if len(parts) == 2:
                            model_group_name, detail_group_name = parts[0].strip(), parts[1].strip()
                            if model_group_name and detail_group_name:
                                # Format: "DetailGroup : ModelGroup" (matching the organize_model_groups output)
                                groups.append("{} : {}".format(detail_group_name, model_group_name))
                        elif len(parts) == 1:
                            # Just model group name, no detail group
                            model_group_name = parts[0].strip()
                            if model_group_name:
                                groups.append("None : {}".format(model_group_name))

            # Process parameter columns
            for col in parameter_columns:
                data = row.get(col, "").strip()
                if data:
                    for entry in data.split(","):
                        parts = entry.strip().split(":", 3)
                        if len(parts) == 4:
                            family_name, type_name, param_name, param_value = [p.strip() for p in parts]
                            if all([family_name, type_name, param_name, param_value]):
                                family_label = "{} : {}".format(type_name, family_name)
                                parameters_dict.setdefault(cad_name, {}).setdefault(family_label, {})[param_name] = param_value

            if families:
                matchings_dict[cad_name] = families
            if groups:
                groups_dict[cad_name] = groups

    return matchings_dict, groups_dict, parameters_dict


def organize_symbols_by_category(fixture_symbols):
    """Organize family symbols by Revit category and create lookup dictionaries."""
    from Autodesk.Revit.DB import BuiltInParameter

    symbol_label_map = {}  # "Type : Family" -> symbol
    symbols_by_category = {}  # category -> ["Type : Family", ...]
    families_by_category = {}  # category -> [family_name, ...]
    types_by_family = {}  # family_name -> [type_name, ...]
    family_symbols = {}  # (family_name, type_name) -> symbol

    for sym in fixture_symbols:
        family_name = getattr(sym.Family, 'Name', 'UnknownFamily') if hasattr(sym, 'Family') else 'UnknownFamily'

        try:
            param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = param.AsString() if param else "UnknownType"
        except Exception:
            type_name = "UnknownType"

        label = "{} : {}".format(type_name, family_name)
        symbol_label_map[label] = sym
        family_symbols[(family_name, type_name)] = sym

        category_name = getattr(sym.Category, 'Name', 'Unknown Category') if hasattr(sym, 'Category') else 'Unknown Category'

        # Add to all dictionaries using setdefault
        symbols_by_category.setdefault(category_name, []).append(label)

        if family_name not in families_by_category.setdefault(category_name, []):
            families_by_category[category_name].append(family_name)

        if type_name not in types_by_family.setdefault(family_name, []):
            types_by_family[family_name].append(type_name)

    # Sort all lists
    for category in symbols_by_category:
        symbols_by_category[category].sort()
        families_by_category[category].sort()
        print("DEBUG: Category '{}' has {} families".format(category, len(families_by_category[category])))

    for family in types_by_family:
        types_by_family[family].sort()

    return (symbol_label_map, symbols_by_category, families_by_category,
            types_by_family, family_symbols)


def determine_family_category(family, symbols_by_category):
    """Determine which category a family belongs to."""
    # Check if family exists in any category
    for cat, families in symbols_by_category.items():
        if family in families:
            return cat
    return None

    # # Try to guess category based on keywords
    # family_lower = family.lower()
    # keyword_map = {
    #     'light': ['light'],
    #     'luminaire': ['light'],
    #     'receptacle': ['electrical'],
    #     'switch': ['electrical'],
    #     'equipment': ['equipment', 'mechanical'],
    #     'mechanical': ['mechanical', 'equipment'],
    #     'plumbing': ['plumbing']
    # }

    # for keyword, category_keywords in keyword_map.items():
    #     if keyword in family_lower:
    #         for cat_name in symbols_by_category.keys():
    #             cat_lower = cat_name.lower()
    #             if any(cat_kw in cat_lower for cat_kw in category_keywords):
    #                 return cat_name

    # # Fallback: use category with most families, or "Unknown Category"
    # if symbols_by_category:
    #     return max(symbols_by_category.keys(), key=lambda k: len(symbols_by_category[k]))
    # return "Unknown Category"


def organize_model_groups(doc):
    """Organize model groups and their attached detail groups."""
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Element

    # Collect model groups using the correct category filter
    model_groups = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_IOSModelGroups) \
        .WhereElementIsElementType() \
        .ToElements()

    # Collect detail groups using the correct category filter
    detail_groups = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_IOSDetailGroups) \
        .WhereElementIsElementType() \
        .ToElements()

    group_label_map = {}  # "DetailGroup : ModelGroup" -> (model_group_type, detail_group_type) OR "None : ModelGroup" -> (model_group_type, None)
    details_by_model_group = {}  # "ModelGroup" -> ["DetailGroup1", "DetailGroup2", "None"]
    model_group_names = []  # List of all model group names
    detail_groups_dict = {}  # Store detail groups by name for lookup

    # Process detail groups
    for detail_group_type in detail_groups:
        try:
            detail_name = Element.Name.__get__(detail_group_type)
            detail_groups_dict[detail_group_type.Id] = detail_group_type
            print("DEBUG: Found detail group: '{}'".format(detail_name))
        except Exception as ex:
            print("DEBUG: Error processing detail group: {}".format(ex))
            continue

    # Now organize model groups and find their attached detail groups
    for model_group_type in model_groups:
        try:
            model_group_name = Element.Name.__get__(model_group_type)
            print("DEBUG: Found model group: '{}'".format(model_group_name))

            model_group_names.append(model_group_name)

            # Initialize with "None" option (no detail group)
            details_by_model_group[model_group_name] = ["None"]

            # Add mapping for just the model group (no detail)
            label_none = "None : {}".format(model_group_name)
            group_label_map[label_none] = (model_group_type, None)

            # Try to find attached detail groups
            # Note: In Revit, detail groups are typically placed separately and associated through naming or manual linking
            # We'll make all detail groups available for all model groups
            for detail_id, detail_group_type in detail_groups_dict.items():
                # Get detail group name using Element.Name
                detail_name = Element.Name.__get__(detail_group_type)

                # Create label for this combination
                label = "{} : {}".format(detail_name, model_group_name)
                group_label_map[label] = (model_group_type, detail_group_type)

                # Add to the list of available details for this model group
                if detail_name not in details_by_model_group[model_group_name]:
                    details_by_model_group[model_group_name].append(detail_name)
        except Exception as ex:
            print("DEBUG: Error organizing model group: {}".format(ex))
            continue

    # Sort lists
    model_group_names.sort()
    for model_group in details_by_model_group:
        details_by_model_group[model_group].sort()

    print("DEBUG: Found {} model groups".format(len(model_group_names)))
    for mg_name in model_group_names:
        print("DEBUG: Model group '{}' has {} detail options".format(mg_name, len(details_by_model_group[mg_name])))

    return group_label_map, details_by_model_group, model_group_names


def set_instance_parameters(instance, parameters, symbol_label):
    """Set parameters on a family instance based on the parameters dictionary."""
    if not parameters or symbol_label not in parameters:
        return

    for param_name, param_value in parameters[symbol_label].items():
        try:
            param = instance.LookupParameter(param_name) or instance.get_Parameter(param_name)

            if not param or param.IsReadOnly:
                print("WARNING: Parameter '{}' not found or is read-only on family '{}'".format(param_name, symbol_label))
                continue

            storage_type = param.StorageType.ToString()

            # Type conversion handlers
            if storage_type == "Integer":
                try:
                    param.Set(int(param_value))
                except ValueError:
                    print("WARNING: Could not convert '{}' to integer for parameter '{}'".format(param_value, param_name))
                    continue
            elif storage_type == "Double":
                try:
                    param.Set(float(param_value))
                except ValueError:
                    print("WARNING: Could not convert '{}' to double for parameter '{}'".format(param_value, param_name))
                    continue
            else:
                param.Set(str(param_value))

            print("DEBUG: Set parameter '{}' = '{}' on instance".format(param_name, param_value))

        except Exception as ex:
            print("ERROR: Failed to set parameter '{}' = '{}': {}".format(param_name, param_value, ex))