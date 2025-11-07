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
        for row in reader:
            # Skip invalid rows
            if row.get("Count", "").strip() != "1" or not row.get("Position X", "").strip():
                continue

            xyz_rows.append(row)
            cad_name = row.get("Name", "").strip()
            if cad_name:
                unique_names.add(cad_name)

    return xyz_rows, unique_names


def read_matchings_json(json_path):
    """Read JSON matchings and return mapping and parameter dictionaries.

    Returns:
        matchings_dict: CAD_Block_Name -> list of family labels ("Type : Family") for example {'HM62B_Up tanning': ['Fused - 30A : EF-U_Disconnect Switch_CED']
        groups_dict: CAD_Block_Name -> list of group labels ("DetailGroup : ModelGroup" or "None : ModelGroup")
        parameters_dict: CAD_Block_Name -> {label -> {param_name -> param_value}} for example {'43 tv': {'Duplex Floor : EF-U_Receptacle_CED': {'dev-Group ID': '43 tv'}}, 'HM62B_Up tanning': {'Fused - 30A : EF-U_Disconnect Switch_CED': {'dev-Group ID': 'HM62B_Up tanning'}},
    """
    import json
    import codecs

    matchings_dict = {}
    groups_dict = {}
    parameters_dict = {}
    offsets_dict = {}

    with codecs.open(json_path, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)

    for cad_name, categories in data.items():
        families = []
        groups = []

        # Ensure categories is a dictionary
        if not isinstance(categories, dict):
            print("WARNING: CAD block '{}' value is not a dictionary (type: {}). Skipping.".format(cad_name, type(categories).__name__))
            continue

        # Process each category
        for category_name, category_data in categories.items():
            # Skip Notes field (Notes is a string, not a list)
            if category_name == "Notes":
                continue

            # category_data should be an array (skip if not)
            if not isinstance(category_data, list):
                print("WARNING: Category '{}' for CAD block '{}' is not an array. Skipping.".format(category_name, cad_name))
                continue

            # Iterate through array of fixture objects
            for fixture_obj in category_data:
                # fixture_obj is a dict with one key (the family name)
                for family_name, types_dict in fixture_obj.items():
                    # Skip entries with "_skip" prefix
                    if family_name.startswith("_skip"):
                        print("DEBUG: Skipping entry with '_skip' prefix: '{}'".format(family_name))
                        continue

                    # types_dict contains type names as keys
                    for type_name, type_data in types_dict.items():
                        # Check if this is a Model Group (accept both "Model_Groups" and "Model")
                        if category_name in ["Model_Groups", "Model"]:
                            # For model groups, use "None : ModelGroupName" format to match organize_model_groups
                            # The family_name is the actual model group name
                            label = "None : {}".format(family_name)
                            groups.append(label)
                            # DEBUG: Show Model_Groups processing
                            print("DEBUG JSON: Found Model_Group: '{}' (type: '{}')".format(family_name, type_name))
                            print("  Full label for UI: '{}'".format(label))
                        else:
                            # For families, use "Type : Family" format
                            label = "{} : {}".format(type_name, family_name)
                            families.append(label)

                        # Extract OFFSET if it exists
                        if "OFFSET" in type_data:
                            offset_data = type_data.get("OFFSET", {})
                            if offset_data:
                                try:
                                    offset_x = float(offset_data.get("x", 0.0))
                                    offset_y = float(offset_data.get("y", 0.0))
                                    offset_z = float(offset_data.get("z", 0.0))
                                    offset_rotation = float(offset_data.get("r", 0.0))
                                    offsets_dict.setdefault(cad_name, {}).setdefault(label, []).append({"x": offset_x, "y": offset_y, "z": offset_z, "r" : offset_rotation})
                                except (ValueError, TypeError):
                                    # If conversion fails, just skip (defaults to no offset)
                                    pass

                        # Extract parameters if they exist
                        if "PARAMETERS" in type_data and type_data["PARAMETERS"]:
                            params = type_data["PARAMETERS"]
                            parameters_dict.setdefault(cad_name, {}).setdefault(label, []).append(params)

        # Store results
        if families:
            matchings_dict[cad_name] = families
            print("DEBUG JSON: CAD block '{}' has {} families: {}".format(cad_name, len(families), families[:2]))
        if groups:
            groups_dict[cad_name] = groups
            print("DEBUG JSON: CAD block '{}' has {} model groups: {}".format(cad_name, len(groups), groups))
    

    return matchings_dict, groups_dict, parameters_dict, offsets_dict


def organize_symbols_by_category(fixture_symbols):
    """Organize family symbols by Revit category and create lookup dictionaries.
    
    symbol_label_map = {}  # "Type : Family" -> symbol THIS IS EVERY FAMILY AND TYPE MAP 
    symbols_by_category = {}  # category -> ["Type : Family", ...]
    families_by_category = {}  # category -> [family_name, ...]
    types_by_family = {}  # family_name -> [type_name, ...]
    family_symbols = {}  # (family_name, type_name) -> symbol

    """
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
        #print("DEBUG: Category '{}' has {} families".format(category, len(families_by_category[category])))

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


def organize_model_groups(doc):
    """Organize model groups and their attached detail groups."""
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Element, BuiltInParameter

    # Collect model groups using the correct category filter
    model_groups = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_IOSModelGroups) \
        .WhereElementIsElementType() \
        .ToElements()

    group_label_map = {}  # "DetailGroup : ModelGroup" -> (model_group_type, detail_group_type) OR "None : ModelGroup" -> (model_group_type, None)
    details_by_model_group = {}  # "ModelGroup" -> ["DetailGroup1", "DetailGroup2", "None"]
    model_group_names = []  # List of all model group names

    print("\n=== FINDING ATTACHED DETAIL GROUPS FROM MODEL GROUP TYPES ===")

    # Now organize model groups and get their attached detail groups directly from the TYPE
    for model_group_type in model_groups:
        try:
            model_group_name = Element.Name.__get__(model_group_type)
            model_group_names.append(model_group_name)

            # Initialize with "None" option (no detail group)
            details_by_model_group[model_group_name] = ["None"]

            # Add mapping for just the model group (no detail)
            label_none = "None : {}".format(model_group_name)
            group_label_map[label_none] = (model_group_type, None)

            # Get attached detail group types directly from the model group TYPE
            try:
                available_detail_ids = model_group_type.GetAvailableAttachedDetailGroupTypeIds()

                if available_detail_ids and len(available_detail_ids) > 0:
                    print("Model group '{}' has {} attached detail groups".format(
                        model_group_name, len(available_detail_ids)))

                    for detail_id in available_detail_ids:
                        detail_type = doc.GetElement(detail_id)
                        if detail_type:
                            detail_name = Element.Name.__get__(detail_type)

                            # Create label for this combination
                            label = "{} : {}".format(detail_name, model_group_name)
                            group_label_map[label] = (model_group_type, detail_type)

                            # Add to the list of available details for this model group
                            if detail_name not in details_by_model_group[model_group_name]:
                                details_by_model_group[model_group_name].append(detail_name)
                                print("  - Added detail option: '{}'".format(detail_name))
                else:
                    print("Model group '{}' has NO attached detail groups".format(model_group_name))
            except Exception as ex:
                print("Error getting attached details for '{}': {}".format(model_group_name, ex))
                print("Model group '{}' has NO attached detail groups".format(model_group_name))
        except Exception as ex:
            #print("DEBUG: Error organizing model group: {}".format(ex))
            continue

    # Sort lists
    model_group_names.sort()
    for model_group in details_by_model_group:
        details_by_model_group[model_group].sort()

    print("\n=== MODEL GROUPS AND THEIR ATTACHED DETAIL GROUPS ===")
    print("Found {} model groups total".format(len(model_group_names)))
    groups_with_details = sum(1 for mg in details_by_model_group if len(details_by_model_group[mg]) > 1)
    print("Groups with attached detail groups: {}".format(groups_with_details))
    print("Groups without attached detail groups: {}".format(len(model_group_names) - groups_with_details))

    return group_label_map, details_by_model_group, model_group_names
