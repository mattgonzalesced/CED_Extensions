json_path = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\REVIT\MG_tools\MG pyTools.extension\MG_Tools.tab\MG_Panel.panel\Let there be JSON.pushbutton\Kevin's profile.json"



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
                    # types_dict contains type names as keys
                    for type_name, type_data in types_dict.items():
                        # Construct the label in "Type : Family" format
                        label = "{} : {}".format(type_name, family_name)

                        # Check if this is a Model Group
                        if category_name == "Model_Groups":
                            groups.append(label)
                        else:
                            families.append(label)

                        # Extract OFFSET if it exists
                        if "OFFSET" in type_data:
                            offset_data = type_data.get("OFFSET", {})
                            if offset_data:
                                try:
                                    offset_x = float(offset_data.get("x", 0.0))
                                    offset_y = float(offset_data.get("y", 0.0))
                                    offsets_dict.setdefault(cad_name, {}).setdefault(label, []).append({"x": offset_x, "y": offset_y})
                                except (ValueError, TypeError):
                                    # If conversion fails, just skip (defaults to no offset)
                                    pass

                        # Extract parameters if they exist
                        if "PARAMETERS" in type_data and type_data["PARAMETERS"]:
                            params = type_data["PARAMETERS"]
                            parameters_dict.setdefault(cad_name, {}).setdefault(label, []).append(params)
        print('FAMILIES:')
        for family in families:
            print(family)
        # Store results
        if families:
            matchings_dict[cad_name] = families
            #print("DEBUG JSON: CAD block '{}' has {} families: {}".format(cad_name, len(families), families[:2]))
        if groups:
            groups_dict[cad_name] = groups
            #print("DEBUG JSON: CAD block '{}' has {} groups: {}".format(cad_name, len(groups), groups[:2]))
    
    print('ORGANIZE SYMBOLS BY CATEGORY FUNCTION')

    print('MATCHINGS_DICT:')
    for item in matchings_dict.items():
        print(item)

    print('GROUPS_DICT:')
    for item in groups_dict.items():
        print(item)

    print('PARAMETERS_DICT:')
    for item in parameters_dict.items():
        print(item)

    print('OFFSETS_DICT:')
    for item in offsets_dict.items():
        print(item)

    return matchings_dict, groups_dict, parameters_dict, offsets_dict

read_matchings_json(json_path)