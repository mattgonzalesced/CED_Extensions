# -*- coding: utf-8 -*-

from pyrevit import revit, DB

# Get the current document
doc = revit.doc

# Define reference subcategory names and their specifications (without line pattern)
subcategory_specs = {
    "Weak Reference": {
        "color": DB.Color(0, 127, 0),  # RGB Green
        "line_weight": 1
    },
    "Not a Reference": {
        "color": DB.Color(255, 183, 255),  # RGB Pink
        "line_weight": 1
    },
    "Strong Reference": {
        "color": DB.Color(206, 0, 0),  # RGB Red
        "line_weight": 3
    },
    "Origin": {
        "color": DB.Color(0, 0, 206),  # RGB Blue
        "line_weight": 3
    }
}



# Function to get or create a line pattern by name
# Function to get or create a line pattern containing "Aligning" in its name
def get_or_create_line_pattern(line_pattern_name="Aligning Line"):
    # Search for any line pattern that contains "Aligning" in its name
    line_patterns = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement)
    for pattern in line_patterns:
        if "Aligning Line" in pattern.Name:
            return pattern.Id

    # If no matching pattern is found, create a new one with the specified name
    # Here, we're creating a simple dashed pattern; you can customize as needed
    new_line_pattern = DB.LinePattern(line_pattern_name)
    return DB.LinePatternElement.Create(doc, new_line_pattern).Id


# Function to create or update subcategory with specified settings, using parent category's line pattern
def get_or_create_subcategory(parent_category, subcategory_name, specs):
    # Get the parent category's line pattern ID
    parent_line_pattern_id = parent_category.GetLinePatternId(DB.GraphicsStyleType.Projection)

    # Check if subcategory exists
    for subcategory in parent_category.SubCategories:
        if subcategory.Name == subcategory_name:
            # Overwrite existing subcategory properties
            subcategory.LineColor = specs["color"]
            subcategory.SetLineWeight(specs["line_weight"], DB.GraphicsStyleType.Projection)
            subcategory.SetLinePatternId(parent_line_pattern_id, DB.GraphicsStyleType.Projection)
            return subcategory.Id

    # Create subcategory if it doesn't exist
    new_subcategory = doc.Settings.Categories.NewSubcategory(parent_category, subcategory_name)
    new_subcategory.LineColor = specs["color"]
    new_subcategory.SetLineWeight(specs["line_weight"], DB.GraphicsStyleType.Projection)
    new_subcategory.SetLinePatternId(parent_line_pattern_id, DB.GraphicsStyleType.Projection)
    return new_subcategory.Id


# Ensure subcategories with specified styles exist
def ensure_subcategories():
    loaded_subcategories = {}
    clines_category = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_CLines)

    with revit.Transaction("Create Subcategories"):
        # Create each subcategory with specified styles if missing
        for subcategory_name, specs in subcategory_specs.items():
            subcategory_id = get_or_create_subcategory(clines_category, subcategory_name, specs)
            loaded_subcategories[subcategory_name] = subcategory_id

    return loaded_subcategories


# Ensure subcategories exist
loaded_subcategories = ensure_subcategories()

# Collect all reference planes in the document
collector = DB.FilteredElementCollector(doc).OfClass(DB.ReferencePlane).WhereElementIsNotElementType()

# Define the unwanted values for the "Is Reference" parameter
unwanted_names = ["Not a Reference", "Weak Reference", "Strong Reference"]

# Start a transaction to modify reference planes
with revit.Transaction("Assign Reference Plane Subcategories"):
    for ref_plane in collector:
        # Retrieve the "Is Reference" and "Defines Origin" parameters
        is_reference_param = ref_plane.LookupParameter("Is Reference")
        is_reference_value = is_reference_param.AsValueString() if is_reference_param else None

        defines_origin_param = ref_plane.LookupParameter("Defines Origin")
        defines_origin_value = defines_origin_param.AsInteger() if defines_origin_param else 0  # 1 = Checked, 0 = Unchecked

        # Get the built-in "CLINE_SUBCATEGORY" parameter
        subcategory_param = ref_plane.get_Parameter(DB.BuiltInParameter.CLINE_SUBCATEGORY)

        # Set the parameter value based on the rules
        if defines_origin_value == 1:
            # Set to Origin if Defines Origin is checked
            subcategory_param.Set(loaded_subcategories["Origin"])

        elif is_reference_value == "Weak Reference":
            # Set to Weak Reference subcategory if Is Reference is Weak Reference
            subcategory_param.Set(loaded_subcategories["Weak Reference"])

        elif is_reference_value == "Not a Reference":
            # Set to Not a Reference subcategory if Is Reference is Not a Reference
            subcategory_param.Set(loaded_subcategories["Not a Reference"])

        elif is_reference_value and is_reference_value != "Weak Reference" and is_reference_value != "Not a Reference":
            # Set to Strong Reference if it's any other valid reference type
            subcategory_param.Set(loaded_subcategories["Strong Reference"])

            # Copy the reference type name to the Name parameter if it’s not "Weak Reference", "Not a Reference", or "Strong Reference"
            if is_reference_value not in unwanted_names:

                name_param = ref_plane.LookupParameter("Name")

                if name_param:
                    name_param.Set(is_reference_value)
# The transaction automatically commits the changes at the end of the 'with' block
