# -*- coding: utf-8 -*-
__title__ = "Place Device on Equipment"

from pyrevit import DB, script, forms, revit
from pyrevit.revit import query

# Initialize logger
logger = script.get_logger()

# Define parameter mapping
parameter_mappings = {
    'ElementID': 'Reference Element ID_CED',  # Special handling for ElementID
    'Mark': 'System Number_CEDT',
    'EMS Item Number_CEDT': 'EMS Item Number_CEDT',
    'FLA_CED': 'FLA Input_CED',
    'Voltage_CED': 'Voltage_CED',
    'Number of Poles_CED': 'Number of Poles_CED'
}

type_parameters = ['EMS Item Number_CEDT', 'FLA_CED', 'Voltage_CED', 'Number of Poles_CED']

copy_params = False  # Toggle to enable/disable parameter copying

app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc = revit.doc
output = script.get_output()

def pick_reference_elements():
    """Prompt user to select reference elements if none are selected."""
    selection = revit.get_selection()
    if not selection:
        selection = revit.pick_elements(message="Please select reference elements to map.")
    return selection

def pick_family():
    """Prompt user to pick a family."""
    families = DB.FilteredElementCollector(doc).OfClass(DB.Family).ToElements()

    # Build a list of formatted family names with categories
    family_display_names = [
        (f.FamilyCategory.Name + ": " + f.Name) if f.FamilyCategory else ("Unknown Category: " + f.Name)
        for f in families
    ]

    # Show the selection list with formatted names
    selected_display_name = forms.SelectFromList.show(
        family_display_names,
        title="Select a Family",
        multiselect=False
    )
    if not selected_display_name:
        return None

    # Extract the selected family name (everything after ': ')
    selected_family_name = selected_display_name.split(": ")[-1]

    # Find and return the matching family
    return next(f for f in families if f.Name == selected_family_name)


def pick_family_type(family):
    """Prompt user to pick a type from the selected family."""
    # Retrieve all family types (symbols) from the family
    family_types = [doc.GetElement(i) for i in family.GetFamilySymbolIds()]

    # Use pyRevit's query.get_name to get the name of the family
    family_name = query.get_name(family)

    # Collect family type names for the selection list using query.get_name
    family_type_names = [query.get_name(ft) for ft in family_types]

    # Prompt user to pick a family type
    selected_type_name = forms.SelectFromList.show(
        family_type_names,
        title="Pick Type from {}".format(family_name),
        multiselect=False
    )

    # Return the selected family type
    if not selected_type_name:
        return None
    return next(ft for ft in family_types if query.get_name(ft) == selected_type_name)


def get_parameter_value(element, param_name):
    """Retrieve a parameter value."""
    param = element.LookupParameter(param_name)
    if param:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == DB.StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == DB.StorageType.ElementId:
            return param.AsElementId()
    return None

def set_parameter_value(element, param_name, value):
    """Set parameter value if writable."""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        if param.StorageType == DB.StorageType.String:
            param.Set(str(value))
        elif param.StorageType == DB.StorageType.Double:
            param.Set(value)
        elif param.StorageType == DB.StorageType.Integer:
            param.Set(value)
        elif param.StorageType == DB.StorageType.ElementId:
            param.Set(value)
    else:
        output.print_md("Warning: Parameter **{}** is read-only or not found.".format(param_name))

def gather_parameters(ref_family):
    """Gather parameters from the reference family."""
    instance_params = {param_name: get_parameter_value(ref_family, param)
                       for param, param_name in parameter_mappings.items() if param != 'ElementID'}
    instance_params['Equipment Description'] = str(ref_family.Id.IntegerValue)
    ref_type_element = doc.GetElement(ref_family.GetTypeId())
    type_params = {parameter_mappings[param]: get_parameter_value(ref_type_element, param)
                   for param in type_parameters if param in parameter_mappings} if ref_type_element else {}
    return instance_params, type_params

# Add a variable at the top to toggle rotation behavior
match_rotation = True  # Set to False if you do not want to match rotation



def place_family_instance(ref_family, target_symbol):
    """Place a new family instance accounting for the project base point and optional rotation."""
    # Ensure the target symbol (family type) is active
    if not target_symbol.IsActive:
        target_symbol.Activate()
        doc.Regenerate()  # Regenerate the document after activation

    # Get reference family location
    ref_location = ref_family.Location

    # Attempt to retrieve the level
    ref_level = doc.GetElement(ref_family.LevelId)
    if ref_level is None or ref_level.Id == DB.ElementId.InvalidElementId:
        logger.debug("Reference element has no valid LevelId. Trying fallback parameter.")
        level_param = ref_family.get_Parameter(DB.BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
        if level_param and level_param.AsElementId() != DB.ElementId.InvalidElementId:
            ref_level = doc.GetElement(level_param.AsElementId())
            logger.debug("Using level from INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM: {}".format(ref_level.Name))
        else:
            raise ValueError("Unable to determine the level for the reference family.")
    else:
        logger.debug("Using level from LevelId: {}".format(ref_level.Name))

    # Handle location depending on the type
    if isinstance(ref_location, DB.LocationPoint):
        # Extract the point location
        ref_point = ref_location.Point

        # Account for the project base point offset
        project_transform = doc.ActiveProjectLocation.GetProjectPosition(DB.XYZ(0, 0, 0))
        base_point_offset = project_transform.Elevation  # Elevation in internal units (feet)

        # Adjust height using project base point offset
        ref_height = ref_point.Z + base_point_offset
        new_point = DB.XYZ(ref_point.X, ref_point.Y, ref_height)  # Maintain same X, Y, and correct Z
    else:
        raise ValueError("Reference element location is not a point. Adjust script for other location types.")

    # Create a new family instance at the correct height
    new_instance = doc.Create.NewFamilyInstance(
        new_point,
        target_symbol,
        ref_level,
        DB.Structure.StructuralType.NonStructural
    )

    # Match rotation if enabled and applicable
    if match_rotation and hasattr(ref_location, "Rotation"):
        ref_rotation = ref_location.Rotation  # Rotation is in radians
        axis = DB.Line.CreateBound(new_point, DB.XYZ(new_point.X, new_point.Y, new_point.Z + 1))  # Vertical axis
        new_instance.Location.Rotate(axis, ref_rotation)

    return new_instance


def copy_family_parameters(new_instance, ref_family):
    """Copy parameters from the reference family to the new instance."""
    instance_params, type_params = gather_parameters(ref_family)
    for param_name, value in instance_params.items():
        if value is not None:
            set_parameter_value(new_instance, param_name, value)
    for param_name, value in type_params.items():
        if value is not None:
            set_parameter_value(new_instance, param_name, value)


from pyrevit import forms, script


def main():
    # Get the current selection
    selection = uidoc.Selection.GetElementIds()

    # If no elements are selected, prompt the user to pick elements
    if not selection:
        selection = pick_reference_elements()
        if not selection:  # Check again if no elements were picked
            forms.alert("No elements selected. Exiting script.", title="Error")
            script.exit()

    # Proceed with the rest of the script
    family = pick_family()
    if not family:
        output.print_md("**No family selected. Exiting script.**")
        return

    family_type = pick_family_type(family)
    if not family_type:
        output.print_md("**No family type selected. Exiting script.**")
        return

    with DB.Transaction(doc, "Place New Family Instances") as trans:
        trans.Start()
        for ref_family_id in selection:
            ref_family = doc.GetElement(ref_family_id)
            new_instance = place_family_instance(ref_family, family_type)
            set_parameter_value(new_instance,"Reference Element ID_CED",ref_family_id)
            if copy_params:
                copy_family_parameters(new_instance, ref_family)
        trans.Commit()
    output.print_md("**Script completed successfully.**")


if __name__ == "__main__":
    main()
