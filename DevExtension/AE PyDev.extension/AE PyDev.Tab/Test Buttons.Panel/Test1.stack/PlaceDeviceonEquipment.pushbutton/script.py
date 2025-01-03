# -*- coding: utf-8 -*-
__title__ = "Place Device on Equipment"

from pyrevit import DB, script, forms, revit, output
from pyrevit.revit import query
import clr
from System.Collections.Generic import List

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

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
output = script.get_output()


def pick_reference_elements():
    """Prompt user to select reference elements if none are selected."""
    selection = revit.get_selection()
    if not selection:
        selection = revit.pick_elements(message="Please select reference elements to map.")
    valid_selection = [
        el for el in selection
        if isinstance(doc.GetElement(el.Id), DB.FamilyInstance)  # Ensure only FamilyInstance elements are selected
    ]
    if not valid_selection:
        logger.error("No valid family instances selected. Exiting.")
        script.exit()
    return valid_selection

def pick_family():
    """Prompt user to pick a family grouped by FamilyCategory."""
    # Collect all families in the document
    fam_collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
    logger.debug("Total families in document: {}".format(fam_collector.GetElementCount()))

    fam_options = {" All": []}  # " All" with a space to ensure it appears first

    for fam in fam_collector:
        fam_category = fam.FamilyCategory

        if not fam_category:
            logger.debug("Skipped family with no category: {}".format(fam.Name))
            continue

        if fam_category.IsTagCategory:
            logger.debug("Skipped tag family: {}".format(fam.Name))
            continue

        fam_name = fam.Name
        fam_cat_name = fam_category.Name

        # Add family to the " All" group
        fam_options[" All"].append(fam)

        # Add family to its category group
        if fam_cat_name not in fam_options:
            fam_options[fam_cat_name] = []
        fam_options[fam_cat_name].append(fam)

        logger.debug("Added family: {} to category: {}".format(fam_name, fam_cat_name))

    grouped_options = {group: [] for group in fam_options}
    for group, families in fam_options.items():
        for fam in families:
            option_text = "{} | {}".format(fam.FamilyCategory.Name, fam.Name)
            grouped_options[group].append(option_text)

    logger.debug("Grouped Options for Selection: {}".format(grouped_options))

    for key in grouped_options:
        grouped_options[key].sort()

    selected_option = forms.SelectFromList.show(
        grouped_options,
        title="Select a Family",
        group_selector_title="Category:",
        multiselect=False
    )

    if not selected_option:
        logger.info("No family selected. Exiting script.")
        return None

    for group, families in fam_options.items():
        for fam in families:
            if "{} | {}".format(fam.FamilyCategory.Name, fam.Name) == selected_option:
                return fam

    return None


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


def get_reference_level_or_view(ref_family):
    """Determine the level associated with the reference family, defaulting to active view level if needed."""
    if hasattr(ref_family, 'LevelId') and ref_family.LevelId != DB.ElementId.InvalidElementId:
        ref_level = doc.GetElement(ref_family.LevelId)
        if ref_level:
            logger.debug("Using Level: {}".format(ref_level.Name))
            return ref_level

    active_view = doc.ActiveView
    if hasattr(active_view, 'GenLevel') and active_view.GenLevel:
        ref_level = doc.GetElement(active_view.GenLevel.Id)
        logger.debug("Fallback to Active View Level: {}".format(ref_level.Name))
        return ref_level

    raise ValueError("Unable to determine a valid level for placement.")

def place_3d_family_instance(ref_point, target_symbol, ref_level):
    """Place a 3D family instance."""
    new_instance = doc.Create.NewFamilyInstance(
        ref_point,
        target_symbol,
        ref_level,
        DB.Structure.StructuralType.NonStructural
    )
    return new_instance

def place_2d_family_instance(ref_point, target_symbol, owner_view):
    """Place a 2D family instance in the specified view."""
    new_instance = doc.Create.NewFamilyInstance(
        ref_point,
        target_symbol,
        owner_view
    )
    return new_instance

def place_family_instance(ref_family, target_symbol):
    """Place a new family instance, accounting for 2D and 3D differences."""
    # Ensure the target symbol (family type) is active
    if not target_symbol.IsActive:
        target_symbol.Activate()
        doc.Regenerate()

    # Get reference family location
    ref_location = ref_family.Location
    if not isinstance(ref_location, DB.LocationPoint):
        raise ValueError("Reference family location is not a point. Adjust script for other location types.")

    ref_point = ref_location.Point

    # Always get a level for placement, even for 2D elements
    associated_level_or_view = get_reference_level_or_view(ref_family)

    if ref_family.ViewSpecific:
        logger.debug("Placing a 2D element.")
        owner_view = doc.GetElement(ref_family.OwnerViewId)
        return place_2d_family_instance(ref_point, target_symbol, owner_view)

    elif isinstance(associated_level_or_view, DB.Level):
        logger.debug("Placing a 3D element.")
        return place_3d_family_instance(ref_point, target_symbol, associated_level_or_view)

    else:
        # Fallback to active view's level for placement
        active_view = doc.ActiveView
        if hasattr(active_view, 'GenLevel') and active_view.GenLevel:
            fallback_level = doc.GetElement(active_view.GenLevel.Id)
            logger.debug("Fallback to Active View Level: {}".format(fallback_level.Name))
            return place_3d_family_instance(ref_point, target_symbol, fallback_level)

        raise ValueError("Unexpected case for level or view determination.")

def copy_family_parameters(new_instance, ref_family):
    """Copy parameters from the reference family to the new instance."""
    instance_params, type_params = gather_parameters(ref_family)
    for param_name, value in instance_params.items():
        if value is not None:
            set_parameter_value(new_instance, param_name, value)
    for param_name, value in type_params.items():
        if value is not None:
            set_parameter_value(new_instance, param_name, value)

def main():
    # Get the current selection
    selection = pick_reference_elements()

    # Proceed with the rest of the script
    family = pick_family()
    if not family:
        logger.info("No family selected. Exiting script.")
        return

    family_type = pick_family_type(family)
    if not family_type:
        logger.info("No family type selected. Exiting script.")
        return

    with DB.Transaction(doc, "Place New Family Instances") as trans:
        trans.Start()
        for ref_family in selection:
            ref_family_id = doc.GetElement(ref_family.Id)
            new_instance = place_family_instance(ref_family, family_type)
            set_parameter_value(new_instance, "Reference Element ID_CED", ref_family_id)
            if copy_params:
                copy_family_parameters(new_instance, ref_family)
        trans.Commit()
    output.print_md("**Script completed successfully.**")

if __name__ == "__main__":
    main()
