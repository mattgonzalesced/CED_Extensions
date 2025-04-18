# -*- coding: utf-8 -*-
__title__ = "Place Device on Equipment 2"

from pyrevit import DB, script, forms, revit, output
from pyrevit.revit import query
import clr
from System.Collections.Generic import List

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

console = script.get_output()
logger = script.get_logger()

#129689 2D
#1534642    3d
class ParentElement:
    """Class to store details about a parent (reference) element."""

    def __init__(self, element_id, location_point=None, facing_orientation=None, is_view_specific=None):
        self.element_id = element_id
        self.location_point = location_point
        self.facing_orientation = facing_orientation
        self.is_view_specific = is_view_specific

    @property
    def owner_view_id(self):
        """
        Get the OwnerViewId if the parent is view-specific (2D).
        Returns None if not view-specific.
        """
        if self.is_view_specific:
            element = doc.GetElement(self.element_id)
            return element.OwnerViewId
        return None

    @property
    def level_id(self):
        """
        Get the LevelId if the parent is not view-specific (3D).
        Returns None if view-specific.
        """
        if not self.is_view_specific:
            element = doc.GetElement(self.element_id)
            if hasattr(element, 'LevelId') and element.LevelId != DB.ElementId.InvalidElementId:
                return element.LevelId
        return None

    @classmethod
    def from_element_id(cls, element_id):
        """
        Create a ParentElement instance from an ElementId.

        Args:
            element_id: A Revit ElementId.

        Returns:
            A ParentElement instance or None if the element is not valid.
        """
        element = doc.GetElement(element_id)
        if element:
            return cls.from_family_instance(element)
        logger.debug("No element found for ElementId: {}".format(element_id))
        return None

    @classmethod
    def from_family_instance(cls, element):
        """
        Create a ParentElement instance from a FamilyInstance.

        Args:
            element: A Revit FamilyInstance object.

        Returns:
            A ParentElement instance or None if the element is not valid.
        """
        if not isinstance(element, DB.FamilyInstance):
            logger.debug("Input is not a FamilyInstance: {}".format(element.Id))
            return None

        # Get element details
        location = element.Location
        if not isinstance(location, DB.LocationPoint):
            logger.debug("Skipping element without valid LocationPoint: {}".format(element.Id))
            return None

        location_point = location.Point
        facing_orientation = element.FacingOrientation if hasattr(element, "FacingOrientation") else None
        is_view_specific = element.ViewSpecific

        return cls(
            element_id=element.Id,
            location_point=location_point,
            facing_orientation=facing_orientation,
            is_view_specific=is_view_specific
        )

    def get_parameter_value(self, parameter_name):
        """
        Retrieve the value of a parameter from the parent element.

        Args:
            parameter_name (str): The name of the parameter to retrieve.

        Returns:
            The value of the parameter, or None if not found.
        """
        # Use the Revit API to fetch the parameter value
        param = doc.GetElement(self.element_id).LookupParameter(parameter_name)
        if param and param.HasValue:
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
        return None

    def __repr__(self):
        return "ParentElement(ID={}, Point={}, Orientation={}, ViewSpecific={})".format(
            self.element_id,
            self.location_point,
            self.facing_orientation,
            self.is_view_specific
        )


class ChildElement:
    """Class to store details about a child (to-be-placed) element."""

    def __init__(
        self,
        family_name,
        symbol_name,
        family_symbol=None,
        parent_element=None,
        view_specific=None,
        location_point=None,
        facing_orientation=None,
        level_id=None,
        owner_view_id=None,
        structural_type=None,
    ):
        self.family_name = family_name
        self.symbol_name = symbol_name
        self.family_symbol = family_symbol
        self.parent_element = parent_element
        self.view_specific = view_specific  # Determined from the FamilySymbol
        self.location_point = location_point
        self.facing_orientation = facing_orientation
        self.level_id = level_id  # Inherited or active view's level
        self.owner_view_id = owner_view_id
        self.structural_type = structural_type
        self.child_id = None  # To store the ID of the placed instance

    @classmethod
    def from_parent_and_symbol(cls, parent, family_symbol, family_name, symbol_name):
        """
        Create a ChildElement instance using data from a ParentElement and a FamilySymbol.

        Args:
            parent: A ParentElement instance.
            family_symbol: A Revit FamilySymbol object.
            family_name: The name of the family.
            symbol_name: The name of the symbol.

        Returns:
            A ChildElement instance.
        """
        # Determine if the FamilySymbol is view-specific
        view_specific = family_symbol.Category.Id in [
            DB.ElementId(DB.BuiltInCategory.OST_GenericAnnotation),
            DB.ElementId(DB.BuiltInCategory.OST_DetailComponents),
        ]

        # Determine the Level ID
        if view_specific:
            level_id = None
        else:
            level_id = parent.level_id

        if level_id is None and not view_specific:
            # Fallback to the level of the active view if the parent lacks a LevelID
            active_view = doc.ActiveView
            if hasattr(active_view, "GenLevel") and active_view.GenLevel:
                level_id = active_view.GenLevel.Id
            else:
                raise ValueError("Unable to determine a valid level for 3D child placement.")

        # For view-specific elements, use the active view ID
        owner_view_id = doc.ActiveView.Id if view_specific else None

        return cls(
            family_name=family_name,
            symbol_name=symbol_name,
            family_symbol=family_symbol,
            parent_element=parent,
            view_specific=view_specific,
            location_point=parent.location_point,
            facing_orientation=parent.facing_orientation,
            level_id=level_id,
            owner_view_id=owner_view_id,
            structural_type=DB.Structure.StructuralType.NonStructural
            if not view_specific
            else None,
        )

    def place(self):
        """
        Place the child element in the Revit model based on its properties.
        """
        try:
            if not self.family_symbol.IsActive:
                self.family_symbol.Activate()
                doc.Regenerate()

            if self.view_specific:
                # Place as a 2D element
                owner_view = doc.GetElement(self.owner_view_id)
                if owner_view is None:
                    raise ValueError("OwnerViewId is invalid for placing 2D elements.")
                placed_element = doc.Create.NewFamilyInstance(
                    self.location_point, self.family_symbol, owner_view
                )
            else:
                # Place as a 3D element
                level = doc.GetElement(self.level_id)
                if level is None:
                    raise ValueError("LevelId is invalid for placing 3D elements.")
                placed_element = doc.Create.NewFamilyInstance(
                    self.location_point, self.family_symbol, level, self.structural_type
                )

            if placed_element:
                self.child_id = placed_element.Id  # Store the placed element ID
                logger.info("Successfully placed element with ID: {}".format(self.child_id))
            return placed_element
        except Exception as e:
            logger.error("Failed to place element: {}".format(e))
            raise

    def copy_parameters(self, parameter_mapping):
        """
        Copy parameter values from the associated parent to the child.

        Args:
            parameter_mapping (dict): A dictionary mapping parent parameters to child parameters.
        """
        if not self.parent_element:
            logger.warning("No parent associated with this child element.")
            return

        for parent_param, child_param in parameter_mapping.items():
            # Get the parameter value from the parent
            parent_value = self.parent_element.get_parameter_value(parent_param)
            if parent_value is None:
                logger.warning("Parent parameter '{}' not found or has no value.".format(parent_param))
                continue

            # Set the parameter value on the child
            if not self.set_parameter_value(child_param, parent_value):
                logger.warning("Child parameter '{}' not found or could not be set.".format(child_param))

    def set_parameter_value(self, parameter_name, value):
        """
        Set a parameter value on the placed child element.

        Args:
            parameter_name (str): The name of the parameter to set.
            value: The value to set for the parameter.
        """
        element = doc.GetElement(self.child_id)  # Retrieve the placed element
        if element is None:
            logger.warning("Child element with ID {} not found.".format(self.child_id))
            return False

        param = element.LookupParameter(parameter_name)
        if param and not param.IsReadOnly:
            try:
                if param.StorageType == DB.StorageType.String:
                    param.Set(str(value))
                elif param.StorageType == DB.StorageType.Double:
                    param.Set(float(value))
                elif param.StorageType == DB.StorageType.Integer:
                    param.Set(int(value))
                elif param.StorageType == DB.StorageType.ElementId:
                    param.Set(value)
                return True
            except Exception as e:
                logger.error("Failed to set parameter '{}': {}".format(parameter_name, e))
        else:
            logger.warning("Parameter '{}' is read-only or not found on child.".format(parameter_name))
        return False

    def __repr__(self):
        return "ChildElement(Family={}, Symbol={}, ParentID={}, ViewSpecific={}, Point={}, Orientation={}, LevelID={}, OwnerViewID={}, StructuralType={}, PlacedElementID={})".format(
            self.family_name,
            self.symbol_name,
            self.parent_element,
            self.view_specific,
            self.location_point,
            self.facing_orientation,
            self.level_id,
            self.owner_view_id,
            self.structural_type,
            self.child_id,
        )



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
    """
    Prompt user to pick a family grouped by FamilyCategory.
    Returns:
        The selected Revit Family object.
    """
    # Collect all families in the document
    fam_collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
    logger.debug("Total families in document: {}".format(fam_collector.GetElementCount()))

    fam_options = {" All": []}  # " All" with a space to ensure it appears first

    for fam in fam_collector:
        fam_category = fam.FamilyCategory

        # Skip families without a category or those classified as tags
        if not fam_category or fam_category.IsTagCategory:
            logger.debug("Skipped family: {}, Category: {}".format(
                fam.Name, fam_category.Name if fam_category else "None"
            ))
            continue

        # Add family to the "All" group
        fam_options[" All"].append(fam)

        # Add family to its category group
        fam_cat_name = fam_category.Name
        if fam_cat_name not in fam_options:
            fam_options[fam_cat_name] = []
        fam_options[fam_cat_name].append(fam)

    # Prepare grouped options for display
    grouped_options = {}
    for group, families in fam_options.items():
        grouped_options[group] = [
            "{} | {}".format(fam.FamilyCategory.Name, fam.Name) for fam in families
        ]
        grouped_options[group].sort()

    logger.debug("Grouped Options for Selection: {}".format(grouped_options))

    # Show selection dialog
    selected_option = forms.SelectFromList.show(
        grouped_options,
        title="Select a Family",
        group_selector_title="Category:",
        multiselect=False
    )

    if not selected_option:
        logger.info("No family selected. Exiting script.")
        script.exit()

    # Find and return the selected family
    for group, families in fam_options.items():
        for fam in families:
            if "{} | {}".format(fam.FamilyCategory.Name, fam.Name) == selected_option:
                logger.debug("Selected Family: {}, Category: {}".format(fam.Name, fam.FamilyCategory.Name))
                return fam

    logger.error("Failed to match the selected option to a family.")
    script.exit()




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

    # Return the selected family type and its name
    if not selected_type_name:
        script.exit()

    selected_type = next(ft for ft in family_types if query.get_name(ft) == selected_type_name)
    return selected_type, selected_type_name

def main():
    # Define parameter mapping
    parameter_mapping = {
        "Mark": "Equipment ID_CEDT",
        "Family and Type": "Equipment Remarks_CEDT",
    }

    # Pick reference elements
    selected_elements = pick_reference_elements()
    parent_instances = [ParentElement.from_element_id(el.Id) for el in selected_elements]

    # Prompt user to pick a family and type
    family = pick_family()
    family_type = pick_family_type(family)

    # Use the FamilySymbol from pick_family_type
    family_symbol = family_type[0]
    family_name = query.get_name(family_symbol.Family)
    symbol_name = query.get_name(family_symbol)

    # Create ChildElement instances for each ParentElement
    child_instances = []
    for parent in parent_instances:
        # Create a child element and associate it with its parent
        child = ChildElement.from_parent_and_symbol(parent, family_symbol, family_name, symbol_name)
        child_instances.append(child)

    # Place and set parameters in a transaction
    with DB.Transaction(doc, "Place and Set Parameters") as trans:
        trans.Start()
        for child in child_instances:
            # Place the child element
            placed_instance = child.place()

            # Copy parameters from parent to child
            child.copy_parameters(parameter_mapping)
        trans.Commit()

    # Print results
    logger.info("Parent Elements:")
    for parent in parent_instances:
        logger.info("{}".format(repr(parent)))

    logger.info("Child Elements:")
    for child in child_instances:
        logger.info("{}".format(repr(child)))

if __name__ == "__main__":
    main()
