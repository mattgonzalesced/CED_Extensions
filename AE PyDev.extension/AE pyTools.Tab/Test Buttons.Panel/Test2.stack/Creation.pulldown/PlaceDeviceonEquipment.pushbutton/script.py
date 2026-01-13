# -*- coding: utf-8 -*-

from pyrevit import DB, script, forms, revit
from pyrevit.revit import query

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

console = script.get_output()
logger = script.get_logger()
class ElementLogger(object):
    def __init__(self, base_logger, element=None, label=None):
        """
        Args:
            base_logger: The pyRevit logger to wrap.
            element: Revit element or ElementId to associate with logs.
            label: Optional label like "Parent" or "Child"
        """
        self.logger = base_logger
        self.element = element
        self.label = label or "Element"

    def _prefix(self):
        if isinstance(self.element, DB.ElementId):
            el_id = self.element.IntegerValue
        elif hasattr(self.element, "Id"):
            el_id = self.element.Id.IntegerValue
        else:
            el_id = "?"

        return "[{} ID={}]".format(self.label, el_id)

    def info(self, message):
        self.logger.info("{} {}".format(self._prefix(), message))

    def warning(self, message):
        self.logger.warning("{} {}".format(self._prefix(), message))

    def debug(self, message):
        self.logger.debug("{} {}".format(self._prefix(), message))

    def error(self, message):
        self.logger.error("{} {}".format(self._prefix(), message))


# 129689 2D
# 1534642    3d
class ParentElement:
    """Class to store details about a parent (reference) element."""

    def __init__(self, element_id, location_point=None, facing_orientation=None, is_view_specific=None):
        self.element_id = element_id
        self.location_point = location_point
        self.facing_orientation = facing_orientation
        self.is_view_specific = is_view_specific

    @property
    def log(self):
        return ElementLogger(logger, self.element_id, label="Parent")

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

    @property
    def element(self):
        """Returns the Revit element object from the stored ID."""
        return doc.GetElement(self.element_id)

    @property
    def symbol(self):
        """Returns the FamilySymbol (type) of the element, if it's a FamilyInstance."""
        if isinstance(self.element, DB.FamilyInstance):
            return self.element.Symbol
        return None

    @property
    def instance_parameters(self):
        """
        Returns a dictionary of parameter name -> DB.Parameter for instance parameters.
        """
        if self.element:
            return {
                param.Definition.Name: param
                for param in self.element.Parameters
            }
        else:
            return {}

    @property
    def type_parameters(self):
        """
        Returns a dictionary of parameter name -> DB.Parameter for type (symbol) parameters.
        """
        if self.symbol:
            return {
                param.Definition.Name: param
                for param in self.symbol.Parameters

            }
        else:
            return {}

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
        Retrieve a value from either an instance parameter or type parameter.

        Args:
            parameter_name (str): The name of the parameter to retrieve.

        Returns:
            The value of the parameter, or None if not found.
        """

        elem = doc.GetElement(self.element_id)

        if elem is None:
            logger.info("[Parent get param]: no element found")
            return None

        # Try instance parameter first
        param = elem.LookupParameter(parameter_name)

        # If not found, check type parameters
        if not param and hasattr(elem, "Symbol"):
            logger.info("[Parent get param]:instance param <{}> not found. trying type".format(parameter_name))
            symbol = elem.Symbol
            if symbol:
                param = symbol.LookupParameter(parameter_name)

        if not param:
            logger.info("[Parent get param]: type param <{}>not found. returning none".format(parameter_name))
            return None

        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == DB.StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == DB.StorageType.ElementId:
            return param.AsElementId()

        return None

    def __repr__(self):
        return "ParentElement(ID={}, Point={}, Orientation={}, ViewSpecific={})".format(
            self.element_id,
            self.location_point,
            self.facing_orientation,
            self.is_view_specific
        )


class ChildElement:
    """Class to place a child Revit element (FamilyInstance or Group) relative to a parent."""

    def __init__(
        self,
        element_type,            # "FamilyInstance" or "Group"
        family_name,
        symbol_name,
        symbol_or_type=None,     # FamilySymbol or GroupType
        parent_element=None,
        placement_info=None,     # PlacementInfo object
        view_specific=False,
        structural_type=None
    ):
        self.element_type = element_type
        self.family_name = family_name
        self.symbol_name = symbol_name
        self.symbol_or_type = symbol_or_type
        self.parent_element = parent_element
        self.placement_info = placement_info
        self.view_specific = view_specific
        self.structural_type = structural_type
        self.child_id = None

    @property
    def log(self):
        return ElementLogger(logger, self.child_id, label="Parent")

    @classmethod
    def from_parent_and_symbol(cls, parent, symbol, family_name, symbol_name, element_type="FamilyInstance"):
        """
        Create a ChildElement from a ParentElement and Revit symbol/type.
        Supports FamilyInstance or Group.
        """
        view_specific = symbol.Category.Id in [
            DB.ElementId(DB.BuiltInCategory.OST_GenericAnnotation),
            DB.ElementId(DB.BuiltInCategory.OST_DetailComponents),
        ] if isinstance(symbol, DB.FamilySymbol) else False

        level_id = parent.level_id if not view_specific else None
        owner_view_id = parent.owner_view_id if view_specific else None

        if not view_specific and level_id is None:
            active_view = doc.ActiveView
            if hasattr(active_view, "GenLevel") and active_view.GenLevel:
                level_id = active_view.GenLevel.Id
            else:
                raise ValueError("Unable to determine a valid level for 3D child placement.")

        placement_info = PlacementInfo(
            location_point=parent.location_point,
            facing_orientation=parent.facing_orientation,
            level_id=level_id,
            owner_view_id=owner_view_id
        )

        return cls(
            element_type=element_type,
            family_name=family_name,
            symbol_name=symbol_name,
            symbol_or_type=symbol,
            parent_element=parent,
            placement_info=placement_info,
            view_specific=view_specific,
            structural_type=DB.Structure.StructuralType.NonStructural if not view_specific else None
        )

    def place(self):
        """Place the element in Revit based on its type."""
        if self.element_type == "FamilyInstance":
            return self._place_family_instance()
        elif self.element_type == "Group":
            return self._place_group()
        else:
            raise ValueError("Unsupported element type: {}".format(self.element_type))

    def _place_family_instance(self):
        if not self.symbol_or_type.IsActive:
            self.symbol_or_type.Activate()
            doc.Regenerate()

        if self.view_specific:
            view = doc.GetElement(self.placement_info.owner_view_id)
            if not view:
                raise ValueError("Invalid view for 2D placement.")
            placed = doc.Create.NewFamilyInstance(
                self.placement_info.location_point, self.symbol_or_type, view
            )
        else:
            level = doc.GetElement(self.placement_info.level_id)
            if not level:
                raise ValueError("Invalid level for 3D placement.")

            # Adjust Z relative to level
            offset_z = self.placement_info.location_point.Z - level.Elevation
            point = DB.XYZ(
                self.placement_info.location_point.X,
                self.placement_info.location_point.Y,
                offset_z
            )
            placed = doc.Create.NewFamilyInstance(
                point, self.symbol_or_type, level, self.structural_type
            )

        self.child_id = placed.Id
        return placed

    def _place_group(self):
        if not isinstance(self.symbol_or_type, DB.GroupType):
            raise TypeError("Expected GroupType for group placement.")
        placed = doc.Create.PlaceGroup(self.placement_info.location_point, self.symbol_or_type)
        self.child_id = placed.Id
        return placed

    def rotate_to_match_parent(self):
        """Rotate the child element to match its parent's facing orientation."""
        if not self.child_id:
            logger.warning("No placed child element to rotate.")
            return False

        child_element = doc.GetElement(self.child_id)
        if child_element is None:
            logger.warning("Child element with ID {} not found.".format(self.child_id))
            return False

        parent_orientation = self.placement_info.facing_orientation
        if not parent_orientation:
            logger.warning("Parent orientation is missing.")
            return False

        default_orientation = DB.XYZ(0, 1, 0)
        angle = default_orientation.AngleTo(parent_orientation)

        cross = default_orientation.CrossProduct(parent_orientation)
        if cross.Z < 0:
            angle = -angle

        axis = DB.Line.CreateBound(
            self.placement_info.location_point,
            DB.XYZ(self.placement_info.location_point.X,
                   self.placement_info.location_point.Y,
                   self.placement_info.location_point.Z + 1)
        )

        try:
            child_element.Location.Rotate(axis, angle)
            logger.info("Rotated element {} by {:.2f} radians.".format(self.child_id, angle))
            return True
        except Exception as e:
            logger.error("Rotation failed: {}".format(e))
            return False

    def copy_parameters(self, parameter_mapping):
        """Copy parameters from the parent element to the child element."""
        if not self.parent_element:
            logger.warning("No parent associated with this child element.")
            return

        for parent_param, child_param in parameter_mapping.items():
            parent_value = self.parent_element.get_parameter_value(parent_param)
            if parent_value is None:
                logger.warning("Parent parameter '{}' not found or has no value.".format(parent_param))
                continue

            if not self.set_parameter_value(child_param, parent_value):
                logger.warning("Failed to set child parameter '{}'.".format(child_param))

    def set_parameter_value(self, parameter_name, value):
        """
        Set a parameter value on the placed child element.
        """
        element = doc.GetElement(self.child_id)
        if not element:
            logger.warning("Child element with ID {} not found.".format(self.child_id))
            return False

        param = element.LookupParameter(parameter_name)
        if not param:
            logger.warning("Child element missing parameter '{}'.".format(parameter_name))
            return False

        if param.IsReadOnly:
            logger.warning("Parameter '{}' is read-only on child.".format(parameter_name))
            return False

        try:
            storage = param.StorageType
            logger.debug("Setting parameter '{}' on child. StorageType: {}, Value: {}".format(
                parameter_name, storage, value
            ))

            if storage == DB.StorageType.String:
                param.Set(str(value))
            elif storage == DB.StorageType.Double:
                param.Set(float(value))
            elif storage == DB.StorageType.Integer:
                param.Set(int(value))
            elif storage == DB.StorageType.ElementId:
                if isinstance(value, DB.ElementId):
                    param.Set(value)
                else:
                    logger.warning(
                        "Value for ElementId parameter '{}' is not a valid ElementId.".format(parameter_name))
                    return False
            else:
                logger.warning("Unhandled StorageType '{}' for parameter '{}'.".format(storage, parameter_name))
                return False

            return True

        except Exception as e:
            logger.error("Failed to set parameter '{}': {}".format(parameter_name, e))
            return False

    def __repr__(self):
        return "ChildElement(Type={}, Family={}, Symbol={}, PlacedID={}, Placement={})".format(
            self.element_type,
            self.family_name,
            self.symbol_name,
            self.child_id,
            self.placement_info
        )


class PlacementInfo(object):
    def __init__(self, location_point, level_id=None, owner_view_id=None, facing_orientation=None):
        self.location_point = location_point
        self.level_id = level_id
        self.owner_view_id = owner_view_id
        self.facing_orientation = facing_orientation

    def __repr__(self):
        return "PlacementInfo(Point={}, LevelID={}, ViewID={}, Orientation={})".format(
            self.location_point, self.level_id, self.owner_view_id, self.facing_orientation
        )

class LeaderData(object):
    def __init__(self, elbow, end):
        self.elbow = elbow
        self.end = end



# ___________________________________________________________________________
# HELPER FUNCTIONS
# ___________________________________________________________________________
def get_selected_generic_annotation_instance():
    """Return the single selected Generic Annotation FamilyInstance, or None."""
    sel = revit.get_selection()
    if not sel or len(sel) != 1:
        return None

    try:
        elem = doc.GetElement(sel[0].Id)
    except Exception:
        return None

    if not isinstance(elem, DB.FamilyInstance):
        return None

    # Must be Generic Annotation + view specific
    try:
        if elem.Category and elem.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericAnnotation):
            return None
    except Exception:
        return None

    if not elem.ViewSpecific:
        return None

    if not elem.Symbol:
        return None

    return elem


def pick_generic_annotation_type(title_text):
    """Pick a Generic Annotation FamilySymbol (type) from the project."""
    symbols = (DB.FilteredElementCollector(doc)
               .OfCategory(DB.BuiltInCategory.OST_GenericAnnotation)
               .OfClass(DB.FamilySymbol)
               .ToElements())

    if not symbols or len(symbols) == 0:
        forms.alert("No Generic Annotation types found in the project.", exitscript=True)

    labels = []
    by_label = {}
    for sym in symbols:
        fam_name = query.get_name(sym.Family)
        sym_name = query.get_name(sym)
        label = "{} | {}".format(fam_name, sym_name)
        labels.append(label)
        by_label[label] = sym

    labels.sort()

    picked = forms.SelectFromList.show(
        labels,
        title=title_text,
        multiselect=False
    )
    if not picked:
        script.exit()

    return by_label[picked]


def collect_generic_annotation_instances_across_all_views(source_symbol_id):
    """
    Collect ALL view-specific Generic Annotation instances (FamilyInstance)
    whose Symbol.Id matches source_symbol_id, across ALL views.
    """
    instances = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_GenericAnnotation)
                 .WhereElementIsNotElementType()
                 .ToElements())

    results = []
    for inst in instances:
        try:
            if not isinstance(inst, DB.FamilyInstance):
                continue
            if not inst.ViewSpecific:
                continue
            if not inst.Symbol:
                continue
            if inst.Symbol.Id != source_symbol_id:
                continue
            if inst.OwnerViewId == DB.ElementId.InvalidElementId:
                continue
            results.append(inst)
        except Exception:
            continue

    return results


def pick_source_and_target_symbols_with_selection_behavior():
    """
    Behavior:
    - If exactly 1 Generic Annotation instance is selected:
        Use its Symbol as SOURCE; prompt for TARGET only.
    - Otherwise:
        Prompt for SOURCE and TARGET.
    Returns:
        (source_symbol, target_symbol)
    """
    selected_inst = get_selected_generic_annotation_instance()

    if selected_inst:
        source_symbol = selected_inst.Symbol
        logger.info("Using selected element as SOURCE: {} | {}".format(
            query.get_name(source_symbol.Family), query.get_name(source_symbol)
        ))

        target_symbol = pick_generic_annotation_type("Pick TARGET Generic Annotation Type (place these)")
        if target_symbol.Id == source_symbol.Id:
            forms.alert("Target type matches the selected source type. Nothing to do.", exitscript=True)

        return source_symbol, target_symbol

    # No valid single selection => pick both
    source_symbol = pick_generic_annotation_type("Pick SOURCE Generic Annotation Type (replace these)")
    target_symbol = pick_generic_annotation_type("Pick TARGET Generic Annotation Type (place these)")

    if target_symbol.Id == source_symbol.Id:
        forms.alert("Source and Target are the same type. Nothing to do.", exitscript=True)

    return source_symbol, target_symbol


def get_annotation_leaders(annotation_instance):
    """
    Extract leader geometry from a Generic Annotation instance.

    Returns:
        list[LeaderData]
    """
    leader_data = []

    try:
        leaders = annotation_instance.GetLeaders()
        if not leaders:
            return leader_data

        for leader in leaders:
            try:
                end = leader.End
                elbow = leader.Elbow
                leader_data.append(LeaderData(elbow,end))
            except Exception:
                continue

    except Exception as ex:
        logger.debug(
            "Failed reading leaders from {}: {}".format(
                annotation_instance.Id.IntegerValue, ex
            )
        )

    return leader_data


def apply_annotation_leaders(target_instance, leader_data_list):
    """
    Apply leaders to a Generic Annotation instance.
    Leaders must be added, regenerated, then modified.
    """
    if not leader_data_list:
        return

    try:
        # Step 1: add leaders (do NOT touch geometry yet)
        for _ in leader_data_list:
            target_instance.addLeader()

        # Step 2: force Revit to actually create them
        doc.Regenerate()

        # Step 3: fetch the newly created leaders
        new_leaders = target_instance.GetLeaders()
        if not new_leaders:
            return

        # Step 4: apply geometry (safe now)
        for leader, ld in zip(new_leaders, leader_data_list):
            try:
                leader.End = ld.end
                if ld.elbow:
                    leader.Elbow = ld.elbow
            except Exception:
                continue

    except Exception as ex:
        logger.warning(
            "Failed applying leaders to {}: {}".format(
                target_instance.Id.IntegerValue, ex
            )
        )

def apply_annotation_leaders_batch(target_to_leaderdata):
    """
    Batch-add and apply leaders to Generic Annotation instances.
    Exactly one regenerate. Correct point ordering.
    """
    if not target_to_leaderdata:
        return

    # 1) Add leaders only
    for target_id, leader_data_list in target_to_leaderdata.items():
        if not leader_data_list:
            continue

        target = doc.GetElement(target_id)
        if not target:
            continue

        for _ in leader_data_list:
            target.addLeader()

    # 2) Single regenerate
    doc.Regenerate()

    # 3) Apply geometry (ORDER MATTERS)
    for target_id, leader_data_list in target_to_leaderdata.items():
        if not leader_data_list:
            continue

        target = doc.GetElement(target_id)
        if not target:
            continue

        leaders = target.GetLeaders()
        if not leaders:
            continue

        for leader, ld in zip(leaders, leader_data_list):
            try:
                if ld.elbow:
                    # Bent leader: elbow FIRST, then end
                    leader.End = ld.end
                    leader.Elbow = ld.elbow
                else:
                    # Straight leader
                    leader.End = ld.end
            except Exception:
                continue



def replace_generic_annotations_across_all_views(parameter_mapping, delete_source=True):
    """
    Replaces ALL instances of a source Generic Annotation TYPE across ALL views
    with a target Generic Annotation TYPE, in their corresponding views,
    copying parameters using existing ChildElement.copy_parameters().

    Source/target picking behavior:
    - If exactly 1 Generic Annotation instance is selected: SOURCE = that instance's type, pick TARGET
    - Else: pick SOURCE and TARGET
    """
    source_symbol, target_symbol = pick_source_and_target_symbols_with_selection_behavior()

    source_instances = collect_generic_annotation_instances_across_all_views(source_symbol.Id)

    if not source_instances:
        forms.alert("No instances of the selected SOURCE type were found in any view.", exitscript=True)

    logger.info("Found {} source annotation(s) across all views.".format(len(source_instances)))

    placed_count = 0
    deleted_count = 0
    failed = []

    family_name = query.get_name(target_symbol.Family)
    symbol_name = query.get_name(target_symbol)

    target_leader_map = {}  # ElementId -> [LeaderData]


    tg = DB.TransactionGroup(doc, "Replace Generic Annotations Across All Views")
    tg.Start()

    # ----------------------------------------------------------------------
    # Transaction 1: Place + copy parameters + delete source
    # ----------------------------------------------------------------------
    t1 = DB.Transaction(doc, "Place Replacement Annotations")
    t1.Start()

    for src in source_instances:
        try:
            parent = ParentElement.from_family_instance(src)
            leader_data = get_annotation_leaders(src)

            if not parent:
                failed.append(src.Id)
                continue

            child = ChildElement.from_parent_and_symbol(
                parent,
                symbol=target_symbol,
                family_name=family_name,
                symbol_name=symbol_name,
                element_type="FamilyInstance"
            )

            child.place()
            child.copy_parameters(parameter_mapping)

            if leader_data:
                target_leader_map[child.child_id] = leader_data
            placed_count += 1

            if delete_source:
                doc.Delete(src.Id)
                deleted_count += 1

        except Exception as ex:
            logger.error(
                "Failed replacing annotation {}: {}".format(
                    src.Id.IntegerValue, ex
                )
            )
            failed.append(src.Id)

    t1.Commit()

    # ----------------------------------------------------------------------
    # Transaction 2: Reapply leaders
    # ----------------------------------------------------------------------
    if target_leader_map:
        t2 = DB.Transaction(doc, "Reapply Annotation Leaders")
        t2.Start()

        apply_annotation_leaders_batch(target_leader_map)

        t2.Commit()

    tg.Assimilate()

    # ----------------------------------------------------------------------
    # Logging
    # ----------------------------------------------------------------------
    logger.info("Replacement complete.")
    logger.info(
        "Placed: {} | Deleted: {} | Failed: {}".format(
            placed_count, deleted_count, len(failed)
        )
    )

    if failed:
        logger.warning("Failed ElementIds:")
        for fid in failed:
            logger.warning(" - {}".format(fid.IntegerValue))


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


def pick_family_or_group():
    """
    Prompt user to pick a Family or GroupType, grouped by category.
    Returns:
        A DB.Family or DB.GroupType object.
    """
    # Collect families
    fam_collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
    group_collector = DB.FilteredElementCollector(doc).OfClass(DB.GroupType)

    logger.debug("Total families: {}, Total group types: {}".format(
        fam_collector.GetElementCount(), group_collector.GetElementCount()))

    fam_options = {" All Families": []}
    group_options = {}

    # -- Handle Families --
    for fam in fam_collector:
        fam_cat = fam.FamilyCategory

        if not fam_cat or fam_cat.IsTagCategory:
            continue
        if fam_cat.Id.IntegerValue in [
            int(DB.BuiltInCategory.OST_MultiCategoryTags),
            int(DB.BuiltInCategory.OST_KeynoteTags)
        ]:
            continue

        fam_options[" All Families"].append(fam)
        cat_name = fam_cat.Name
        if cat_name not in fam_options:
            fam_options[cat_name] = []
        fam_options[cat_name].append(fam)

    # -- Handle GroupTypes --
    for group_type in group_collector:
        if "Model Groups" not in group_options:
            group_options["Model Groups"] = []
        group_options["Model Groups"].append(group_type)

    # Combine families and groups into one UI structure
    grouped_options = {}

    for group, fams in fam_options.items():
        grouped_options[group] = ["[Family] {} | {}".format(f.FamilyCategory.Name, f.Name) for f in fams]

    for group, groups in group_options.items():
        grouped_options[group] = [
            "[Group] {}".format(DB.Element.Name.__get__(g)) for g in groups
        ]

    # Sort entries for each group
    for key in grouped_options:
        grouped_options[key].sort()

    # Show UI
    selected = forms.SelectFromList.show(
        grouped_options,
        title="Select a Family or Group",
        group_selector_title="Category:",
        multiselect=False
    )

    if not selected:
        logger.info("No selection made. Exiting.")
        script.exit()

    # Match the selection back to the element
    for fams in fam_options.values():
        for f in fams:
            label = "[Family] {} | {}".format(f.FamilyCategory.Name, f.Name)
            if selected == label:
                return f

    for groups in group_options.values():
        for g in groups:
            label = "[Group] {}".format(DB.Element.Name.__get__(g))
            if selected == label:
                return g

    logger.error("Selected item not matched to any element.")
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


def inspect_parent_parameters():
    """Prompt user to select a parent and print its instance and type parameters with values."""
    selected_elements = pick_reference_elements()
    if not selected_elements:
        logger.warning("No element selected.")
        return
    for el in selected_elements:
        parent = ParentElement.from_element_id(el.Id)
        if not parent:
            logger.error("Could not create ParentElement.")
            return

        logger.info("Inspecting ParentElement: {}".format(repr(parent)))

        logger.info("---- Instance Parameters ----")
        for name, param in parent.instance_parameters.items():
            value = parent.get_parameter_value(name)
            logger.info("[Instance] {} = {}".format(name, value))

        logger.info("---- Type Parameters ----")
        for name, param in parent.type_parameters.items():
            if name in parent.instance_parameters:
                continue  # already logged
            value = parent.get_parameter_value(name)
            logger.info("[Type] {} = {}".format(name, value))

def main_replace_across_all_views():
    parameter_mapping = {
        "CED-G-NOTE #": "Keynote Value",
        "CED-G-NOTE TEXT": "Keynote Description",
        "CED-G-SHEET #/DISCIPLINE": "Keynote Category",

    }

    replace_generic_annotations_across_all_views(parameter_mapping, delete_source=True)


def main():
    parameter_mapping = {
        "Note Number": "Symbol Label_CEDT",
        "Family and Type": "Equipment Remarks_CEDT",
        "Voltage_CED":"Voltage_CED",
        "Number of Poles_CED":"Number of Poles_CED",
        "FLA_CED":"FLA Input_CED",
    }

    # Select parents
    selected_elements = pick_reference_elements()
    parent_instances = [ParentElement.from_element_id(el.Id) for el in selected_elements]



    # Prompt user to select Family or Group
    selected_child = pick_family_or_group()

    child_instances = []

    if isinstance(selected_child, DB.Family):
        family_type = pick_family_type(selected_child)
        family_symbol = family_type[0]
        family_name = query.get_name(family_symbol.Family)
        symbol_name = query.get_name(family_symbol)

        for parent in parent_instances:
            child = ChildElement.from_parent_and_symbol(
                parent,
                symbol=family_symbol,
                family_name=family_name,
                symbol_name=symbol_name,
                element_type="FamilyInstance"
            )
            child_instances.append(child)

    elif isinstance(selected_child, DB.GroupType):
        group_type = selected_child
        group_name = query.get_name(group_type)

        for parent in parent_instances:
            child = ChildElement.from_parent_and_symbol(
                parent,
                symbol=group_type,
                family_name="Model Group",
                symbol_name=group_name,
                element_type="Group"
            )
            child_instances.append(child)

    else:
        logger.error("Unsupported selection type: {}".format(type(selected_child)))
        script.exit()

    # Place elements
    with DB.Transaction(doc, "Place and Rotate Child Elements") as trans:
        trans.Start()
        for child in child_instances:
            child.place()
            if child.element_type == "FamilyInstance":
                child.rotate_to_match_parent()
                child.copy_parameters(parameter_mapping)
        trans.Commit()

    # Log results
    logger.info("Parent Elements:")
    for parent in parent_instances:
        logger.info("{}".format(repr(parent)))

    logger.info("Child Elements:")
    for child in child_instances:
        logger.info("{}".format(repr(child)))




if __name__ == "__main__":
    main_replace_across_all_views()