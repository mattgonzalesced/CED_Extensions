# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script
from pyrevit.revit import query
import re
from collections import defaultdict

doc = revit.doc
logger = script.get_logger()
uidoc = revit.uidoc
PARAM_NAME = "Refrigeration Circuit Number_CEDT"
GROUP_TYPE_NAME = "Case Power - 1 Case, 3 Ckts"

class ParentElement:
    """Store details about a parent (reference) element."""

    def __init__(self, element_id, location_point, facing_orientation):
        self.element_id = element_id
        self.location_point = location_point
        self.facing_orientation = facing_orientation

    @property
    def circuit_number(self):
        return self.get_parameter_value("Circuit #")

    @classmethod
    def from_family_instance(cls, element):
        if not isinstance(element, DB.FamilyInstance):
            logger.debug("Input is not a FamilyInstance: {}".format(element.Id))
            return None
        loc = element.Location
        if not isinstance(loc, DB.LocationPoint):
            logger.debug("Skipping element without valid LocationPoint: {}".format(element.Id))
            return None
        orientation = element.FacingOrientation if hasattr(element, "FacingOrientation") else None
        return cls(element.Id, loc.Point, orientation)

    def get_parameter_value(self, name):
        param = doc.GetElement(self.element_id).LookupParameter(name)
        if not param or not param.HasValue:
            return None
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString()
        if st == DB.StorageType.Double:
            return param.AsDouble()
        if st == DB.StorageType.Integer:
            return param.AsInteger()
        return None


class ChildGroup:
    DETAIL_NAME = "Case Power Tags - BOX ONLY"

    def __init__(self, parent, group_type):
        self.parent = parent
        self.group_type = group_type
        self.location = parent.location_point
        self.orientation = parent.facing_orientation
        self.child_id = None

    def place(self):
        inst = doc.Create.PlaceGroup(self.location, self.group_type)
        self.child_id = inst.Id
        logger.info("Placed model-group ID {}".format(self.child_id))
        return inst

    def rotate_to_match_parent(self):
        if not self.child_id or not self.orientation:
            return False
        default = DB.XYZ(0, 1, 0)
        angle = default.AngleTo(self.orientation)
        if default.CrossProduct(self.orientation).Z < 0:
            angle = -angle
        axis = DB.Line.CreateBound(
            self.location,
            self.location + DB.XYZ(0, 0, 1)
        )
        try:
            grp = doc.GetElement(self.child_id)
            grp.Location.Rotate(axis, angle)
            logger.info("Rotated model-group ID {} by {}".format(self.child_id, angle))
            return True
        except Exception as e:
            logger.error("Rotation failed: {}".format(e))
            return False

    def copy_parameters(self, mapping):
        grp = doc.GetElement(self.child_id)
        for p_name, c_name in mapping.items():
            val = self.parent.get_parameter_value(p_name)
            if val is not None:
                param = grp.LookupParameter(c_name)
                if param and not param.IsReadOnly:
                    try:
                        st = param.StorageType
                        if st == DB.StorageType.String:
                            param.Set(str(val))
                        elif st == DB.StorageType.Double:
                            param.Set(float(val))
                        elif st == DB.StorageType.Integer:
                            param.Set(int(val))
                        elif st == DB.StorageType.ElementId:
                            param.Set(val)
                    except Exception as e:
                        logger.error("Setting {} failed: {}".format(c_name, e))

    def attach_detail_group_by_type(self, detail_group_type):
        if not self.child_id:
            logger.warning("Model group not placed yet.")
            return False
        try:
            group = doc.GetElement(self.child_id)
            group.ShowAttachedDetailGroups(doc.ActiveView, detail_group_type.Id)
            logger.info("Attached detail group '{}' to group {}".format(query.get_name(detail_group_type), self.child_id))
            return True
        except Exception as e:
            logger.error("Failed to attach detail group '{}': {}".format(query.get_name(detail_group_type), e))
            return False

    @classmethod
    def from_existing_group(cls, group):
        if not isinstance(group, DB.Group):
            logger.error("Provided element is not a Group: {}".format(group.Id))
            return None
        loc = group.Location
        if not isinstance(loc, DB.LocationPoint):
            logger.warning("Group {} has no LocationPoint".format(group.Id))
            return None

        dummy_parent = type('Dummy', (), {
            'location_point': loc.Point,
            'facing_orientation': DB.XYZ(0, 1, 0),
            'circuit_number': group.LookupParameter(PARAM_NAME).AsString() if group.LookupParameter(PARAM_NAME) else None
        })()

        instance = cls(dummy_parent, group.GroupType)
        instance.child_id = group.Id
        return instance

    def ungroup_and_propagate(self):
        group = doc.GetElement(self.child_id)
        if not isinstance(group, DB.Group):
            logger.warning("Element is not a group: {}".format(self.child_id))
            return []

        circuit_number = None
        param = group.LookupParameter(PARAM_NAME)
        if param and param.HasValue:
            circuit_number = param.AsString()

        if not circuit_number:
            logger.warning("Group {} has no circuit number.".format(group.Id))
            return

        try:
            doc.Regenerate()
        except Exception as e:
            logger.error("Failed to show attached detail groups: {}".format(e))
            return

        ungrouped_ids = []

        def is_attached_to_group(x):
            return hasattr(x, "AttachedParentId") and x.AttachedParentId == group.Id

        detail_groups = DB.FilteredElementCollector(doc, doc.ActiveView.Id) \
            .OfCategory(DB.BuiltInCategory.OST_IOSAttachedDetailGroups) \
            .WhereElementIsNotElementType().ToElements()

        for dg in filter(is_attached_to_group, detail_groups):
            try:
                ids = dg.UngroupMembers()
                ungrouped_ids.extend(ids)
                logger.info("Ungrouped detail group {}".format(dg.Id))
            except Exception as e:
                logger.error("Failed to ungroup detail group {}: {}".format(dg.Id, e))

        try:
            model_ids = group.UngroupMembers()
            ungrouped_ids.extend(model_ids)
            logger.info("Ungrouped model group {}".format(group.Id))
        except Exception as e:
            logger.error("Failed to ungroup model group {}: {}".format(group.Id, e))

        for eid in ungrouped_ids:
            el = doc.GetElement(eid)
            if isinstance(el, DB.FamilyInstance):
                if el.Category and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_ElectricalFixtures):
                    param = el.LookupParameter(PARAM_NAME)
                    if param and not param.IsReadOnly:
                        try:
                            param.Set(str(circuit_number))
                            logger.info("Wrote circuit number to fixture ID {}".format(eid))
                        except Exception as e:
                            logger.error("Failed to set circuit number on ID {}: {}".format(eid, e))


    @classmethod
    def collect_target_groups(cls):
        provider = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.SYMBOL_NAME_PARAM))
        evaluator = DB.FilterStringEquals()  # no fourth param
        rule = DB.FilterStringRule(provider, evaluator, GROUP_TYPE_NAME)
        filter_ = DB.ElementParameterFilter(rule)

        return DB.FilteredElementCollector(doc) \
            .OfClass(DB.Group) \
            .WherePasses(filter_) \
            .ToElements()


def collect_reference_tags():
    selected_ids = revit.get_selection().element_ids
    if selected_ids:
        selected_elements = [doc.GetElement(eid) for eid in selected_ids]
        tags = [
            inst for inst in selected_elements
            if isinstance(inst, DB.FamilyInstance)
               and inst.Symbol.Family.Name == "Refrigeration Case Tag - EMS"
               and query.get_name(inst.Symbol) == "EMS Circuit Label"
        ]
        if tags:
            logger.info("Using {} selected EMS tags.".format(len(tags)))
            return tags
        else:
            logger.warning("Selection has no matching EMS tags; falling back to view scan.")

    view_id = doc.ActiveView.Id
    collector = DB.FilteredElementCollector(doc, view_id).OfClass(DB.FamilyInstance)
    tags = [
        inst for inst in collector
        if inst.Symbol.Family.Name == "Refrigeration Case Tag - EMS"
           and query.get_name(inst.Symbol) == "EMS Circuit Label"
    ]
    if not tags:
        logger.error("No matching refrigeration tags found in active view.")
        script.exit()
    return tags


def get_model_group_type(name):
    for gt in DB.FilteredElementCollector(doc).OfClass(DB.GroupType):
        if query.get_name(gt) == name:
            return gt
    logger.error("GroupType not found: {!r}".format(name))
    script.exit()


def get_attached_detail_types(group_type):
    detail_types = []
    for dt_id in group_type.GetAvailableAttachedDetailGroupTypeIds():
        dt = doc.GetElement(dt_id)
        detail_types.append(dt)
    return detail_types


def extract_system_id(circuit_number):
    if not circuit_number:
        return None
    match = re.match(r"^([A-Z]+\d+)", circuit_number)
    if match:
        return match.group(1)
    return None

