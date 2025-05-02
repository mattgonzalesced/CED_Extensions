# -*- coding: utf-8 -*-
from pyrevit import DB, revit, script
from pyrevit.revit import query

doc    = revit.doc
logger = script.get_logger()

class ParentElement:
    """Store details about a parent (reference) element."""
    def __init__(self, element_id, location_point, facing_orientation):
        self.element_id         = element_id
        self.location_point     = location_point
        self.facing_orientation = facing_orientation

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


def collect_reference_tags():
    """Collect all Refrigeration Case Tag – EMS / EMS Circuit Label in active view."""
    view_id   = doc.ActiveView.Id
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
    """
    Finds a loaded model GroupType by its exact name, using query.get_name to avoid .Name errors.
    """
    for gt in DB.FilteredElementCollector(doc).OfClass(DB.GroupType):
        if query.get_name(gt) == name:
            return gt
    logger.error("GroupType not found: {!r}".format(name))
    script.exit()


class ChildGroup:
    """Place, rotate, copy parameters & attach a specific detail group to a Group instance."""
    DETAIL_NAME = "Case Power Tags - 1 Case, No Tags"

    def __init__(self, parent, group_type):
        self.parent      = parent
        self.group_type  = group_type
        self.location    = parent.location_point
        self.orientation = parent.facing_orientation
        self.child_id    = None

    def place(self):
        inst = doc.Create.PlaceGroup(self.location, self.group_type)
        self.child_id = inst.Id
        logger.info("Placed model-group ID {}".format(self.child_id))
        return inst

    def rotate_to_match_parent(self):
        if not self.child_id or not self.orientation:
            return False
        default = DB.XYZ(0, 1, 0)
        angle   = default.AngleTo(self.orientation)
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

    def attach_detail_group(self):
        """
        Attach only the specified detail-group type to this group in the active view.
        """
        view = doc.ActiveView
        inst = doc.GetElement(self.child_id)
        for dt_id in self.group_type.GetAvailableAttachedDetailGroupTypeIds():
            dt = doc.GetElement(dt_id)
            if query.get_name(dt) == ChildGroup.DETAIL_NAME:
                try:
                    inst.ShowAttachedDetailGroups(view, dt_id)
                    logger.info("Attached detail-group '{}' to model-group {}".format(ChildGroup.DETAIL_NAME, self.child_id))
                    return True
                except Exception as e:
                    logger.error("Failed to attach detail-group '{}': {}".format(ChildGroup.DETAIL_NAME, e))
                    return False
        logger.error("Detail-group '{}' not available for this model-group".format(ChildGroup.DETAIL_NAME))
        return False


def main():
    parameter_mapping = {
        "Circuit #": "Refrigeration Circuit Number_CEDT"
    }

    tags       = collect_reference_tags()
    parents    = [ParentElement.from_family_instance(t) for t in tags]
    model_type = get_model_group_type("Case Power - 1 Case, 3 Ckts")  # use the model group type
    children   = [ChildGroup(p, model_type) for p in parents if p]

    with DB.Transaction(doc, "Place Case Power Groups & Attach Details") as trans:
        trans.Start()
        for c in children:
            c.place()
            c.rotate_to_match_parent()
            c.copy_parameters(parameter_mapping)
            c.attach_detail_group()
        trans.Commit()

    logger.info("Placed {} model-groups (with '{}' detail).".format(len(children), ChildGroup.DETAIL_NAME))


if __name__ == "__main__":
    main()
