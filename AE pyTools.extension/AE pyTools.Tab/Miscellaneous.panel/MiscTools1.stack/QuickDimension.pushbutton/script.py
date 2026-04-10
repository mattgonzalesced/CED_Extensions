# coding: utf8

from Autodesk.Revit import Exceptions
from Autodesk.Revit.DB import (
    Reference,
    ReferenceArray,
    XYZ,
    Line,
    FamilyInstance,
    FamilyInstanceReferenceType,
    Edge, ReferencePlane, DetailCurve, ModelCurve, Grid
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, forms, script

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()


# ------------------------------
# Selection Filters
# ------------------------------

class CustomFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return True

    def AllowReference(self, reference, position):
        return True


class CurveLineFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem.Location.Curve, Line)
        except:
            return False

    def AllowReference(self, reference, position):
        try:
            elem = doc.GetElement(reference.ElementId)
            try:
                return isinstance(elem.Location.Curve, Line)
            except:
                return False
        except:
            return False


class LeadDirectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        # Allow element picks for these types
        try:
            if isinstance(elem, ReferencePlane):
                return True
        except:
            pass

        try:
            if isinstance(elem, Grid):
                return isinstance(elem.Curve, Line)
        except:
            pass

        try:
            if isinstance(elem, DetailCurve) or isinstance(elem, ModelCurve):
                return isinstance(elem.GeometryCurve, Line)
        except:
            pass

        try:
            return isinstance(elem.Location.Curve, Line)
        except:
            return False

    def AllowReference(self, reference, position):
        # With PointOnElement, many clicks arrive here.
        # We must allow references that belong to allowed elements, and linear edges.
        try:
            elem = doc.GetElement(reference.ElementId)
            if elem is None:
                return False

            # If it's one of our allowed element types, allow the reference too
            if self.AllowElement(elem):
                return True

            # Otherwise allow linear edges explicitly
            try:
                geom = elem.GetGeometryObjectFromReference(reference)
                if isinstance(geom, Edge):
                    crv = geom.AsCurve()
                    return isinstance(crv, Line)
            except:
                pass

        except:
            pass

        return False






# ------------------------------
# Selection Helpers
# ------------------------------

def get_selection():
    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    if selection:
        return selection

    try:
        return [
            doc.GetElement(reference)
            for reference in uidoc.Selection.PickObjects(
                ObjectType.Element, CustomFilter(), "Pick"
            )
        ]
    except Exceptions.OperationCanceledException:
        logger.info("User cancelled selection.")
        return None


def get_lead_direction(selection):

    # 1) Infer from current selection first
    for element in selection:
        d = resolve_direction(element, None)
        if d:
            logger.debug("Lead direction inferred from selection element {}".format(element.Id))
            return d

    # 2) Prompt user
    with forms.WarningBar(title="Pick a line, grid, detail line, or reference plane to define lead direction."):
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                LeadDirectionFilter()
            )
        except Exceptions.OperationCanceledException:
            logger.info("User cancelled lead direction pick.")
            return None

    element = doc.GetElement(ref.ElementId)
    if element is None:
        logger.warning("Picked reference had no element.")
        return None

    d = resolve_direction(element, ref)
    if d:
        logger.debug("Lead direction resolved from picked element {}".format(element.Id))
        return d

    logger.warning("Could not determine lead direction from picked object (ElementId: {}).".format(element.Id))
    return None


def resolve_direction(element, ref=None):
    """Return normalized XYZ direction for element or geometry reference, or None."""
    # 1) Grid
    try:
        if isinstance(element, Grid):
            crv = element.Curve
            if isinstance(crv, Line):
                return crv.Direction.Normalize()
    except:
        pass

    # 2) DetailCurve / ModelCurve (both expose GeometryCurve)
    try:
        if isinstance(element, DetailCurve) or isinstance(element, ModelCurve):
            crv = element.GeometryCurve
            if isinstance(crv, Line):
                return crv.Direction.Normalize()
    except:
        pass

    # 3) Location.Curve elements (pipes/ducts/conduit/etc.)
    try:
        crv = element.Location.Curve
        if isinstance(crv, Line):
            return crv.Direction.Normalize()
    except:
        pass

    # 4) ReferencePlane
    try:
        if isinstance(element, ReferencePlane):
            return element.Direction.Normalize()
    except:
        pass

    # 5) Picked Edge (only if ref provided)
    if ref is not None:
        try:
            geom = element.GetGeometryObjectFromReference(ref)
            if isinstance(geom, Edge):
                crv = geom.AsCurve()
                if isinstance(crv, Line):
                    return crv.Direction.Normalize()
        except:
            pass

    return None


def create_dimension_line(lead_direction):
    try:
        pt1 = uidoc.Selection.PickPoint()
    except Exceptions.OperationCanceledException:
        logger.info("User cancelled dimension placement.")
        return None, None

    dim_dir = XYZ(-lead_direction.Y, lead_direction.X, 0).Normalize()
    pt2 = pt1 + dim_dir

    logger.debug("Lead Direction: {}".format(lead_direction))
    logger.debug("Dimension Direction: {}".format(dim_dir))

    return Line.CreateBound(pt1, pt2), dim_dir


# ------------------------------
# Reference Resolution
# ------------------------------

def get_family_reference(fi, dim_dir):

    t = fi.GetTransform()

    axis_map = {
        FamilyInstanceReferenceType.CenterLeftRight: t.BasisX,
        FamilyInstanceReferenceType.Left: t.BasisX,
        FamilyInstanceReferenceType.Right: t.BasisX,

        FamilyInstanceReferenceType.CenterFrontBack: t.BasisY,
        FamilyInstanceReferenceType.Front: t.BasisY,
        FamilyInstanceReferenceType.Back: t.BasisY,

        FamilyInstanceReferenceType.CenterElevation: t.BasisZ,
        FamilyInstanceReferenceType.Top: t.BasisZ,
        FamilyInstanceReferenceType.Bottom: t.BasisZ,
    }

    priority = [
        FamilyInstanceReferenceType.CenterLeftRight,
        FamilyInstanceReferenceType.CenterFrontBack,
        FamilyInstanceReferenceType.CenterElevation,

        FamilyInstanceReferenceType.Left,
        FamilyInstanceReferenceType.Right,
        FamilyInstanceReferenceType.Front,
        FamilyInstanceReferenceType.Back,
        FamilyInstanceReferenceType.Top,
        FamilyInstanceReferenceType.Bottom,
    ]

    dim_dir = dim_dir.Normalize()

    for ref_type in priority:

        refs = fi.GetReferences(ref_type)
        if not (refs and refs.Count > 0):
            continue

        axis = axis_map.get(ref_type)
        if axis is None:
            continue

        axis = axis.Normalize()
        alignment = abs(dim_dir.DotProduct(axis))

        logger.debug(
            "Instance {} | {} alignment: {}".format(fi.Id, ref_type, alignment)
        )

        if alignment > 0.999:
            logger.info(
                "Instance {} using {}".format(fi.Id, ref_type)
            )
            return refs[0]

    logger.warning(
        "Instance {} no aligned stable plane found. Using fallback.".format(fi.Id)
    )

    for ref_type in FamilyInstanceReferenceType.GetValues(
        FamilyInstanceReferenceType
    ):
        refs = fi.GetReferences(ref_type)
        if refs and refs.Count > 0:
            return refs[0]

    return None


def get_reference(element, dim_dir):

    if isinstance(element, FamilyInstance):
        return get_family_reference(element, dim_dir)

    try:
        return Reference(element)
    except:
        logger.warning("Could not create reference for element {}".format(element.Id))
        return None


# ------------------------------
# Main
# ------------------------------

def main():
    output = script.get_output()
    output.close_others()
    selection = get_selection()
    if not selection:
        return

    lead_direction = get_lead_direction(selection)
    if not lead_direction:
        return

    line, dim_dir = create_dimension_line(lead_direction)
    if not line:
        return

    discarded = []
    reference_array = ReferenceArray()

    for element in selection:
        ref = get_reference(element, dim_dir)
        if ref:
            reference_array.Append(ref)
        else:
            discarded.append(element)

    if reference_array.Size < 2:
        forms.alert(
            "Not enough valid references to create a dimension.",
            exitscript=True
        )
        return

    with revit.Transaction("Quick Dimension"):

        dim = doc.Create.NewDimension(
            doc.ActiveView,
            line,
            reference_array
        )

        if dim:
            logger.info("Dimension created: {}".format(dim.Id))
            revit.get_selection().set_to(dim)
        else:
            logger.warning("Dimension creation failed.")

        if discarded:
            output = script.get_output()
            output.print_md("### ⚠ Elements Not Dimensioned")

            for el in discarded:
                output.print_md(
                    "- {}".format(output.linkify(el.Id))
                )


# ------------------------------
main()
