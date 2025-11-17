# -*- coding: utf-8 -*-
from pyrevit import revit, DB, script

output = script.get_output()
doc = revit.doc

PI = 3.141592653589793
RAD_TO_DEG = 180.0 / PI


# ============================================================
# MATH HELPERS
# ============================================================

def signed_angle_deg_from_basis(basis):
    """Signed rotation angle (deg) between Project North X and given basis."""
    if basis is None:
        return 0.0
    ang = DB.XYZ.BasisX.AngleTo(basis)
    cross = DB.XYZ.BasisX.CrossProduct(basis)
    if cross.Z < 0:
        ang = -ang
    return ang * RAD_TO_DEG



def format_xyz(xyz,unit="feet"):
    if unit == "meters":
        return "({:.3f}, {:.3f}, {:.3f})".format(
            xyz.X * 3.28084,
            xyz.Y * 3.28084,
            xyz.Z * 3.28084
        )
    else:
        return "({:.3f}, {:.3f}, {:.3f})".format(
            xyz.X ,
            xyz.Y ,
            xyz.Z
        )

# ============================================================
# COORDINATE SYSTEM POINT SET
# ============================================================

class CoordSystemPointSet(object):
    def __init__(self, internal_xyz, survey_xyz, base_xyz):
        self.internal = internal_xyz
        self.survey = survey_xyz
        self.base = base_xyz

    def transform(self, transform):
        """Return a NEW CoordSystemPointSet mapped by a Transform."""
        return CoordSystemPointSet(
            transform.OfPoint(self.internal),
            transform.OfPoint(self.survey),
            transform.OfPoint(self.base)
        )


# ============================================================
# DOCUMENT COORDINATE FRAME
# ============================================================

class DocumentCoordFrame(object):
    def __init__(self, document):
        self.document = document
        self.title = document.Title

        # Local (Project North) points
        self.internal_local = DB.XYZ(0.0, 0.0, 0.0)
        self.survey_local = self._get_survey_params(document)
        self.base_local = self._get_base_params(document)

        # Active Project Location (True North / Shared coords)
        self.project_location_transform = self._get_project_location_transform(document)
        self.rotation_true_north_deg = signed_angle_deg_from_basis(
            self.project_location_transform.BasisX
        )

    def _get_project_location_transform(self, document):
        try:
            return document.ActiveProjectLocation.GetTransform()
        except:
            return DB.Transform.Identity

    def _get_survey_params(self, document):
        """Shared Base Point element parameters (Survey Point)."""
        try:
            sps = DB.FilteredElementCollector(document) \
                    .OfCategory(DB.BuiltInCategory.OST_SharedBasePoint) \
                    .WhereElementIsNotElementType()
            for sp in sps:
                e = sp.get_Parameter(DB.BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
                n = sp.get_Parameter(DB.BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
                z = sp.get_Parameter(DB.BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
                return DB.XYZ(e, n, z)
        except:
            pass
        return DB.XYZ(0.0, 0.0, 0.0)

    def _get_base_params(self, document):
        """Project Base Point element parameters."""
        try:
            bps = DB.FilteredElementCollector(document) \
                    .OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint) \
                    .WhereElementIsNotElementType()
            for bp in bps:
                e = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
                n = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
                z = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
                return DB.XYZ(e, n, z)
        except:
            pass
        return DB.XYZ(0.0, 0.0, 0.0)

    # ---- public helpers ----

    def get_local_pointset(self):
        """Project North (document internal coordinates)."""
        return CoordSystemPointSet(
            self.internal_local,
            self.survey_local,
            self.base_local
        )

    def get_truenorth_pointset(self):
        """
        True North coordinates: ActiveProjectLocation transform
        applied to the local points.
        """
        return self.get_local_pointset().transform(self.project_location_transform)


# ============================================================
# OUTPUT FOR ACTIVE DOCUMENT
# ============================================================

def print_active_doc_report(frame):
    output.print_md("# Active Document: **{}**".format(frame.title))

    local = frame.get_local_pointset()
    tn = frame.get_truenorth_pointset()

    # Local = Project North frame (doc internal)
    output.print_md("## Project North (Document Local Coordinates)")
    output.print_md("- Internal Origin (PN): {}".format(format_xyz(local.internal)))
    output.print_md("- Survey Point (PN params): {}".format(format_xyz(local.survey)))
    output.print_md("- Base Point (PN params): {}".format(format_xyz(local.base)))
    output.print_md("")

    # True North = ProjectLocation transform applied
    output.print_md("## True North (Project Location Transform Applied)")
    output.print_md("- Internal Origin (TN): {}".format(format_xyz(tn.internal)))
    output.print_md("- Survey Point (TN): {}".format(format_xyz(tn.survey)))
    output.print_md("- Base Point (TN): {}".format(format_xyz(tn.base)))
    output.print_md("- Angle to True North: {:.3f}°".format(frame.rotation_true_north_deg))
    output.print_md("")


# ============================================================
# RUN (ACTIVE DOC ONLY FOR NOW)
# ============================================================

host_frame = DocumentCoordFrame(doc)
print_active_doc_report(host_frame)