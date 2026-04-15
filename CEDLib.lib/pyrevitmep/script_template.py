# coding: utf8
from pyrevit import script

__title__ = "Title"
__author__ = "Cyril Waechter"
__doc__ = "Description"

def get_active_doc():
    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    return getattr(uidoc, "Document", None)

logger = script.get_logger()
