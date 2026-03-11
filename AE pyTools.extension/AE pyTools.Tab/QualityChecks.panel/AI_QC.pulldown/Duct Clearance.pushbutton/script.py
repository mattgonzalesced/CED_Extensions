# -*- coding: utf-8 -*-
__title__ = "Duct\nClearance"
__doc__ = "Duct clearances vs structure, duct, pipe, cable tray."

import os
import sys
from pyrevit import forms, revit

def _lib_root():
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.normpath(os.path.join(here, "..", "..", "..", "..", ".."))
    return os.path.join(root, "CEDLib.lib")

def main():
    doc = getattr(revit, "doc", None)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        forms.alert("Open a project model before running this check.", title=__title__)
        return
    if _lib_root() not in sys.path:
        sys.path.insert(0, _lib_root())
    from QualityChecks import mech_duct_clearance
    mech_duct_clearance.run_check(doc, show_ui=True, show_empty=True, options=None)

if __name__ == "__main__":
    main()
