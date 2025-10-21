# -*- coding: utf-8 -*-
try:
    # rpw optional but nice
    from rpw import revit
    def get_doc(): return revit.doc
    def get_uidoc(): return revit.uidoc
except:
    import __revit__
    def get_doc(): return __revit__.ActiveUIDocument.Document
    def get_uidoc(): return __revit__.ActiveUIDocument