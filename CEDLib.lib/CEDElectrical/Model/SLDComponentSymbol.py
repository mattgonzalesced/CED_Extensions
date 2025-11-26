# -*- coding: utf-8 -*-

from pyrevit import DB, script, revit

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

console = script.get_output()
logger = script.get_logger()



class SLDComponentSymbol:
    """Represents a detail item symbol (FamilySymbol)."""
    def __init__(self, symbol):
        self.symbol = symbol


    @staticmethod
    def get_symbol(doc, family_name, type_name):
        """Retrieve a symbol by family name and type name."""
        collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DetailComponents).WhereElementIsElementType()
        for symbol in collector:
            if isinstance(symbol, DB.FamilySymbol):
                family_name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                symbol_family_name = family_name_param.AsString() if family_name_param else None
                if symbol_family_name == family_name and symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == type_name:
                    return SLDComponentSymbol(symbol)
        return None

    def activate(self, doc):
        """Activate the symbol if not already active."""
        if not self.symbol.IsActive:
            self.symbol.Activate()
            doc.Regenerate()

    def place(self, doc, view, location, rotation=0):
        """Place the symbol at a location in the view."""
        try:
            detail_item = doc.Create.NewFamilyInstance(location, self.symbol, view)
            if rotation:
                detail_item.Location.Rotate(DB.Line.CreateBound(location, location + DB.XYZ(0, 0, 1)), rotation)
            return detail_item
        except:
            return None