# -*- coding: utf-8 -*-
from pyrevit import revit, DB, script, forms
from pyrevit.interop import xl as pyxl
from pyrevit.revit import query

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc
active_view = uidoc.ActiveView
level = active_view.GenLevel
output = script.get_output()
output.close_others()
HEADERS = [
    "Id", "Family", "Type", "Level: Name",
    "Coordinates - Internal: Point_X",
    "Coordinates - Internal: Point_Y",
    "Coordinates - Internal: Point_Z",
    "CKT_Panel_CEDT",
    "CKT_Circuit Number_CEDT"
]

FAMILY_COL = "Family"
TYPE_COL = "Type"
X_COL = "Coordinates - Internal: Point_X"
Y_COL = "Coordinates - Internal: Point_Y"
Z_COL = "Coordinates - Internal: Point_Z"
PNL_COL = "CKT_Panel_CEDT"
CKT_COL = "CKT_Circuit Number_CEDT"


# Use LookupParameter to set shared parameters
def set_parameter(element, param_name, value):
    param = element.LookupParameter(param_name)
    if param and value is not None:
        param.Set(value)
def load_excel_rows():
    path = forms.pick_file(title="Select Excel File")
    if not path:
        script.exit()
    data = pyxl.load(path, sheets=None, columns=HEADERS)
    sheet_name = list(data.keys())[0]
    return [r for r in data[sheet_name]["rows"] if r.get(FAMILY_COL) and r.get(TYPE_COL)]

# Build a map of (Family, Type) -> FamilySymbol
def get_family_symbol_map():
    symbol_map = {}
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        fam_name = query.get_name(symbol.Family)
        typ_name = query.get_name(symbol)
        symbol_map[(fam_name, typ_name)] = symbol
    return symbol_map

rows = load_excel_rows()
symbol_lookup = get_family_symbol_map()
elements_created = []

with revit.Transaction("Place Families from Excel"):
    for row in rows:
        fam = row[FAMILY_COL]
        typ = row[TYPE_COL]
        x = float(row.get(X_COL, "0"))
        y = float(row.get(Y_COL, "0"))
        z = float(row.get(Z_COL, "0"))
        id = str(row.get("Id","XXXX"))
        pnl = str(row.get(PNL_COL,""))
        ckt = str(row.get(CKT_COL,""))
        symbol = symbol_lookup.get((fam, typ))
        if not symbol:
            logger.warning("Family symbol not found: {} : {}".format(fam, typ))
            continue

        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        pt = DB.XYZ(x, y, z)
        inst = doc.Create.NewFamilyInstance(pt, symbol, level, DB.Structure.StructuralType.NonStructural)
        set_parameter(inst,"CKT_Panel_CEDT",pnl)
        set_parameter(inst, "CKT_Circuit Number_CEDT", ckt)
        set_parameter(inst, "Mark", str(id))


        elements_created.append(inst.Id)
selection = revit.get_selection()
selection.set_to(elements_created)

script.get_output().print_md("### Placed {} element(s).".format(len(elements_created)))
