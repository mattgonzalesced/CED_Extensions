# -*- coding: utf-8 -*-
import clr
import csv
from pyrevit import script, revit, DB, forms
from pyrevit.revit.db import query
from pyrevit.interop import xl as pyxl
from wmlib import *
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from collections import defaultdict


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
output.close_others()
output.set_width(800)
logger = script.get_logger()



def get_panel_surfaces(panel_names):
    surfaces = {}
    for eq in DB.FilteredElementCollector(revit.doc)\
            .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)\
            .WhereElementIsNotElementType():
        param = eq.LookupParameter("Panel Name_CEDT")
        if param and param.HasValue:
            value = param.AsString().strip()
            if value in panel_names:
                surfaces[value] = EquipmentSurface(eq.Id.IntegerValue)
    return surfaces

def get_family_symbol(family_name, type_name):
    for fs in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        if query.get_name(fs.Family) == family_name and query.get_name(fs) == type_name:
            return fs
    output.print_md("❌ Symbol not found for Family: '{}' | Type: '{}'".format(family_name, type_name))
    return None


def set_instance_parameters(inst, row):
    skip = ['Family', 'Type', 'INCLUDE CIRCUIT', 'CIRCUIT SORT']
    for key, val in row.items():
        if key in skip or not val:
            continue
        param = inst.LookupParameter(key)
        if param and not param.IsReadOnly:
            try:
                if param.StorageType == DB.StorageType.String:
                    param.Set(str(val))
                elif param.StorageType == DB.StorageType.Integer:
                    param.Set(int(float(val)))
                elif param.StorageType == DB.StorageType.Double:
                    # Special handling for Apparent Load keys
                    if key in ["Apparent Load Ph 1_CED", "Apparent Load Ph 2_CED", "Apparent Load Ph 3_CED"]:
                        logger.debug("Setting Apparent Load Units")
                        forge_type_va = DB.ForgeTypeId("autodesk.unit.unit:voltAmperes-1.0.1")
                        converted = DB.UnitUtils.ConvertToInternalUnits(float(val), forge_type_va)
                        param.Set(converted)
                        logger.debug("Original Val: {}, Converted: {}".format(val, converted))
                    else:
                        logger.debug("No unit conversion, regular double")
                        param.Set(float(val))
            except Exception as e:
                output.print_md("⚠️ Failed to set {}: {}".format(key, e))





def create_circuit(doc, instance, panel):
    if not instance or not panel:
        return None
    return DB.Electrical.ElectricalSystem.Create(doc, List[DB.ElementId]([instance.Id]), DB.Electrical.ElectricalSystemType.PowerCircuit)

def main():
    loader = ExcelCircuitLoader()
    loader.pick_excel_file()
    sheetnames = loader.get_valid_sheet_names()
    if not sheetnames:
        forms.alert("No valid sheets found.")
        return

    selected = loader.pick_sheet_names(sheetnames)
    if not selected:
        return

    ordered_rows = loader.get_ordered_rows(selected)
    if not ordered_rows:
        output.print_md("⚠️ No valid circuit rows found in selected sheets.")
        return

    panel_names = set(row["CKT_Panel_CED"] for row in ordered_rows)
    surface_map = get_panel_surfaces(panel_names)

    output.print_md("### 📋 Matched {} panel names from Excel".format(len(panel_names)))
    output.print_md("### 🧱 Found {} panel surfaces in model".format(len(surface_map)))

    all_instances = []
    system_data = []
    activated_symbols = set()

    with DB.TransactionGroup(doc, "Place + Wire + Param Families") as tg:
        tg.Start()

        with revit.Transaction("Activate Symbols", doc):
            for row in ordered_rows:
                symbol = get_family_symbol(row["Family"], row["Type"])
                if symbol and not symbol.IsActive and symbol.Id.IntegerValue not in activated_symbols:
                    symbol.Activate()
                    activated_symbols.add(symbol.Id.IntegerValue)
            output.print_md("✅ Activated {} symbols.".format(len(activated_symbols)))

        with revit.Transaction("Place and Parameterize", doc, swallow_errors=True):
            for row in ordered_rows:
                panel_name = row["CKT_Panel_CED"]
                circuit_number = row["CKT_Circuit Number_CEDT"]

                surface = surface_map.get(panel_name)
                if not surface or not surface.face:
                    output.print_md("❌ Panel '{}' missing or has no placeable face.".format(panel_name))
                    continue

                symbol = get_family_symbol(row["Family"], row["Type"])
                if not symbol:
                    output.print_md("❌ Symbol not found for {} / {}".format(row["Family"], row["Type"]))
                    continue

                ref = surface.face
                point = surface.location + 0.25 * surface.facing
                instance = doc.Create.NewFamilyInstance(ref, point, DB.XYZ(1, 0, 0), symbol)

                set_instance_parameters(instance, row)

                all_instances.append(instance)
                system_data.append({
                    "instance": instance,
                    "panel": surface.element,
                    "circuit_number": circuit_number,
                    "load_name": row.get("CKT_Load Name_CEDT", ""),
                    "rating": row.get("CKT_Rating_CED", ""),
                    "frame": row.get("CKT_Frame_CED", "")
                })

            doc.Regenerate()

        created_systems = {}
        with revit.Transaction("Create Circuits", doc, swallow_errors=True):
            for data in system_data:
                system = create_circuit(doc, data["instance"], data["panel"])
                if system:
                    system.SelectPanel(data["panel"])
                    created_systems[system.Id] = data

        with revit.Transaction("Set Circuit Parameters", doc):
            for sys_id, data in created_systems.items():
                sys = doc.GetElement(sys_id)
                if not sys:
                    continue
                if data["load_name"]:
                    sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).Set(data["load_name"])
                if data["rating"]:
                    try:
                        sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM).Set(float(data["rating"]))
                    except:
                        pass
                if data["frame"]:
                    param = sys.LookupParameter("CKT_Frame_CED")
                    if param and not param.IsReadOnly:
                        param.Set(data["frame"])

        tg.Assimilate()

    output.print_md("### ✅ Placement & Circuiting Complete")
    output.print_md("🧱 Total placed: {}".format(len(all_instances)))
    output.print_md("🔌 Total circuits: {}".format(len(created_systems)))



if __name__ == "__main__":
    main()
