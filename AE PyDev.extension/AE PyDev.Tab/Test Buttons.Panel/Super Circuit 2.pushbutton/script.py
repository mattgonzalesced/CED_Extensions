# -*- coding: utf-8 -*-
import clr
import csv
from pyrevit import script, revit, DB, forms
from pyrevit.revit.db import query
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from collections import defaultdict
from Snippets.family_utils import FamilyLoader

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
output.close_others()
output.set_width(800)
logger = script.get_logger()


# === Load CSV ===
def load_csv_table(filepath):
    with open(filepath, 'r') as file:
        return list(csv.DictReader(file))


# === Pick face ===
def pick_face_with_reference():
    ref = uidoc.Selection.PickObject(ObjectType.Face, "Pick a face")
    face = doc.GetElement(ref.ElementId).GetGeometryObjectFromReference(ref)
    face_norm = face.FaceNormal
    location = doc.GetElement(ref.ElementId).Location.Point
    normal = face.ComputeNormal(DB.UV(0.5, 0.5)).Normalize()
    bbox = face.GetBoundingBox()
    logger.debug(
        "ref:{}, Face: {}, face_norm:{}, norm:{}, loc:{} ".format(ref.ElementId, face, face_norm, normal, location))
    return face, ref, normal, bbox, location


# === Get direction in plane of face ===
def get_reference_direction(normal):
    return DB.XYZ(1, 0, 0) if abs(normal.Z) > 0.9 else DB.XYZ(0, 0, 1)


# === Generate placement points ===
def generate_face_split_points(face, bbox, data_rows):
    odds = [r for r in data_rows if int(r['CKT_Circuit Number_CEDT']) % 2 == 1]
    evens = [r for r in data_rows if int(r['CKT_Circuit Number_CEDT']) % 2 == 0]

    min_u, max_u = bbox.Min.U, bbox.Max.U
    min_v, max_v = bbox.Min.V, bbox.Max.V
    mid_u = (min_u + max_u) / 2.0

    def create_uvs(start_u, count):
        return [DB.UV(start_u, max_v - i * ((max_v - min_v) / max(1, count - 1))) for i in range(count)]

    left = [face.Evaluate(uv) for uv in create_uvs(min_u + 0.05, len(odds))]
    right = [face.Evaluate(uv) for uv in create_uvs(mid_u + 0.05, len(evens))]

    return {"odds": odds, "evens": evens, "left_points": left, "right_points": right}

def get_element_location_point_from_face_ref(face_ref):
    element = doc.GetElement(face_ref.ElementId)
    location = element.Location

    if isinstance(location, DB.LocationPoint):
        return location.Point
    elif isinstance(location, DB.LocationCurve):
        return location.Curve.Evaluate(0.5, True)
    else:
        raise Exception("Cannot extract placement point from selected element.")

# === Family + Param utils ===
def get_family_symbol(row):
    fam, typ = row['Family'].strip(), row['Type'].strip()
    symbols = query.get_family_symbol(fam, typ, doc)
    return symbols[0] if symbols else None


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


def create_electrical_system(doc, element_ids, panel_element):
    if not element_ids or not panel_element:
        return None
    elist = List[DB.ElementId](element_ids)
    system = DB.Electrical.ElectricalSystem.Create(doc, elist, DB.Electrical.ElectricalSystemType.PowerCircuit)
    if system:
        system.SelectPanel(panel_element)
        doc.Regenerate()
    return system


# === MAIN EXECUTION ===
csv_path = r"C:\Users\Aevelina\OneDrive - CoolSys Inc\Documents\FilteredDataExport2.csv"
filepath = forms.pick_file(file_ext="csv", multi_file=False, title="Pick CSV File")
table = load_csv_table(csv_path)

face, ref, normal, bbox, location = pick_face_with_reference()
ref_dir = get_reference_direction(normal)


# Collect panels and build panel name lookup
panel_lookup = {
    query.get_param_value(query.get_param(p, "Panel Name")): p
    for p in DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType()
}

# === PLACEMENT & WIRING ===
circuit_rows = []

with DB.TransactionGroup(doc, "Place + Wire + Param Families") as tg:
    tg.Start()

    with revit.Transaction("Activate Symbols", doc):
        activated = set()
        for row in table:
            symbol = get_family_symbol(row)
            if symbol and not symbol.IsActive and symbol.Id.IntegerValue not in activated:
                symbol.Activate()
                activated.add(symbol.Id.IntegerValue)

    with revit.Transaction("Place and Parameterize", doc):
        for row in table:
            circuit_number = row['CKT_Circuit Number_CEDT'].strip()
            panel_name = row['CKT_Panel_CEDT'].strip()
            panel = panel_lookup.get(panel_name)

            symbol = get_family_symbol(row)
            if not symbol:
                output.print_md("❌ Missing symbol for {} / {}".format(row['Family'], row['Type']))
                continue

            instance = doc.Create.NewFamilyInstance(ref, location, ref_dir, symbol)
            set_instance_parameters(instance, row)

            circuit_rows.append({
                "instance": instance,
                "panel": panel,
                "rating": row.get("CKT_Rating_CED", "").strip(),
                "load_name": row.get("CKT_Load Name_CEDT", "").strip(),
                "notes": row.get("CKT_Schedule Notes_CEDT", "").strip(),
                "row_id": circuit_number
            })

    created_systems = {}

    with revit.Transaction("Create Circuits", doc):
        for row_data in circuit_rows:
            instance = row_data["instance"]
            panel = row_data["panel"]
            if not panel:
                output.print_md("⚠️ Skipping {} — missing panel".format(row_data["row_id"]))
                continue

            system = create_electrical_system(doc, [instance.Id], panel)
            if system:
                created_systems[system.Id] = row_data

    with revit.Transaction("Set Circuit Parameters", doc):
        for sys_id, data in created_systems.items():
            sys = doc.GetElement(sys_id)
            if not sys:
                continue

            if data["load_name"]:
                sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).Set(data["load_name"])
            if data["rating"]:
                try:
                    rating_val = float(data["rating"])
                    sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM).Set(rating_val)
                    val = sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM).AsDouble()
                    output.print_md("✅ Rating set for {}: {}".format(data["row_id"], val))
                except Exception as e:
                    output.print_md("⚠️ Failed to set rating for {}: {}".format(data["row_id"], e))
            if data["notes"]:
                sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM).Set(data["notes"])

    tg.Assimilate()

output.print_md("### ✅ Placement & Circuiting Complete")
