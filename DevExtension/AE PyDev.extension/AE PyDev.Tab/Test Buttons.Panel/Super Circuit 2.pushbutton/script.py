# -*- coding: utf-8 -*-
import clr
import csv
from pyrevit import script, revit, DB
from pyrevit.revit.db import query
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from collections import defaultdict

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
    location = doc.GetElement(ref.ElementId).Location
    normal = face.ComputeNormal(DB.UV(0.5, 0.5)).Normalize()
    bbox = face.GetBoundingBox()
    logger.debug("ref:{}, Face: {}, face_norm:{}, norm:{}, loc:{} ".format(ref.ElementId,face,face_norm,normal,location))
    return face, ref, normal, bbox

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
                        logger.debug("Original Val: {}, Converted: {}".format(val,converted))
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
table = load_csv_table(csv_path)

face, ref, normal, bbox = pick_face_with_reference()
ref_dir = get_reference_direction(normal)
placement = generate_face_split_points(face, bbox, table)

# Collect panels and build panel name lookup
panel_lookup = {
    query.get_param_value(query.get_param(p, "Panel Name")): p
    for p in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType()
}

# === PLACEMENT & GROUPING ===
instance_row_pairs = []
circuit_groups = {}
circuit_group_keys_in_order = []

with DB.TransactionGroup(doc, "Place + Wire + Param Families") as tg:
    tg.Start()

    with revit.Transaction("Activate Symbols", doc):
        activated = set()
        for row in table:
            sym = get_family_symbol(row)
            if sym and not sym.IsActive and sym.Id.IntegerValue not in activated:
                sym.Activate()
                activated.add(sym.Id.IntegerValue)

    # Map circuit number to face points (still needed)
    odds_map = {r['CKT_Circuit Number_CEDT']: pt for r, pt in zip(placement['odds'], placement['left_points'])}
    evens_map = {r['CKT_Circuit Number_CEDT']: pt for r, pt in zip(placement['evens'], placement['right_points'])}

    with revit.Transaction("Place and Parameterize", doc):
        for row in table:
            circuit_number = row['CKT_Circuit Number_CEDT'].strip()
            panel = row['CKT_Panel_CEDT'].strip()
            pt = odds_map.get(circuit_number) or evens_map.get(circuit_number)
            output.print_md("ckt: {}, Point: {}".format(circuit_number,pt))
            if not pt:
                output.print_md("⚠️ No point found for circuit {}".format(circuit_number))
                continue

            symbol = get_family_symbol(row)
            if not symbol:
                output.print_md("❌ Missing symbol for {} / {}".format(row['Family'], row['Type']))
                continue

            instance = doc.Create.NewFamilyInstance(ref, pt, ref_dir, symbol)
            set_instance_parameters(instance, row)
            instance_row_pairs.append((instance, row))

            key = (panel, circuit_number)
            if key not in circuit_groups:
                circuit_groups[key] = {
                    "elements": [],
                    "panel": panel_lookup.get(panel),
                    "rating": row.get("CKT_Rating_CED", "").strip(),
                    "load_name": row.get("CKT_Load Name_CEDT", "").strip(),
                    "notes": row.get("CKT_Schedule Notes_CEDT", "").strip()
                }
                circuit_group_keys_in_order.append(key)

            circuit_groups[key]["elements"].append(instance)

    created_systems = {}

    with revit.Transaction("Create Circuits", doc):
        for key in circuit_group_keys_in_order:  # maintain CSV order
            data = circuit_groups[key]
            ids = [e.Id for e in data["elements"]]
            panel = data["panel"]
            if not panel or not ids:
                output.print_md("⚠️ Skipping circuit {} — missing panel or elements".format(key))
                continue
            system = create_electrical_system(doc, ids, panel)
            if system:
                created_systems[system.Id] = data

    with revit.Transaction("Set Circuit Parameters", doc):
        for sys_id, data in created_systems.items():
            sys = doc.GetElement(sys_id)
            if not sys:
                continue

            if data["load_name"]:
                sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).Set(data["load_name"])
            if data["rating"]:
                sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM).Set(data["rating"])
            if data["notes"]:
                sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM).Set(data["notes"])

    tg.Assimilate()

output.print_md("### ✅ Placement & Circuiting Complete")
output.print_md("**Odds:** {}".format([r['CKT_Circuit Number_CEDT'] for r in placement['odds']]))
output.print_md("**Evens:** {}".format([r['CKT_Circuit Number_CEDT'] for r in placement['evens']]))