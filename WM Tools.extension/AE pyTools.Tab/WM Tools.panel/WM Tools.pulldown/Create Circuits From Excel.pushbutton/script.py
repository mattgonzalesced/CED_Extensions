# -*- coding: utf-8 -*-
import Autodesk.Revit.DB.Electrical as DBE
import clr
from System.Collections.Generic import List

from wmlib import *

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
output.close_others()
output.set_width(800)
logger = script.get_logger()

def get_panel_surfaces(panel_names):
    surfaces = {}
    for eq in DB.FilteredElementCollector(revit.doc) \
            .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
            .WhereElementIsNotElementType():
        param = eq.LookupParameter("Panel Name_CEDT")
        if param and param.HasValue:
            value = param.AsString().strip()
            if value in panel_names:
                surfaces[value] = EquipmentSurface(eq.Id.Value)
    return surfaces


def get_panel_elements(panel_names):
    panels = {}
    for eq in DB.FilteredElementCollector(revit.doc) \
            .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
            .WhereElementIsNotElementType():
        param = eq.LookupParameter("Panel Name_CEDT")
        if param and param.HasValue:
            value = param.AsString().strip()
            if value in panel_names:
                panels[value] = eq
    return panels   # ✅ return the dict, not just eq

# --- Load Families ---
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = True
        return True


def find_symbol_in_doc(family_name, type_name):
    for fs in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        if query.get_name(fs.Family) == family_name and query.get_name(fs) == type_name:
            return fs
    return None


def load_family_from_disk(family_name):
    import os
    family_path = os.path.join(os.path.dirname(__file__), "{}.rfa".format(family_name))
    if not os.path.exists(family_path):
        output.print_md("❌ Family file not found: `{}`".format(family_path))
        return None

    output.print_md("📥 Loading family from file: `{}`".format(family_path))
    handler = FamilyLoaderOptionsHandler()
    family_loaded = clr.Reference[DB.Family]()
    success = doc.LoadFamily(family_path, handler, family_loaded)
    return family_loaded.Value if success else None


def get_family_symbol(family_name, type_name):
    symbol = find_symbol_in_doc(family_name, type_name)
    if symbol:
        return symbol

    loaded_family = load_family_from_disk(family_name)
    if loaded_family:
        for sym_id in loaded_family.GetFamilySymbolIds():
            symbol = doc.GetElement(sym_id)
            if query.get_name(symbol) == type_name:
                return symbol

        output.print_md("❌ Type '{}' not found in newly loaded family '{}'.".format(type_name, family_name))
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
                    # --- Apparent Load (VA) ---
                    if key in ["Apparent Load Ph 1_CED",
                               "Apparent Load Ph 2_CED",
                               "Apparent Load Ph 3_CED",
                               "Apparent Load Input_CED"]:
                        logger.debug("Setting Apparent Load Units (VA)")
                        forge_type_va = DB.ForgeTypeId("autodesk.unit.unit:voltAmperes-1.0.1")
                        converted = DB.UnitUtils.ConvertToInternalUnits(float(val), forge_type_va)
                        param.Set(converted)
                        logger.debug("Original Val: {}, Converted (VA): {}".format(val, converted))

                    # --- Voltage (V) ---
                    elif "Voltage" in key:   # or use a stricter list if needed
                        logger.debug("Setting Voltage Units (V)")
                        forge_type_v = DB.ForgeTypeId("autodesk.unit.unit:volts-1.0.1")
                        converted = DB.UnitUtils.ConvertToInternalUnits(float(val), forge_type_v)
                        param.Set(converted)
                        logger.debug("Original Val: {}, Converted (V): {}".format(val, converted))

                    # --- Regular doubles ---
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
    loader.build_dict_rows()
    output.freeze()
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

    panel_names = set(row["CKT_Panel_CEDT"] for row in ordered_rows)
    surface_map = get_panel_surfaces(panel_names)
    panels = get_panel_elements(panel_names)
    output.print_md("### 📋 Matched {} panel names from Excel".format(len(panel_names)))
    output.print_md("### 🧱 Found {} panel surfaces in model".format(len(surface_map)))

    all_instances = []
    system_data = []
    activated_symbols = set()

    missing_families = set()
    for row in ordered_rows:
        fam = row["Family"]
        typ = row["Type"]
        if not find_symbol_in_doc(fam, typ):
            missing_families.add(fam)

    with DB.TransactionGroup(doc, "Create Circuits From Excel") as tg:
        tg.Start()

        if missing_families:
            with revit.Transaction("Load Missing Families", doc):
                for fam in missing_families:
                    load_family_from_disk(fam)
        output.unfreeze()
        with revit.Transaction("Set Ckt Sequence", doc):
            try:
                electrical_settings = DBE.ElectricalSetting.GetElectricalSettings(doc)

                current_sequence = electrical_settings.CircuitSequence
                current_value = current_sequence.ToString()

                logger.info("Current circuit sequence setting: {}".format(current_value))

                # Only update if not already OddThenEven
                if current_value != "OddThenEven":
                    electrical_settings.CircuitSequence = current_sequence.__class__.OddThenEven

                    logger.info("Circuit sequence was updated to 'OddThenEven'.")
                else:
                    logger.info("No changes needed. Already set to 'OddThenEven'.")

            except Exception as e:
                logger.info("Failed to read or update electrical settings:\n{}".format(str(e)))

        with revit.Transaction("Activate Symbols", doc):
            for row in ordered_rows:
                symbol = get_family_symbol(row["Family"], row["Type"])
                if symbol and not symbol.IsActive and symbol.Id.Value not in activated_symbols:
                    symbol.Activate()
                    activated_symbols.add(symbol.Id.Value)

        spacing = 3.0  # feet
        panel_offsets = {}  # track per-panel placement counts
        panel_map = get_panel_elements(panel_names)

        with revit.Transaction("Place and Parameterize", doc, swallow_errors=True):
            for row in ordered_rows:
                panel_name = row["CKT_Panel_CEDT"]
                circuit_number = row["CKT_Circuit Number_CEDT"]

                symbol = get_family_symbol(row["Family"], row["Type"])
                if not symbol:
                    output.print_md("❌ Symbol not found for {} / {}".format(row["Family"], row["Type"]))
                    continue

                placement_type = symbol.Family.FamilyPlacementType

                if placement_type in (
                        DB.FamilyPlacementType.OneLevelBasedHosted,
                        DB. FamilyPlacementType.WorkPlaneBased,  # sometimes still face-based
                ):
                    # --- FACE-BASED LOGIC ---
                    surface = surface_map.get(panel_name)
                    if not surface or not surface.face:
                        output.print_md("❌ Panel '{}' missing or has no placeable face.".format(panel_name))
                        continue

                    ref = surface.face
                    point = surface.location + 0.25 * surface.facing
                    instance = doc.Create.NewFamilyInstance(ref, point, DB.XYZ(1, 0, 0), symbol)

                else:
                    # --- UNHOSTED LOGIC ---
                    # Track per-panel column index
                    if panel_name not in panel_offsets:
                        panel_offsets[panel_name] = 0

                    # Row index for this circuit in the panel
                    row_index = panel_offsets[panel_name]
                    panel_offsets[panel_name] += 1

                    # Column index = position of this panel among all panels
                    if "panel_columns" not in globals():
                        panel_columns = {}
                    if panel_name not in panel_columns:
                        panel_columns[panel_name] = len(panel_columns)  # assign new column for each panel

                    col_index = panel_columns[panel_name]

                    # Placement point: X = panel column, Y = circuit row
                    x_spacing = 6.0  # feet between panel columns
                    y_spacing = 3.0  # feet between circuits
                    base_point = DB.XYZ(col_index * x_spacing, row_index * y_spacing, 0)

                    instance = doc.Create.NewFamilyInstance(
                        base_point,
                        symbol,
                        doc.ActiveView.GenLevel,
                        DB.Structure.StructuralType.NonStructural
                    )

                # --- Shared parameterization ---
                set_instance_parameters(instance, row)

                param = instance.LookupParameter("CKT_Schedule Notes_CEDT")
                if param and not param.IsReadOnly:
                    param.Set("EX")

                all_instances.append(instance)
                system_data.append({
                    "instance": instance,
                    "panel": panel_map.get(panel_name),  # ✅ now looks up from dict
                    "circuit_number": circuit_number,
                    "load_name": row.get("CKT_Load Name_CEDT", ""),
                    "rating": row.get("CKT_Rating_CED", ""),
                    "frame": row.get("CKT_Frame_CED", "")
                })

            doc.Regenerate()

        created_systems = {}
        with revit.Transaction("Create Circuits", doc, swallow_errors=True):
            for data in system_data:
                inst = data["instance"]
                panel = data["panel"]
                ckt_num = data["circuit_number"]

                # Create the electrical system
                system = create_circuit(doc, inst, panel)
                if system:
                    logger.debug("Created system for circuit {} (Instance Id: {})"
                                 .format(ckt_num, inst.Id.IntegerValue))

                    if panel:
                        try:
                            system.SelectPanel(panel)

                            # Gather panel info for debug
                            poles = inst.LookupParameter("Number of Poles").AsInteger() \
                                if inst.LookupParameter("Number of Poles") else None
                            voltage_param = panel.LookupParameter("Voltage_CED") or panel.LookupParameter(
                                "Panel Voltage")
                            volts = voltage_param.AsDouble() if voltage_param else None

                            # Convert volts from internal units
                            if volts:
                                volts = DB.UnitUtils.ConvertFromInternalUnits(
                                    volts,
                                    DB.ForgeTypeId("autodesk.unit.unit:volts-1.0.1")
                                )

                            logger.debug("Assigned to panel '{}': Circuit {} | Poles={} | Volts={}"
                                         .format(query.get_name(panel),
                                                 ckt_num,
                                                 poles if poles is not None else "N/A",
                                                 volts if volts is not None else "N/A"))

                        except Exception as e:
                            logger.warning("Failed to assign system for circuit {} to panel '{}': {}"
                                           .format(ckt_num, query.get_name(panel), e))

                    created_systems[system.Id] = data

        with revit.Transaction("Set Circuit Parameters", doc, swallow_errors=True):
            for sys_id, data in created_systems.items():
                sys = doc.GetElement(sys_id)
                if not sys:
                    continue

                # Transfer 'EX' note from instance to circuit-level parameter
                notes_param = data["instance"].LookupParameter("CKT_Schedule Notes_CEDT")
                if notes_param and notes_param.HasValue:
                    val = notes_param.AsString()
                    target_param = sys.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
                    if target_param and not target_param.IsReadOnly:
                        target_param.Set(val)

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

        deleted_count = 0
        with revit.Transaction("Remove Space Circuits", doc, swallow_errors=True):
            for data in system_data:
                load_name = data.get("load_name", "").lower()
                rating = data.get("rating", "")
                try:
                    rating_val = float(rating)
                except:
                    rating_val = -1  # fallback

                if "space" in load_name and rating_val == 0:
                    try:
                        doc.Delete(data["instance"].Id)
                        deleted_count += 1
                    except Exception as e:
                        logger.warning("Failed to delete instance: {}".format(e))


        tg.Assimilate()

    output.print_md("### ✅ Placement & Circuiting Complete")
    output.print_md("🧱 Total placed: {}".format(len(all_instances)))
    output.print_md("🔌 Total circuits: {}".format(len(created_systems)))
    output.print_md("🧹 Removed {} space circuits (load name contains 'space' and rating = 0)".format(deleted_count))



if __name__ == "__main__":
    main()
