# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script, output
from pyrevit.interop import xl as pyxl
from pyrevit.revit import query
from Autodesk.Revit.DB.Electrical import DistributionSysType, ElectricalPhase

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
out = output.get_output()
out.set_title("Panel Placement Report")

HEADERS = [
    "Column1", "Family", "Type", "Distribution System", "Panel Name_CEDT",
    "Max Number of Circuits_CED", "Max Number of Single Pole Breakers_CED",
    "Mains Rating_CED", "Mains Type_CEDT", "Main Breaker Rating_CED",
    "Short Circuit Rating_CEDT", "Comments"
]

FAMILY_COL = "Family"
TYPE_COL = "Type"
DIST_SYS_COL = "Distribution System"
PARAM_EXCLUDE = [FAMILY_COL, TYPE_COL, "Column1"]
DELTA_X = 5


def load_excel_rows():
    path = forms.pick_file(title="Select Excel File")
    if not path:
        script.exit()
    data = pyxl.load(path, sheets=["Panel Creation"], columns=HEADERS)
    return [r for r in data["Panel Creation"]["rows"] if r.get(FAMILY_COL) and r.get(TYPE_COL)]


def get_distribution_lookup():
    dist_elements = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElecDistributionSys) \
        .WhereElementIsElementType().ToElements()
    return {query.get_name(d): d.Id for d in dist_elements}


def get_dist_system_data(ds):
    dist_data = {
        'name': query.get_name(ds),
        'phase': str(ds.ElectricalPhase),
        'num_wires': str(ds.NumWires),
        'phase_config': str(ds.ElectricalPhaseConfiguration),
        'lg_voltage': '-',
        'll_voltage': '-'
    }
    if ds.VoltageLineToGround:
        p = ds.VoltageLineToGround.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
        if p: dist_data['lg_voltage'] = round(p.AsDouble(), 2)
    if ds.VoltageLineToLine:
        p = ds.VoltageLineToLine.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
        if p: dist_data['ll_voltage'] = round(p.AsDouble(), 2)
    return dist_data


def resolve_missing_distribution_systems(rows, dist_lookup):
    dist_names = sorted(set(r.get(DIST_SYS_COL, '').strip() for r in rows if r.get(DIST_SYS_COL)))
    unresolved = [name for name in dist_names if name not in dist_lookup]
    if not unresolved:
        return dist_lookup

    logger.info("Prompting user to resolve missing distribution systems:")
    updated = {}

    all_dist_elements = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElecDistributionSys) \
        .WhereElementIsElementType().ToElements()

    display_map = {}
    for ds in all_dist_elements:
        name = query.get_name(ds)
        phase = str(ds.ElectricalPhase)
        wires = str(ds.NumWires)
        config = str(ds.ElectricalPhaseConfiguration)
        label = "{} | {} | {} wires | {}".format(name, phase, wires, config)
        display_map[label] = ds

    for name in unresolved:
        choices = sorted(display_map.keys())
        selected = forms.SelectFromList.show(
            choices,
            title="Select replacement for missing distribution system '{}':".format(name),
            multiselect=False
        )
        if selected:
            updated[name] = display_map[selected].Id
        else:
            logger.warning("No replacement selected for '{}'. It will be skipped.".format(name))

    dist_lookup.update(updated)
    return dist_lookup

    logger.info("Prompting user to resolve missing distribution systems:")
    updated = {}

    all_dist_elements = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElecDistributionSys) \
        .WhereElementIsElementType().ToElements()

    display_map = {}
    for ds in all_dist_elements:
        data = get_dist_system_data(ds)
        label = "{name} | {phase} | {num_wires} wires | LG:{lg_voltage} | LL:{ll_voltage}".format(**data)
        display_map[label] = ds

    for name in unresolved:
        choices = sorted(display_map.keys())
        selected = forms.SelectFromList.show(
            choices,
            title="Select replacement for missing distribution system '{}':".format(name),
            multiselect=False
        )
        if selected:
            updated[name] = display_map[selected].Id
        else:
            logger.warning("No replacement selected for '{}'. It will be skipped.".format(name))

    dist_lookup.update(updated)
    return dist_lookup

    logger.info("Prompting user to resolve missing distribution systems:")
    updated = {}

    all_dist_elements = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalDistributionSystems) \
        .WhereElementIsElementType().ToElements()

    def describe_dist_system(ds):
        try:
            phase = str(ds.ElectricalPhase)
            wires = str(ds.NumWires)
            config = str(ds.ElectricalPhaseConfiguration)
            lg = ds.VoltageLineToGround
            ll = ds.VoltageLineToLine

            def get_voltage(v):
                if not v:
                    return "-"
                p = v.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
                return round(p.AsDouble(), 2) if p else "-"

            lg_v = get_voltage(lg)
            ll_v = get_voltage(ll)
            return "{} | {}Φ | {} wires | LG:{} | LL:{}".format(query.get_name(ds), phase, wires, lg_v, ll_v)
        except:
            return query.get_name(ds)

    name_map = {describe_dist_system(d): d for d in all_dist_elements}

    for name in unresolved:
        choices = sorted(name_map.keys())
        selected = forms.SelectFromList.show(
            choices,
            title="Select replacement for missing distribution system '{}':".format(name),
            multiselect=False
        )
        if selected:
            updated[name] = name_map[selected].Id
        else:
            logger.warning("No replacement selected for '{}'. It will be skipped.".format(name))

    dist_lookup.update(updated)
    return dist_lookup  # Nothing missing

    logger.info("Prompting user to resolve missing distribution systems:")
    updated = {}
    all_dist_elements = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalDistributionSystems) \
        .WhereElementIsElementType().ToElements()
    all_options = {query.get_name(e): e for e in all_dist_elements}

    for name in unresolved:
        choices = sorted(all_options.keys())
        selected = forms.SelectFromList.show(
            choices,
            title="Select distribution system to use for: '{}'".format(name),
            multiselect=False
        )
        if selected:
            updated[name] = all_options[selected].Id
        else:
            logger.warning("No replacement selected for '{}'. It will be skipped.".format(name))

    dist_lookup.update(updated)
    return dist_lookup



def get_start_location():
    return uidoc.Selection.PickPoint("Select placement point")


def get_active_view_level():
    lvl = doc.ActiveView.GenLevel
    if not lvl:
        logger.error("Active view has no associated level.")
        script.exit()
    return lvl


def find_family_symbol(family_name, type_name):
    for fs in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        if query.get_name(fs.Family) == family_name and query.get_name(fs) == type_name:
            return fs
    return None


def set_instance_parameters(panel, row, dist_lookup):
    for param_name, value in row.items():
        if param_name in PARAM_EXCLUDE:
            continue

        param = panel.LookupParameter(param_name)
        if not param or param.IsReadOnly:
            continue

        try:
            if param_name == DIST_SYS_COL:
                sys_id = dist_lookup.get(str(value).strip())
                if sys_id:
                    param.Set(sys_id)
                else:
                    logger.warning("Distribution system '{}' not found".format(value))
            elif param.StorageType == DB.StorageType.String:
                param.Set(str(value))
            elif param.StorageType == DB.StorageType.Double:
                param.Set(float(value))
            elif param.StorageType == DB.StorageType.Integer:
                param.Set(int(value))
        except Exception as e:
            logger.warning("Could not set param '{}': {}".format(param_name, e))


def print_panel_report(row):
    panel_name = row.get("Panel Name_CEDT", "[Unnamed Panel]")
    out.print_md("### **{}**".format(panel_name))

    ordered_keys = [
        FAMILY_COL, TYPE_COL, DIST_SYS_COL,
        "Max Number of Circuits_CED",
        "Max Number of Single Pole Breakers_CED",
        "Mains Rating_CED", "Mains Type_CEDT",
        "Main Breaker Rating_CED",
        "Short Circuit Rating_CEDT",
        "Comments"
    ]

    for key in ordered_keys:
        val = row.get(key)
        if val is not None:
            out.print_md("**{}**: {}".format(key, val))

    out.print_md("---")


def main():
    rows = load_excel_rows()
    start_point = get_start_location()
    view_level = get_active_view_level()
    dist_lookup = resolve_missing_distribution_systems(rows, get_distribution_lookup())

    with DB.Transaction(doc, "Place Panels from Excel") as t:
        t.Start()
        for i, row in enumerate(rows):
            family_name = row.get(FAMILY_COL)
            type_name = row.get(TYPE_COL)

            family_symbol = find_family_symbol(family_name, type_name)
            if not family_symbol:
                logger.warning("Family/type not found: {} / {}".format(family_name, type_name))
                continue

            if not family_symbol.IsActive:
                family_symbol.Activate()
                doc.Regenerate()

            point = DB.XYZ(start_point.X + (i * DELTA_X), start_point.Y, start_point.Z)
            panel = doc.Create.NewFamilyInstance(point, family_symbol, view_level,
                                                 DB.Structure.StructuralType.NonStructural)
            set_instance_parameters(panel, row, dist_lookup)
            print_panel_report(row)
        t.Commit()


if __name__ == "__main__":
    main()
