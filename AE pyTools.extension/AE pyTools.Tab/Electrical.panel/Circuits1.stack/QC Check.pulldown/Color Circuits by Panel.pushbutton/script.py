# -*- coding: utf-8 -*-
# Color Circuits by Panel  |  Revit Python 2.7 – pyRevit / Revit API
# - Creates/updates: per-panel multi-category OR filters (per your screenshot)
# - Colors current (or new) plan/RCP/3D view using overrides (unique for included, grey+halftone for others)
# - Creates/updates legend drafting view with swatches + labels (swatches keyed via Comments)
# - DOES NOT modify or create any View Templates

from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script

logger = script.get_logger()
output = script.get_output()
doc = revit.doc

# -----------------------------------------------------------------------------
# 0) Constants
# -----------------------------------------------------------------------------
FILTER_PREFIX = "PanelChecker - {0} - {1}"      # part_type, panel_name
LEGEND_VIEW_NAME = "PanelChecker Legend"
FILLED_REGION_TYPE_NAME = "PanelChecker - Solid Fill"
VIEW_TEMPLATE_NAME = "E_PanelChecker View"
NO_PANEL_FLAG = "__NO PANEL__"
BIP_PANEL_ON_DEVICES = DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM
BIP_COMMENTS = DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS

DEVICE_FIXTURE_OR_CATS = {
    DB.BuiltInCategory.OST_DetailComponents: DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
    DB.BuiltInCategory.OST_ElectricalEquipment: DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    DB.BuiltInCategory.OST_ElectricalFixtures: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_DataDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_FireAlarmDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_LightingDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_LightingFixtures: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_MechanicalControlDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_SecurityDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
}

SYSTEMS_DEVICES = {
    DB.BuiltInCategory.OST_DataDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_FireAlarmDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_SecurityDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
}

POWER_FIXTURE_EQUIP = {
    DB.BuiltInCategory.OST_DetailComponents: DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
    DB.BuiltInCategory.OST_ElectricalEquipment: DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    DB.BuiltInCategory.OST_ElectricalFixtures: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_LightingDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_LightingFixtures: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    DB.BuiltInCategory.OST_MechanicalControlDevices: DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
}

COLOR_PALETTE = [
    (0, 108, 153),  # teal blue
    (255, 200, 0),  # golden yellow
    (84, 0, 153),  # dark violet
    (255, 80, 0),  # bright orange
    (0, 220, 180),  # aqua green
    (200, 0, 150),  # magenta
    (100, 220, 0),  # lime green
    (255, 0, 195),  # hot pink
    (0, 60, 120),  # navy blue
    (255, 255, 0),  # yellow
    (182, 0, 255),  # bright purple
    (153, 48, 0),  # reddish brown
    (0, 234, 255),  # light cyan
    (120, 0, 90),  # plum
    (255, 104, 0),  # pumpkin orange
    (36, 108, 132),  # muted teal
    (255, 0, 234),  # neon pink
    (60, 132, 0),  # forest green
    (140, 0, 255),  # violet
    (220, 120, 0),  # amber
    # ---- remaining, still sequenced for variety ----
    (0, 100, 200),
    (153, 120, 0),
    (0, 132, 108),
    (78, 234, 255),
    (132, 72, 0),
    (255, 156, 0),
    (24, 96, 36),
    (40, 160, 60),
    (52, 208, 78),
    (153, 72, 120),
    (255, 120, 200),
    (255, 156, 255),
    (96, 120, 24),
    (160, 200, 40),
    (208, 255, 52),
    (90, 0, 0),
    (150, 0, 0),
    (195, 0, 0),
    (108, 12, 24),
    (180, 20, 40),
    (83, 41, 11),
    (139, 69, 19),
    (180, 89, 24),
    (126, 63, 18),
    (123, 79, 37),
    (205, 133, 63),
    (255, 172, 81),
    (96, 49, 27),
    (160, 82, 45),
    (208, 106, 58),
    (126, 108, 84),
    (210, 180, 140),
]

GREY = (160, 160, 160)
RED = (255,0,0)

# -----------------------------------------------------------------------------
# 1) Small helpers
# -----------------------------------------------------------------------------
def _rgb(r, g, b):
    return DB.Color(bytearray([r])[0], bytearray([g])[0], bytearray([b])[0])

def _as_sorted_unique(items):
    seen = set()
    out = []
    for s in items:
        if s and s not in seen:
            seen.add(s)
            out.append(s)
    out.sort()
    return out



# -----------------------------------------------------------------------------
# 2) Panel discovery
# -----------------------------------------------------------------------------
def get_electrical_equipment(doc):
    return list(DB.FilteredElementCollector(doc)
                .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)
                .WhereElementIsNotElementType())

def get_all_panel_names(doc):
    names = []
    for eq in get_electrical_equipment(doc):
        ok = True
        try:
            ds = eq.LookupParameter("Distribution System")
            if ds is None or not ds.HasValue or not ds.AsElementId() \
               or ds.AsElementId() == DB.ElementId.InvalidElementId:
                ok = False
        except Exception:
            ok = False
        if not ok:
            continue
        pn = None
        try:
            p = eq.LookupParameter("Panel Name")
            if p and p.HasValue:
                pn = p.AsString()
        except Exception as ex:
            logger.debug("Could not read Panel Name on equipment id {0}: {1}"
                         .format(eq.Id.IntegerValue, ex))
        if pn:
            names.append(pn.strip())
    return _as_sorted_unique(names)

# -----------------------------------------------------------------------------
# 3) Filter creation (ONE filter per panel, OR across category-scoped rules)
# -----------------------------------------------------------------------------
def _param_element_id_for_shared_name(doc, exact_name):
    try:
        for pe in DB.FilteredElementCollector(doc).OfClass(DB.ParameterElement):
            nm = None
            try:
                nm = pe.Name
            except Exception:
                d = pe.GetDefinition()
                if d:
                    nm = d.Name
            if nm and nm == exact_name:
                return pe.Id
    except Exception:
        pass
    return None

def rect_curveloops(x, y, w, h):
    p0 = DB.XYZ(x,     y,     0.0)
    p1 = DB.XYZ(x + w, y,     0.0)
    p2 = DB.XYZ(x + w, y + h, 0.0)
    p3 = DB.XYZ(x,     y + h, 0.0)
    loop = DB.CurveLoop()
    loop.Append(DB.Line.CreateBound(p0, p1))
    loop.Append(DB.Line.CreateBound(p1, p2))
    loop.Append(DB.Line.CreateBound(p2, p3))
    loop.Append(DB.Line.CreateBound(p3, p0))
    outer = List[DB.CurveLoop]()
    outer.Add(loop)
    return outer

def get_or_create_filled_region_type(doc, type_name):
    for frt in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType):
        if DB.Element.Name.__get__(frt) == type_name:
            return frt
    base = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
    if base is None:
        logger.debug("No base FilledRegionType available to duplicate.")
        return None
    try:
        new_frt = base.Duplicate(type_name)  # returns FilledRegionType, NOT ElementId
        return new_frt
    except Exception as ex:
        logger.debug("Failed to duplicate FilledRegionType: {}".format(ex))
        return None


def pick_text_type():
    types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    preferred = None
    fallback = types[0] if types else None
    for t in types:
        try:
            name = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            font = t.get_Parameter(DB.BuiltInParameter.TEXT_FONT).AsString() or ""
            if ("3/32" in name) and ("Arial" in font):
                return t
        except Exception:
            continue
    return preferred or fallback

# -----------------------------------------------------------------------------
# 4) Overrides + apply to a single view (no templates touched)
# -----------------------------------------------------------------------------
def get_or_create_view_template(template_name):
    active_view = revit.active_view
    if not isinstance(active_view, DB.ViewPlan) or active_view.ViewType != DB.ViewType.FloorPlan:
        forms.alert("Active view must be a floor plan to create a view template.",script)
        return None

    # Check for existing template
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType():
        if v.IsTemplate and v.Name == template_name:
            return v

    try:
        template = active_view.CreateViewTemplate()
        template.Name = template_name
        logger.debug("Created new view template from active view: {}".format(active_view.Name))
        return template
    except Exception as ex:
        logger.debug("Failed to create template from active view: {}".format(ex))
        return None

def enforce_template_controls_only_filters(template_view):
    try:
        filters_param_id = None
        param_ids_to_exclude = List[DB.ElementId]()

        for param in template_view.Parameters:
            try:
                defn = param.Definition
                if defn is None:
                    continue

                bip = getattr(defn, "BuiltInParameter", None)
                if bip == DB.BuiltInParameter.VIS_GRAPHICS_FILTERS:
                    filters_param_id = param.Id
                    logger.debug("Identified 'V/G Overrides Filters' correctly.")
                else:
                    param_ids_to_exclude.Add(param.Id)
                    logger.debug("Excluding from control: '{}' (BuiltIn: {})"
                                 .format(defn.Name, bip))

            except Exception as inner_ex:
                logger.debug("Error evaluating parameter: {}".format(inner_ex))

        if filters_param_id is None:
            logger.debug("Failed to find 'V/G Overrides Filters' built-in parameter.")
            return

        template_view.SetNonControlledTemplateParameterIds(param_ids_to_exclude)
        logger.debug("Successfully restricted view template to only control V/G Filters")

    except Exception as ex:
        logger.debug("Failed to restrict template controls: {}".format(ex))


def apply_filters_to_template(view_template, filter_override_data):
    try:
        current_filters = list(view_template.GetFilters())
        for f_id in current_filters:
            filt = doc.GetElement(f_id)
            if not filt or not filt.Name.startswith("PanelChecker -"):
                view_template.RemoveFilter(f_id)
    except Exception as ex:
        logger.debug("Failed to clear old filters: {0}".format(ex))

    for pfe, ogs in filter_override_data:
        try:
            if not view_template.IsFilterApplied(pfe.Id):
                view_template.AddFilter(pfe.Id)
            view_template.SetFilterOverrides(pfe.Id, ogs)
            view_template.SetFilterVisibility(pfe.Id, True)
        except Exception as ex:
            logger.debug("Failed to apply filter '{0}' to template: {1}".format(pfe.Name, ex))

def activate_temp_view_mode(view, template):
    try:
        if view.IsTemporaryViewPropertiesModeEnabled():
            view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryViewProperties)
        #view.ViewTemplateId = DB.ElementId.InvalidElementId  # Clear existing
        view.EnableTemporaryViewPropertiesMode(template.Id)
        logger.debug("Activated temporary view mode with template: {}".format(template.Name))
    except Exception as ex:
        logger.debug("Failed to activate temporary view mode: {}".format(ex))


def build_overrides(color_tuple, use_halftone, solid_fill_pattern_id, lineweight=None):
    r, g, b = color_tuple
    col = _rgb(r, g, b)
    ogs = DB.OverrideGraphicSettings()
    try:
        ogs.SetProjectionLineColor(col)
    except Exception:
        pass
    try:
        ogs.SetSurfaceForegroundPatternId(solid_fill_pattern_id)
        ogs.SetSurfaceForegroundPatternColor(col)
    except Exception:
        pass
    try:
        ogs.SetHalftone(bool(use_halftone))
    except Exception:
        pass
    if lineweight is not None:
        try:
            ogs.SetProjectionLineWeight(lineweight)
        except Exception:
            pass
    return ogs



# -----------------------------------------------------------------------------
# 5) Legend view + filled regions (swatches)
# -----------------------------------------------------------------------------

def get_solid_fill_pattern_id(doc):
    for fpe in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp.Target == DB.FillPatternTarget.Drafting and fp.IsSolidFill:
                return fpe.Id
        except Exception:
            pass
    raise Exception("No Solid Fill pattern found in document.")


def get_filled_region_type(name):
    # Just return the first one available, regardless of name
    fr_type = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
    if not fr_type:
        logger.debug("No FilledRegionType available in document.")
        return None
    return fr_type

def get_filled_region_type_thisisbroken(name):
    for frt in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType):
        if DB.Element.Name.__get__(frt) == name:
            logger.debug("get filled region type - first return")
            return frt

    base = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
    if not base:
        logger.debug("No base FilledRegionType available to duplicate.")
        return None

    new_type = None
    new_type_id = base.Duplicate(name)
    new_ref = DB.Reference(new_type_id)
    new_type = doc.GetElement(new_ref)

    # Must be done *inside* transaction
    new_type.ForegroundPatternId = base.ForegroundPatternId
    new_type.ForegroundPatternColor = base.ForegroundPatternColor
    new_type.BackgroundPatternId = base.BackgroundPatternId
    new_type.BackgroundPatternColor = base.BackgroundPatternColor
    logger.debug("get filled region type - first return")
    return new_type



def get_or_create_drafting_view(view_name, scale_int, template):
    legend_view = next((v for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewDrafting)
                        if v.Name == view_name), None)
    if legend_view:
        logger.debug("legend exists already!")
        return legend_view

    vft = next((x for x in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
                if x.ViewFamily == DB.ViewFamily.Drafting), None)
    if not vft:
        logger.debug("No Drafting ViewFamilyType found.")
        return None

    legend_view = DB.ViewDrafting.Create(doc, vft.Id)
    legend_view.Name = LEGEND_VIEW_NAME
    legend_view.Scale = 48
    legend_view.ViewTemplateId = template.Id
    logger.debug("I CREATED A NEW DRAFTING VIEW!")
    return legend_view

def create_or_update_legend_drafting_view(legend_view, color_map, template, fr_type):

    if not legend_view:
        logger.debug("Legend view creation failed.")
        return

    # Set name, scale, and view template
    try:
        legend_view.Name = LEGEND_VIEW_NAME
        legend_view.Scale = 48
        if template:
            legend_view.ViewTemplateId = template.Id
    except Exception as ex:
        logger.debug("Failed to set legend view props: {}".format(ex))

    if not fr_type:
        logger.debug("No filled region type found.")
        return

    ttype = pick_text_type()
    if not ttype:
        logger.debug("No valid TextNoteType found.")
        return

    view_items = DB.FilteredElementCollector(doc, legend_view.Id)
    deletable_items = [el for el in view_items
                       if isinstance(el,DB.FilledRegion) or isinstance(el,DB.TextNote)]
    for el in deletable_items:
        try:
            doc.Delete(el.Id)
        except Exception as ex:
            logger.debug("Failed to delete element {}: {}".format(el.Id, ex))


    width, height, spacing, text_dx = 1.0, 0.5, 1.0, 1.0

    for idx, (panel_name, rgb) in enumerate(color_map):
        x, y = 0.0, idx * spacing
        loops = rect_curveloops(x, y, width, height)

        try:
            filled = DB.FilledRegion.Create(doc, fr_type.Id, legend_view.Id, loops)
            comment_param = filled.LookupParameter("Comments")
            if comment_param:
                comment_param.Set(panel_name)
        except Exception as ex:
            logger.debug("Failed to create filled region for {}: {}".format(panel_name, ex))

        try:
            pt = DB.XYZ(x + width + text_dx, y + height * 0.5, 0.0)
            opts = DB.TextNoteOptions()
            opts.TypeId = ttype.Id
            opts.VerticalAlignment = DB.VerticalTextAlignment.Middle
            DB.TextNote.Create(doc, legend_view.Id, pt, panel_name, opts)
        except Exception as ex:
            logger.debug("Failed to create text note for {}: {}".format(panel_name, ex))



def create_or_update_panel_filter_logical_or(panel_name):
    """Creates or updates a ParameterFilterElement with a Logical OR of per-category ElementParameterFilters.
       Special handling: if panel_name == "__NO_PANEL__", then:
         - Electrical Equipment: Supply From == ""
         - Fixtures/Devices: Panel == ""
         - Detail Items: Comments == "__NO PANEL__"
    """
    is_no_panel = (panel_name == NO_PANEL_FLAG)
    filter_label = "No Panel" if is_no_panel else panel_name
    filter_name = "Panel Checker - {0}".format(filter_label)
    logger.debug("Building filter: {0}".format(filter_name))


    or_filters = []
    cat_ids = []

    for bic, bip in POWER_FIXTURE_EQUIP.items():
        try:
            cat = doc.Settings.Categories.get_Item(bic)
            if not cat:
                continue
            cat_id = cat.Id
            cat_ids.append(cat_id)

            # --- Special case handling ---
            if is_no_panel:
                if bic == DB.BuiltInCategory.OST_ElectricalEquipment:
                    bip = DB.BuiltInParameter.RBS_ELEC_PANEL_SUPPLY_FROM_PARAM
                    match_val = ""
                elif bic == DB.BuiltInCategory.OST_DetailComponents:
                    match_val = NO_PANEL_FLAG
                else:
                    match_val = ""
            else:
                # Normal case: match panel name
                match_val = panel_name

            # Build rules (same style as before)
            cat_rule = DB.FilterCategoryRule(List[DB.ElementId]([cat_id]))
            val_rule = DB.FilterStringRule(
                DB.ParameterValueProvider(DB.ElementId(bip)),
                DB.FilterStringEquals(),
                match_val,
            )
            filter_rules = List[DB.FilterRule]([cat_rule, val_rule])
            epf = DB.ElementParameterFilter(filter_rules)

            or_filters.append(epf)
            logger.debug("Added rule for category: {0}".format(bic.ToString()))

        except Exception as ex:
            logger.debug("Failed to build rule for category {0}: {1}".format(bic.ToString(), ex))

    if not or_filters:
        logger.debug("No valid filters for panel '{0}'".format(filter_label))
        return None

    try:
        final_filter = DB.LogicalOrFilter(List[DB.ElementFilter](or_filters))
    except Exception as ex:
        logger.debug("Failed to create LogicalOrFilter for panel '{0}': {1}".format(filter_label, ex))
        return None

    # Reuse or create filter element
    existing = None
    for pfe in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
        if pfe.Name == filter_name:
            existing = pfe
            break

    try:
        if existing:
            existing.SetCategories(List[DB.ElementId](cat_ids))
            existing.SetElementFilter(final_filter)
            logger.debug("Updated existing filter: {0}".format(filter_name))
            return existing
        else:
            new_pfe = DB.ParameterFilterElement.Create(doc, filter_name, List[DB.ElementId](cat_ids))
            new_pfe.SetElementFilter(final_filter)
            logger.debug("Created new filter: {0}".format(filter_name))
            return new_pfe
    except Exception as ex:
        logger.debug("Failed to create or update filter '{0}': {1}".format(filter_name, ex))
        return None




# -----------------------------------------------------------------------------
# 7) Main
# -----------------------------------------------------------------------------
def main():
    active_view = revit.active_view
    if not isinstance(active_view, DB.ViewPlan) or active_view.ViewType != DB.ViewType.FloorPlan:
        forms.alert("Active view must be a floor plan to create a view template.",exitscript=True)

    all_panels = get_all_panel_names(doc)
    if not all_panels:
        forms.alert("No panel names found on Electrical Equipment with a Distribution System. Nothing to do.")
        return

    selection = forms.SelectFromList.show(
        all_panels, multiselect=True,
        title="Color Circuits by Panel: choose included panels",
        width=500, height=400, button_name="Use Selected"
    )
    if not selection:
        return

    include_set = set(selection)
    include_set.add(NO_PANEL_FLAG)
    solid_id = get_solid_fill_pattern_id(doc)

    template = None
    legend_view = None
    color_map = {}
    filter_override_data = []

    with DB.TransactionGroup(doc, "Color Circuits by Panel") as tg:
        tg.Start()

        with DB.Transaction(doc, "Create/Update Filters") as tx1:
            tx1.Start()
            color_index = 0  # counter for included panels

            # --- Normal panels ---
            for pn in all_panels:
                if pn in include_set:
                    rgb = COLOR_PALETTE[color_index % len(COLOR_PALETTE)]
                    color_index += 1
                    use_halftone = False
                else:
                    rgb = GREY
                    use_halftone = True

                color_map[pn] = rgb
                pfe = create_or_update_panel_filter_logical_or(pn)
                if pfe:
                    ogs = build_overrides(rgb, use_halftone, solid_id)
                    filter_override_data.append((pfe, ogs))

            # --- Special "No Panel" filter ---
            if NO_PANEL_FLAG in include_set:
                rgb = RED
                use_halftone = False
                color_map[NO_PANEL_FLAG] = rgb
                pfe = create_or_update_panel_filter_logical_or(NO_PANEL_FLAG)
                if pfe:
                    ogs = build_overrides(rgb, use_halftone, solid_id, lineweight=8)
                    filter_override_data.append((pfe, ogs))

            tx1.Commit()

        with DB.Transaction(doc, "View Template Setup") as tx2:
            tx2.Start()
            template = get_or_create_view_template(VIEW_TEMPLATE_NAME)
            if template:
                enforce_template_controls_only_filters(template)
                apply_filters_to_template(template, filter_override_data)
            tx2.Commit()

        with DB.Transaction(doc, "Legend Creation") as tx21:
            tx21.Start()
            legend_view = get_or_create_drafting_view(LEGEND_VIEW_NAME, 48, template)
            tx21.Commit()

        with DB.Transaction(doc, "Legend Creation2") as tx3:
            tx3.Start()
            # Before calling legend view function
            fr_type = get_filled_region_type("PanelChecker - Solid Fill")

            if not fr_type:
                logger.debug("No filled region type could be created.")
                return
            create_or_update_legend_drafting_view(legend_view, list(color_map.items()), template, fr_type)

            tx3.Commit()

        with DB.Transaction(doc, "Activate View Template") as tx4:
            tx4.Start()
            if template:
                activate_temp_view_mode(revit.active_view, template)
            tx4.Commit()

        tg.Assimilate()

    forms.alert("Filters applied to view template '{0}'. View switched to Temporary Template Mode.".format(VIEW_TEMPLATE_NAME))

if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        logger.exception("Color Circuits by Panel failed: {0}".format(ex))
