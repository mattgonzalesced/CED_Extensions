# -*- coding: utf-8 -*-
import os
from pyrevit import revit, DB, script, UI
from pyrevit import forms
from pyrevit.interop import xl as pyxl
from pyrevit.revit import create
from System.Collections.Generic import List
import Autodesk.Revit.DB.Electrical as DBE
from pyrevit.framework import clr


# Setup
uidoc = revit.uidoc
doc = revit.doc
app = doc.Application
uiapp = revit.uidoc.Application

logger = script.get_logger()
output = script.get_output()

# Paths
acc_web = r"https://acc.autodesk.com/docs/files/projects/c0252f4b-fd26-43c1-b4ac-0849b65ec8c7?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.SD11yNeBTBeiIII3QK5NoA&viewModel=detail&moduleId=folders"
user_folder = os.path.expanduser('~')
content_folder = "Content"
script_dir = os.path.dirname(__file__)
content_dir = os.path.join(script_dir, content_folder)


def safely_load_shared_parameter_file(app, shared_param_txt):
    """
    Load a shared parameter file without overwriting user's config.
    Returns the opened shared_param_file and original file path.
    Exits if loading fails.
    """
    original_file = app.SharedParametersFilename

    if not os.path.exists(shared_param_txt):
        logger.warning("❌ Shared parameter file missing: {}".format(shared_param_txt))
        script.exit()

    if not original_file:
        logger.info("No shared parameter file currently set. Setting shared param path to project file.")

    app.SharedParametersFilename = shared_param_txt
    shared_param_file = app.OpenSharedParameterFile()

    if not shared_param_file:
        logger.warning("❌ Failed to load shared parameter file: {}".format(shared_param_txt))
        script.exit()

    return shared_param_file, original_file


def load_excel(config_path):
    # Load Excel and define expected columns
    COLUMNS = ["GUID", "UniqueId", "Parameter Name", "Discipline", "Type of Parameter", "Group Under", "Instance/Type",
               "Categories", "Groups"]
    logger.info("Loading parameter table from: {}".format(config_path))
    xldata = pyxl.load(config_path, headers=False)
    sheet = xldata.get("Parameter List")
    if not sheet:
        forms.alert("Sheet 'Parameter List' not found in Excel file.")
        script.exit()

    sheetdata = [dict(zip(COLUMNS, row)) for row in sheet['rows'][1:] if len(row) >= len(COLUMNS)]
    logger.info("✅ Loaded {} rows from 'Parameter List'.".format(len(sheetdata)))

    # Sort by UniqueId
    sheetdata = sorted(sheetdata, key=lambda row: row.get("UniqueId", ""))

    return sheetdata



# Group mapping

GROUP_MAP = {
    'Electrical': DB.GroupTypeId.Electrical,
    'Identity Data': DB.GroupTypeId.IdentityData,
    'Electrical - Circuiting': DB.GroupTypeId.ElectricalCircuiting
}
cat_map = {cat.Name: cat for cat in doc.Settings.Categories}


# --- Helpers ---
def get_definition(name, shared_param_file):
    for group in shared_param_file.Groups:
        definition = group.Definitions.get_Item(name)
        if definition:
            return definition
    return None


def get_existing_binding(defn):
    iterator = doc.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        if iterator.Key.Name == defn.Name:
            return iterator.Current
    return None


def build_category_set(names):
    cats = DB.CategorySet()
    for name in names:
        cat = cat_map.get(name)
        if cat:
            cats.Insert(cat)
        else:
            logger.warning("⚠ Category '{}' not found.".format(name))
    return cats


# --- Main Operation ---
def process_param_row(row, shared_param_file):
    name = row['Parameter Name']
    group_label = row['Group Under']
    is_instance = row['Instance/Type'].lower() == 'instance'
    categories = [c.strip() for c in row['Categories'].split(',') if c.strip()]

    output.print_md("*Adding parameter* **{}**...".format(name))

    definition = get_definition(name, shared_param_file)
    if not definition:
        output.print_md("❌ **Shared param '{}' not found.**".format(name))
        return

    param_group = GROUP_MAP.get(group_label, None)
    expected_group_id = param_group
    catset = build_category_set(categories)
    bindmap = doc.ParameterBindings

    existing_binding = get_existing_binding(definition)
    if existing_binding:
        is_current_instance = isinstance(existing_binding, DB.InstanceBinding)
        current_cats = set([c.Id for c in existing_binding.Categories])
        target_cats = set([c.Id for c in catset])

        needs_update = not (
                is_current_instance == is_instance and
                current_cats == target_cats
        )
    else:
        needs_update = True

    if not needs_update:
        output.print_md("☑️ Parameter **{}** already configured correctly. Skipping.".format(name))
        return

    # --- Transaction: Insert or Update Binding ---
    try:
        binding = DB.InstanceBinding(catset) if is_instance else DB.TypeBinding(catset)
        if existing_binding:
            output.print_md("🔁 Parameter **{}** exists. Updating binding...".format(name))
            bindmap.ReInsert(definition, binding, param_group)
        else:
            bindmap.Insert(definition, binding, param_group)
        output.print_md("✅ Parameter **{}** bound successfully.".format(name))
    except Exception as e:
        logger.error("❌ Exception binding '{}': {}".format(name, e))

        return


# --- Load Families ---
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = True
        return True


def load_rfa_families_from_content_folder(content_dir):
    output.print_md("## Loading Families")
    loaded_count = 0

    # Collect existing family names before loading
    existing_names = set(f.Name for f in DB.FilteredElementCollector(doc).OfClass(DB.Family))

    family_files = [f for f in os.listdir(content_dir) if f.lower().endswith(".rfa")]

    for fname in family_files:
        family_path = os.path.join(content_dir, fname)
        fam_name = os.path.splitext(fname)[0]

        if fam_name in existing_names:
            output.print_md("🔁 Family **{}** already loaded. Skipping.".format(fam_name))
            continue
        with DB.Transaction(doc, "Load RFA Families"):
            try:

                success = doc.LoadFamily(family_path)
                # Confirm it actually got added
                if success:
                    # Re-check that it's really now in the doc
                    just_loaded = next(
                        (f for f in DB.FilteredElementCollector(doc).OfClass(DB.Family)
                         if f.Name == fam_name), None)

                    if just_loaded:
                        output.print_md("✅ Family **{}** loaded successfully.".format(fam_name))
                        loaded_count += 1
                    else:
                        output.print_md("❌ Load reported success but family not found: **{}**".format(fam_name))
                else:
                    output.print_md("⚠ Failed to load family **{}**.".format(fam_name))

            except Exception as e:
                logger.error("❌ Error loading '{}': {}".format(fname, e))

    output.print_md("📦 **Total families loaded: {}**".format(loaded_count))


def collect_schedule_ids(source_doc):
    source = DB.FilteredElementCollector(source_doc) \
        .OfClass(DB.ViewSchedule) \
        .WhereElementIsNotElementType() \
        .ToElements()

    existing_names = set([
        s.Name for s in DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSchedule)
        .WhereElementIsNotElementType()
    ])

    collected_ids = List[DB.ElementId]()
    for s in source:
        if s.Name in existing_names:
            output.print_md("🔁 Schedule **{}** already exists. Skipping. (ID: `{}`)".format(s.Name, s.Id.Value))

        else:
            output.print_md("✅ Schedule **{}** will be copied.".format(s.Name))
            collected_ids.Add(s.Id)

    return collected_ids


def build_name_filter(name):
    provider = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME))
    rule = DB.FilterStringRule(provider, DB.FilterStringContains(), name)
    return DB.ElementParameterFilter(rule)


def collect_panel_template_ids(source_doc):
    output.print_md("## 🔍 Scanning Panel Schedule Templates")

    source_templates = [
        t for t in DB.FilteredElementCollector(source_doc)
        .OfClass(DBE.PanelScheduleTemplate)
        if "_CED" in t.Name
    ]

    output.print_md("📋 **Source panel templates matching '_CED':**")
    for t in source_templates:
        output.print_md("- {}".format(t.Name))

    target_templates = [
        t for t in DB.FilteredElementCollector(doc)
        .OfClass(DBE.PanelScheduleTemplate)
        if "_CED" in t.Name
    ]

    existing_names = [t.Name for t in target_templates]
    output.print_md("📂 **Existing templates in current project:**")
    for name in existing_names:
        output.print_md("- {}".format(name))

    combined = List[DB.ElementId]()
    for t in source_templates:
        if t.Name in existing_names:
            output.print_md("🔁 Panel Template **{}** already exists. Skipping.".format(t.Name))
        else:
            output.print_md("✅ Panel Template **{}** will be copied.".format(t.Name))
            combined.Add(t.Id)

    output.print_md("📋 **Total panel templates copied: {}**".format(combined.Count))
    return combined


def collect_filter_ids(source_doc):
    output.print_md("## 🔍 Scanning View Filters")

    source_filters = [
        f for f in DB.FilteredElementCollector(source_doc)
        .OfClass(DB.ParameterFilterElement)
        if "_CED" in f.Name
    ]

    output.print_md("🧾 **Source filters matching '_CED':**")
    for f in source_filters:
        output.print_md("- {}".format(f.Name))

    target_filters = [
        f for f in DB.FilteredElementCollector(doc)
        .OfClass(DB.ParameterFilterElement)
        if "_CED" in f.Name
    ]

    existing_names = [f.Name for f in target_filters]
    output.print_md("📂 **Existing filters in current project:**")
    for f in existing_names:
        output.print_md("- {}".format(f))

    combined = List[DB.ElementId]()
    for f in source_filters:
        if f.Name in existing_names:
            output.print_md("🔁 View Filter **{}** already exists. Skipping.".format(f.Name))
        else:
            output.print_md("✅ View Filter **{}** will be copied.".format(f.Name))
            combined.Add(f.Id)

    output.print_md("🧱 **Total view filters copied: {}**".format(combined.Count))
    return combined


def copy_elements_from_document(source_doc, element_ids, description="Copied Elements"):
    if not element_ids or element_ids.Count == 0:
        output.print_md("⚠ No elements to copy: {}".format(description))
        return

    with revit.Transaction(description, doc, swallow_errors=True):
        try:
            create.copy_elements(element_ids, source_doc, doc)
            output.print_md("✅ {}: **{}** elements copied.".format(description, element_ids.Count))
        except Exception as e:
            logger.error("❌ Failed to copy {}: {}".format(description, e))


def get_elements_in_view(doc_source, view):
    """Returns filtered ElementIds from a view in the given document."""
    allowed_categories = [
        DB.BuiltInCategory.OST_IOSModelGroups,
        DB.BuiltInCategory.OST_IOSAttachedDetailGroups,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_KeynoteTags,
        DB.BuiltInCategory.OST_ElectricalFixtureTags
    ]
    category_filters = [DB.ElementCategoryFilter(cat) for cat in allowed_categories]
    combined_filter = DB.LogicalOrFilter(category_filters)
    collector = DB.FilteredElementCollector(doc_source, view.Id) \
        .WhereElementIsNotElementType() \
        .WherePasses(combined_filter) \
        .ToElementIds()
    return collector


def get_destination_view(doc_target):
    """Ensures destination is a floor plan view. Returns the valid View."""
    active_view = doc_target.ActiveView
    if active_view.ViewType == DB.ViewType.FloorPlan:
        return active_view

    # Find fallback
    floorplans = DB.FilteredElementCollector(doc_target) \
        .OfClass(DB.View) \
        .WhereElementIsNotElementType() \
        .ToElements()

    floorplans = [v for v in floorplans if v.ViewType == DB.ViewType.FloorPlan and not v.IsTemplate]

    if not floorplans:
        forms.alert("No floor plan views found in the current document.", exitscript=True)
        return None

    fallback = floorplans[0]
    output.print_md("⚠️ Active view is not a floor plan. Using fallback view: **{}**".format(fallback.Name))
    return fallback


def copy_groups(source_view, element_ids, dest_doc, dest_view):
    """Copies elements from source_doc to dest_doc, then deletes them."""
    options = DB.CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(create.CopyUseDestination())
    output.print_md("## Copying Groups")
    with revit.Transaction("Copy Elements From Starter", doc=dest_doc, swallow_errors=True):
        copied_ids = DB.ElementTransformUtils.CopyElements(
            source_view,  # Document to copy from
            element_ids,  # ICollection<ElementId>
            dest_view,  # Document to copy to
            None,  # No transform
            options
        )
        output.print_md("✅ Copied {} groups into view: **{}**".format(len(copied_ids), dest_view.Name))
    return copied_ids


def delete_feeder_key_schedule(doc):
    """
    Deletes the schedule named '_Electrical Feeder Key Sched1' safely.
    """
    try:
        schedule = DB.FilteredElementCollector(doc) \
            .OfClass(DB.ViewSchedule) \
            .WhereElementIsNotElementType() \
            .ToElements()

        target_sched = next((s for s in schedule if s.Name == "_Electrical Feeder Key Sched1"), None)

        if not target_sched:
            output.print_md("⚠️ Schedule '_Electrical Feeder Key Sched1' not found — nothing to delete.")
            return

        if target_sched.IsTemplate:
            output.print_md("⚠️ Schedule is a template and cannot be deleted.")
            return

        if doc.ActiveView.Id == target_sched.Id:
            output.print_md("⚠️ Cannot delete schedule that is currently active.")
            return

        with revit.Transaction("Delete Feeder Key Schedule", doc):
            deleted = doc.Delete(target_sched.Id)
            output.print_md("🗑️ Deleted schedule copied schedule")
    except Exception as e:
        output.print_md("❌ Exception during schedule deletion: {}".format(e))
        logger.error("Schedule deletion failed: {}".format(e))


def set_leader_arrowheads_for_tags(doc):
    """
    Sets the LEADER_ARROWHEAD parameter for specific tag symbols
    to the arrowhead named 'Dot Open 1/16"'.
    Prints all arrowhead types for inspection.
    """
    arrowhead = None

    # Step 1: Inspect all ElementTypes for Arrowheads
    for elem in DB.FilteredElementCollector(doc).OfClass(DB.ElementType).WhereElementIsElementType():
        family_name = DB.ElementType.FamilyName
        name = DB.Element.Name.__get__(elem)

        if name == 'Dot Open 1/16"':
            arrowhead = elem
            output.print_md("🎯 Found match: '{}'".format(name))
            break

    if not arrowhead:
        output.print_md("❌ Arrowhead named 'Dot Open 1/16\"' not found.")
        return

    arrow_id = arrowhead.Id

    # Step 2: Define target tag symbols (Family Name, Type Name)
    target_tags = [
        ("EF-Tag_Wire Size_CED", "Wire Size ID (Boxed)"),
        ("EF-Tag_Panel & Circuit_CED", "Panel & Circuit Number (Open Dot)")
    ]

    updated = []

    with revit.Transaction("Set Leader Arrowheads", doc):
        for fam_name, type_name in target_tags:
            tag_symbol = next(
                (t for t in DB.FilteredElementCollector(doc)
                .OfClass(DB.FamilySymbol)
                .WhereElementIsElementType()
                 if DB.ElementType.FamilyName.__get__(t) == fam_name and DB.Element.Name.__get__(t) == type_name),
                None
            )

            if tag_symbol:
                param = tag_symbol.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD)
                if param and param.StorageType == DB.StorageType.ElementId:
                    param.Set(arrow_id)
                    updated.append("{} : {}".format(fam_name, type_name))
            else:
                output.print_md("⚠️ Tag symbol not found: '{} : {}'".format(fam_name, type_name))

    if updated:
        output.print_md("✅ Updated tag symbols:\n- " + "\n- ".join(updated))
    else:
        output.print_md("⚠️ No tag symbols updated.")


def delete_groups(dest_doc, copied_ids):
    with revit.Transaction("Delete Copied Elements", doc=dest_doc):
        deleted_ids = dest_doc.Delete(copied_ids)
        output.print_md("🗑️ Deleted {} copied elements.".format(len(deleted_ids)))


def main():
    shared_param_txt = os.path.join(content_dir, 'ELEC SHARED PARAMS.txt')
    config_excel_path = os.path.join(content_dir, 'ELEC SHARED PARAM TABLE.xlsx')



    forms.alert(
        title="Load Electrical Parameters",
        msg="This tool loads all required parameters for the 'Calculate Circuits' tool."
            "\n\nPlease Sync before starting!"
            "\n\nDo you wish to continue?",
        ok=True,
        cancel=True,
        warn_icon=True,
        exitscript=True
    )

    output.close_others()
    output.show()

    sheetdata = load_excel(config_excel_path)

    with DB.TransactionGroup(doc, "Load Shared Parameters and Content") as tg:
        tg.Start()
        new_param_file, original_param_file = safely_load_shared_parameter_file(app, shared_param_txt)
        output.print_md("## Adding Parameters")
        with revit.Transaction("Bind Shared Parameters", doc):
            for row in sheetdata:
                process_param_row(row, new_param_file)

        if original_param_file:
            app.SharedParametersFilename = original_param_file
            logger.info("🔄 Restored shared parameter file: {}".format(original_param_file))
        else:
            logger.info("🔄No original parameter file. keeping new one set: {}".format(new_param_file))

        tg.Assimilate()


# === Entry Point ===
if __name__ == '__main__':
    main()
