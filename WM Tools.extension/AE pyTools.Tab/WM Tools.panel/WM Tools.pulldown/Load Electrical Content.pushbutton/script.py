# -*- coding: utf-8 -*-
import os
from pyrevit import revit, DB, script, UI
from pyrevit import forms
from pyrevit.interop import xl as pyxl
from pyrevit.revit import create
from System.Collections.Generic import List
import Autodesk.Revit.DB.Electrical as DBE

# Setup
uidoc = revit.uidoc
doc = revit.doc
app = doc.Application
uiapp = revit.uidoc.Application

logger = script.get_logger()
output = script.get_output()



# Paths
client_std_web = r"https://coolsysinc.sharepoint.com/sites/Teams-CEDClientStandardsandAdmin/Shared%20Documents/Forms/AllItems.aspx"
user_folder = os.path.expanduser('~')
content_folder = r"CoolSys Inc\Teams-CED Client Standards and Admin - Documents\Walmart\Revit Startup Tools\CED Elec Tools\Content"
content_folder2 = r"\OneDrive - CoolSys Inc\Documents - Teams-CED Client Standards and Admin\Walmart\Revit Startup Tools\CED Elec Tools\Content"
script_dir = os.path.dirname(__file__)



# --- Content Folder Path Resolution ---
def resolve_content_folder():
    candidates = [
        os.path.join(user_folder, content_folder),
        os.path.join(user_folder, content_folder2)
    ]

    for path in candidates:
        if os.path.exists(path):
            logger.info("✅ Found content folder at: {}".format(path))
            return path
    example_path = os.path.join(user_folder, content_folder)
    logger.warning("❌ Could not find content folder in either known OneDrive path.")
    show_sharepoint_sync_instructions(example_path)
    script.exit()



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
    COLUMNS = ["GUID", "UniqueId", "Parameter Name","Discipline","Type of Parameter", "Group Under", "Instance/Type", "Categories", "Groups"]
    logger.info("Loading parameter table from: {}".format(config_path))
    xldata = pyxl.load(config_path, headers=False)
    sheet = xldata.get("Parameter List")
    if not sheet:
        forms.alert("Sheet 'Parameter List' not found in Excel file.")
        script.exit()

    sheetdata = [dict(zip(COLUMNS, row)) for row in sheet['rows'][1:] if len(row) >= len(COLUMNS)]
    logger.info("✅ Loaded {} rows from 'Parameter List'.".format(len(sheetdata)))
    return sheetdata


# Group mapping
GROUP_MAP = {
    'Electrical': DB.BuiltInParameterGroup.PG_ELECTRICAL,
    'Identity Data': DB.BuiltInParameterGroup.PG_IDENTITY_DATA,
    'Electrical - Circuiting': DB.BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING
}

cat_map = {cat.Name: cat for cat in doc.Settings.Categories}


# --- Helpers ---
def get_definition(name,shared_param_file):
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
def process_param_row(row,shared_param_file):
    name = row['Parameter Name']
    group_label = row['Group Under']
    is_instance = row['Instance/Type'].lower() == 'instance'
    varies = row['Groups'].lower().startswith('values can vary')
    categories = [c.strip() for c in row['Categories'].split(',') if c.strip()]

    output.print_md("### Adding Parameter")
    output.print_md("*Adding parameter* **{}**...".format(name))

    definition = get_definition(name,shared_param_file)
    if not definition:
        output.print_md("❌ **Shared param '{}' not found.**".format(name))
        return

    param_group = GROUP_MAP.get(group_label, DB.BuiltInParameterGroup.INVALID)
    expected_group_id = DB.ElementId(param_group.value__)

    catset = build_category_set(categories)
    bindmap = doc.ParameterBindings

    existing_binding = get_existing_binding(definition)
    if existing_binding:
        is_current_instance = isinstance(existing_binding, DB.InstanceBinding)
        current_cats = set([c.Id.IntegerValue for c in existing_binding.Categories])
        target_cats = set([c.Id.IntegerValue for c in catset])

        vary_matches = True
        if is_instance and varies:
            try:
                vary_matches = definition.InternalDefinition.VariesAcrossGroups
            except:
                vary_matches = True

        needs_update = not (
                is_current_instance == is_instance and
                current_cats == target_cats and
                vary_matches
        )
    else:
        needs_update = True

    if not needs_update:
        output.print_md("☑️ Parameter **{}** already configured correctly. Skipping.".format(name))
        return

    # --- Transaction: Insert or Update Binding ---
    t = DB.Transaction(doc, "Bind: {}".format(name))
    t.Start()
    try:
        binding = DB.InstanceBinding(catset) if is_instance else DB.TypeBinding(catset)
        if existing_binding:
            output.print_md("🔁 Parameter **{}** exists. Updating binding...".format(name))
            bindmap.ReInsert(definition, binding, param_group)
        else:
            bindmap.Insert(definition, binding, param_group)
        t.Commit()
        output.print_md("✅ Parameter **{}** bound successfully.".format(name))
    except Exception as e:
        logger.error("❌ Exception binding '{}': {}".format(name, e))
        t.RollBack()
        return

    # --- Transaction: Set Vary Across Groups ---
    if is_instance and varies:
        t2 = DB.Transaction(doc, "Set Vary: {}".format(name))
        t2.Start()
        try:
            internal_def = definition.InternalDefinition if hasattr(definition, "InternalDefinition") else None
            if internal_def and hasattr(internal_def, "SetAllowVaryBetweenGroups"):
                failures = internal_def.SetAllowVaryBetweenGroups(doc, True)
                if failures and failures.Count > 0:
                    logger.info("⚠ Some failed to set vary for '{}'".format(name))
                else:
                    logger.info("✔ Set 'Can Vary by Group' for '{}'".format(name))
            else:
                logger.info("❌ Cannot set vary for '{}': No InternalDefinition".format(name))
            t2.Commit()
        except Exception as vary_error:
            logger.error("Failed setting vary on '{}': {}".format(name, vary_error))
            t2.RollBack()


# --- Load Families ---
def load_rfa_families_from_content_folder(content_dir):
    output.print_md("### Loading Families")
    loaded_count = 0

    # Collect existing family names
    existing_names = set([
        f.Name for f in DB.FilteredElementCollector(doc).OfClass(DB.Family)
    ])

    for fname in os.listdir(content_dir):
        if not fname.lower().endswith(".rfa"):
            continue

        family_path = os.path.join(content_dir, fname)
        fam_name = os.path.splitext(fname)[0]  # Strip .rfa

        if fam_name in existing_names:
            output.print_md("🔁 Family **{}** already loaded. Skipping.".format(fam_name))
            continue

        try:
            with DB.Transaction(doc, "Load Family: {}".format(fam_name)) as t:
                t.Start()
                success = doc.LoadFamily(family_path)
                if success:
                    output.print_md("✅ Family **{}** loaded successfully.".format(fam_name))
                    loaded_count += 1
                else:
                    output.print_md("⚠ Failed to load family **{}**.".format(fam_name))
                t.Commit()
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
            output.print_md("🔁 Schedule **{}** already exists. Skipping. (ID: `{}`)".format(s.Name, s.Id.IntegerValue))

        else:
            output.print_md("✅ Schedule **{}** will be copied.".format(s.Name))
            collected_ids.Add(s.Id)

    return collected_ids


def build_name_filter(name):
    provider = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME))
    rule = DB.FilterStringRule(provider, DB.FilterStringContains(), name)
    return DB.ElementParameterFilter(rule)


def collect_panel_template_ids(source_doc):
    output.print_md("### 🔍 Scanning Panel Schedule Templates with '_CED' in the name")

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
    output.print_md("### 🔍 Scanning View Filters with '_CED' in the name")

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

    t = DB.Transaction(doc, description)
    t.Start()
    try:
        create.copy_elements(element_ids,source_doc,doc)
        t.Commit()
        output.print_md("✅ {}: **{}** elements copied.".format(description, element_ids.Count))
    except Exception as e:
        logger.error("❌ Failed to copy {}: {}".format(description, e))
        t.RollBack()


def show_sharepoint_sync_instructions(content_dir):
    output.freeze()
    output.print_md("### ⚠️ SharePoint Content Not Synced\n")

    output.print_md(
        "This tool cannot proceed because the required content folder is missing from your machine.\n\n"
        "The content should be synced from the **CED Client Standards and Admin** SharePoint site."
    )

    output.print_md("## 🔧 How to fix this issue:\n")

    output.print_md(
        "1. Click the link below to open the SharePoint folder in your browser:\n"
        "    📁 [Open SharePoint Folder in your Web Browser]({})\n\n"
        "2. In the browser, click **'Sync'** (usually found near the top menu).\n\n"
        "3. Once sync completes, confirm that the following folder now exists:\n"
        "    📂 [{}]({})\n\n"
        "4. Once verified, rerun the **Load Electrical Content** tool.\n".format(
            client_std_web,
            content_dir, content_dir.replace("\\", "/")
        )
    )

    output.print_md("\n\n#### 🔍 Example: Where to click 'Sync' in SharePoint\n")
    sync_img = os.path.join(script_dir, "sync_instruction.png")
    if os.path.exists(sync_img):
        output.print_image(sync_img)
    else:
        logger.warning("Sync help image not found: {}".format(sync_img))

    output.print_md("\n\n#### 🗂 Example: Synced folder location in File Explorer\n")
    explorer_img = os.path.join(script_dir, "file_explorer_instruction.png")
    if os.path.exists(explorer_img):
        output.print_image(explorer_img)
    else:
        logger.warning("File explorer instruction image not found: {}".format(explorer_img))
    output.unfreeze()

def main():

    content_dir = resolve_content_folder()

    shared_param_txt = os.path.join(content_dir, 'WM ELEC SHARED PARAMS.txt')
    config_excel_path = os.path.join(content_dir, 'WM ELEC SHARED PARAM TABLE.xlsx')

    new_param_file, original_param_file = safely_load_shared_parameter_file(app,shared_param_txt)

    forms.alert(
        title="WM Load Content",
        msg="This tool loads all required electrical content for WM projects."
            "\n\nWhen complete, it will open a starter project so you can manually copy model groups."
            "\n\nDo you wish to continue?",
        ok=True,
        cancel=True,
        warn_icon=True,
        exitscript=True
    )

    starter_path = os.path.join(content_dir, "WM ELEC STARTER.rvt")
    output.close_others()
    output.show()

    if not os.path.exists(starter_path):
        output.print_md("❌ Could not find **WM ELEC STARTER.rvt** in content folder.")
        return

    sheetdata = load_excel(config_excel_path)
    starter_doc = app.OpenDocumentFile(starter_path)

    with DB.TransactionGroup(doc, "Load Shared Parameters and Content") as tg:
        tg.Start()

        for row in sheetdata:
            process_param_row(row,new_param_file)

        load_rfa_families_from_content_folder(content_dir)
        schedule_ids = collect_schedule_ids(starter_doc)
        panel_template_ids = collect_panel_template_ids(starter_doc)
        view_filter_ids = collect_filter_ids(starter_doc)

        all_ids = List[DB.ElementId]()
        all_ids.AddRange(schedule_ids)
        all_ids.AddRange(panel_template_ids)
        all_ids.AddRange(view_filter_ids)
        copy_elements_from_document(starter_doc, all_ids, "Import Schedules, Templates, Filters")
        tg.Assimilate()

    if original_param_file:
        app.SharedParametersFilename = original_param_file
        logger.info("🔄 Restored shared parameter file: {}".format(original_param_file))
    else:
        logger.info("🔄No original parameter file. keeping new one set: {}".format(new_param_file))

    try:
        uiapp.OpenAndActivateDocument(starter_path)
        output.print_md("📂 Opened **WM ELEC STARTER.rvt**. You can now manually copy the groups.")
    except Exception as e:
        logger.error("❌ Failed to open starter file interactively: {}".format(e))

    output.show()
def main2():

    content_dir = resolve_content_folder()

    shared_param_txt = os.path.join(content_dir, 'WM ELEC SHARED PARAMS.txt')
    config_excel_path = os.path.join(content_dir, 'WM ELEC SHARED PARAM TABLE.xlsx')

    new_param_file, original_param_file = safely_load_shared_parameter_file(app,shared_param_txt)


    starter_path = os.path.join(content_dir, "WM ELEC STARTER.rvt")
    output.close_others()
    output.show()

    if not os.path.exists(starter_path):
        output.print_md("❌ Could not find **WM ELEC STARTER.rvt** in content folder.")
        return

    sheetdata = load_excel(config_excel_path)

    with DB.TransactionGroup(doc, "Load Shared Parameters and Content") as tg:
        tg.Start()

        for row in sheetdata:
            process_param_row(row,new_param_file)

        tg.Assimilate()

    if original_param_file:
        app.SharedParametersFilename = original_param_file
        logger.info("🔄 Restored shared parameter file: {}".format(original_param_file))
    else:
        logger.info("🔄No original parameter file. keeping new one set: {}".format(new_param_file))

    output.show()


# === Entry Point ===
if __name__ == '__main__':
    param_only = 0
    if param_only == 1:
        main2()
    else:
        main()
        script.exit()
