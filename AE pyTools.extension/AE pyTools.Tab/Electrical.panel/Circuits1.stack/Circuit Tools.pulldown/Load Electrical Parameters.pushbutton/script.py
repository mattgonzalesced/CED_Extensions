# -*- coding: utf-8 -*-
import os

from pyrevit import forms
from pyrevit import revit, DB, script
from pyrevit.interop import xl as pyxl

# Setup
uidoc = revit.uidoc
doc = revit.doc
app = doc.Application
uiapp = revit.uidoc.Application

logger = script.get_logger()
output = script.get_output()

# Paths
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
