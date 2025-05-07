# -*- coding: utf-8 -*-
import clr
import os

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager

# --- Robust Context Setup ---
uidoc = DocumentManager.Instance.CurrentUIDocument
doc = DocumentManager.Instance.CurrentDBDocument

# Fallback if PyRevit context fails
if not uidoc or not doc:
    try:
        uidoc = __revit__.ActiveUIDocument
        doc = uidoc.Document
    except:
        raise Exception("This script must be run from within an open Revit project.")

# Try getting the Application object safely
try:
    app = uidoc.Application.Application
except:
    raise Exception("Failed to access Revit application context. Restart Revit or re-run script.")

# --- Shared Parameters File ---
script_dir = os.path.dirname(__file__)
shared_param_file = os.path.join(script_dir, "SharedParameters.txt")

if not os.path.exists(shared_param_file):
    raise Exception("Shared parameter file not found:\n{}".format(shared_param_file))

app.SharedParametersFilename = shared_param_file
shared_param_file_obj = app.OpenSharedParameterFile()
if not shared_param_file_obj:
    raise Exception("Could not open shared parameter file. Ensure format and path are valid.")

# --- Target Categories ---
category_names = ["Electrical Equipment", "Lighting Fixtures"]
categories = CategorySet()
for name in category_names:
    cat = doc.Settings.Categories.get_Item(name)
    if cat:
        categories.Insert(cat)
    else:
        print("⚠️ Category not found: {}".format(name))

# --- Bind Parameters Safely with Native Transactions ---
bound_params = doc.ParameterBindings

for group in shared_param_file_obj.Groups:
    for definition in group.Definitions:
        param_name = definition.Name

        if bound_params.Contains(definition):
            print("✅ Already bound: {}".format(param_name))
            continue

        binding = app.Create.NewInstanceBinding(categories)
        t = Transaction(doc, "Bind Shared Parameter: {}".format(param_name))

        try:
            t.Start()
            success = bound_params.Insert(definition, binding, BuiltInParameterGroup.PG_DATA)
            t.Commit()
            print("✅ Bound: {}".format(param_name) if success else "❌ Failed to bind: {}".format(param_name))
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            print("❌ Error binding '{}': {}".format(param_name, str(ex)))
