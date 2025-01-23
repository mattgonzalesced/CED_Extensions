# -*- coding: utf-8 -*-
from pyrevit import forms, script
from pyrevit import revit, DB

# Initialize the configuration to save default parameters
config = script.get_config("revision_parameters_config")

# Get one sample Revision Cloud element to fetch its parameters
doc = revit.doc
sample_cloud = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
    .WhereElementIsNotElementType() \
    .FirstElement()

if not sample_cloud:
    forms.alert("No Revision Clouds found in the project.")
    script.exit()


def pick_parameters(sample_element):
    """Prompt the user to select parameters from a single sample element, excluding 'Comments'."""
    # Gather unique parameter names
    parameters = {param.Definition.Name for param in sample_element.Parameters if param.Definition}

    # Exclude the 'Comments' parameter as it is included by default
    parameters.discard("Comments")

    # Sort the parameters alphabetically
    sorted_params = sorted(parameters)

    # Use PyRevit's selection form
    selected_params = forms.SelectFromList.show(
        sorted_params,
        title="Select Additional Parameters",
        button_name="OK",
        multiselect=True
    )

    return selected_params


# Prompt the user to select parameters using `pick_parameters`
selected_parameters = pick_parameters(sample_cloud)

if selected_parameters:
    # Save parameter names in the config
    config.selected_param_names = ",".join(selected_parameters)  # Save parameters as a comma-separated string
    script.save_config()
    forms.alert("Parameters saved successfully!")
else:
    # Reset to default (no additional parameters)
    config.selected_param_names = ""
    script.save_config()
    forms.alert("No parameters selected. Defaults will be used (Sheet Number, Sheet Name, Comments).")
