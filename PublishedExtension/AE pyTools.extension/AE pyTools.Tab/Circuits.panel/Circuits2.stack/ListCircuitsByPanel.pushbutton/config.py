# -*- coding: utf-8 -*-
from pyrevit import forms, script
from pyrevit import revit, DB



config = script.get_config("list_circuits_by_panel_config")
doc = revit.doc

sample_ckt = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit) \
    .WhereElementIsNotElementType() \
    .FirstElement()

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

# Store default parameters as names (strings) for compatibility with JSON serialization
DEFAULT_PARAMETERS = [
    "RBS_ELEC_PANEL_NAME",
    "RBS_ELEC_CIRCUIT_NUMBER",
    "RBS_ELEC_CIRCUIT_NAME",
    "RBS_ELEC_CIRCUIT_RATING_PARAM",
    "RBS_ELEC_CIRCUIT_FRAME_PARAM",
    "RBS_ELEC_APPARENT_CURRENT_PARAM",
    "RBS_ELEC_VOLTAGE",
    "RBS_ELEC_NUMBER_OF_POLES"
]

# Load user-selected parameters or set defaults if none exist
if not config.has_option("user_selected_parameters"):
    config.user_selected_parameters = DEFAULT_PARAMETERS
