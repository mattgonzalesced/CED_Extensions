# config_list_circuits_by_panel.py
from pyrevit import script

config = script.get_config("list_circuits_by_panel_config")

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
