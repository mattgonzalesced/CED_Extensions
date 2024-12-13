from pyrevit import script, forms

# Initialize configuration to store the zoom factor
config = script.get_config("zoom_selection_config")

# Default zoom factor
DEFAULT_ZOOM_FACTOR = 1

# Retrieve the current zoom factor, or use the default if it doesn't exist
current_zoom_factor = config.get_option("zoom_factor", DEFAULT_ZOOM_FACTOR)

# Prompt the user to select a new zoom factor using a slider
new_zoom_factor = forms.ask_for_number_slider(
    min=1,
    max=10,
    interval=1,
    prompt="Set a new zoom factor (1 = exact fit, >1 = zoom out):",
    default=current_zoom_factor,
)

# Save the new zoom factor to the configuration
if new_zoom_factor is not None:  # Ensure the user selected a value
    config.set_option("zoom_factor", new_zoom_factor)
    script.save_config()
