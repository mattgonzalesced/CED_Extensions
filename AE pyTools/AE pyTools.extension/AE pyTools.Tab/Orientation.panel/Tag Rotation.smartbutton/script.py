# -*- coding: utf-8 -*-
from pyrevit import forms, script
from pyrevit.revit import ui
import pyrevit.extensions as exts

# Create or load the script configuration
config = script.get_config("orientation_config")

# Ensure the tag_angle setting exists in the config
if not hasattr(config, 'tag_angle'):
    config.tag_angle = False  # Default to OFF

# Function to toggle the state and update the icon
def toggle_tag_angle():
    # Toggle the state
    config.tag_angle = not config.tag_angle

    # Update the button icon
    script.toggle_icon(config.tag_angle)

    # Save the updated config
    script.save_config()

    # Display feedback
    state = "ON" if config.tag_angle else "OFF"
    balloon_msg = "Tag Position toggled to: {}".format(state)
    forms.show_balloon("Orientation Config",balloon_msg)

# Initialization function for the smart button
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    # Resolve the icon file based on the current state
    off_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_OFF_ICON_FILE)
    on_icon = ui.resolve_icon_file(script_cmp.directory, "on.png")

    # Set the appropriate icon
    icon_path = on_icon if config.tag_angle else off_icon
    ui_button_cmp.set_icon(icon_path)

# Main script logic
if __name__ == '__main__':
    toggle_tag_angle()
