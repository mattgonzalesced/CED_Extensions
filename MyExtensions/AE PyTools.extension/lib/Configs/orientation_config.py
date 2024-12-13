# -*- coding: utf-8 -*-
"""Configuration handler for Orientation Panel Tools."""

from pyrevit import script, forms, revit

# Create a configuration object for the Orientation Panel tools
config = script.get_config("orientation_panel")

# Default Settings
DEFAULT_INCLUDE_TAGS = True
DEFAULT_ADJUST_TAG_ANGLE = False

def load_configs():
    """Load user settings or apply defaults if not set."""
    # Check if SHIFT+click was used to trigger configuration change
    include_tags = config.get_option("include_tags", DEFAULT_INCLUDE_TAGS)
    adjust_tag_angle = config.get_option("adjust_tag_angle", DEFAULT_ADJUST_TAG_ANGLE)
    return include_tags, adjust_tag_angle

def save_configs(include_tags, adjust_tag_angle):
    """Save user settings to the configuration file."""
    config.include_tags = include_tags
    config.adjust_tag_angle = adjust_tag_angle
    config.write()

def reset_to_defaults():
    """Reset settings to default values."""
    config.include_tags = DEFAULT_INCLUDE_TAGS
    config.adjust_tag_angle = DEFAULT_ADJUST_TAG_ANGLE
    config.write()

def change_settings_ui():
    """Display UI for users to change configuration settings."""
    selected_option, switches = forms.CommandSwitchWindow.show(
        ['Run Script'],
        switches=['Include Tags', 'Adjust Tag Angle'],
        message='Select options for Better Orient Up:',
        recognize_access_key=True
    )

    if selected_option:
        # Update configuration based on user choice
        include_tags = switches['Include Tags']
        adjust_tag_angle = switches['Adjust Tag Angle']

        # Save the settings for future runs
        save_configs(include_tags, adjust_tag_angle)
