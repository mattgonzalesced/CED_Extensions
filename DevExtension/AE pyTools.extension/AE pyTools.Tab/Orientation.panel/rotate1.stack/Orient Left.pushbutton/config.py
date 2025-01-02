from pyrevit import forms, script

# Create a configuration object for the Orientation Panel tools
config = script.get_config("orientation_config")

# Default Settings
DEFAULT_INCLUDE_TAG_POSITION = True
DEFAULT_INCLUDE_TAG_ANGLE = True

# Ensure config attributes exist and set defaults if not present
if not hasattr(config, 'tag_position'):
    config.tag_position = DEFAULT_INCLUDE_TAG_POSITION
if not hasattr(config, 'tag_angle'):
    config.tag_angle = DEFAULT_INCLUDE_TAG_ANGLE

# Convert config values to boolean if they are strings
config.tag_position = config.tag_position in [True, "True", "true"]
config.tag_angle = config.tag_angle in [True, "True", "true"]

# Define the configuration for the CommandSwitchWindow
switch_config = {
    'Tag Position': {'state': config.tag_position},  # State reflects config
    'Tag Rotation': {'state': config.tag_angle}
}

# Display the CommandSwitchWindow
selected_option, switches = forms.CommandSwitchWindow.show(
    ['Finish'],  # Context options
    switches=['Tag Position', 'Tag Rotation'],  # Switches
    message='Select Tag Options to Include:',
    config=switch_config  # Properly formatted config with initial states
)

if not selected_option:
    script.exit()
else:
    # Update configuration based on user choice
    config.tag_position = switches['Tag Position']
    config.tag_angle = switches['Tag Rotation']
    script.save_config()
    forms.alert("Tag modification options saved successfully!")
