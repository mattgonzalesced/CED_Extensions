from pyrevit import script

# Get the configuration section for the script
config = script.get_config("my_custom_config")

# Add "bananas" to the configuration
config.bananas = "true"  # This will appear as `bananas = true` in the config file

# Save the updated configuration
script.save_config()

# Notify the user
from pyrevit import forms
forms.alert("'bananas' has been written to the config file.", title="Success")
