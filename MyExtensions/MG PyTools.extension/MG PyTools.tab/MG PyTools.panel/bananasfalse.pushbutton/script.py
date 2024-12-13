from pyrevit import script, forms

# Update "bananas" in the "my_custom_config" section
custom_config = script.get_config("my_custom_config")
custom_config.bananas = "false"
script.save_config()

# Update "disabled" in the "MG_PyTools.extension" section
mg_tools_config = script.get_config("MG PyTools.extension")
mg_tools_config.disabled = "true"
script.save_config()

# Notify the user
forms.alert("Configuration updated:\n- 'bananas' set to false\n- 'MG PyTools.extension' disabled set to true.", title="Success")
