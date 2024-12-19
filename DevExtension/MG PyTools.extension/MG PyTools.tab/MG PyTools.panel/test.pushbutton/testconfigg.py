import os

# Path to the pyRevit config file
config_file_path = os.path.expanduser(r"~\AppData\Roaming\pyRevit\pyRevit_config")

# Print the resolved path for debugging
print(f"Resolved config file path: {config_file_path}")

# Check if the config file exists
if os.path.exists(config_file_path):
    print(f"Config file found: {config_file_path}")
    # Read the config file
    with open(config_file_path, "r") as file:
        lines = file.readlines()

    # Modify the "disabled" value for MG PyTools
    with open(config_file_path, "w") as file:
        inside_section = False
        for line in lines:
            # Check if we are inside the MG PyTools section
            if line.strip().startswith("[MG PyTools.extension]"):
                inside_section = True
                file.write(line)  # Write the section header
                continue
            if inside_section and line.strip().startswith("disabled ="):
                file.write("disabled = true\n")  # Update the disabled value
                inside_section = False  # Exit the section after modifying
                continue
            file.write(line)  # Write other lines unchanged
    print("Updated 'disabled' to true for MG PyTools.extension")
else:
    print(f"Config file not found: {config_file_path}")
