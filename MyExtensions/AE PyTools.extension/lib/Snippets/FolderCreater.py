import os

# Define the base path you provided
base_path = r"C:\Users\Aevelina\OneDrive - CoolSys Inc\AE Pyrevit\MyExtensions\MyFirstExtenson.extension\AE Tools.Tab"

# Define the folder structure starting from the three panels
folders = [
    "Circuits",
    "Rotation.panel",
    "Circuits/CleanModelTags.pushbutton",
    "Circuits/Pushbutton2.pushbutton",
    "Rotation.panel/Reports.pulldown",
    "Rotation.panel/Reports.pulldown/Rotate CCW.pushbutton",
    "Rotation.panel/Reports.pulldown/Orient Down.pushbutton",
    "Rotation.panel/rotate1.stack",
    "Rotation.panel/rotate1.stack/CleanModelTags.pushbutton",
    "Rotation.panel/rotate1.stack/Pushbutton2.pushbutton",
    "Rotation.panel/rotate1.stack/Rotate CCW.pushbutton"
]

# Create the folders starting from the base path
for folder in folders:
    full_path = os.path.join(base_path, folder)
    os.makedirs(full_path, exist_ok=True)

print("Folder structure created successfully!")
