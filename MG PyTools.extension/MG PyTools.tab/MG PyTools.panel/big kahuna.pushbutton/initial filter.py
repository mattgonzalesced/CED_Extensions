import csv
import os
import subprocess
from Autodesk.Revit import DB

# File path to save the CSV
csv_file_path = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\mechanical_equipment_data.csv"

# Access the current document
doc = __revit__.ActiveUIDocument.Document

# Parameter names to extract for mechanical equipment
mech_param_names = ["CED-E-MCA", "CED-E-MOCP"]

# Parameter names to extract for electrical fixtures
electrical_param_names = ["MCA_CED", "MOCP_CED"]

# Collect all mechanical equipment in the document
mech_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()

# Collect all electrical fixtures in the document
electrical_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType()

# Prepare data for CSV
data = []
data.append(["Element Id", "Category", "Type Name", "Family Name", "X", "Y", "Z"] + mech_param_names + electrical_param_names)  # Add headers

# Function to get XYZ coordinates
def get_xyz_coordinates(element):
    location = element.Location
    if location and isinstance(location, DB.LocationPoint):
        point = location.Point
        return point.X, point.Y, point.Z
    return "", "", ""  # Return empty values if no location point

# Process mechanical equipment
for equipment in mech_collector:
    row = [equipment.Id.IntegerValue, "Mechanical Equipment"]  # Start with the element ID and category
    type_name = equipment.Name if hasattr(equipment, "Name") else ""  # Get the type name
    family_name = equipment.Symbol.Family.Name if hasattr(equipment, "Symbol") and equipment.Symbol else ""  # Get the family name
    row.extend([type_name, family_name])
    x, y, z = get_xyz_coordinates(equipment)  # Get coordinates
    row.extend([x, y, z])
    for param_name in mech_param_names:
        param = equipment.LookupParameter(param_name)
        if param:
            if param.StorageType == DB.StorageType.Double:  # Numeric values with units
                row.append(param.AsDouble())
            elif param.StorageType == DB.StorageType.String:  # Text values
                row.append(param.AsString())
            elif param.StorageType == DB.StorageType.Integer:  # Integer values
                row.append(param.AsInteger())
            else:
                row.append("N/A")
        else:
            row.append("")  # Add an empty value if the parameter is not found
    row.extend(["" for _ in electrical_param_names])  # Add empty columns for electrical-specific parameters
    data.append(row)

# Process electrical fixtures
for fixture in electrical_collector:
    row = [fixture.Id.IntegerValue, "Electrical Fixture"]  # Start with the element ID and category
    type_name = fixture.Name if hasattr(fixture, "Name") else ""  # Get the type name
    family_name = fixture.Symbol.Family.Name if hasattr(fixture, "Symbol") and fixture.Symbol else ""  # Get the family name
    row.extend([type_name, family_name])
    x, y, z = get_xyz_coordinates(fixture)  # Get coordinates
    row.extend([x, y, z])
    row.extend(["" for _ in mech_param_names])  # Add empty columns for mech-specific parameters
    for param_name in electrical_param_names:
        param = fixture.LookupParameter(param_name)
        if param:
            # Handle different storage types
            if param.StorageType == DB.StorageType.Double:  # Numeric values with units
                row.append(param.AsDouble())  # Convert units if necessary
            elif param.StorageType == DB.StorageType.String:  # Text values
                row.append(param.AsString())
            elif param.StorageType == DB.StorageType.Integer:  # Integer values
                row.append(param.AsInteger())
            else:
                row.append("N/A")  # Unexpected type
        else:
            row.append("N/A")  # Placeholder for missing parameters

    data.append(row)

# Ensure the folder exists
output_dir = os.path.dirname(csv_file_path)
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Write data to CSV
try:
    with open(csv_file_path, mode="w") as file:  # Removed 'newline' for IronPython compatibility
        writer = csv.writer(file)
        writer.writerows(data)


    # Call the external script after writing the CSV
    python_exe = r"C:\\Users\\m.gonzales\\AppData\\Local\\Programs\\Python\\Python313\\python.exe"  # Adjust this path to your CPython environment
    external_script_path = r"C:\\Users\\m.gonzales\\OneDrive - CoolSys Inc\\Desktop\\CED\\DevExtension\\MG PyTools.extension\\MG PyTools.tab\\MG PyTools.panel\\big kahuna.pushbutton\\EXT pandas.py"

    # Run the external script with the CSV file as an argument
    try:
        subprocess.call([python_exe, external_script_path, csv_file_path])
    except Exception as e:
        print("Error while executing the external script: {}".format(e))
except IOError as e:
    print("Failed to write to file: {}".format(e))
