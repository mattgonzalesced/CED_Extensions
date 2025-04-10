import csv
from pyrevit import revit, DB
from pyrevit import script

# Initialize output manager
output = script.get_output()
output.close_others()

# Path to the CSV file
csv_path = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\filtered_mechanical_equipment_data.csv"

# Read CSV and extract IDs from column "Mech Element Id" if "Match" column is False
try:
    with open(csv_path, mode='r') as csvfile:  # Removed encoding argument for IronPython compatibility
        reader = csv.DictReader(csvfile, delimiter=',')

        # Debugging: Print detected headers
        #print("Detected headers:", reader.fieldnames)

        # Ensure required columns exist
        required_columns = ["Mech Element Id", "Elec Element Id", "Match", "Mech CED-E-MCA", "Mech CED-E-MOCP", "Elec MCA_CED", "Elec MOCP_CED"]
        if not all(col in reader.fieldnames for col in required_columns):
            raise ValueError("CSV file does not have the required columns.")

        elements_data = [
            {
                "mech_id": int(row["Mech Element Id"].strip()),
                "elec_id": int(row["Elec Element Id"].strip()),
                "ced_e_mca": row["Mech CED-E-MCA"].strip(),
                "mech_ced_e_mocp": row["Mech CED-E-MOCP"].strip(),
                "elec_mca_ced": row["Elec MCA_CED"].strip(),
                "elec_mocp_ced": row["Elec MOCP_CED"].strip()
            }
            for row in reader
            if row["Match"].strip().lower() == "false" and row["Mech Element Id"].strip().isdigit() and row["Elec Element Id"].strip().isdigit()
        ]
except Exception as e:
    script.get_logger().error("Error reading CSV file: {0}".format(str(e)))
    raise

# Process each element ID
def process_element(element_data):
    mech_element_id = element_data["mech_id"]
    elec_element_id = element_data["elec_id"]
    mech_element = revit.doc.GetElement(DB.ElementId(mech_element_id))
    elec_element = revit.doc.GetElement(DB.ElementId(elec_element_id))

    if mech_element:
        try:
            mech_element_name = mech_element.Name if hasattr(mech_element, "Name") else "(No Name)"
        except Exception as e:
            mech_element_name = "(Error retrieving name)"
            script.get_logger().warning("Error retrieving name for mechanical element ID {0}: {1}".format(mech_element_id, str(e)))

        # Create a clickable link to the mechanical element
        try:
            clickable_mech_id = output.linkify([mech_element.Id])
        except Exception as e:
            clickable_mech_id = "(Error creating link)"
            script.get_logger().warning("Error creating link for mechanical element ID {0}: {1}".format(mech_element_id, str(e)))

    else:
        clickable_mech_id = "(Not Found)"
        mech_element_name = "(Not Found)"

    if elec_element:
        try:
            elec_element_name = elec_element.Name if hasattr(elec_element, "Name") else "(No Name)"
        except Exception as e:
            elec_element_name = "(Error retrieving name)"
            script.get_logger().warning("Error retrieving name for electrical element ID {0}: {1}".format(elec_element_id, str(e)))

        # Create a clickable link to the electrical element
        try:
            clickable_elec_id = output.linkify([elec_element.Id])
        except Exception as e:
            clickable_elec_id = "(Error creating link)"
            script.get_logger().warning("Error creating link for electrical element ID {0}: {1}".format(elec_element_id, str(e)))

    else:
        clickable_elec_id = "(Not Found)"
        elec_element_name = "(Not Found)"

    # Output element details
    print("Mech ID: {0}\tMech Name: {1}\tElec ID: {2}\tElec Name: {3}\tCED-E-MCA: {4}\tMech CED-E-MOCP: {5}\tElec MCA_CED: {6}\tElec MOCP_CED: {7}".format(
        clickable_mech_id, mech_element_name, clickable_elec_id, elec_element_name, element_data["ced_e_mca"], element_data["mech_ced_e_mocp"], element_data["elec_mca_ced"], element_data["elec_mocp_ced"]
    ))

# Loop through elements data and process
for element_data in elements_data:
    process_element(element_data)

print("\nProcessing complete.")
