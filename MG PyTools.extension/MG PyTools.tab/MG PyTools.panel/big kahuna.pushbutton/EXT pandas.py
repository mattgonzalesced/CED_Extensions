import pandas as pd
import numpy as np
import json
import os

# File paths
csv_file_path = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\mechanical_equipment_data.csv"
filtered_filepath_csv = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\filtered_mechanical_equipment_data.csv"
filtered_filepath_json = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\filtered_mechanical_equipment_data.json"

# Parameters to include in the output
mech_param_names = ["CED-E-MCA", "CED-E-MOCP"]
electrical_param_names = ["MCA_CED", "MOCP_CED"]

# Check if the input CSV file exists and create if it doesn't
if not os.path.exists(csv_file_path):
    print(f"The file {csv_file_path} does not exist. Creating a new one.")
    # Create an empty DataFrame with expected columns
    empty_columns = ["Type Name", "Category", "X", "Y", "Z", "Element Id", "CED-E-MCA", "CED-E-MOCP", "MCA_CED", "MOCP_CED"]
    pd.DataFrame(columns=empty_columns).to_csv(csv_file_path, index=False)
    print(f"Empty file created at {csv_file_path}. Please populate it and rerun the script.")
    exit(0)

# Load the data from the CSV file
df = pd.read_csv(csv_file_path)

# Load data and filter specific types
filtered_df = df[df["Type Name"].str.contains("CED-M-4-6 TON AC_Condenser|ZJ150|Capacity 5 ton|25.0 ton|ACCU|Equip Connection - 480/3ph|CED-M-EQPM Generic (R019) ELECTRIC HEATER|Equip Connection - 208/1ph|ELECTRIC HEATER|UPS|HCUc8040 S71074-070", na=False)].copy()

# Split data into Mechanical Equipment and Electrical Fixtures
mechanical_df = filtered_df[filtered_df["Category"] == "Mechanical Equipment"].reset_index()
electrical_df = filtered_df[filtered_df["Category"] == "Electrical Fixture"].reset_index()

# Ensure coordinates are numeric
mechanical_df[["X", "Y", "Z"]] = mechanical_df[["X", "Y", "Z"]].apply(pd.to_numeric, errors="coerce")
electrical_df[["X", "Y", "Z"]] = electrical_df[["X", "Y", "Z"]].apply(pd.to_numeric, errors="coerce")

# Initialize pair counter and new DataFrame for results
pair_count = 1
pair_results = []

# Loop until one of the DataFrames is empty
while not mechanical_df.empty and not electrical_df.empty:
    min_distance = float('inf')  # Reset minimum distance for each iteration
    closest_pair = None          # Reset closest pair for each iteration

    # Iterate over all rows in mechanical_df
    for i, mech_batch in mechanical_df.iterrows():
        # Iterate over all rows in electrical_df
        for j, elec_batch in electrical_df.iterrows():
            # Calculate the Euclidean distance
            distance = np.sqrt(
                (elec_batch["X"] - mech_batch["X"])**2 +
                (elec_batch["Y"] - mech_batch["Y"])**2 +
                (elec_batch["Z"] - mech_batch["Z"])
            )
            # Update minimum distance and closest pair if this distance is smaller
            if distance < min_distance:
                min_distance = distance
                closest_pair = (i, j)

    # Double-check closest pair and distance before proceeding
    if closest_pair:
        mech_index, elec_index = closest_pair
        mech_row = mechanical_df.loc[mech_index]
        elec_row = electrical_df.loc[elec_index]

        # Determine if the parameters match
        mech_mca = mech_row.get("CED-E-MCA", None)
        elec_mca = elec_row.get("MCA_CED", None)
        mech_mocp = mech_row.get("CED-E-MOCP", None)
        elec_mocp = elec_row.get("MOCP_CED", None)
        
        is_match = (str(mech_mca) == str(elec_mca)) and (str(mech_mocp) == str(elec_mocp))

        # Append pair information to the results
        pair_results.append({
            "Pair Count": pair_count,
            "Mech Element Id": mech_row["Element Id"],
            "Mech Type Name": mech_row["Type Name"],
            "Mech Category": mech_row["Category"],
            "Mech CED-E-MCA": mech_mca,
            "Mech CED-E-MOCP": mech_mocp,
            "Elec Element Id": elec_row["Element Id"],
            "Elec Type Name": elec_row["Type Name"],
            "Elec Category": elec_row["Category"],
            "Elec MCA_CED": elec_mca,
            "Elec MOCP_CED": elec_mocp,
            "Distance": min_distance,
            "Unique Identifier": f"{mech_row['Element Id']}{elec_row['Element Id']}",
            "Match": is_match
        })

        # Remove the matched rows from both DataFrames to avoid duplicate pairing
        mechanical_df = mechanical_df.drop(mech_index).reset_index(drop=True)
        electrical_df = electrical_df.drop(elec_index).reset_index(drop=True)

        # Increment pair counter
        pair_count += 1

# Create a new DataFrame from the pair results
paired_df = pd.DataFrame(pair_results)

# Save the paired data to a CSV file
paired_df.to_csv(filtered_filepath_csv, index=False)

# Save the paired data to a JSON file
with open(filtered_filepath_json, 'w') as json_file:
    json.dump(pair_results, json_file, indent=4, default=lambda o: int(o) if isinstance(o, np.integer) else str(o) if isinstance(o, np.floating) else o)

print("CSV and JSON files have been successfully created.")
