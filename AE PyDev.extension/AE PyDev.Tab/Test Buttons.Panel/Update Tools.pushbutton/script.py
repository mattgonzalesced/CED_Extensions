import subprocess
import sys

# Build command
command = ["pyrevit", "extensions", "update", "--all"]

# Run command
try:
    subprocess.run(command, check=True)
    print("✅ All pyRevit extensions updated successfully.")
except subprocess.CalledProcessError as e:
    print("❌ Failed to update extensions.")
    print(e)
