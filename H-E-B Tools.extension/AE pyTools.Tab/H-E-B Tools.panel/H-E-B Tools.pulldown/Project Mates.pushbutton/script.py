# -*- coding: utf-8 -*-
"""
Launches an external executable (app.exe) located in the same folder as this script.
"""

__doc__ = "Launches web browser and navigates through project mates."

import os
import subprocess
from pyrevit import script

def main():
    # Determine the folder this script lives in
    script_dir = os.path.dirname(__file__)
    exe_path = os.path.join(script_dir, "app.exe")

    # Check for existence
    if not os.path.isfile(exe_path):
        script.exit("[ERROR] Could not find app.exe at:\n    " + exe_path)

    # Launch the executable with correct working directory
    try:
        subprocess.Popen([exe_path], cwd=script_dir, shell=True)
    except Exception as e:
        script.exit("[ERROR] Failed to launch app.exe:\n    " + str(e))

if __name__ == "__main__":
    main()