# -*- coding: utf-8 -*-
__doc__ = "Launches the Updater_pyrevit.exe in the same folder as this script"
__title__ = "Run Updater"

import os
import clr

# Bring in .NET Process API
clr.AddReference('System')
from System.Diagnostics import Process

def main():
    # Get folder where this script lives
    script_dir = os.path.dirname(__file__)
    
    # Name of your EXE
    exe_name = "Updater_pyrevit.exe"
    exe_path = os.path.join(script_dir, exe_name)
    
    # If the EXE isn't found, alert and exit
    if not os.path.isfile(exe_path):
        from pyrevit import forms
        forms.alert(
            "Could not find:\n{}".format(exe_path),
            exitscript=True
        )
    
    # Launch the EXE
    try:
        Process.Start(exe_path)
    except Exception as e:
        from pyrevit import forms
        forms.alert(
            "Failed to start:\n{}\n\n{}".format(exe_path, e),
            exitscript=True
        )

if __name__ == "__main__":
    main()
