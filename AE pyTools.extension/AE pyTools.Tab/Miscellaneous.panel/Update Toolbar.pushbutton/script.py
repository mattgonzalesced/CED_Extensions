# -*- coding: utf-8 -*-
__doc__ = "Copies & launches Updater_pyrevit.exe from %TEMP% (no UAC needed)"
__title__ = "Run Updater (Temp)"

import os
import shutil
import tempfile
import clr
from pyrevit import forms

# Bring in the .NET Process API
clr.AddReference('System')
from System.Diagnostics import Process

def main():
    # 1. Locate source EXE (next to this script)
    script_dir = os.path.dirname(__file__)
    src_exe   = os.path.join(script_dir, "Updater_pyrevit.exe")
    if not os.path.isfile(src_exe):
        forms.alert("Could not find:\n{}".format(src_exe), exitscript=True)

    # 2. Copy to a temp location
    temp_dir = tempfile.gettempdir()
    dst_exe  = os.path.join(temp_dir, "Updater_pyrevit.exe")
    try:
        shutil.copy2(src_exe, dst_exe)
    except Exception as copy_err:
        forms.alert(
            "Failed to copy updater to temp:\n{}\n\n{}".format(dst_exe, copy_err),
            exitscript=True
        )

    # 3. Run the EXE from %TEMP%
    try:
        Process.Start(dst_exe)
    except Exception as run_err:
        forms.alert(
            "Failed to start updater:\n{}\n\n{}".format(dst_exe, run_err),
            exitscript=True
        )

if __name__ == "__main__":
    main()
