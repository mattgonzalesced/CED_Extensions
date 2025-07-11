# -*- coding: utf-8 -*-
__doc__   = "Copies updater out first, wipes old extension, then runs updater from %TEMP%"
__title__ = "Safe Clean & Run Updater"

import os
import shutil
import tempfile
import clr
from pyrevit import forms

# .NET Process API
clr.AddReference('System')
from System.Diagnostics import Process

def main():
    # 1) Where this script lives
    script_dir = os.path.dirname(__file__)
    
    # 2) Locate the updater EXE next to this script
    src_exe = os.path.join(script_dir, "Updater_pyrevit.exe")
    if not os.path.isfile(src_exe):
        forms.alert("Could not find updater EXE:\n{}".format(src_exe), exitscript=True)
    
    # 3) Copy it into a truly writable spot first
    tmp_exe = os.path.join(tempfile.gettempdir(), "Updater_pyrevit.exe")
    try:
        shutil.copy2(src_exe, tmp_exe)
    except Exception as copy_err:
        forms.alert(
            "Failed to copy updater to temp:\n{}\n\n{}".format(tmp_exe, copy_err),
            exitscript=True
        )

    # 4) Now locate and remove the old .extension folder
    parts = script_dir.split(os.sep)
    ext_root = None
    for idx, p in enumerate(parts):
        if p.lower().endswith('.extension'):
            ext_root = os.sep.join(parts[:idx+1])
            break

    if ext_root and os.path.isdir(ext_root):
        try:
            shutil.rmtree(ext_root)
        except Exception as rm_err:
            forms.alert(
                "Failed to remove old extension:\n{}\n\n{}".format(ext_root, rm_err),
                exitscript=True
            )
    else:
        # Not found? Just warn and keep going
        forms.alert(
            "Warning: Could not locate old extension at:\n{}\nContinuing anyway...".format(ext_root or script_dir),
            title="Warning",
            ok=True
        )

    # 5) Finally, run the updater from %TEMP%
    try:
        Process.Start(tmp_exe)
    except Exception as run_err:
        forms.alert(
            "Failed to start updater:\n{}\n\n{}".format(tmp_exe, run_err),
            exitscript=True
        )

if __name__ == "__main__":
    main()
