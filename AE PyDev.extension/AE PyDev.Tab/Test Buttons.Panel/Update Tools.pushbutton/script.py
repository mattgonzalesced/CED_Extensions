import subprocess

from pyrevit import forms

try:
    # Run PowerShell script in background (no visible window)
    ps_script = r'C:\Users\Aevelina\CED_Extensions\UpdateCEDTools.ps1'
    p = subprocess.Popen(
        ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_script],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=True
    )
    p.wait()  # Wait for PowerShell to finish

    # Show final alert
    forms.alert("Update complete!\nPlease manually click the 'Reload' button in pyRevit to finalize the update.", title="Update Complete")

except Exception as e:
    forms.alert("Update failed:\n{}".format(e), title="Update Error")
