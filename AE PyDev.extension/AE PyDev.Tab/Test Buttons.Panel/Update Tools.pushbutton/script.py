import subprocess

from pyrevit import script, forms

logger = script.get_logger()

try:
    # Launch PowerShell script and wait for it to complete
    ps_script = r'C:\Users\Aevelina\CED_Extensions\UpdateCEDTools.ps1'
    p = subprocess.Popen(
        ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_script],
        shell=True
    )
    p.wait()  # Wait for PowerShell to complete

    # Inform the user to manually reload
    forms.alert(
        "Update complete.\nPlease manually click the 'Reload' button in pyRevit to avoid crashes.",
        title="Update Complete"
    )

except Exception as e:
    forms.alert("Update failed:\n{}".format(e), title="Update Error")
