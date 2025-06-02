import subprocess

from pyrevit import forms

try:
    # Step 1: Check status
    ps_check = r'C:\Users\Aevelina\AppData\Roaming\pyRevit\Extensions\CED_Extensions\Updater\CheckCEDTools.ps1'
    p1 = subprocess.Popen(
        ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_check],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True
    )
    stdout, stderr = p1.communicate()
    status = stdout.decode('utf-8').strip()

    # Step 2: Act on status
    if "status: up-to-date" in status:
        forms.alert("✅ Extensions are already up to date!\nNo further action needed.", title="Update Status")
    elif "status: clone-needed" in status:
        forms.alert("⚠️ Extensions folder is missing!\nPlease re-clone manually.", title="Update Status")
    elif "status: updates-available" in status:
        proceed = forms.alert("🔄 Updates available.\nClick OK to continue updating.", ok=True, cancel=True)
        if proceed:
            # Step 3: Perform the update
            ps_update = r'C:\Users\Aevelina\AppData\Roaming\pyRevit\Extensions\CED_Extensions\Updater\UpdateCEDTools.ps1'
            p2 = subprocess.Popen(
                ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_update],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True
            )
            p2.wait()  # Wait for update to finish
            forms.alert("✅ Update complete.\nPlease manually click the 'Reload' button in pyRevit.", title="Update Complete")
        else:
            forms.alert("Update cancelled by user.", title="Update Cancelled")
    else:
        forms.alert("❌ Unexpected status:\n{}".format(status), title="Update Status")

except Exception as e:
    forms.alert("Update process failed:\n{}".format(e), title="Update Error")
