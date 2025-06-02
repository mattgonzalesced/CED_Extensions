import subprocess

from pyrevit import forms

try:
    subprocess.Popen(
        r'cmd.exe /c start "" powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Aevelina\CED_Extensions\UpdateCEDTools.ps1"',
        shell=True
    )
    forms.alert("Update script launched in a new PowerShell window.\nFollow the prompts there.", title="Update Started")
except Exception as e:
    forms.alert("Update failed:\n{}".format(e), title="Update Error")
