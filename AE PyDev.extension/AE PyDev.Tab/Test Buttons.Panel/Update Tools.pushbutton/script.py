import subprocess

from pyrevit import forms

try:
    subprocess.Popen(r'cmd.exe /c start "" "C:\Users\Aevelina\CED_Extensions\UpdateCEDTools.cmd"', shell=True)
    forms.alert("Update script launched in a new command window.\nFollow the prompts there.", title="Update Started")
except Exception as e:
    forms.alert("Update failed:\n{}".format(e), title="Update Error")
