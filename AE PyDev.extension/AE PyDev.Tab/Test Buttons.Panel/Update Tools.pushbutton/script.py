import subprocess

from pyrevit import forms

try:
    subprocess.Popen(r'"C:\Program Files\CEDApp\CED_Extensions\UpdateCEDTools.cmd"', shell=True)
    forms.alert("Update script launched.\nFollow the prompts in the command window.", title="Update Started")
except Exception as e:
    forms.alert("Update failed:\n{}".format(e), title="Update Error")
