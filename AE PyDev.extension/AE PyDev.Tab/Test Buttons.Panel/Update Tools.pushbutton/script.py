# -*- coding: utf-8 -*-
import subprocess

from pyrevit import forms

try:
    ps_update = r'C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.ps1'
    update_cmd = [
        'powershell.exe',
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ps_update
    ]
    print("=== Running update script ===")
    p = subprocess.Popen(update_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    stdout, stderr = p.communicate()
    print("=== RAW STDOUT ===")
    print(stdout.decode('utf-8').strip())
    print("=== RAW STDERR ===")
    print(stderr.decode('utf-8').strip())
    print("=== Update script exited with code: {} ===".format(p.returncode))
    forms.alert("✅ Update/Clone complete.\nPlease manually click the 'Reload' button in pyRevit.", title="Complete")
except Exception as e:
    print("❌ Update process failed:")
    print(e)
