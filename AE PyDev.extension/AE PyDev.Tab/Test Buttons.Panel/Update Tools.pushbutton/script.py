# -*- coding: utf-8 -*-
import subprocess

from pyrevit import forms

try:
    # Use the CMD file instead of PowerShell directly
    cmd_file = r'C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.cmd'
    update_cmd = [
        'cmd.exe',
        '/c',
        'start',
        '',  # Start a new window
        cmd_file
    ]
    print("=== Launching PowerShell window via CMD ===")
    p = subprocess.Popen(update_cmd, shell=True)
    print("=== CMD launched ===")
    forms.alert("🔍 PowerShell window launched.\nPlease watch the console for progress.\nRemember to manually click 'Reload' in pyRevit when done.", title="Update Launched")
except Exception as e:
    print("❌ Update process failed:")
    print(e)
