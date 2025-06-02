# -*- coding: utf-8 -*-
import subprocess

from pyrevit import forms

try:
    # Step 1: Check status
    ps_check = r'C:\Users\Aevelina\CED_Extensions\Updater\CheckCEDTools.ps1'
    check_cmd = [
        'powershell.exe',
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ps_check
    ]
    print("=== Running check script ===")
    p1 = subprocess.Popen(check_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    stdout, stderr = p1.communicate()
    status_output = stdout.decode('utf-8').strip()
    print("=== RAW STDOUT ===")
    print(status_output)
    print("=== RAW STDERR ===")
    print(stderr.decode('utf-8').strip())

    status_line = None
    for line in status_output.splitlines():
        if line.startswith("status:"):
            status_line = line.strip()
            break
    print("=== Extracted Status Line ===")
    print(status_line)

    if status_line == "status: up-to-date":
        print("✅ Tools already up to date.")
    elif status_line in ["status: updates-available", "status: clone-needed"]:
        proceed = forms.alert("🔄 Updates or clone needed. Do you want to continue?", ok=True, cancel=True)
        if proceed:
            ps_update = r'C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.ps1'
            update_cmd = [
                'powershell.exe',
                '-NoLogo',
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', ps_update
            ]
            print("=== Running update script ===")
            p2 = subprocess.Popen(update_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
            stdout2, stderr2 = p2.communicate()
            print("=== RAW STDOUT (Update) ===")
            print(stdout2.decode('utf-8').strip())
            print("=== RAW STDERR (Update) ===")
            print(stderr2.decode('utf-8').strip())
            print("=== Update script exited with code: {} ===".format(p2.returncode))
            forms.alert("✅ Update/Clone complete.\nPlease manually click the 'Reload' button in pyRevit.", title="Complete")
        else:
            print("❌ Update/clone cancelled by user.")
    else:
        print("❌ Unexpected status detected.")
        print(status_output)

except Exception as e:
    print("❌ Update process failed:")
    print(e)
