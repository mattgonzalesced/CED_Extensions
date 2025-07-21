# -*- coding: utf-8 -*-
import os
import subprocess

from pyrevit import forms, script

logger = script.get_logger()

try:
    # Dynamically build the path to the PowerShell check script
    appdata_dir = os.getenv('APPDATA')
    ps_check = os.path.join(appdata_dir, r'pyRevit\Extensions\CED_Extensions\Updater\CheckCEDTools.ps1')
    check_cmd = [
        'powershell.exe',
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ps_check
    ]
    logger.debug("=== Running check script ===")
    p1 = subprocess.Popen(check_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    stdout, stderr = p1.communicate()
    status_output = stdout.decode('utf-8').strip()
    logger.debug("=== RAW STDOUT ===")
    logger.debug(status_output)
    logger.debug("=== RAW STDERR ===")
    logger.debug(stderr.decode('utf-8').strip())

    # Parse status line
    status_line = None
    for line in status_output.splitlines():
        if line.startswith("status:"):
            status_line = line.strip()
            break
    logger.debug("=== Extracted Status Line ===")
    logger.debug(status_line)

    # Decision logic
    if status_line == "status: up-to-date":
        logger.debug("✅ Tools already up to date. No update needed.")
        forms.alert("✅ Tools are already up to date.\nNo further action is needed.", title="Status")
    elif status_line in ["status: updates-available", "status: clone-needed"]:
        proceed = forms.alert(
            "🔄 Updates Available! \nDo you want to continue?",
            ok=True, cancel=True
        )
        if proceed:
            # Dynamically build the path to the update script
            ps_update = os.path.join(appdata_dir, r'pyRevit\Extensions\CED_Extensions\Updater\UpdateCEDTools.cmd')
            update_cmd = [
                'cmd.exe',
                '/c',
                'start',
                '',  # Start a new window
                ps_update
            ]
            logger.debug("=== Launching PowerShell window via CMD ===")
            subprocess.Popen(update_cmd, shell=True)
            logger.debug("=== CMD launched ===")
            forms.alert(
                "🔍 PowerShell window launched.\nPlease watch the console for progress.\n Click 'Reload' in pyRevit when Finished!",
                title="Update Launched"
            )
        else:
            logger.debug("❌ Update/clone cancelled by user.")

    else:
        logger.warning("❌ Unexpected status detected.")
        logger.warning(status_output)


except Exception as e:
    logger.error("❌ Update process failed: {}".format(e))
    forms.alert("❌ Update process failed:\n{}".format(e), title="Error")
