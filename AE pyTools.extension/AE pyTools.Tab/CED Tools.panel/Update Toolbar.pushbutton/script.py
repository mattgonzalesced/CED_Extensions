# -*- coding: utf-8 -*-
__doc__ = "Updates Toolbar by pulling from Github Repo (safe-scoped version with backup)"

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

import clr
from pyrevit import forms

# .NET Process API
clr.AddReference('System')
from System.Diagnostics import Process


def find_extension_root(script_dir):
    """Find the nearest parent folder whose name ends with .extension."""
    p = Path(script_dir).resolve()
    for parent in [p] + list(p.parents):
        if parent.name.lower().endswith('.extension'):
            return parent
    return None


def is_under(path, root):
    """Return True if path is under root (after resolving), else False."""
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except Exception:
        return False


def main():
    # 1) Where this script lives
    script_dir = os.path.dirname(__file__)

    # 2) Locate the updater EXE next to this script
    src_exe = os.path.join(script_dir, "Updater_pyrevit.exe")
    if not os.path.isfile(src_exe):
        forms.alert("Could not find updater EXE:\n{}".format(src_exe), exitscript=True)

    # 3) Copy it into a truly writable spot first
    tmp_exe = os.path.join(tempfile.gettempdir(), "Updater_pyrevit.exe")
    try:
        shutil.copy2(src_exe, tmp_exe)
    except Exception as copy_err:
        forms.alert(
            "Failed to copy updater to temp:\n{}\n\n{}".format(tmp_exe, copy_err),
            exitscript=True
        )

    # 4) Resolve extension root to remove (scoped)
    ext_root = find_extension_root(script_dir)

    if not ext_root or not ext_root.exists():
        forms.alert(
            "Warning: Could not locate an '.extension' folder above:\n{}\nContinuing anyway..."
            .format(script_dir),
            title="Warning",
            ok=True
        )
    else:
        # Allowed roots (scope)
        appdata = os.environ.get('APPDATA', '')
        programdata = os.environ.get('PROGRAMDATA', r'C:\ProgramData')
        allowed_roots = [
            Path(appdata) / 'pyRevit' / 'Extensions',
            Path(programdata) / 'pyRevit' / 'Extensions'
        ]

        # Check ext_root is under one of the allowed roots
        if not any(is_under(ext_root, r) for r in allowed_roots):
            forms.alert(
                "Safety stop:\n\nThe target extension folder is outside the expected pyRevit directories.\n\n"
                "Extension:\n  {}\n\nAllowed roots:\n  {}\n\nNo changes were made."
                .format(ext_root, "\n  ".join(str(r) for r in allowed_roots)),
                exitscript=True
            )

        # Confirm with the user
        proceed = forms.alert(
            "About to remove (with backup) the extension folder:\n\n  {}\n\n"
            "It will be MOVED to your temp folder (not permanently deleted), "
            "then the updater will run.\n\nProceed?"
            .format(ext_root),
            yes=True, no=True
        )
        if not proceed:
            return

        # 4a) Move to backup in %TEMP% instead of deleting outright
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_dir = Path(tempfile.gettempdir()) / ("pyrevit_backup_{}_{}".format(ext_root.name, timestamp))

        try:
            shutil.move(str(ext_root), str(backup_dir))
        except Exception as move_err:
            forms.alert(
                "Failed to move extension to backup:\n{}\n\n{}\n\n"
                "No changes were made.".format(backup_dir, move_err),
                exitscript=True
            )

    # 5) Finally, run the updater from %TEMP%
    try:
        Process.Start(tmp_exe)
    except Exception as run_err:
        forms.alert(
            "Failed to start updater:\n{}\n\n{}".format(tmp_exe, run_err),
            exitscript=True
        )


if __name__ == "__main__":
    main()
