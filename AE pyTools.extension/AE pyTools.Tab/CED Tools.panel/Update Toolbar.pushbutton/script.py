# -*- coding: utf-8 -*-
__doc__ = "Updates Toolbar from Github Repo (scoped to %APPDATA%\\pyRevit\\Extensions only)"

import os
import shutil
import tempfile
from pathlib import Path

import clr
from pyrevit import forms

# .NET Process API
clr.AddReference('System')
from System.Diagnostics import Process


def find_extension_root(script_dir):
   """Return the nearest parent ending with .extension (Path) or None."""
   p = Path(script_dir).resolve()
   for parent in [p] + list(p.parents):
       if parent.name.lower().endswith('.extension'):
           return parent
   return None


def is_under(path, root):
   """True if path is inside root (after resolving)."""
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

   # 3) Copy it into a writable temp location
   tmp_exe = os.path.join(tempfile.gettempdir(), "Updater_pyrevit.exe")
   try:
       shutil.copy2(src_exe, tmp_exe)
   except Exception as copy_err:
       forms.alert(
           "Failed to copy updater to temp:\n{}\n\n{}".format(tmp_exe, copy_err),
           exitscript=True
       )

   # 4) Resolve extension root and **enforce APPDATA\\pyRevit\\Extensions scope**
   ext_root = find_extension_root(script_dir)
   appdata = os.environ.get('APPDATA', '')
   appdata_ext_root = Path(appdata) / 'pyRevit' / 'Extensions'

   if not ext_root or not ext_root.exists():
       forms.alert(
           "Warning: Could not locate an '.extension' folder above:\n{}\nNo changes made."
           .format(script_dir),
           title="Out of scope",
           ok=True
       )
       return

   # HARD SCOPE: only proceed if the extension is under %APPDATA%\pyRevit\Extensions
   if not is_under(Path(ext_root), appdata_ext_root):
       forms.alert(
           "Safety stop: The target extension is not under:\n  {}\n\n"
           "Extension found at:\n  {}\n\nNo changes were made."
           .format(appdata_ext_root, ext_root),
           exitscript=True
       )

   # 5) Remove ONLY the managed extensions (permanent delete, but only within scope)
   MANAGED_EXTENSIONS = [
       'AE pyTools.extension',
       'WM Tools.extension',
       'H-E-B Tools.extension'
   ]
   
   for ext_name in MANAGED_EXTENSIONS:
       ext_path = appdata_ext_root / ext_name
       if ext_path.exists():
           try:
               shutil.rmtree(str(ext_path))
           except Exception as rm_err:
               forms.alert(
                   "Failed to remove extension:\n{}\n\n{}".format(ext_path, rm_err),
                   exitscript=True
               )

   # 6) Run the updater from %TEMP%
   try:
       Process.Start(tmp_exe)
   except Exception as run_err:
       forms.alert(
           "Failed to start updater:\n{}\n\n{}".format(tmp_exe, run_err),
           exitscript=True
       )


if __name__ == "__main__":
   main()