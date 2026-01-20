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
from System.Diagnostics import Process, ProcessStartInfo


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
   # 1) Kill app.exe and msedgedriver.exe processes if running
   processes_to_kill = ["app", "msedgedriver"]  # Note: without .exe extension
   
   for process_name in processes_to_kill:
       try:
           for proc in Process.GetProcessesByName(process_name):
               try:
                   proc.Kill()
                   proc.WaitForExit(5000)  # Wait up to 5 seconds for it to close
               except:
                   pass  # Process might have already closed
       except Exception as kill_err:
           forms.alert(
               "Warning: Could not terminate {}:\n{}".format(process_name, kill_err),
               title="Process Warning",
               ok=True
           )
   
   # 2) Where this script lives
   script_dir = os.path.dirname(__file__)

   # 3) Locate the updater EXE next to this script
   src_exe = os.path.join(script_dir, "Updater_pyrevit.exe")
   if not os.path.isfile(src_exe):
       forms.alert("Could not find updater EXE:\n{}".format(src_exe), exitscript=True)

   # 4) Copy the entire onedir bundle to a writable temp location
   src_dir = os.path.dirname(src_exe)
   tmp_dir = os.path.join(tempfile.gettempdir(), "Updater_pyrevit")
   tmp_exe = os.path.join(tmp_dir, "Updater_pyrevit.exe")
   try:
       if os.path.isdir(tmp_dir):
           shutil.rmtree(tmp_dir)
       shutil.copytree(src_dir, tmp_dir)
   except Exception as copy_err:
       forms.alert(
           "Failed to copy updater bundle to temp:\n{}\n\n{}".format(tmp_dir, copy_err),
           exitscript=True
       )

   # 5) Resolve extension root and **enforce APPDATA\\pyRevit\\Extensions scope**
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

   # 6) Remove ONLY the managed extensions (permanent delete, but only within scope)
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

   # 7) Run the updater from %TEMP% and capture output
   try:
       psi = ProcessStartInfo()
       psi.FileName = tmp_exe
       psi.UseShellExecute = False
       psi.CreateNoWindow = True
       psi.RedirectStandardOutput = True
       psi.RedirectStandardError = True

       proc = Process.Start(psi)
       stdout = proc.StandardOutput.ReadToEnd()
       stderr = proc.StandardError.ReadToEnd()
       proc.WaitForExit()

       if proc.ExitCode != 0:
           forms.alert(
               "Updater failed (exit {}).\n\nSTDOUT:\n{}\n\nSTDERR:\n{}".format(
                   proc.ExitCode, stdout, stderr
               ),
               exitscript=True
           )
   except Exception as run_err:
       forms.alert(
           "Failed to start updater:\n{}\n\n{}".format(tmp_exe, run_err),
           exitscript=True
       )


if __name__ == "__main__":
   main()
